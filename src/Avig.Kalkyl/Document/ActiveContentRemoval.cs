using System.IO.Compression;
using System.Xml;
using System.Xml.Linq;
using Avig.Kalkyl.Internal;
using Avig.Kalkyl.Reading;
using Avig.OpenXml;

namespace Avig.Kalkyl.Document;

/// <summary>
/// What a save after <see cref="Workbook.RemoveActiveContent"/> leaves out of an opened package.
/// <para>
/// First the relationships to active content are marked: to the VBA project, macro and dialog
/// sheets, ActiveX controls, OLE objects, external links, data connections, query tables and
/// Power Query parts, and every external relationship but hyperlinks. In each worksheet the
/// &lt;oleObjects&gt; and the ActiveX &lt;control&gt; elements go too, with the relationships only
/// they used (the icon pictures) and their shapes in the sheet's drawing and VML drawing.
/// </para>
/// <para>
/// Then the parts that are kept are found by following the relationships that are left from the
/// package root. Whatever can no longer be reached is left out, so a part that only active content
/// pointed to (activeX1.bin, oleObject1.bin, the queryTable part) goes with it, and so does a part
/// nothing pointed to in the first place. The parts that are kept but referred to something that
/// is gone (a drawing, a chart, a table fed by a query) are cleaned when they are copied.
/// </para>
/// </summary>
internal sealed class ActiveContentRemoval
{
    private const string ShapeIdPrefix = "_x0000_s";
    private const string OfficeNamespace = "urn:schemas-microsoft-com:office:office";

    /// <summary>The relationship kinds whose targets are active content.</summary>
    private static readonly HashSet<string> RemovedKinds = new(StringComparer.Ordinal)
    {
        "vbaProject", "xlMacrosheet", "xlIntlMacrosheet", "dialogsheet", "control", "activeXControlBinary",
        "oleObject", "package", "externalLink", "connections", "queryTable",
    };

    /// <summary>Elements that cannot be without the relationship they refer to, and so go with it.</summary>
    private static readonly HashSet<string> DropWhenReferenced = new(StringComparer.Ordinal)
    {
        "control", "oleObject", "externalData", "picture", "oleLink",
    };

    private static readonly HashSet<string> NoIds = new(StringComparer.Ordinal);

    private readonly Dictionary<string, HashSet<string>> removedRelationships;
    private readonly Dictionary<string, HashSet<string>> removedShapes;
    private readonly HashSet<string> keptParts;

    private ActiveContentRemoval(Dictionary<string, HashSet<string>> removedRelationships, Dictionary<string, HashSet<string>> removedShapes, HashSet<string> keptParts, bool isActive)
    {
        this.removedRelationships = removedRelationships;
        this.removedShapes = removedShapes;
        this.keptParts = keptParts;
        IsActive = isActive;
    }

    /// <summary>For a save without <see cref="Workbook.RemoveActiveContent"/>: nothing is left out.</summary>
    public static ActiveContentRemoval None { get; } = new(new(StringComparer.OrdinalIgnoreCase), new(StringComparer.OrdinalIgnoreCase), new(StringComparer.OrdinalIgnoreCase), false);

    /// <summary>True when active content is being removed.</summary>
    public bool IsActive { get; }

    /// <summary>The ids of the relationships of a part that are left out.</summary>
    public IReadOnlySet<string> RemovedIds(string sourcePart)
    {
        if (removedRelationships.TryGetValue(sourcePart, out HashSet<string>? ids) == true)
        {
            return ids;
        }

        return NoIds;
    }

    /// <summary>The shape ids (_x0000_s1025) of the objects and controls left out of a drawing or VML drawing.</summary>
    public IReadOnlySet<string> RemovedShapesOf(string partName)
    {
        if (removedShapes.TryGetValue(partName, out HashSet<string>? shapes) == true)
        {
            return shapes;
        }

        return NoIds;
    }

    /// <summary>True for a part of the original package that is written again.</summary>
    public bool Keeps(string partName)
    {
        return IsActive == false || keptParts.Contains(partName);
    }

    /// <summary>True for a kept part that referred to something left out, and so has to be cleaned when it is copied.</summary>
    public bool Rewrites(string partName)
    {
        return RemovedIds(partName).Count > 0 || RemovedShapesOf(partName).Count > 0;
    }

    /// <summary>
    /// True when an external relationship in the part can be taken out: the part goes anyway, or
    /// it is one the removal cleans (the package root, the workbook, a worksheet, a drawing or a chart).
    /// </summary>
    public static bool CanCleanExternalReferencesIn(string sourcePart, string contentType)
    {
        if (sourcePart.Length == 0 || ActiveContentRules.OfPart(sourcePart, contentType) != ActiveContent.None)
        {
            return true;
        }

        if (contentType.EndsWith("main+xml", StringComparison.OrdinalIgnoreCase) == true || string.Equals(contentType, Ooxml.WorksheetContentType, StringComparison.OrdinalIgnoreCase) == true)
        {
            return true;
        }

        string path = sourcePart.Replace('\\', '/');

        return path.Contains("/worksheets/", StringComparison.OrdinalIgnoreCase) == true
            || path.Contains("/drawings/", StringComparison.OrdinalIgnoreCase) == true
            || path.Contains("/charts/", StringComparison.OrdinalIgnoreCase) == true;
    }

    /// <summary>Works out what to leave out of <paramref name="original"/>, which <paramref name="scan"/> was made from.</summary>
    public static async Task<ActiveContentRemoval> CreateAsync(ZipArchive original, PackageScan scan, string workbookPath, CancellationToken cancellationToken)
    {
        Dictionary<string, HashSet<string>> removed = new(StringComparer.OrdinalIgnoreCase);
        Dictionary<string, HashSet<string>> shapes = new(StringComparer.OrdinalIgnoreCase);

        foreach ((string source, IReadOnlyList<OpcRelationship> relationships) in scan.Relationships)
        {
            foreach (OpcRelationship relationship in relationships)
            {
                if (ShouldRemove(relationship, source, scan) == true)
                {
                    IdsOf(removed, source).Add(relationship.Id);
                }
            }
        }

        foreach ((string source, IReadOnlyList<OpcRelationship> relationships) in scan.Relationships)
        {
            if (IsWorksheet(source, scan) == false || XmlParts.TryFindEntry(original, source, out ZipArchiveEntry? sheetEntry) == false)
            {
                continue;
            }

            HashSet<string> sheetRemoved = IdsOf(removed, source);
            SheetObjects objects = await ReadSheetObjectsAsync(sheetEntry, sheetRemoved, cancellationToken);
            sheetRemoved.UnionWith(objects.ReferencedByRemoved.Except(objects.ReferencedElsewhere));

            if (objects.Shapes.Count == 0)
            {
                continue;
            }

            foreach (OpcRelationship relationship in relationships)
            {
                if (relationship.IsExternal == false && (OpcPath.IsRelationship(relationship.Type, "drawing") == true || OpcPath.IsRelationship(relationship.Type, "vmlDrawing") == true))
                {
                    IdsOf(shapes, OpcPath.ResolveTarget(source, relationship.Target)).UnionWith(objects.Shapes);
                }
            }
        }

        // The pictures the removed shapes showed (VML's o:relid) go with them, unless another shape uses them.
        foreach ((string drawing, HashSet<string> drawingShapes) in shapes)
        {
            if (XmlParts.TryFindEntry(original, drawing, out ZipArchiveEntry? drawingEntry) == true)
            {
                IdsOf(removed, drawing).UnionWith(await ReadShapeReferencesAsync(drawingEntry, drawingShapes, cancellationToken));
            }
        }

        return new ActiveContentRemoval(removed, shapes, Reachable(scan, removed, workbookPath), true);
    }

    /// <summary>
    /// Writes a kept part cleaned of what was left out: elements that need a removed relationship
    /// go, references to removed relationships go, the drawing and VML shapes of removed objects
    /// go, and a table that was filled by a query becomes an ordinary table. A part that is not
    /// XML (VML is not always) is copied as it was.
    /// </summary>
    public async Task WriteCleanedAsync(ZipArchiveEntry entry, Stream target, CancellationToken cancellationToken)
    {
        string name = entry.FullName.Replace('\\', '/');
        XDocument document;

        try
        {
            await using Stream stream = await XmlParts.OpenEntryAsync(entry, cancellationToken);
            using XmlReader reader = XmlParts.CreateReader(stream);
            document = await XDocument.LoadAsync(reader, LoadOptions.None, cancellationToken);
        }
        catch (XmlException)
        {
            await using Stream stream = await XmlParts.OpenEntryAsync(entry, cancellationToken);
            await stream.CopyToAsync(target, cancellationToken);

            return;
        }

        if (document.Root is not null)
        {
            CleanChildren(document.Root, RemovedIds(name), RemovedShapesOf(name));
            MakePlainTable(document.Root);
        }

        await using XmlWriter xml = XmlWriter.Create(target, XlsxWriter.XmlSettings);
        await document.SaveAsync(xml, cancellationToken);
        await xml.FlushAsync();
    }

    /// <summary>Copies the Relationship elements of a .rels part except those whose Id is in <paramref name="removedIds"/>.</summary>
    public static async Task CopyRelationshipsAsync(ZipArchiveEntry entry, XmlWriter xml, IReadOnlySet<string> removedIds, CancellationToken cancellationToken)
    {
        await using Stream stream = await XmlParts.OpenEntryAsync(entry, cancellationToken);
        using XmlReader reader = XmlParts.CreateReader(stream);

        while (reader.EOF == false)
        {
            if (XmlParts.IsElement(reader, "Relationship") == true)
            {
                if (reader.GetAttribute("Id") is string id && removedIds.Contains(id) == true)
                {
                    await reader.SkipAsync();

                    continue;
                }

                await xml.WriteNodeAsync(reader, true);

                continue;
            }

            await reader.ReadAsync();
        }
    }

    /// <summary>
    /// Cleans an element and what is in it; false when the element itself has to go. Used for the
    /// top-level elements of worksheets and drawings and for whole parts.
    /// </summary>
    public static bool Clean(XElement element, IReadOnlySet<string> removedIds, IReadOnlySet<string> removedShapes)
    {
        string name = element.Name.LocalName;

        // Every embedded or linked object in a worksheet goes, and so does the list of them.
        if (name == "oleObjects")
        {
            return false;
        }

        if (removedShapes.Count > 0 && IsShapeOf(element, removedShapes) == true)
        {
            return false;
        }

        List<XAttribute> references = [.. element.Attributes().Where(attribute => IsReference(attribute) == true && removedIds.Contains(attribute.Value) == true)];

        if (references.Count > 0 && DropWhenReferenced.Contains(name) == true)
        {
            return false;
        }

        foreach (XAttribute reference in references)
        {
            reference.Remove();
        }

        CleanChildren(element, removedIds, removedShapes);

        if (name == "controls" && element.HasElements == false)
        {
            return false;
        }

        // An mc:AlternateContent whose choices are all empty now said only what was removed.
        if (name == "AlternateContent" && element.Elements().Any(choice => choice.HasElements == true) == false)
        {
            return false;
        }

        return true;
    }

    /// <summary>True for an attribute that holds a relationship id: r:id, r:embed, r:link … in either namespace, and VML's o:relid.</summary>
    public static bool IsReference(XAttribute attribute)
    {
        string space = attribute.Name.NamespaceName;

        if (space.EndsWith("/relationships", StringComparison.Ordinal) == true && space.Contains("/package/", StringComparison.Ordinal) == false)
        {
            return true;
        }

        return space == OfficeNamespace && attribute.Name.LocalName == "relid";
    }

    private static void CleanChildren(XElement element, IReadOnlySet<string> removedIds, IReadOnlySet<string> removedShapes)
    {
        foreach (XElement child in element.Elements().ToList())
        {
            if (Clean(child, removedIds, removedShapes) == false)
            {
                child.Remove();
            }
        }
    }

    /// <summary>
    /// A drawing anchor that holds the shape Excel drew for an OLE object or control (its
    /// a14:compatExt names the VML shape), or that VML shape itself.
    /// </summary>
    private static bool IsShapeOf(XElement element, IReadOnlySet<string> removedShapes)
    {
        string name = element.Name.LocalName;

        if (name == "shape")
        {
            string? id = (string?)element.Attribute("id");
            string? spid = (string?)element.Attribute(XName.Get("spid", OfficeNamespace));

            return (id is not null && removedShapes.Contains(id) == true) || (spid is not null && removedShapes.Contains(spid) == true);
        }

        if (name is not ("twoCellAnchor" or "oneCellAnchor" or "absoluteAnchor"))
        {
            return false;
        }

        return element.Descendants().Any(descendant => descendant.Name.LocalName == "compatExt" && (string?)descendant.Attribute("spid") is string spid && removedShapes.Contains(spid) == true);
    }

    /// <summary>A table that a query filled (tableType="queryTable") becomes an ordinary table once the query table is gone.</summary>
    private static void MakePlainTable(XElement root)
    {
        if (root.Name.LocalName != "table" || (string?)root.Attribute("tableType") != "queryTable")
        {
            return;
        }

        root.Attribute("tableType")?.Remove();
        root.Attribute("connectionId")?.Remove();

        foreach (XElement column in root.Descendants().Where(element => element.Name.LocalName == "tableColumn"))
        {
            column.Attribute("queryTableFieldId")?.Remove();
            column.Attribute("uniqueName")?.Remove();
        }
    }

    private static bool ShouldRemove(OpcRelationship relationship, string source, PackageScan scan)
    {
        if (relationship.IsExternal == true)
        {
            return OpcPath.IsRelationship(relationship.Type, "hyperlink") == false;
        }

        string kind = relationship.Type[(relationship.Type.LastIndexOf('/') + 1)..];

        if (RemovedKinds.Contains(kind) == true)
        {
            return kind != "package" || ActiveContentRules.IsChart(source) == false;
        }

        string target = OpcPath.ResolveTarget(source, relationship.Target);

        if (scan.DataMashupParts.Contains(target) == true)
        {
            return true;
        }

        return ActiveContentRules.OfPart(target, scan.ContentTypeOf(target)) != ActiveContent.None;
    }

    private static bool IsWorksheet(string part, PackageScan scan)
    {
        string type = scan.ContentTypeOf(part);

        if (type.Length > 0)
        {
            return string.Equals(type, Ooxml.WorksheetContentType, StringComparison.OrdinalIgnoreCase) == true;
        }

        return part.Contains("/worksheets/", StringComparison.OrdinalIgnoreCase) == true;
    }

    /// <summary>
    /// The parts reached through the relationships that are left, with their .rels parts and the
    /// content types. The search starts from the package root and from the workbook part, which
    /// the reader also finds in a package whose root relationships are missing.
    /// </summary>
    private static HashSet<string> Reachable(PackageScan scan, Dictionary<string, HashSet<string>> removed, string workbookPath)
    {
        HashSet<string> kept = new(StringComparer.OrdinalIgnoreCase);
        HashSet<string> visited = new(StringComparer.OrdinalIgnoreCase);
        kept.Add("[Content_Types].xml");
        visited.Add(string.Empty);
        visited.Add(workbookPath);
        Queue<string> pending = new();
        pending.Enqueue(string.Empty);
        pending.Enqueue(workbookPath);

        while (pending.Count > 0)
        {
            string part = pending.Dequeue();

            if (part.Length > 0)
            {
                kept.Add(part);
            }

            if (scan.Relationships.TryGetValue(part, out IReadOnlyList<OpcRelationship>? relationships) == false)
            {
                continue;
            }

            kept.Add(OpcPath.RelationshipsPath(part));
            removed.TryGetValue(part, out HashSet<string>? removedHere);

            foreach (OpcRelationship relationship in relationships)
            {
                if (relationship.IsExternal == true || (removedHere is not null && removedHere.Contains(relationship.Id) == true))
                {
                    continue;
                }

                string target = OpcPath.ResolveTarget(part, relationship.Target);

                if (visited.Add(target) == true)
                {
                    pending.Enqueue(target);
                }
            }
        }

        return kept;
    }

    private static HashSet<string> IdsOf(Dictionary<string, HashSet<string>> map, string key)
    {
        if (map.TryGetValue(key, out HashSet<string>? ids) == false)
        {
            ids = new HashSet<string>(StringComparer.Ordinal);
            map[key] = ids;
        }

        return ids;
    }

    /// <summary>
    /// Reads the elements of a worksheet outside its cells and finds the OLE objects and the
    /// ActiveX controls (a &lt;control&gt; whose relationship is being removed): their shape ids,
    /// and the relationship ids used inside them and elsewhere in the sheet.
    /// </summary>
    private static async Task<SheetObjects> ReadSheetObjectsAsync(ZipArchiveEntry entry, IReadOnlySet<string> removedIds, CancellationToken cancellationToken)
    {
        SheetObjects objects = new();
        await using Stream stream = await XmlParts.OpenEntryAsync(entry, cancellationToken);
        using XmlReader reader = XmlParts.CreateReader(stream);
        await reader.MoveToContentAsync();

        if (reader.IsEmptyElement == true)
        {
            return objects;
        }

        await reader.ReadAsync();

        while (reader.EOF == false && XmlParts.IsEnd(reader, 0) == false)
        {
            if (reader.NodeType != XmlNodeType.Element || reader.Depth != 1)
            {
                await reader.ReadAsync();

                continue;
            }

            if (reader.LocalName == "sheetData")
            {
                await reader.SkipAsync();

                continue;
            }

            XElement element = (XElement)await XNode.ReadFromAsync(reader, cancellationToken);
            Collect(element, false, removedIds, objects);
        }

        return objects;
    }

    /// <summary>
    /// The relationship ids used only inside the removed shapes of a drawing or VML drawing. A
    /// VML drawing that is not well-formed XML (Excel's sometimes is not) gives none, and is copied as it was.
    /// </summary>
    private static async Task<IReadOnlyList<string>> ReadShapeReferencesAsync(ZipArchiveEntry entry, IReadOnlySet<string> removedShapes, CancellationToken cancellationToken)
    {
        XDocument document;

        try
        {
            await using Stream stream = await XmlParts.OpenEntryAsync(entry, cancellationToken);
            using XmlReader reader = XmlParts.CreateReader(stream);
            document = await XDocument.LoadAsync(reader, LoadOptions.None, cancellationToken);
        }
        catch (XmlException)
        {
            return [];
        }

        HashSet<string> inside = new(StringComparer.Ordinal);
        HashSet<string> elsewhere = new(StringComparer.Ordinal);

        foreach (XElement element in document.Descendants())
        {
            bool removedShape = element.AncestorsAndSelf().Any(ancestor => IsShapeOf(ancestor, removedShapes) == true);

            foreach (XAttribute attribute in element.Attributes().Where(attribute => IsReference(attribute) == true))
            {
                if (removedShape == true)
                {
                    inside.Add(attribute.Value);
                }
                else
                {
                    elsewhere.Add(attribute.Value);
                }
            }
        }

        return [.. inside.Except(elsewhere)];
    }

    private static void Collect(XElement element, bool insideRemoved, IReadOnlySet<string> removedIds, SheetObjects objects)
    {
        string name = element.Name.LocalName;
        bool removing = insideRemoved || name == "oleObjects";

        if (name == "control" && element.Attributes().Any(attribute => IsReference(attribute) == true && removedIds.Contains(attribute.Value) == true) == true)
        {
            removing = true;
        }

        if (removing == true && (name is "oleObject" or "control") && (string?)element.Attribute("shapeId") is string shapeId)
        {
            objects.Shapes.Add(ShapeIdPrefix + shapeId);
        }

        foreach (XAttribute attribute in element.Attributes().Where(attribute => IsReference(attribute) == true))
        {
            if (removing == true)
            {
                objects.ReferencedByRemoved.Add(attribute.Value);
            }
            else
            {
                objects.ReferencedElsewhere.Add(attribute.Value);
            }
        }

        foreach (XElement child in element.Elements())
        {
            Collect(child, removing, removedIds, objects);
        }
    }

    /// <summary>What <see cref="ReadSheetObjectsAsync"/> finds in a worksheet.</summary>
    private sealed class SheetObjects
    {
        public HashSet<string> Shapes { get; } = new(StringComparer.Ordinal);

        public HashSet<string> ReferencedByRemoved { get; } = new(StringComparer.Ordinal);

        public HashSet<string> ReferencedElsewhere { get; } = new(StringComparer.Ordinal);
    }
}
