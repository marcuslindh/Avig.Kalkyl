using System.IO.Compression;
using System.Xml;
using System.Xml.Linq;
using Avig.Kalkyl.Document;
using Avig.Kalkyl.Reading;

namespace Avig.Kalkyl.Internal;

/// <summary>
/// Writes the parts behind charts and pictures into a package, picking names no other part has,
/// and remembers the content types they need. Targets are written from the package root
/// (<c>/xl/charts/chart1.xml</c>), so they are right whichever folder the drawing is in.
/// </summary>
internal sealed class DrawingPackage
{
    private readonly ZipArchive archive;
    private readonly HashSet<string> paths;

    public DrawingPackage(ZipArchive archive, IEnumerable<string> existingPaths)
    {
        this.archive = archive;
        paths = new HashSet<string>(existingPaths, StringComparer.OrdinalIgnoreCase);
    }

    /// <summary>What a save after <see cref="Workbook.RemoveActiveContent"/> leaves out, which also applies to the drawings added to.</summary>
    public ActiveContentRemoval Removal { get; init; } = ActiveContentRemoval.None;

    /// <summary>Overrides for [Content_Types].xml: part name (with a leading /) and content type.</summary>
    public List<(string PartName, string ContentType)> Overrides { get; } = [];

    /// <summary>Defaults for [Content_Types].xml: image extension and content type.</summary>
    public SortedDictionary<string, string> ImageTypes { get; } = new(StringComparer.OrdinalIgnoreCase);

    /// <summary>Writes a new drawing part with its charts and pictures, and returns the drawing's path.</summary>
    public async Task<string> WriteDrawingAsync(IReadOnlyList<DrawingItem> items, CancellationToken cancellationToken)
    {
        string drawingPath = Unused("xl/drawings/drawing", ".xml");
        (List<string> ids, List<(string Id, string Type, string Target)> relationships) = await WriteTargetsAsync(items, [], cancellationToken);
        await WritePartAsync(drawingPath, xml => DrawingXml.WriteDrawingAsync(xml, items, ids, 2), cancellationToken);
        await WritePartAsync(Ooxml.RelationshipsPath(drawingPath), xml => DrawingXml.WriteRelationshipsAsync(xml, relationships), cancellationToken);
        Overrides.Add(("/" + drawingPath, DrawingXml.DrawingContentType));

        return drawingPath;
    }

    /// <summary>
    /// Writes an opened file's drawing part with new anchors after its own, and its .rels with
    /// the new relationships after its own. Returns the two paths, which the caller must not copy.
    /// </summary>
    public async Task<string[]> AppendToDrawingAsync(ZipArchive original, ZipArchiveEntry drawing, IReadOnlyList<DrawingItem> items, CancellationToken cancellationToken)
    {
        string drawingPath = drawing.FullName.Replace('\\', '/');
        string relsPath = Ooxml.RelationshipsPath(drawingPath);
        HashSet<string> usedIds = await ReadRelationshipIdsAsync(original, relsPath, cancellationToken);
        int firstShapeId;

        // The drawing is read here for the first time: the package copies drawings it does not
        // add to without looking at them. Reading it whole before writing anything means the
        // copy below cannot fail half-way.
        try
        {
            firstShapeId = await MaxShapeIdAsync(drawing, cancellationToken) + 1;
        }
        catch (XmlException exception)
        {
            throw new InvalidDataException($"{KalkylError.DamagedPrefix}: the drawing {drawingPath} is not valid XML ({exception.Message}), so nothing can be added to it.", exception);
        }

        (List<string> ids, List<(string Id, string Type, string Target)> relationships) = await WriteTargetsAsync(items, usedIds, cancellationToken);

        await WritePartAsync(drawingPath, async xml =>
        {
            await using Stream stream = await XmlParts.OpenEntryAsync(drawing, cancellationToken);
            using XmlReader reader = XmlParts.CreateReader(stream);
            await reader.MoveToContentAsync();
            await xml.WriteStartDocumentAsync(true);
            await xml.WriteStartElementAsync(reader.Prefix, reader.LocalName, reader.NamespaceURI);
            await xml.WriteAttributesAsync(reader, true);

            if (reader.IsEmptyElement == false)
            {
                await reader.ReadAsync();

                IReadOnlySet<string> removedIds = Removal.RemovedIds(drawingPath);
                IReadOnlySet<string> removedShapes = Removal.RemovedShapesOf(drawingPath);

                while (reader.EOF == false && XmlParts.IsEnd(reader, 0) == false)
                {
                    // When active content is removed, the anchors are cleaned as the copy would have cleaned them.
                    if (Removal.IsActive == true && reader.NodeType == XmlNodeType.Element)
                    {
                        XElement anchor = (XElement)await XNode.ReadFromAsync(reader, cancellationToken);

                        if (ActiveContentRemoval.Clean(anchor, removedIds, removedShapes) == true)
                        {
                            await anchor.WriteToAsync(xml, cancellationToken);
                        }

                        continue;
                    }

                    await xml.WriteNodeAsync(reader, true);
                }
            }

            await DrawingXml.WriteAnchorsAsync(xml, items, ids, firstShapeId);
            await xml.WriteEndElementAsync();
            await xml.WriteEndDocumentAsync();
        }, cancellationToken);

        await WritePartAsync(relsPath, async xml =>
        {
            await xml.WriteStartDocumentAsync(true);
            await xml.WriteStartElementAsync(null, "Relationships", Ooxml.PackageRelationshipsNamespace);
            await CopyRelationshipsAsync(original, relsPath, xml, Removal.RemovedIds(drawingPath), cancellationToken);
            await DrawingXml.WriteRelationshipElementsAsync(xml, relationships);
            await xml.WriteEndElementAsync();
            await xml.WriteEndDocumentAsync();
        }, cancellationToken);

        paths.Add(relsPath);

        return [drawingPath, relsPath];
    }

    /// <summary>The Relationship elements of an opened .rels part, copied as they are, except those whose Id is in <paramref name="removedIds"/>.</summary>
    public static async Task CopyRelationshipsAsync(ZipArchive original, string relsPath, XmlWriter xml, IReadOnlySet<string> removedIds, CancellationToken cancellationToken)
    {
        if (XmlParts.TryFindEntry(original, relsPath, out ZipArchiveEntry? entry) == false)
        {
            return;
        }

        await ActiveContentRemoval.CopyRelationshipsAsync(entry, xml, removedIds, cancellationToken);
    }

    /// <summary>Every Id in a .rels part, external relationships included, so that none is reused.</summary>
    public static async Task<HashSet<string>> ReadRelationshipIdsAsync(ZipArchive original, string relsPath, CancellationToken cancellationToken)
    {
        HashSet<string> ids = new(StringComparer.Ordinal);

        if (XmlParts.TryFindEntry(original, relsPath, out ZipArchiveEntry? entry) == false)
        {
            return ids;
        }

        await using Stream stream = await XmlParts.OpenEntryAsync(entry, cancellationToken);
        using XmlReader reader = XmlParts.CreateReader(stream);

        while (await reader.ReadAsync() == true)
        {
            if (XmlParts.IsElement(reader, "Relationship") == true && reader.GetAttribute("Id") is string id)
            {
                ids.Add(id);
            }
        }

        return ids;
    }

    /// <summary>The first "rIdN" not in the set, which is then added to it.</summary>
    public static string NextId(HashSet<string> used)
    {
        int number = 1;

        while (used.Contains("rId" + SheetXml.Text(number)) == true)
        {
            number++;
        }

        string id = "rId" + SheetXml.Text(number);
        used.Add(id);

        return id;
    }

    /// <summary>Writes the chart parts and images and returns, per item, the id the drawing refers to it by.</summary>
    private async Task<(List<string> Ids, List<(string Id, string Type, string Target)> Relationships)> WriteTargetsAsync(IReadOnlyList<DrawingItem> items, HashSet<string> usedIds, CancellationToken cancellationToken)
    {
        List<string> ids = [];
        List<(string Id, string Type, string Target)> relationships = [];

        foreach (DrawingItem item in items)
        {
            string id = NextId(usedIds);
            ids.Add(id);

            if (item.Chart is not null)
            {
                string chartPath = Unused("xl/charts/chart", ".xml");
                await WritePartAsync(chartPath, xml => DrawingXml.WriteChartAsync(xml, item), cancellationToken);
                relationships.Add((id, DrawingXml.ChartRelationship, "/" + chartPath));
                Overrides.Add(("/" + chartPath, DrawingXml.ChartContentType));

                continue;
            }

            Picture picture = item.Picture!;
            string imagePath = Unused("xl/media/image", "." + picture.Extension);

            // Images are compressed already; deflating them again costs time and saves nothing.
            ZipArchiveEntry entry = archive.CreateEntry(imagePath, CompressionLevel.NoCompression);

            await using (Stream entryStream = await entry.OpenAsync(cancellationToken))
            {
                await entryStream.WriteAsync(picture.Data, cancellationToken);
            }

            relationships.Add((id, DrawingXml.ImageRelationship, "/" + imagePath));
            ImageTypes[picture.Extension] = picture.ContentType;
        }

        return (ids, relationships);
    }

    /// <summary>The highest cNvPr id in a drawing: shape ids must be unique within it.</summary>
    private static async Task<int> MaxShapeIdAsync(ZipArchiveEntry drawing, CancellationToken cancellationToken)
    {
        int max = 1;
        await using Stream stream = await XmlParts.OpenEntryAsync(drawing, cancellationToken);
        using XmlReader reader = XmlParts.CreateReader(stream);

        while (await reader.ReadAsync() == true)
        {
            if (XmlParts.IsElement(reader, "cNvPr") == true
                && int.TryParse(reader.GetAttribute("id"), System.Globalization.NumberStyles.None, System.Globalization.CultureInfo.InvariantCulture, out int id) == true)
            {
                max = Math.Max(max, id);
            }
        }

        return max;
    }

    private string Unused(string prefix, string suffix)
    {
        int number = 1;

        while (paths.Contains(prefix + SheetXml.Text(number) + suffix) == true)
        {
            number++;
        }

        string path = prefix + SheetXml.Text(number) + suffix;
        paths.Add(path);

        return path;
    }

    private async Task WritePartAsync(string path, Func<XmlWriter, Task> write, CancellationToken cancellationToken)
    {
        ZipArchiveEntry entry = archive.CreateEntry(path, CompressionLevel.Optimal);
        await using Stream entryStream = await entry.OpenAsync(cancellationToken);
        await using XmlWriter xml = XmlWriter.Create(entryStream, XlsxWriter.XmlSettings);
        await write(xml);
        await xml.FlushAsync();
    }
}
