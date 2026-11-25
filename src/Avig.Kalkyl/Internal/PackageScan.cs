using System.IO.Compression;
using System.Xml;
using Avig.Kalkyl.Document;
using Avig.Kalkyl.Reading;
using Avig.OpenXml;

namespace Avig.Kalkyl.Internal;

/// <summary>
/// One pass over an opened package that finds its active content: the content types of the
/// parts, every relationship, the DDE links in the external link parts and the Power Query
/// parts. The document model keeps it, to report what it found and to know what
/// <see cref="Workbook.RemoveActiveContent"/> has to take out; the streaming reader makes it
/// when it is asked. Only the .rels parts, [Content_Types].xml and a few small parts are read.
/// </summary>
internal sealed class PackageScan
{
    private const string DataMashupRoot = "DataMashup";

    private readonly string removalBlocker;

    private PackageScan(
        ActiveContent found,
        IReadOnlyList<OpcExternalRelationship> externalReferences,
        Dictionary<string, IReadOnlyList<OpcRelationship>> relationships,
        Dictionary<string, string> contentTypes,
        HashSet<string> dataMashupParts,
        string removalBlocker)
    {
        Found = found;
        ExternalReferences = externalReferences;
        Relationships = relationships;
        ContentTypes = contentTypes;
        DataMashupParts = dataMashupParts;
        this.removalBlocker = removalBlocker;
    }

    /// <summary>A scan of a workbook made in memory, which has nothing in it.</summary>
    public static PackageScan Empty { get; } = new(ActiveContent.None, [], new(StringComparer.OrdinalIgnoreCase), new(StringComparer.OrdinalIgnoreCase), new(StringComparer.OrdinalIgnoreCase), string.Empty);

    /// <summary>The kinds of active content in the package.</summary>
    public ActiveContent Found { get; }

    /// <summary>The external relationships other than hyperlinks, and the DDE links.</summary>
    public IReadOnlyList<OpcExternalRelationship> ExternalReferences { get; }

    /// <summary>Every relationship, by the part it belongs to (the empty name for the package).</summary>
    public IReadOnlyDictionary<string, IReadOnlyList<OpcRelationship>> Relationships { get; }

    /// <summary>The content type of every part, by its name without a leading slash.</summary>
    public IReadOnlyDictionary<string, string> ContentTypes { get; }

    /// <summary>The custom XML parts that hold Power Query queries.</summary>
    public IReadOnlySet<string> DataMashupParts { get; }

    /// <summary>Why the active content cannot be removed without breaking the workbook, if there is a reason.</summary>
    public bool TryGetRemovalBlocker(out string reason)
    {
        reason = removalBlocker;

        return removalBlocker.Length > 0;
    }

    /// <summary>The content type of a part; empty when the package does not give one.</summary>
    public string ContentTypeOf(string partName)
    {
        if (ContentTypes.TryGetValue(partName, out string? type) == true)
        {
            return type;
        }

        return string.Empty;
    }

    /// <summary>
    /// Scans a package. Throws XmlException for a part it reads that is not XML, and
    /// InvalidDataException for one that cannot be unpacked: a file that hides its active
    /// content in a broken part is a damaged file.
    /// </summary>
    public static async Task<PackageScan> ReadAsync(ZipArchive archive, CancellationToken cancellationToken)
    {
        string reading = "[Content_Types].xml";

        try
        {
            return await ScanAsync(archive, part => reading = part, cancellationToken);
        }
        catch (XmlException exception)
        {
            throw new InvalidDataException($"{KalkylError.DamagedPrefix}: {reading} has broken XML ({exception.Message}).", exception);
        }
    }

    private static async Task<PackageScan> ScanAsync(ZipArchive archive, Action<string> reading, CancellationToken cancellationToken)
    {
        (Dictionary<string, string> defaults, Dictionary<string, string> overrides) = await ReadContentTypesAsync(archive, cancellationToken);
        Dictionary<string, string> contentTypes = new(StringComparer.OrdinalIgnoreCase);
        Dictionary<string, IReadOnlyList<OpcRelationship>> relationships = new(StringComparer.OrdinalIgnoreCase);
        HashSet<string> dataMashupParts = new(StringComparer.OrdinalIgnoreCase);
        List<OpcExternalRelationship> external = [];
        List<ZipArchiveEntry> charts = [];
        List<ZipArchiveEntry> pivotCaches = [];
        ActiveContent found = ActiveContent.None;
        bool dataModel = false;

        foreach (ZipArchiveEntry entry in archive.Entries)
        {
            string name = entry.FullName.Replace('\\', '/');

            if (name.EndsWith('/') == true)
            {
                continue;
            }

            string type = TypeOf(name, defaults, overrides);
            contentTypes[name] = type;
            reading(name);
            found |= ActiveContentRules.OfPart(name, type);

            if (OpcPath.TryGetSourcePart(name, out string? source) == true)
            {
                cancellationToken.ThrowIfCancellationRequested();
                IReadOnlyList<OpcRelationship> all = await OpcZip.ReadAllRelationshipsAsync(archive, name, cancellationToken);
                relationships[source] = all;

                foreach (OpcRelationship relationship in all)
                {
                    found |= ActiveContentRules.OfRelationship(relationship, source);

                    if (relationship.IsExternal == true && OpcPath.IsRelationship(relationship.Type, "hyperlink") == false)
                    {
                        external.Add(new OpcExternalRelationship(source, relationship));
                    }
                }

                continue;
            }

            if (string.Equals(type, ActiveContentRules.ExternalLinkContentType, StringComparison.OrdinalIgnoreCase) == true || name.Contains("/externalLinks/", StringComparison.OrdinalIgnoreCase) == true)
            {
                foreach (string ddeTarget in await ReadDdeLinksAsync(entry, cancellationToken))
                {
                    external.Add(new OpcExternalRelationship(name, new OpcRelationship(string.Empty, ActiveContentRules.DdeLinkType, ddeTarget, true)));
                }
            }
            else if (IsCustomXmlItem(name) == true && await ReadRootNameAsync(entry, cancellationToken) == DataMashupRoot)
            {
                dataMashupParts.Add(name);
                found |= ActiveContent.DataConnections;
            }
            else if (string.Equals(type, ActiveContentRules.ChartContentType, StringComparison.OrdinalIgnoreCase) == true)
            {
                charts.Add(entry);
            }
            else if (string.Equals(type, ActiveContentRules.PivotCacheDefinitionContentType, StringComparison.OrdinalIgnoreCase) == true)
            {
                pivotCaches.Add(entry);
            }

            dataModel |= string.Equals(type, ActiveContentRules.DataModelContentType, StringComparison.OrdinalIgnoreCase) == true || name.StartsWith("xl/model/", StringComparison.OrdinalIgnoreCase) == true;
        }

        if (external.Count > 0)
        {
            found |= ActiveContent.ExternalReferences;
        }

        string blocker = await FindRemovalBlockerAsync(found, external, contentTypes, charts, pivotCaches, dataModel, reading, cancellationToken);

        return new PackageScan(found, external, relationships, contentTypes, dataMashupParts, blocker);
    }

    /// <summary>
    /// What <see cref="Workbook.RemoveActiveContent"/> could not take out without breaking the
    /// workbook or changing what it shows: a pivot table or the data model fed by a connection, a
    /// chart drawn from another workbook, or an external relationship in a part it does not clean.
    /// Empty when there is nothing in the way.
    /// </summary>
    private static async Task<string> FindRemovalBlockerAsync(
        ActiveContent found,
        List<OpcExternalRelationship> external,
        Dictionary<string, string> contentTypes,
        List<ZipArchiveEntry> charts,
        List<ZipArchiveEntry> pivotCaches,
        bool dataModel,
        Action<string> reading,
        CancellationToken cancellationToken)
    {
        if (found.HasFlag(ActiveContent.DataConnections) == true)
        {
            if (dataModel == true)
            {
                return "a data model (Power Pivot)";
            }

            foreach (ZipArchiveEntry pivotCache in pivotCaches)
            {
                reading(pivotCache.FullName);

                if (await HasElementWithAttributeAsync(pivotCache, "cacheSource", "connectionId", cancellationToken) == true)
                {
                    return "a pivot table that gets its data from a data connection";
                }
            }
        }

        foreach (OpcExternalRelationship reference in external)
        {
            contentTypes.TryGetValue(reference.SourcePart, out string? sourceType);

            if (ActiveContentRemoval.CanCleanExternalReferencesIn(reference.SourcePart, sourceType ?? string.Empty) == false)
            {
                return $"an external reference in '{reference.SourcePart}' (to '{reference.Target}')";
            }
        }

        if (contentTypes.Any(pair => ActiveContentRules.OfPart(pair.Key, pair.Value) == ActiveContent.ExternalReferences) == true)
        {
            foreach (ZipArchiveEntry chart in charts)
            {
                reading(chart.FullName);

                if (await HasExternalFormulaAsync(chart, cancellationToken) == true)
                {
                    return "a chart whose data comes from another workbook";
                }
            }
        }

        return string.Empty;
    }

    private static string TypeOf(string name, Dictionary<string, string> defaults, Dictionary<string, string> overrides)
    {
        if (overrides.TryGetValue(name, out string? type) == true)
        {
            return type;
        }

        string file = name[(name.LastIndexOf('/') + 1)..];
        int dot = file.LastIndexOf('.');

        if (dot >= 0 && defaults.TryGetValue(file[(dot + 1)..], out string? byExtension) == true)
        {
            return byExtension;
        }

        return string.Empty;
    }

    private static bool IsCustomXmlItem(string name)
    {
        if (name.StartsWith("customXml/", StringComparison.OrdinalIgnoreCase) == false || name.EndsWith(".xml", StringComparison.OrdinalIgnoreCase) == false)
        {
            return false;
        }

        string file = name["customXml/".Length..];

        return file.Contains('/') == false && file.StartsWith("itemProps", StringComparison.OrdinalIgnoreCase) == false;
    }

    private static async Task<(Dictionary<string, string> Defaults, Dictionary<string, string> Overrides)> ReadContentTypesAsync(ZipArchive archive, CancellationToken cancellationToken)
    {
        Dictionary<string, string> defaults = new(StringComparer.OrdinalIgnoreCase);
        Dictionary<string, string> overrides = new(StringComparer.OrdinalIgnoreCase);

        if (XmlParts.TryFindEntry(archive, "[Content_Types].xml", out ZipArchiveEntry? entry) == false)
        {
            return (defaults, overrides);
        }

        await using Stream stream = await XmlParts.OpenEntryAsync(entry, cancellationToken);
        using XmlReader reader = XmlParts.CreateReader(stream);

        while (await reader.ReadAsync() == true)
        {
            if (reader.NodeType != XmlNodeType.Element)
            {
                continue;
            }

            string? contentType = reader.GetAttribute("ContentType");

            if (contentType is null)
            {
                continue;
            }

            if (reader.LocalName == "Default" && reader.GetAttribute("Extension") is string extension)
            {
                defaults[extension] = contentType;
            }
            else if (reader.LocalName == "Override" && reader.GetAttribute("PartName") is string partName)
            {
                overrides[partName.TrimStart('/')] = contentType;
            }
        }

        return (defaults, overrides);
    }

    /// <summary>The DDE links in an external link part, as service|topic.</summary>
    private static async Task<List<string>> ReadDdeLinksAsync(ZipArchiveEntry entry, CancellationToken cancellationToken)
    {
        List<string> links = [];
        await using Stream stream = await XmlParts.OpenEntryAsync(entry, cancellationToken);
        using XmlReader reader = XmlParts.CreateReader(stream);

        while (await reader.ReadAsync() == true)
        {
            if (XmlParts.IsElement(reader, "ddeLink") == true)
            {
                string service = reader.GetAttribute("ddeService") ?? string.Empty;
                string topic = reader.GetAttribute("ddeTopic") ?? string.Empty;
                links.Add(service + "|" + topic);
            }
        }

        return links;
    }

    private static async Task<string> ReadRootNameAsync(ZipArchiveEntry entry, CancellationToken cancellationToken)
    {
        await using Stream stream = await XmlParts.OpenEntryAsync(entry, cancellationToken);
        using XmlReader reader = XmlParts.CreateReader(stream);

        if (await reader.MoveToContentAsync() == XmlNodeType.Element)
        {
            return reader.LocalName;
        }

        return string.Empty;
    }

    private static async Task<bool> HasElementWithAttributeAsync(ZipArchiveEntry entry, string element, string attribute, CancellationToken cancellationToken)
    {
        await using Stream stream = await XmlParts.OpenEntryAsync(entry, cancellationToken);
        using XmlReader reader = XmlParts.CreateReader(stream);

        while (await reader.ReadAsync() == true)
        {
            if (XmlParts.IsElement(reader, element) == true && reader.GetAttribute(attribute) is not null)
            {
                return true;
            }
        }

        return false;
    }

    /// <summary>True when a chart's series or labels come from another workbook: a c:f with [1] in it.</summary>
    private static async Task<bool> HasExternalFormulaAsync(ZipArchiveEntry entry, CancellationToken cancellationToken)
    {
        await using Stream stream = await XmlParts.OpenEntryAsync(entry, cancellationToken);
        using XmlReader reader = XmlParts.CreateReader(stream);

        while (reader.EOF == false)
        {
            if (XmlParts.IsElement(reader, "f") == true)
            {
                string formula = await reader.ReadElementContentAsStringAsync();

                if (ActiveContentRules.HasExternalReference(formula) == true)
                {
                    return true;
                }

                continue;
            }

            await reader.ReadAsync();
        }

        return false;
    }
}
