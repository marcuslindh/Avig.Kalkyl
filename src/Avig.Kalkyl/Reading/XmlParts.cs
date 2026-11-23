using System.Diagnostics.CodeAnalysis;
using System.IO.Compression;
using System.Text;
using System.Xml;

namespace Avig.Kalkyl.Reading;

/// <summary>
/// Helpers for reading the XML parts of the package. Elements are compared by local name, not namespace,
/// so that both transitional and Strict OOXML can be read.
/// </summary>
internal static class XmlParts
{
    public static readonly XmlReaderSettings Settings = new()
    {
        Async = true,

        // No DTD and no resolver: an xlsx often comes from outside, and a DTD is the way in
        // for XXE and "billion laughs".
        DtdProcessing = DtdProcessing.Prohibit,
        XmlResolver = null,
        IgnoreComments = true,
        IgnoreProcessingInstructions = true,
        CloseInput = false,
    };

    public static bool TryFindEntry(ZipArchive archive, string path, [NotNullWhen(true)] out ZipArchiveEntry? entry)
    {
        entry = archive.GetEntry(path);

        if (entry is not null)
        {
            return true;
        }

        // Part names in OPC are case-insensitive, and some programs write \ instead of /.
        foreach (ZipArchiveEntry candidate in archive.Entries)
        {
            if (string.Equals(candidate.FullName.Replace('\\', '/'), path, StringComparison.OrdinalIgnoreCase) == true)
            {
                entry = candidate;

                return true;
            }
        }

        return false;
    }

    public static bool IsElement(XmlReader reader, string localName)
    {
        return reader.NodeType == XmlNodeType.Element && string.Equals(reader.LocalName, localName, StringComparison.Ordinal) == true;
    }

    public static bool IsEnd(XmlReader reader, int depth)
    {
        return reader.NodeType == XmlNodeType.EndElement && reader.Depth == depth;
    }

    /// <summary>The r:id attribute, whose namespace differs between transitional and Strict OOXML.</summary>
    public static bool TryGetRelationshipId(XmlReader reader, [NotNullWhen(true)] out string? id)
    {
        id = null;

        if (reader.MoveToFirstAttribute() == true)
        {
            do
            {
                if (string.Equals(reader.LocalName, "id", StringComparison.Ordinal) == true
                    && reader.NamespaceURI.Contains("relationships", StringComparison.Ordinal) == true)
                {
                    id = reader.Value;
                }
            }
            while (reader.MoveToNextAttribute() == true);

            reader.MoveToElement();
        }

        return id is not null;
    }

    /// <summary>
    /// Reads the text in an &lt;si&gt; or &lt;is&gt;: either one &lt;t&gt; or several &lt;r&gt;&lt;t&gt;
    /// (formatted text). Phonetic text (&lt;rPh&gt;, for Japanese) is left out. Afterwards the reader
    /// is on the element's end.
    /// </summary>
    public static async Task<string> ReadRichTextAsync(XmlReader reader)
    {
        if (reader.IsEmptyElement == true)
        {
            return string.Empty;
        }

        int depth = reader.Depth;
        StringBuilder text = new();
        await reader.ReadAsync();

        while (IsEnd(reader, depth) == false)
        {
            if (IsElement(reader, "t") == true)
            {
                text.Append(await reader.ReadElementContentAsStringAsync());

                continue;
            }

            if (IsElement(reader, "rPh") == true || IsElement(reader, "phoneticPr") == true)
            {
                await reader.SkipAsync();

                continue;
            }

            if (await reader.ReadAsync() == false)
            {
                break;
            }
        }

        return Internal.CellText.Decode(text.ToString());
    }

    /// <summary>The relationships in a .rels file: id to (type, target).</summary>
    public static async Task<Dictionary<string, (string Type, string Target)>> ReadRelationshipsAsync(ZipArchive archive, string relsPath, CancellationToken cancellationToken)
    {
        Dictionary<string, (string Type, string Target)> relationships = new(StringComparer.Ordinal);

        if (TryFindEntry(archive, relsPath, out ZipArchiveEntry? entry) == false)
        {
            return relationships;
        }

        await using Stream stream = await entry.OpenAsync(cancellationToken);
        using XmlReader reader = XmlReader.Create(stream, Settings);

        while (await reader.ReadAsync() == true)
        {
            if (IsElement(reader, "Relationship") == true)
            {
                string? id = reader.GetAttribute("Id");
                string? type = reader.GetAttribute("Type");
                string? target = reader.GetAttribute("Target");
                bool external = string.Equals(reader.GetAttribute("TargetMode"), "External", StringComparison.OrdinalIgnoreCase);

                if (id is not null && type is not null && target is not null && external == false)
                {
                    relationships[id] = (type, target);
                }
            }
        }

        return relationships;
    }
}
