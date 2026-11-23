using System.Diagnostics.CodeAnalysis;
using System.IO.Compression;
using System.Text;
using System.Xml;

namespace Avig.Kalkyl.Reading;

/// <summary>
/// Hjälp för att läsa XML-delarna i paketet. Element jämförs på lokalt namn, inte namnrymd,
/// så att både vanlig och Strict OOXML går att läsa.
/// </summary>
internal static class XmlParts
{
    public static readonly XmlReaderSettings Settings = new()
    {
        Async = true,

        // Ingen DTD och ingen resolver: en xlsx kommer ofta utifrån, och en DTD är vägen in
        // för XXE och "billion laughs".
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

        // Delnamn i OPC är skiftlägesokänsliga, och vissa program skriver \ i stället för /.
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

    /// <summary>r:id-attributet, vars namnrymd skiljer sig mellan vanlig och Strict OOXML.</summary>
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
    /// Läser texten i en &lt;si&gt; eller &lt;is&gt;: antingen en &lt;t&gt; eller flera &lt;r&gt;&lt;t&gt;
    /// (formaterad text). Fonetisk text (&lt;rPh&gt;, för japanska) tas inte med. Läsaren står
    /// efteråt på elementets slut.
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

    /// <summary>Relationerna i en .rels-fil: id till (typ, mål).</summary>
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
