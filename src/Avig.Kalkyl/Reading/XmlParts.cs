using System.Diagnostics.CodeAnalysis;
using System.IO.Compression;
using System.Text;
using System.Xml;
using Avig.OpenXml;

namespace Avig.Kalkyl.Reading;

/// <summary>
/// Helpers for reading the XML parts of the package. The package-level work (zip checks, size
/// limits, the depth-limited XML reader, relationships) lives in Avig.OpenXml, which this
/// package shares with Avig.Ord; what is left here is the spreadsheet's own (rich text) and the
/// wording of the limit message. Elements are compared by local name, not namespace, so that
/// both transitional and Strict OOXML can be read.
/// </summary>
internal static class XmlParts
{
    /// <summary>No DTD and no resolver: an xlsx often comes from outside, and a DTD is the way in for XXE and "billion laughs".</summary>
    public static XmlReaderSettings Settings => OpcXml.Settings;

    public static bool TryFindEntry(ZipArchive archive, string path, [NotNullWhen(true)] out ZipArchiveEntry? entry)
    {
        return OpcZip.TryFindEntry(archive, path, out entry);
    }

    /// <summary>
    /// Invalid for a zip directory no spreadsheet program writes: duplicate names, names that
    /// climb out of the package, and entries that cannot be unpacked (see <see cref="OpcZip.CheckEntriesAsync"/>).
    /// </summary>
    public static Task<Result> CheckEntriesAsync(ZipArchive archive, CancellationToken cancellationToken)
    {
        return OpcZip.CheckEntriesAsync(archive, cancellationToken);
    }

    /// <summary>A reader for a part of a file, with <see cref="Settings"/> and the nesting limit of <see cref="OpcXml.DefaultMaxDepth"/>.</summary>
    public static XmlReader CreateReader(Stream stream)
    {
        return OpcXml.CreateReader(stream);
    }

    /// <summary>As <see cref="CreateReader(Stream)"/>, for a fragment such as one cell's XML.</summary>
    public static XmlReader CreateReader(TextReader text, XmlReaderSettings settings, XmlParserContext context)
    {
        return OpcXml.CreateReader(text, settings, context);
    }

    /// <summary>Opens a part of a file being read, held to the size its zip header gives. Every part that is read goes through here.</summary>
    public static Task<Stream> OpenEntryAsync(ZipArchiveEntry entry, CancellationToken cancellationToken)
    {
        return OpcZip.OpenEntryAsync(entry, cancellationToken);
    }

    /// <summary>
    /// Invalid when the parts of the file add up to more than <paramref name="limits"/> allows
    /// unpacked. Checked from the zip directory before anything is read.
    /// </summary>
    public static Result CheckSize(ZipArchive archive, KalkylLimits limits)
    {
        return OpcZip.CheckSize(archive, limits.MaxUncompressedBytes, "KalkylLimits.MaxUncompressedBytes");
    }

    public static bool IsElement(XmlReader reader, string localName)
    {
        return OpcXml.IsElement(reader, localName);
    }

    public static bool IsEnd(XmlReader reader, int depth)
    {
        return OpcXml.IsEnd(reader, depth);
    }

    /// <summary>The r:id attribute, whose namespace differs between transitional and Strict OOXML.</summary>
    public static bool TryGetRelationshipId(XmlReader reader, [NotNullWhen(true)] out string? id)
    {
        return OpcXml.TryGetRelationshipId(reader, out id);
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

    /// <summary>The relationships in a .rels file: id to (type, target). External ones are left out.</summary>
    public static async Task<Dictionary<string, (string Type, string Target)>> ReadRelationshipsAsync(ZipArchive archive, string relsPath, CancellationToken cancellationToken)
    {
        return new Dictionary<string, (string Type, string Target)>(await OpcZip.ReadRelationshipsAsync(archive, relsPath, cancellationToken), StringComparer.Ordinal);
    }
}
