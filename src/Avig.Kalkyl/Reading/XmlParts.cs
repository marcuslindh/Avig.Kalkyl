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

    /// <summary>
    /// Invalid for a zip directory no spreadsheet program writes: two entries with the same name
    /// (case ignored, \ read as /), a name that climbs out of the package (..), starts at a root
    /// or a drive, or holds control characters, and an entry that cannot be unpacked (a
    /// compression method .NET does not have, or encryption). Duplicates let two readers see two
    /// different files; such names would be passed on when the file is saved, to whoever unpacks
    /// it next; and an entry that cannot be opened would otherwise fail only when it is saved.
    /// </summary>
    public static async Task<Result> CheckEntriesAsync(ZipArchive archive, CancellationToken cancellationToken)
    {
        HashSet<string> names = new(StringComparer.OrdinalIgnoreCase);

        foreach (ZipArchiveEntry entry in archive.Entries)
        {
            string name = entry.FullName.Replace('\\', '/');

            if (names.Add(name) == false)
            {
                return Result.Invalid($"The file is damaged: the part '{entry.FullName}' is in it twice.");
            }

            // An empty name opened fine and then failed the save, where the entry is written again.
            if (name.Length == 0)
            {
                return Result.Invalid("The file is damaged: it has a part without a name.");
            }

            bool climbs = name.Split('/').Contains("..");
            bool rooted = name.StartsWith('/') == true || (name.Length >= 2 && name[1] == ':');

            if (climbs == true || rooted == true || name.Any(char.IsControl) == true)
            {
                return Result.Invalid($"The file is damaged: the part name '{entry.FullName}' points outside the package.");
            }

            try
            {
                await using Stream stream = await entry.OpenAsync(cancellationToken);
            }
            catch (Exception exception) when (exception is NotSupportedException or InvalidDataException)
            {
                return Result.Invalid($"The file is damaged: the part '{entry.FullName}' cannot be unpacked ({exception.Message}).");
            }
        }

        return Result.Success();
    }

    /// <summary>A reader for a part of a file, with <see cref="Settings"/> and the nesting limit of <see cref="Internal.DepthLimitedXmlReader"/>.</summary>
    public static XmlReader CreateReader(Stream stream)
    {
        return new Internal.DepthLimitedXmlReader(XmlReader.Create(stream, Settings));
    }

    /// <summary>As <see cref="CreateReader(Stream)"/>, for a fragment such as one cell's XML.</summary>
    public static XmlReader CreateReader(TextReader text, XmlReaderSettings settings, XmlParserContext context)
    {
        return new Internal.DepthLimitedXmlReader(XmlReader.Create(text, settings, context));
    }

    /// <summary>
    /// Opens a part of a file being read, held to the size its zip header gives. Every part
    /// that is read goes through here (see <see cref="Internal.LimitedStream"/>).
    /// </summary>
    public static async Task<Stream> OpenEntryAsync(ZipArchiveEntry entry, CancellationToken cancellationToken)
    {
        Stream stream = await entry.OpenAsync(cancellationToken);

        return new Internal.LimitedStream(stream, entry.Length, entry.FullName);
    }

    /// <summary>
    /// Invalid when the parts of the file add up to more than <paramref name="limits"/> allows
    /// unpacked. Checked from the zip directory before anything is read.
    /// </summary>
    public static Result CheckSize(ZipArchive archive, KalkylLimits limits)
    {
        long total = 0;

        foreach (ZipArchiveEntry entry in archive.Entries)
        {
            total += entry.Length;

            if (total > limits.MaxUncompressedBytes || total < 0)
            {
                return Result.Invalid($"The file unpacks to more than {limits.MaxUncompressedBytes:N0} bytes, the limit set in KalkylLimits.MaxUncompressedBytes. It may be a zip bomb; if it is a file you trust, open it with larger limits.");
            }
        }

        return Result.Success();
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

        await using Stream stream = await OpenEntryAsync(entry, cancellationToken);
        using XmlReader reader = CreateReader(stream);

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
