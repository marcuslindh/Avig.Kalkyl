namespace Avig.Kalkyl.Internal;

/// <summary>Namnrymder, relationstyper och innehållstyper i en xlsx-fil.</summary>
internal static class Ooxml
{
    public const string SpreadsheetNamespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    public const string RelationshipsNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    public const string PackageRelationshipsNamespace = "http://schemas.openxmlformats.org/package/2006/relationships";
    public const string ContentTypesNamespace = "http://schemas.openxmlformats.org/package/2006/content-types";

    public const string OfficeDocumentRelationship = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
    public const string WorksheetRelationship = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";
    public const string StylesRelationship = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";

    public const string WorkbookContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";
    public const string WorksheetContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";
    public const string StylesContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml";
    public const string RelationshipsContentType = "application/vnd.openxmlformats-package.relationships+xml";

    /// <summary>
    /// Relationstyper jämförs på sista delen, eftersom Strict OOXML (ISO 29500 Strict) har
    /// andra namnrymder (purl.oclc.org) men samma namn på slutet.
    /// </summary>
    public static bool IsRelationship(string type, string kind)
    {
        return type.EndsWith("/" + kind, StringComparison.Ordinal);
    }

    /// <summary>Sökvägen till .rels-filen för en del, till exempel xl/_rels/workbook.xml.rels.</summary>
    public static string RelationshipsPath(string partPath)
    {
        int slash = partPath.LastIndexOf('/');

        if (slash < 0)
        {
            return "_rels/" + partPath + ".rels";
        }

        return partPath[..(slash + 1)] + "_rels/" + partPath[(slash + 1)..] + ".rels";
    }

    /// <summary>
    /// Löser ett mål i en relation mot delen som pekar på det. Ett mål som börjar med / gäller
    /// från paketets rot; annars relativt delens mapp, med .. tillåtet.
    /// </summary>
    public static string ResolveTarget(string sourcePartPath, string target)
    {
        string cleaned = target.Replace('\\', '/');

        if (cleaned.StartsWith('/') == true)
        {
            return cleaned[1..];
        }

        int slash = sourcePartPath.LastIndexOf('/');
        List<string> segments = [];

        if (slash >= 0)
        {
            segments.AddRange(sourcePartPath[..slash].Split('/', StringSplitOptions.RemoveEmptyEntries));
        }

        foreach (string segment in cleaned.Split('/', StringSplitOptions.RemoveEmptyEntries))
        {
            if (segment == "..")
            {
                if (segments.Count > 0)
                {
                    segments.RemoveAt(segments.Count - 1);
                }

                continue;
            }

            if (segment == ".")
            {
                continue;
            }

            segments.Add(segment);
        }

        return string.Join('/', segments);
    }
}
