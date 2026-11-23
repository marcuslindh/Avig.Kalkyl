namespace Avig.Kalkyl.Internal;

/// <summary>Namespaces, relationship types and content types in an xlsx file.</summary>
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
    /// Relationship types are compared on their last segment, because Strict OOXML (ISO 29500 Strict)
    /// has other namespaces (purl.oclc.org) but the same name at the end.
    /// </summary>
    public static bool IsRelationship(string type, string kind)
    {
        return type.EndsWith("/" + kind, StringComparison.Ordinal);
    }

    /// <summary>The path of a part's .rels file, for example xl/_rels/workbook.xml.rels.</summary>
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
    /// Resolves a relationship target against the part that points to it. A target starting with /
    /// is from the package root; otherwise relative to the part's folder, with .. allowed.
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
