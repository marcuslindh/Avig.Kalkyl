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

    /// <summary>Relationship types compared on their last segment, so that Strict OOXML matches (see <see cref="Avig.OpenXml.OpcPath.IsRelationship"/>).</summary>
    public static bool IsRelationship(string type, string kind)
    {
        return Avig.OpenXml.OpcPath.IsRelationship(type, kind);
    }

    /// <summary>The path of a part's .rels file, for example xl/_rels/workbook.xml.rels.</summary>
    public static string RelationshipsPath(string partPath)
    {
        return Avig.OpenXml.OpcPath.RelationshipsPath(partPath);
    }

    /// <summary>Resolves a relationship target against the part that points to it.</summary>
    public static string ResolveTarget(string sourcePartPath, string target)
    {
        return Avig.OpenXml.OpcPath.ResolveTarget(sourcePartPath, target);
    }
}
