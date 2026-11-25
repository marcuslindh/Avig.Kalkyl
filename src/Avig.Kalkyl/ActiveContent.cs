using Avig.OpenXml;

namespace Avig.Kalkyl;

/// <summary>
/// What in an Excel file can act on its own when Excel opens it, or fetch something from
/// elsewhere: the ways malware and phishing get into workbooks. The package never runs, follows
/// or refreshes any of it, but it keeps it when a workbook is saved, so a file passed on is no
/// safer for having been opened here. <see cref="Workbook.RemoveActiveContent"/> takes it out.
/// The values that Avig.Ord also has are the same numbers there.
/// </summary>
[Flags]
public enum ActiveContent
{
    /// <summary>None of it.</summary>
    None = 0,

    /// <summary>
    /// A VBA project (an .xlsm or .xltm), Excel 4.0 (XLM) macro sheets, and dialog sheets, the
    /// Excel 5 forms that XLM macros show. XLM macros run from a defined name such as Auto_Open
    /// and are still a common way into a workbook.
    /// </summary>
    Macros = 1,

    /// <summary>ActiveX controls, which run code when the workbook is opened if Excel lets them.</summary>
    ActiveX = 2,

    /// <summary>Embedded OLE objects and packages: another document, a program, any file shown as an icon.</summary>
    EmbeddedObjects = 4,

    /// <summary>
    /// Links to other workbooks and DDE links (which start a program, as in <c>=cmd|'/c calc'!A0</c>),
    /// and every other relationship to something outside the file that Excel follows on its own,
    /// unlike hyperlinks (see <see cref="Workbook.ExternalReferences"/>).
    /// </summary>
    ExternalReferences = 16,

    /// <summary>
    /// Data connections (connections.xml), the query tables that fill cells from them, and Power
    /// Query queries (the DataMashup). Refreshing them reads files, databases and web pages.
    /// </summary>
    DataConnections = 32,
}

/// <summary>How a part or a relationship counts as active content; one place for the document model, the streaming reader and the removal.</summary>
internal static class ActiveContentRules
{
    public const string MacrosheetContentType = "application/vnd.ms-excel.macrosheet+xml";
    public const string IntlMacrosheetContentType = "application/vnd.ms-excel.intlmacrosheet+xml";
    public const string DialogsheetContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.dialogsheet+xml";
    public const string ExternalLinkContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml";
    public const string ConnectionsContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.connections+xml";
    public const string QueryTableContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.queryTable+xml";
    public const string PivotCacheDefinitionContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml";
    public const string ChartContentType = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";
    public const string DataModelContentType = "application/vnd.openxmlformats-officedocument.model+data";

    /// <summary>
    /// The type given to a DDE link in <see cref="Workbook.ExternalReferences"/>. A DDE link is not
    /// a relationship but an element in an externalLink part, so it gets a made-up type whose
    /// last segment, <see cref="OpcExternalRelationship.Kind"/>, is ddeLink.
    /// </summary>
    public const string DdeLinkType = "http://schemas.openxmlformats.org/spreadsheetml/2006/main/ddeLink";

    /// <summary>The relationship kinds (last segments of the type) whose target is active content, and which kind.</summary>
    private static readonly Dictionary<string, ActiveContent> RelationshipKinds = new(StringComparer.Ordinal)
    {
        ["vbaProject"] = ActiveContent.Macros,
        ["xlMacrosheet"] = ActiveContent.Macros,
        ["xlIntlMacrosheet"] = ActiveContent.Macros,
        ["dialogsheet"] = ActiveContent.Macros,
        ["control"] = ActiveContent.ActiveX,
        ["activeXControlBinary"] = ActiveContent.ActiveX,
        ["oleObject"] = ActiveContent.EmbeddedObjects,
        ["package"] = ActiveContent.EmbeddedObjects,
        ["externalLink"] = ActiveContent.ExternalReferences,
        ["connections"] = ActiveContent.DataConnections,
        ["queryTable"] = ActiveContent.DataConnections,
    };

    /// <summary>What a part is, from its name and content type. Files without content types are recognized from the folders Excel uses.</summary>
    public static ActiveContent OfPart(string name, string? contentType)
    {
        string type = contentType ?? string.Empty;
        string path = name.Replace('\\', '/');

        if (type.StartsWith("application/vnd.ms-office.vba", StringComparison.OrdinalIgnoreCase) == true || Workbook.IsVbaPart(path) == true)
        {
            return ActiveContent.Macros;
        }

        if (IsType(type, MacrosheetContentType, IntlMacrosheetContentType, DialogsheetContentType) == true
            || path.Contains("/macrosheets/", StringComparison.OrdinalIgnoreCase) == true
            || path.Contains("/dialogsheets/", StringComparison.OrdinalIgnoreCase) == true)
        {
            return ActiveContent.Macros;
        }

        if (IsType(type, OpcNamespaces.ActiveXContentType, OpcNamespaces.ActiveXBinaryContentType) == true || path.Contains("/activeX/", StringComparison.OrdinalIgnoreCase) == true)
        {
            return ActiveContent.ActiveX;
        }

        if (IsType(type, OpcNamespaces.OleObjectContentType) == true)
        {
            return ActiveContent.EmbeddedObjects;
        }

        if (IsType(type, ExternalLinkContentType) == true || path.Contains("/externalLinks/", StringComparison.OrdinalIgnoreCase) == true)
        {
            return ActiveContent.ExternalReferences;
        }

        if (IsType(type, ConnectionsContentType, QueryTableContentType) == true || path.Contains("/queryTables/", StringComparison.OrdinalIgnoreCase) == true)
        {
            return ActiveContent.DataConnections;
        }

        return ActiveContent.None;
    }

    /// <summary>
    /// What a relationship is: one to a macro sheet, a control, an embedded object, an external
    /// link or a data connection, or an external target Excel follows. <paramref name="sourcePart"/>
    /// is the part it belongs to: a chart's package is its data, not an embedded object.
    /// </summary>
    public static ActiveContent OfRelationship(OpcRelationship relationship, string sourcePart)
    {
        if (relationship.IsExternal == true)
        {
            if (OpcPath.IsRelationship(relationship.Type, "hyperlink") == true)
            {
                return ActiveContent.None;
            }

            return ActiveContent.ExternalReferences;
        }

        string kind = relationship.Type[(relationship.Type.LastIndexOf('/') + 1)..];

        if (RelationshipKinds.TryGetValue(kind, out ActiveContent found) == false)
        {
            return ActiveContent.None;
        }

        if (kind == "package" && IsChart(sourcePart) == true)
        {
            return ActiveContent.None;
        }

        return found;
    }

    /// <summary>True for a part in a charts folder, whose embedded package holds the chart's data.</summary>
    public static bool IsChart(string partName)
    {
        return partName.Replace('\\', '/').Contains("/charts/", StringComparison.OrdinalIgnoreCase) == true;
    }

    /// <summary>
    /// True when a formula refers to another workbook or a DDE link: a <c>[1]</c> before a sheet
    /// or name, as Excel stores <c>'C:\x\[Book.xlsx]Sheet1'!A1</c>. Brackets inside text, after a
    /// table name (<c>Table1[Column]</c>) or inside other brackets are not.
    /// </summary>
    public static bool HasExternalReference(string formula)
    {
        bool inText = false;
        int depth = 0;

        for (int index = 0; index < formula.Length; index++)
        {
            char current = formula[index];

            if (current == '"')
            {
                inText = inText == false;

                continue;
            }

            if (inText == true)
            {
                continue;
            }

            if (current == ']')
            {
                depth = Math.Max(0, depth - 1);

                continue;
            }

            if (current != '[')
            {
                continue;
            }

            bool startsToken = index == 0 || (char.IsLetterOrDigit(formula[index - 1]) == false && formula[index - 1] is not ('_' or '.' or ']' or '\\'));

            if (depth == 0 && startsToken == true && IsBookIndex(formula, index) == true)
            {
                return true;
            }

            depth++;
        }

        return false;
    }

    private static bool IsBookIndex(string formula, int open)
    {
        int index = open + 1;

        while (index < formula.Length && char.IsAsciiDigit(formula[index]) == true)
        {
            index++;
        }

        return index > open + 1 && index < formula.Length && formula[index] == ']';
    }

    private static bool IsType(string type, params string[] candidates)
    {
        foreach (string candidate in candidates)
        {
            if (string.Equals(type, candidate, StringComparison.OrdinalIgnoreCase) == true)
            {
                return true;
            }
        }

        return false;
    }
}
