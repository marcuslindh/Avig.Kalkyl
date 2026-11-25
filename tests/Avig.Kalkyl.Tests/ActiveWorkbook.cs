using System.Globalization;
using System.IO.Compression;
using System.Text;

namespace Avig.Kalkyl.Tests;

/// <summary>
/// Builds a workbook with active content the way Excel 365 writes it: a VBA project, Excel 4.0
/// macro sheets, dialog sheets, ActiveX controls, OLE objects, links to other workbooks, DDE
/// links, data connections, query tables and Power Query. The XML follows files Excel saved
/// (trimmed of revision ids). Also compiled into tools/Avig.Kalkyl.ExcelCheck, which writes the
/// cleaned workbook for Excel to open.
/// </summary>
internal sealed class ActiveWorkbook
{
    public const string Main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    public const string Relationships = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    public const string PackageRelationships = "http://schemas.openxmlformats.org/package/2006/relationships";
    public const string Office2006 = "http://schemas.microsoft.com/office/2006/relationships";
    public const string Compatibility = "http://schemas.openxmlformats.org/markup-compatibility/2006";
    public const string SheetDrawing = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
    public const string Drawing = "http://schemas.openxmlformats.org/drawingml/2006/main";
    public const string X14 = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
    public const string A14 = "http://schemas.microsoft.com/office/drawing/2010/main";

    private readonly List<string> overrides = [];
    private readonly Dictionary<string, string> defaults = new(StringComparer.OrdinalIgnoreCase)
    {
        ["rels"] = "application/vnd.openxmlformats-package.relationships+xml",
        ["xml"] = "application/xml",
    };

    private readonly List<string> workbookRelationships = [];
    private readonly List<string> sheetRelationships = [];
    private readonly List<string> drawingRelationships = [];
    private readonly List<string> vmlRelationships = [];
    private readonly List<string> sheets = [];
    private readonly List<string> externalReferences = [];
    private readonly List<string> definedNames = [];
    private readonly List<string> vmlShapes = [];
    private readonly List<string> anchors = [];
    private readonly List<string> controls = [];
    private readonly List<string> oleObjects = [];
    private readonly List<string> tableParts = [];
    private readonly List<(string Path, string Content)> parts = [];
    private readonly List<(string Path, byte[] Content)> binaryParts = [];
    private int nextWorkbookId = 2;
    private int nextSheetId = 1;
    private int nextDrawingId = 1;
    private int nextVmlId = 1;
    private int nextLink = 1;
    private int nextShape = 1025;
    private int nextImage = 1;
    private int nextConnection = 1;

    public ActiveWorkbook()
    {
        sheets.Add("<sheet name=\"Data\" sheetId=\"1\" r:id=\"rId1\"/>");
        workbookRelationships.Add($"<Relationship Id=\"rId1\" Type=\"{Relationships}/worksheet\" Target=\"worksheets/sheet1.xml\"/>");
        overrides.Add("<Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>");
    }

    /// <summary>The content type of the workbook part; macro-enabled once a VBA project is added.</summary>
    public string MainContentType { get; set; } = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";

    /// <summary>The rows of the worksheet Data.</summary>
    public string SheetData { get; set; } = "<row r=\"1\"><c r=\"A1\"><v>5</v></c></row>";

    /// <summary>The workbook's bookViews element.</summary>
    public string BookViews { get; set; } = "<bookViews><workbookView xWindow=\"0\" yWindow=\"0\" windowWidth=\"21600\" windowHeight=\"11295\"/></bookViews>";

    /// <summary>A VBA project with its signature.</summary>
    public ActiveWorkbook AddVbaProject()
    {
        MainContentType = "application/vnd.ms-excel.sheet.macroEnabled.main+xml";
        string id = WorkbookId();
        workbookRelationships.Add($"<Relationship Id=\"{id}\" Type=\"{Office2006}/vbaProject\" Target=\"vbaProject.bin\"/>");
        overrides.Add("<Override PartName=\"/xl/vbaProject.bin\" ContentType=\"application/vnd.ms-office.vbaProject\"/>");
        overrides.Add("<Override PartName=\"/xl/vbaProjectSignature.bin\" ContentType=\"application/vnd.ms-office.vbaProjectSignature\"/>");
        parts.Add(("xl/_rels/vbaProject.bin.rels", Rels($"<Relationship Id=\"rId1\" Type=\"{Office2006}/vbaProjectSignature\" Target=\"vbaProjectSignature.bin\"/>")));
        binaryParts.Add(("xl/vbaProject.bin", [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1, 1, 2, 3]));
        binaryParts.Add(("xl/vbaProjectSignature.bin", [9, 9, 9]));

        return this;
    }

    /// <summary>An Excel 4.0 macro sheet before the worksheet, as Excel inserts it, started by Auto_Open.</summary>
    public ActiveWorkbook AddMacroSheet(string name = "Makro1")
    {
        string id = WorkbookId();
        workbookRelationships.Add($"<Relationship Id=\"{id}\" Type=\"{Office2006}/xlMacrosheet\" Target=\"macrosheets/sheet1.xml\"/>");
        overrides.Add("<Override PartName=\"/xl/macrosheets/sheet1.xml\" ContentType=\"application/vnd.ms-excel.macrosheet+xml\"/>");
        sheets.Insert(0, $"<sheet name=\"{name}\" sheetId=\"{sheets.Count + 1}\" r:id=\"{id}\"/>");
        definedNames.Add($"<definedName name=\"Auto_Open\">{name}!$A$1</definedName>");
        parts.Add(("xl/macrosheets/sheet1.xml", $"""<xm:macrosheet xmlns="{Main}" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main" xmlns:r="{Relationships}"><dimension ref="A1:A2"/><sheetViews><sheetView showFormulas="1" workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15"/><sheetData><row r="1"><c r="A1" t="b"><f ca="1">ALERT("Hello")</f><v>0</v></c></row><row r="2"><c r="A2" t="b"><f ca="1">RETURN()</f><v>0</v></c></row></sheetData><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></xm:macrosheet>"""));

        return this;
    }

    /// <summary>An Excel 5 dialog sheet before the worksheet, with its VML drawing of the dialog.</summary>
    public ActiveWorkbook AddDialogSheet(string name = "Dialog1")
    {
        string id = WorkbookId();
        workbookRelationships.Add($"<Relationship Id=\"{id}\" Type=\"{Relationships}/dialogsheet\" Target=\"dialogsheets/sheet1.xml\"/>");
        overrides.Add("<Override PartName=\"/xl/dialogsheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.dialogsheet+xml\"/>");
        defaults["vml"] = "application/vnd.openxmlformats-officedocument.vmlDrawing";
        sheets.Insert(0, $"<sheet name=\"{name}\" sheetId=\"{sheets.Count + 1}\" r:id=\"{id}\"/>");
        parts.Add(("xl/dialogsheets/sheet1.xml", $"""<dialogsheet xmlns="{Main}" xmlns:r="{Relationships}"><sheetViews><sheetView showRowColHeaders="0" showZeros="0" tabSelected="1" showOutlineSymbols="0" workbookViewId="0"/></sheetViews><sheetFormatPr defaultColWidth="1" defaultRowHeight="5.25" customHeight="1"/><sheetProtection sheet="1"/><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/><legacyDrawing r:id="rId1"/></dialogsheet>"""));
        parts.Add(("xl/dialogsheets/_rels/sheet1.xml.rels", Rels($"<Relationship Id=\"rId1\" Type=\"{Relationships}/vmlDrawing\" Target=\"../drawings/vmlDrawing9.vml\"/>")));
        parts.Add(("xl/drawings/vmlDrawing9.vml", """<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel"><o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="2"/></o:shapelayout><v:shapetype id="_x0000_t201" coordsize="21600,21600" o:spt="201" path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/><v:path shadowok="f" o:extrusionok="f" strokeok="f" fillok="f" o:connecttype="rect"/><o:lock v:ext="edit" shapetype="t"/></v:shapetype><v:shape id="_x0000_s2049" type="#_x0000_t201" style="position:absolute;margin-left:68.25pt;margin-top:21pt;width:262.5pt;height:147pt;z-index:1" fillcolor="window [65]" strokecolor="windowText [64]" o:insetmode="auto"><v:textbox o:singleclick="f"><div style="text-align:left"><font face="Tahoma" size="160" color="auto">Dialog Caption</font></div></v:textbox><x:ClientData ObjectType="Dialog"><x:SizeWithCells/><x:Anchor>13, 0, 4, 0, 63, 0, 32, 0</x:Anchor></x:ClientData></v:shape><v:shape id="_x0000_s2050" type="#_x0000_t201" style="position:absolute;margin-left:273pt;margin-top:42pt;width:52.5pt;height:15.75pt;z-index:2" fillcolor="buttonFace [67]" o:insetmode="auto"><x:ClientData ObjectType="Button"><x:Anchor>52, 0, 8, 0, 62, 0, 11, 0</x:Anchor><x:Default/><x:Dismiss/></x:ClientData></v:shape></xml>"""));

        return this;
    }

    /// <summary>An ActiveX command button: a control part, its binary, an icon picture, a VML shape and a drawing anchor.</summary>
    public ActiveWorkbook AddActiveXControl()
    {
        int shape = nextShape++;
        string controlId = SheetId();
        string imageId = SheetId();
        string image = AddImage();
        sheetRelationships.Add($"<Relationship Id=\"{controlId}\" Type=\"{Relationships}/control\" Target=\"../activeX/activeX1.xml\"/>");
        sheetRelationships.Add($"<Relationship Id=\"{imageId}\" Type=\"{Relationships}/image\" Target=\"../media/{image}\"/>");
        overrides.Add($"<Override PartName=\"/xl/activeX/activeX1.xml\" ContentType=\"application/vnd.ms-office.activeX+xml\"/>");
        overrides.Add($"<Override PartName=\"/xl/activeX/activeX1.bin\" ContentType=\"application/vnd.ms-office.activeX\"/>");
        parts.Add(("xl/activeX/activeX1.xml", $"""<ax:ocx ax:classid="{"{D7053240-CE69-11CD-A777-00DD01143C57}"}" ax:persistence="persistStreamInit" r:id="rId1" xmlns:ax="http://schemas.microsoft.com/office/2006/activeX" xmlns:r="{Relationships}"/>"""));
        parts.Add(("xl/activeX/_rels/activeX1.xml.rels", Rels($"<Relationship Id=\"rId1\" Type=\"{Office2006}/activeXControlBinary\" Target=\"activeX1.bin\"/>")));
        binaryParts.Add(("xl/activeX/activeX1.bin", [0x40, 0x41, 0x42]));
        controls.Add(ControlXml(shape, controlId, "CommandButton1", imageId));
        AddShape(shape, "Pict", image, true);

        return this;
    }

    /// <summary>A form control (a button from the Forms toolbar), which is not ActiveX and stays.</summary>
    public ActiveWorkbook AddFormControl()
    {
        int shape = nextShape++;
        string id = SheetId();
        sheetRelationships.Add($"<Relationship Id=\"{id}\" Type=\"{Relationships}/ctrlProp\" Target=\"../ctrlProps/ctrlProp1.xml\"/>");
        overrides.Add("<Override PartName=\"/xl/ctrlProps/ctrlProp1.xml\" ContentType=\"application/vnd.ms-excel.controlproperties+xml\"/>");
        parts.Add(("xl/ctrlProps/ctrlProp1.xml", $"""<formControlPr xmlns="{X14}" objectType="Button" lockText="1"/>"""));
        controls.Add(ControlXml(shape, id, "Button 2", null));
        AddShape(shape, "Button", null, false);

        return this;
    }

    /// <summary>An embedded OLE package (a file shown as an icon): the object, its icon, a VML shape and a drawing anchor.</summary>
    public ActiveWorkbook AddOleObject()
    {
        int shape = nextShape++;
        string objectId = SheetId();
        string imageId = SheetId();
        string image = AddImage();
        sheetRelationships.Add($"<Relationship Id=\"{objectId}\" Type=\"{Relationships}/oleObject\" Target=\"../embeddings/oleObject1.bin\"/>");
        sheetRelationships.Add($"<Relationship Id=\"{imageId}\" Type=\"{Relationships}/image\" Target=\"../media/{image}\"/>");
        overrides.Add("<Override PartName=\"/xl/embeddings/oleObject1.bin\" ContentType=\"application/vnd.openxmlformats-officedocument.oleObject\"/>");
        binaryParts.Add(("xl/embeddings/oleObject1.bin", [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1, 7, 7]));
        oleObjects.Add($"""<mc:AlternateContent xmlns:mc="{Compatibility}"><mc:Choice Requires="x14"><oleObject progId="Package" shapeId="{shape}" r:id="{objectId}"><objectPr defaultSize="0" r:id="{imageId}"><anchor moveWithCells="1"><from><xdr:col>4</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></from><to><xdr:col>4</xdr:col><xdr:colOff>485775</xdr:colOff><xdr:row>2</xdr:row><xdr:rowOff>133350</xdr:rowOff></to></anchor></objectPr></oleObject></mc:Choice><mc:Fallback><oleObject progId="Package" shapeId="{shape}" r:id="{objectId}"/></mc:Fallback></mc:AlternateContent>""");
        AddShape(shape, "Pict", image, true);

        return this;
    }

    /// <summary>A link to another workbook: the external link part, a formula that uses it with its cached value, and a defined name.</summary>
    public ActiveWorkbook AddExternalWorkbookLink(string formulaCell = "<c r=\"B1\"><f>[1]Blad1!A1*2</f><v>84</v></c>")
    {
        int link = AddExternalLinkPart($"""<externalLink xmlns="{Main}"><externalBook xmlns:r="{Relationships}" r:id="rId1"><sheetNames><sheetName val="Blad1"/></sheetNames><sheetDataSet><sheetData sheetId="0"><row r="1"><cell r="A1"><v>42</v></cell></row></sheetData></sheetDataSet></externalBook></externalLink>""",
            Rels($"<Relationship Id=\"rId1\" Type=\"{Relationships}/externalLinkPath\" Target=\"file:///\\\\fileserver\\share\\other.xlsx\" TargetMode=\"External\"/>"));
        SheetData = SheetData.Replace("</row>", formulaCell.Replace("[1]", "[" + link + "]", StringComparison.Ordinal) + "</row>", StringComparison.Ordinal);
        definedNames.Add($"<definedName name=\"Rate\">[{link}]Blad1!$A$1</definedName>");

        return this;
    }

    /// <summary>A DDE link that starts a program, as =cmd|'/c calc'!A0 leaves it, with the formula that uses it.</summary>
    public ActiveWorkbook AddDdeLink()
    {
        int link = AddExternalLinkPart($"""<externalLink xmlns="{Main}"><ddeLink ddeService="cmd" ddeTopic="/c calc"><ddeItems><ddeItem name="A0" advise="1"><values rows="1" cols="1"><value t="e"><val>#N/A</val></value></values></ddeItem></ddeItems></ddeLink></externalLink>""", null);
        SheetData = SheetData.Replace("</row>", $"<c r=\"C1\" t=\"e\"><f>[{link}]!A0</f><v>#N/A</v></c></row>", StringComparison.Ordinal);

        return this;
    }

    /// <summary>A picture in the sheet's drawing that is fetched from a web server when the workbook opens.</summary>
    public ActiveWorkbook AddLinkedPicture()
    {
        string id = "rId" + nextDrawingId++;
        drawingRelationships.Add($"<Relationship Id=\"{id}\" Type=\"{Relationships}/image\" Target=\"https://tracker.example/pixel.png\" TargetMode=\"External\"/>");
        anchors.Add($"""<xdr:oneCellAnchor><xdr:from><xdr:col>6</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>6</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:ext cx="952500" cy="952500"/><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="{nextShape++}" name="Picture 1"/><xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr><xdr:blipFill><a:blip xmlns:r="{Relationships}" r:link="{id}"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/></xdr:oneCellAnchor>""");

        return this;
    }

    /// <summary>A hyperlink, which only goes somewhere when it is clicked, and so is not active content.</summary>
    public ActiveWorkbook AddHyperlink()
    {
        string id = SheetId();
        sheetRelationships.Add($"<Relationship Id=\"{id}\" Type=\"{Relationships}/hyperlink\" Target=\"https://example.com/\" TargetMode=\"External\"/>");
        Hyperlinks = $"<hyperlinks><hyperlink ref=\"A1\" r:id=\"{id}\"/></hyperlinks>";

        return this;
    }

    /// <summary>A text-file connection, as Data > From Text leaves it.</summary>
    public ActiveWorkbook AddConnection()
    {
        if (nextConnection > 1)
        {
            return this;
        }

        nextConnection++;
        string id = WorkbookId();
        workbookRelationships.Add($"<Relationship Id=\"{id}\" Type=\"{Relationships}/connections\" Target=\"connections.xml\"/>");
        overrides.Add("<Override PartName=\"/xl/connections.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.connections+xml\"/>");
        parts.Add(("xl/connections.xml", $"""<connections xmlns="{Main}"><connection id="1" name="data" type="6" refreshedVersion="8" background="1" saveData="1"><textPr prompt="0" sourceFile="\\fileserver\share\data.csv" comma="1"><textFields><textField/></textFields></textPr></connection></connections>"""));

        return this;
    }

    /// <summary>A table filled by a query table on the connection, as a query loaded to a table leaves it.</summary>
    public ActiveWorkbook AddQueryTable()
    {
        AddConnection();
        SheetData = "<row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t>N</t></is></c></row><row r=\"2\"><c r=\"A2\"><v>1</v></c></row><row r=\"3\"><c r=\"A3\"><v>2</v></c></row>";
        string id = SheetId();
        sheetRelationships.Add($"<Relationship Id=\"{id}\" Type=\"{Relationships}/table\" Target=\"../tables/table1.xml\"/>");
        overrides.Add("<Override PartName=\"/xl/tables/table1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml\"/>");
        overrides.Add("<Override PartName=\"/xl/queryTables/queryTable1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.queryTable+xml\"/>");
        parts.Add(("xl/tables/table1.xml", $"""<table xmlns="{Main}" id="1" name="Table_data" displayName="Table_data" ref="A1:A3" tableType="queryTable" totalsRowShown="0"><autoFilter ref="A1:A3"/><tableColumns count="1"><tableColumn id="1" uniqueName="1" name="N" queryTableFieldId="1"/></tableColumns><tableStyleInfo name="TableStyleMedium7" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/></table>"""));
        parts.Add(("xl/tables/_rels/table1.xml.rels", Rels($"<Relationship Id=\"rId1\" Type=\"{Relationships}/queryTable\" Target=\"../queryTables/queryTable1.xml\"/>")));
        parts.Add(("xl/queryTables/queryTable1.xml", $"""<queryTable xmlns="{Main}" name="ExternalData_1" connectionId="1" autoFormatId="0" applyNumberFormats="0" applyBorderFormats="0" applyFontFormats="1" applyPatternFormats="1" applyAlignmentFormats="0" applyWidthHeightFormats="0"><queryTableRefresh nextId="2"><queryTableFields count="1"><queryTableField id="1" name="N" tableColumnId="1"/></queryTableFields></queryTableRefresh></queryTable>"""));
        tableParts.Add($"<tablePart r:id=\"{id}\"/>");
        definedNames.Add("<definedName name=\"ExternalData_1\" localSheetId=\"" + (sheets.Count - 1) + "\" hidden=\"1\">Data!$A$1:$A$3</definedName>");

        return this;
    }

    /// <summary>A Power Query query, kept as a DataMashup in a custom XML part.</summary>
    public ActiveWorkbook AddDataMashup()
    {
        string id = WorkbookId();
        workbookRelationships.Add($"<Relationship Id=\"{id}\" Type=\"{Relationships}/customXml\" Target=\"../customXml/item1.xml\"/>");
        overrides.Add("<Override PartName=\"/customXml/itemProps1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.customXmlProperties+xml\"/>");
        parts.Add(("customXml/item1.xml", "<?xml version=\"1.0\" encoding=\"utf-8\"?><DataMashup xmlns=\"http://schemas.microsoft.com/DataMashup\">AAAAAEQDAABQSwMEFAACAAgA</DataMashup>"));
        parts.Add(("customXml/itemProps1.xml", """<ds:datastoreItem ds:itemID="{EA45853C-2A2C-4E94-93AF-0C245C315B5A}" xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml"><ds:schemaRefs><ds:schemaRef ds:uri="http://schemas.microsoft.com/DataMashup"/></ds:schemaRefs></ds:datastoreItem>"""));
        parts.Add(("customXml/_rels/item1.xml.rels", Rels($"<Relationship Id=\"rId1\" Type=\"{Relationships}/customXmlProps\" Target=\"itemProps1.xml\"/>")));

        return this;
    }

    /// <summary>A pivot cache on the data connection: a pivot table that would break without it.</summary>
    public ActiveWorkbook AddPivotCacheOnConnection()
    {
        AddConnection();
        string id = WorkbookId();
        workbookRelationships.Add($"<Relationship Id=\"{id}\" Type=\"{Relationships}/pivotCacheDefinition\" Target=\"pivotCache/pivotCacheDefinition1.xml\"/>");
        overrides.Add("<Override PartName=\"/xl/pivotCache/pivotCacheDefinition1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml\"/>");
        parts.Add(("xl/pivotCache/pivotCacheDefinition1.xml", $"""<pivotCacheDefinition xmlns="{Main}" refreshOnLoad="1" recordCount="0"><cacheSource type="external" connectionId="1"/><cacheFields count="0"/></pivotCacheDefinition>"""));

        return this;
    }

    /// <summary>A chart whose series comes from another workbook.</summary>
    public ActiveWorkbook AddChartOnExternalLink()
    {
        AddExternalWorkbookLink();
        overrides.Add("<Override PartName=\"/xl/charts/chart1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.drawingml.chart+xml\"/>");
        parts.Add(("xl/charts/chart1.xml", """<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea><c:barChart><c:ser><c:val><c:numRef><c:f>[1]Blad1!$A$1:$A$3</c:f></c:numRef></c:val></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>"""));

        return this;
    }

    /// <summary>A part that nothing refers to.</summary>
    public ActiveWorkbook AddStrayPart(string path)
    {
        binaryParts.Add((path, Encoding.ASCII.GetBytes("MZ stray payload")));

        return this;
    }

    private string Hyperlinks { get; set; } = string.Empty;

    /// <summary>Writes the package.</summary>
    public MemoryStream Build()
    {
        List<(string Path, string Content)> all = [.. parts];
        StringBuilder sheet = new();
        sheet.Append(CultureInfo.InvariantCulture, $"""<worksheet xmlns="{Main}" xmlns:r="{Relationships}" xmlns:xdr="{SheetDrawing}" xmlns:x14="{X14}" xmlns:mc="{Compatibility}" mc:Ignorable="x14"><sheetViews><sheetView tabSelected="1" workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15"/><sheetData>{SheetData}</sheetData>{Hyperlinks}<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>""");

        if (anchors.Count > 0)
        {
            string id = SheetId();
            sheetRelationships.Add($"<Relationship Id=\"{id}\" Type=\"{Relationships}/drawing\" Target=\"../drawings/drawing1.xml\"/>");
            overrides.Add("<Override PartName=\"/xl/drawings/drawing1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.drawing+xml\"/>");
            all.Add(("xl/drawings/drawing1.xml", $"""<xdr:wsDr xmlns:xdr="{SheetDrawing}" xmlns:a="{Drawing}">{string.Concat(anchors)}</xdr:wsDr>"""));

            if (drawingRelationships.Count > 0)
            {
                all.Add(("xl/drawings/_rels/drawing1.xml.rels", Rels(string.Concat(drawingRelationships))));
            }

            sheet.Append(CultureInfo.InvariantCulture, $"<drawing r:id=\"{id}\"/>");
        }

        if (vmlShapes.Count > 0)
        {
            string id = SheetId();
            defaults["vml"] = "application/vnd.openxmlformats-officedocument.vmlDrawing";
            sheetRelationships.Add($"<Relationship Id=\"{id}\" Type=\"{Relationships}/vmlDrawing\" Target=\"../drawings/vmlDrawing1.vml\"/>");
            all.Add(("xl/drawings/vmlDrawing1.vml", $"""<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel"><o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout><v:shapetype id="_x0000_t75" coordsize="21600,21600" o:spt="75" o:preferrelative="t" path="m@4@5l@4@11@9@11@9@5xe" filled="f" stroked="f"><v:stroke joinstyle="miter"/><v:path o:extrusionok="f" gradientshapeok="t" o:connecttype="rect"/><o:lock v:ext="edit" aspectratio="t"/></v:shapetype>{string.Concat(vmlShapes)}</xml>"""));

            if (vmlRelationships.Count > 0)
            {
                all.Add(("xl/drawings/_rels/vmlDrawing1.vml.rels", Rels(string.Concat(vmlRelationships))));
            }

            sheet.Append(CultureInfo.InvariantCulture, $"<legacyDrawing r:id=\"{id}\"/>");
        }

        if (oleObjects.Count > 0)
        {
            sheet.Append("<oleObjects>").Append(string.Concat(oleObjects)).Append("</oleObjects>");
        }

        if (controls.Count > 0)
        {
            sheet.Append("<controls>").Append(string.Concat(controls)).Append("</controls>");
        }

        if (tableParts.Count > 0)
        {
            sheet.Append(CultureInfo.InvariantCulture, $"<tableParts count=\"{tableParts.Count}\">").Append(string.Concat(tableParts)).Append("</tableParts>");
        }

        sheet.Append("</worksheet>");
        all.Add(("xl/worksheets/sheet1.xml", sheet.ToString()));

        if (sheetRelationships.Count > 0)
        {
            all.Add(("xl/worksheets/_rels/sheet1.xml.rels", Rels(string.Concat(sheetRelationships))));
        }

        string external = string.Empty;

        if (externalReferences.Count > 0)
        {
            external = "<externalReferences>" + string.Concat(externalReferences) + "</externalReferences>";
        }

        string names = string.Empty;

        if (definedNames.Count > 0)
        {
            names = "<definedNames>" + string.Concat(definedNames) + "</definedNames>";
        }

        all.Add(("xl/workbook.xml", $"""<workbook xmlns="{Main}" xmlns:r="{Relationships}"><fileVersion appName="xl" lastEdited="7" lowestEdited="7" rupBuild="30625"/><workbookPr defaultThemeVersion="202300"/>{BookViews}<sheets>{string.Concat(sheets)}</sheets>{external}{names}<calcPr calcId="191029"/></workbook>"""));
        all.Add(("xl/_rels/workbook.xml.rels", Rels(string.Concat(workbookRelationships))));
        all.Add(("_rels/.rels", Rels($"<Relationship Id=\"rId1\" Type=\"{Relationships}/officeDocument\" Target=\"xl/workbook.xml\"/>")));
        string defaultTypes = string.Concat(defaults.Select(pair => $"<Default Extension=\"{pair.Key}\" ContentType=\"{pair.Value}\"/>"));
        all.Add(("[Content_Types].xml", $"""<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">{defaultTypes}<Override PartName="/xl/workbook.xml" ContentType="{MainContentType}"/>{string.Concat(overrides)}</Types>"""));

        MemoryStream stream = new();

        using (ZipArchive archive = new(stream, ZipArchiveMode.Create, true))
        {
            foreach ((string path, string content) in all)
            {
                Write(archive, path, Encoding.UTF8.GetBytes(content));
            }

            foreach ((string path, byte[] content) in binaryParts)
            {
                Write(archive, path, content);
            }
        }

        stream.Position = 0;

        return stream;
    }

    private static void Write(ZipArchive archive, string path, byte[] content)
    {
        ZipArchiveEntry entry = archive.CreateEntry(path);
        using Stream target = entry.Open();
        target.Write(content, 0, content.Length);
    }

    private static string Rels(string relationships)
    {
        return $"<Relationships xmlns=\"{PackageRelationships}\">{relationships}</Relationships>";
    }

    private static string ControlXml(int shape, string id, string name, string? imageId)
    {
        string image = string.Empty;

        if (imageId is not null)
        {
            image = $" r:id=\"{imageId}\"";
        }

        return $"""<mc:AlternateContent xmlns:mc="{Compatibility}"><mc:Choice Requires="x14"><control shapeId="{shape}" r:id="{id}" name="{name}"><controlPr defaultSize="0" autoLine="0"{image}><anchor moveWithCells="1"><from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>{shape - 1024}</xdr:row><xdr:rowOff>0</xdr:rowOff></from><to><xdr:col>2</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>{shape - 1023}</xdr:row><xdr:rowOff>0</xdr:rowOff></to></anchor></controlPr></control></mc:Choice><mc:Fallback><control shapeId="{shape}" r:id="{id}" name="{name}"/></mc:Fallback></mc:AlternateContent>""";
    }

    private string AddImage()
    {
        string name = "image" + nextImage++ + ".emf";
        defaults["emf"] = "image/x-emf";
        binaryParts.Add(("xl/media/" + name, [1, 0, 0, 0, 0x6C, 0, 0, 0]));

        return name;
    }

    /// <summary>The VML shape Excel draws for an object or control, and the drawing anchor that names it (a14:compatExt).</summary>
    private void AddShape(int shape, string objectType, string? image, bool picture)
    {
        string imageData = string.Empty;

        if (image is not null)
        {
            string id = "rId" + nextVmlId++;
            vmlRelationships.Add($"<Relationship Id=\"{id}\" Type=\"{Relationships}/image\" Target=\"../media/{image}\"/>");
            imageData = $"<v:imagedata o:relid=\"{id}\" o:title=\"\"/>";
        }

        string type = "#_x0000_t75";

        if (picture == false)
        {
            type = "#_x0000_t201";
        }

        vmlShapes.Add($"""<v:shape id="_x0000_s{shape}" type="{type}" style="position:absolute;margin-left:0;margin-top:0;width:38.25pt;height:40.5pt;z-index:{shape - 1024}" filled="t" fillcolor="window [65]" stroked="t" strokecolor="windowText [64]" o:insetmode="auto">{imageData}<x:ClientData ObjectType="{objectType}"><x:SizeWithCells/><x:Anchor>0, 0, 0, 0, 0, 51, 2, 14</x:Anchor><x:CF>Pict</x:CF><x:AutoPict/></x:ClientData></v:shape>""");
        anchors.Add($"""<mc:AlternateContent xmlns:mc="{Compatibility}"><mc:Choice xmlns:a14="{A14}" Requires="a14"><xdr:twoCellAnchor editAs="oneCell"><xdr:from><xdr:col>4</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>{shape - 1025}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>5</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>{shape - 1023}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="{shape}" name="Object {shape - 1024}" hidden="1"><a:extLst><a:ext uri="{"{63B3BB69-23CF-44E3-9099-C40C66FF867C}"}"><a14:compatExt spid="_x0000_s{shape}"/></a:ext></a:extLst></xdr:cNvPr><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr bwMode="auto"><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:sp><xdr:clientData/></xdr:twoCellAnchor></mc:Choice><mc:Fallback/></mc:AlternateContent>""");
    }

    private int AddExternalLinkPart(string xml, string? relationships)
    {
        int link = nextLink++;
        string id = WorkbookId();
        workbookRelationships.Add($"<Relationship Id=\"{id}\" Type=\"{Relationships}/externalLink\" Target=\"externalLinks/externalLink{link}.xml\"/>");
        overrides.Add($"<Override PartName=\"/xl/externalLinks/externalLink{link}.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml\"/>");
        externalReferences.Add($"<externalReference r:id=\"{id}\"/>");
        parts.Add(($"xl/externalLinks/externalLink{link}.xml", xml));

        if (relationships is not null)
        {
            parts.Add(($"xl/externalLinks/_rels/externalLink{link}.xml.rels", relationships));
        }

        return link;
    }

    private string WorkbookId()
    {
        return "rId" + nextWorkbookId++;
    }

    private string SheetId()
    {
        return "rId" + nextSheetId++;
    }
}
