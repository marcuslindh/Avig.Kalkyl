using System.Globalization;
using System.Text.RegularExpressions;

namespace Avig.Kalkyl.Tests;

/// <summary>
/// Pivot tables and slicers. The workbook <see cref="ExcelLikeWorkbook"/> is the one Excel 365
/// built through COM for docs/testplan.md (PIV-01): the same data, the same pivot tables at the
/// same places, and two slicers. The expected cells and rowItems/colItems are what Excel saved,
/// with its Swedish captions in English.
/// </summary>
public partial class PivotTests
{
    /// <summary>Year, Month, Region, Product, Sales, Units: 24 rows, as in the COM script.</summary>
    internal static Workbook SalesWorkbook(out Worksheet data, out Worksheet pivots)
    {
        Workbook book = Workbook.Create();
        Assert.True(book.AddSheet("Data").IsSuccess(out Worksheet? dataSheet));
        Assert.True(book.AddSheet("P").IsSuccess(out Worksheet? pivotSheet));
        data = dataSheet;
        pivots = pivotSheet;
        string[] headings = ["Year", "Month", "Region", "Product", "Sales", "Units"];

        for (int column = 0; column < headings.Length; column++)
        {
            data[0, column] = headings[column];
        }

        int row = 1;

        foreach (int year in new[] { 2025, 2026 })
        {
            foreach (string month in new[] { "Jan", "Feb", "Mar" })
            {
                foreach (string region in new[] { "North", "South" })
                {
                    foreach (string product in new[] { "A", "B" })
                    {
                        data[row, 0] = year;
                        data[row, 1] = month;
                        data[row, 2] = region;
                        data[row, 3] = product;
                        data[row, 4] = (row + 1) * 10.5;
                        data[row, 5] = (row + 1) % 7;
                        row++;
                    }
                }
            }
        }

        return book;
    }

    /// <summary>The seven pivot tables of the COM file, and the Year slicer on PA and PB with 2026 selected.</summary>
    internal static Workbook ExcelLikeWorkbook()
    {
        Workbook book = SalesWorkbook(out _, out Worksheet pivots);
        Assert.True(book.AddSheet("Dash").IsSuccess(out Worksheet? dash));
        Add(pivots, new PivotTable("PA", "Data!A1:F25") { Rows = ["Month"], Values = [new PivotValue("Sales")] }, "A1");
        Add(pivots, new PivotTable("PB", "Data!A1:F25") { Rows = ["Region", "Product"], Columns = ["Year"], Values = [new PivotValue("Sales"), new PivotValue("Units", PivotAggregation.Count)] }, "E1");
        Add(pivots, new PivotTable("PC", "Data!A1:F25") { Rows = ["Region", "Product"], Columns = ["Year"], Values = [new PivotValue("Sales")], Layout = PivotLayout.Outline }, "A30");
        Add(pivots, new PivotTable("PD", "Data!A1:F25") { Rows = ["Region", "Product"], Columns = ["Year"], Values = [new PivotValue("Sales")], Layout = PivotLayout.Tabular }, "O30");
        Add(pivots, new PivotTable("PE", "Data!A1:F25")
        {
            Filters = ["Region"],
            Rows = ["Month"],
            Values = [new PivotValue("Sales", PivotAggregation.Average) { NumberFormat = "#,##0.00" }, new PivotValue("Sales", PivotAggregation.Min), new PivotValue("Sales", PivotAggregation.Max)],
            GrandTotalRow = false,
            GrandTotalColumn = false,
        }, "A58");
        Add(pivots, new PivotTable("PF", "Data!A1:F25") { Values = [new PivotValue("Units")] }, "J60");
        Add(pivots, new PivotTable("PG", "Data!A1:F25") { Columns = ["Month"], Values = [new PivotValue("Units")] }, "J70");

        Result year = dash.AddSlicer(new Slicer("Year", ["PA", "PB"]) { Name = "YearSlicer", Columns = 2, Style = "SlicerStyleDark2", Selected = ["2026"] }, "A1:E6");
        Assert.True(year.IsSuccess(out _, out _), year.Error);
        Result month = dash.AddSlicer(new Slicer("Month", ["PA", "PB"]) { Name = "MonthSlicer" }, "F1:H14");
        Assert.True(month.IsSuccess(out _, out _), month.Error);
        dash["H1"] = Cell.FromFormula("GETPIVOTDATA(\"Sum of Sales\",P!$A$1,\"Month\",\"Feb\")");

        return book;
    }

    [Fact]
    public async Task The_pivot_tables_show_the_cells_Excel_shows()
    {
        Workbook book = ExcelLikeWorkbook();
        Worksheet sheet = Sheet(book, "P");

        // PA: Month down the side, filtered to 2026 by the slicer.
        AssertCells(sheet, "A1=Row Labels B1=Sum of Sales A2=Jan B2=651 A3=Feb B3=819 A4=Mar B4=987 A5=Grand Total B5=2457");

        // PB: two row fields (compact), Year across, two values, 2026 only.
        AssertCells(sheet, "F1=Column Labels F2=2026 H2=Total Sum of Sales I2=Total Count of Units E3=Row Labels F3=Sum of Sales G3=Count of Units");
        AssertCells(sheet, "E4=North F4=1165.5 G4=6 H4=1165.5 I4=6 E5=A F5=567 G5=3 E6=B F6=598.5 E7=South F7=1291.5 E8=A F8=630 E9=B F9=661.5 I9=3 E10=Grand Total F10=2457 G10=12 H10=2457 I10=12");
        Assert.True(sheet["E1"].Value.IsEmpty);

        // PC: outline.
        AssertCells(sheet, "A30=Sum of Sales C30=Year A31=Region B31=Product C31=2025 D31=2026 E31=Grand Total A32=North C32=409.5 D32=1165.5 E32=1575 B33=A C33=189 D33=567 E33=756 B34=B A35=South B36=A B37=B E37=945 A38=Grand Total C38=945 D38=2457 E38=3402");

        // PD: tabular, the subtotal under each group.
        AssertCells(sheet, "O30=Sum of Sales Q30=Year O31=Region P31=Product Q31=2025 R31=2026 S31=Grand Total O32=North P32=A Q32=189 R32=567 S32=756 P33=B O34=North Total Q34=409.5 R34=1165.5 S34=1575 O35=South P35=A O37=South Total O38=Grand Total S38=3402");
        Assert.True(sheet["O33"].Value.IsEmpty);

        // PE: a filter field above, three values, no grand totals.
        AssertCells(sheet, "A58=Region B58=(All) A60=Row Labels B60=Average of Sales C60=Min of Sales D60=Max of Sales A61=Jan B61=99.75 C61=21 D61=178.5 A62=Feb B62=141.75 C62=63 D62=220.5 A63=Mar B63=183.75 C63=105 D63=262.5");
        Assert.True(sheet["A64"].Value.IsEmpty);
        Assert.Equal("#,##0.00", sheet["B61"].Style.NumberFormat);

        // PF and PG: no row fields.
        AssertCells(sheet, "J60=Sum of Units J61=72 K70=Column Labels K71=Jan L71=Feb M71=Mar N71=Grand Total J72=Sum of Units K72=20 L72=24 M72=28 N72=72");
        Assert.True(sheet["J70"].Value.IsEmpty);
        Assert.True(sheet["J71"].Value.IsEmpty);

        await book.RecalculateAsync(RecalculationScope.All);
        Assert.Equal((CellValue)819, Sheet(book, "Dash")["H1"].Value);
    }

    [Fact]
    public async Task The_definitions_have_the_rows_columns_and_items_Excel_writes()
    {
        await using MemoryStream file = await SaveAsync(ExcelLikeWorkbook());
        Dictionary<string, string> tables = PivotTableParts(file);

        // The rowItems and colItems Excel 365 saved for the same pivot tables.
        Assert.Contains("<location ref=\"A1:B5\" firstHeaderRow=\"1\" firstDataRow=\"1\" firstDataCol=\"1\" />", tables["PA"], StringComparison.Ordinal);
        Assert.Contains("<pivotField showAll=\"0\"><items count=\"3\"><item h=\"1\" x=\"0\" /><item x=\"1\" /><item t=\"default\" /></items></pivotField><pivotField axis=\"axisRow\" showAll=\"0\"><items count=\"4\"><item x=\"0\" /><item x=\"1\" /><item x=\"2\" /><item t=\"default\" /></items></pivotField>", tables["PA"], StringComparison.Ordinal);
        Assert.Contains("<rowItems count=\"4\"><i><x /></i><i><x v=\"1\" /></i><i><x v=\"2\" /></i><i t=\"grand\"><x /></i></rowItems><colItems count=\"1\"><i /></colItems>", tables["PA"], StringComparison.Ordinal);

        Assert.Contains("<location ref=\"E1:I10\" firstHeaderRow=\"1\" firstDataRow=\"3\" firstDataCol=\"1\" />", tables["PB"], StringComparison.Ordinal);
        Assert.Contains("<rowItems count=\"7\"><i><x /></i><i r=\"1\"><x /></i><i r=\"1\"><x v=\"1\" /></i><i><x v=\"1\" /></i><i r=\"1\"><x /></i><i r=\"1\"><x v=\"1\" /></i><i t=\"grand\"><x /></i></rowItems>", tables["PB"], StringComparison.Ordinal);
        Assert.Contains("<colFields count=\"2\"><field x=\"0\" /><field x=\"-2\" /></colFields><colItems count=\"4\"><i><x v=\"1\" /><x /></i><i r=\"1\" i=\"1\"><x v=\"1\" /></i><i t=\"grand\"><x /></i><i t=\"grand\" i=\"1\"><x /></i></colItems>", tables["PB"], StringComparison.Ordinal);
        Assert.Contains("<dataField name=\"Count of Units\" fld=\"5\" subtotal=\"count\" baseField=\"0\" baseItem=\"0\" />", tables["PB"], StringComparison.Ordinal);

        Assert.Contains("compact=\"0\" outline=\"1\" outlineData=\"1\" compactData=\"0\"", tables["PC"], StringComparison.Ordinal);
        Assert.Contains("<location ref=\"A30:E38\" firstHeaderRow=\"1\" firstDataRow=\"2\" firstDataCol=\"2\" />", tables["PC"], StringComparison.Ordinal);

        Assert.Contains("<location ref=\"O30:S38\" firstHeaderRow=\"1\" firstDataRow=\"2\" firstDataCol=\"2\" />", tables["PD"], StringComparison.Ordinal);
        Assert.Contains("<pivotField axis=\"axisRow\" compact=\"0\" outline=\"0\" showAll=\"0\">", tables["PD"], StringComparison.Ordinal);
        Assert.Contains("<rowItems count=\"7\"><i><x /><x /></i><i r=\"1\"><x v=\"1\" /></i><i t=\"default\"><x /></i><i><x v=\"1\" /><x /></i><i r=\"1\"><x v=\"1\" /></i><i t=\"default\"><x v=\"1\" /></i><i t=\"grand\"><x /></i></rowItems>", tables["PD"], StringComparison.Ordinal);

        Assert.Contains("rowGrandTotals=\"0\" colGrandTotals=\"0\"", tables["PE"], StringComparison.Ordinal);
        Assert.Contains("<location ref=\"A60:D63\" firstHeaderRow=\"0\" firstDataRow=\"1\" firstDataCol=\"1\" rowPageCount=\"1\" colPageCount=\"1\" />", tables["PE"], StringComparison.Ordinal);
        Assert.Contains("<rowItems count=\"3\"><i><x /></i><i><x v=\"1\" /></i><i><x v=\"2\" /></i></rowItems><colFields count=\"1\"><field x=\"-2\" /></colFields><colItems count=\"3\"><i><x /></i><i i=\"1\"><x v=\"1\" /></i><i i=\"2\"><x v=\"2\" /></i></colItems><pageFields count=\"1\"><pageField fld=\"2\" hier=\"-1\" /></pageFields>", tables["PE"], StringComparison.Ordinal);
        Assert.Contains("<dataField name=\"Average of Sales\" fld=\"4\" subtotal=\"average\" baseField=\"0\" baseItem=\"0\" numFmtId=\"4\" />", tables["PE"], StringComparison.Ordinal);

        Assert.Contains("<location ref=\"J60:J61\" firstHeaderRow=\"1\" firstDataRow=\"1\" firstDataCol=\"0\" />", tables["PF"], StringComparison.Ordinal);
        Assert.Contains("<rowItems count=\"1\"><i /></rowItems><colItems count=\"1\"><i /></colItems>", tables["PF"], StringComparison.Ordinal);

        Assert.Contains("<location ref=\"J70:N72\" firstHeaderRow=\"1\" firstDataRow=\"2\" firstDataCol=\"1\" />", tables["PG"], StringComparison.Ordinal);
        Assert.Contains("<rowItems count=\"1\"><i /></rowItems><colFields count=\"1\"><field x=\"1\" /></colFields><colItems count=\"4\"><i><x /></i><i><x v=\"1\" /></i><i><x v=\"2\" /></i><i t=\"grand\"><x /></i></colItems>", tables["PG"], StringComparison.Ordinal);
    }

    /// <summary>Nine more layouts: values without rows, column subtotals with and without values, two filter fields, three outline levels.</summary>
    internal static Workbook LayoutWorkbook()
    {
        Workbook book = SalesWorkbook(out _, out Worksheet sheet);
        PivotTable[] tables =
        [
            new("H", "Data!A1:F25") { Values = [new PivotValue("Sales"), new PivotValue("Units")] },
            new("I", "Data!A1:F25") { Rows = ["Region"], Columns = ["Year"], Values = [new PivotValue("Sales")] },
            new("K", "Data!A1:F25") { Rows = ["Region", "Product"], Values = [new PivotValue("Sales"), new PivotValue("Units", PivotAggregation.Min)], Layout = PivotLayout.Tabular },
            new("L", "Data!A1:F25") { Rows = ["Region"], Columns = ["Year", "Product"], Values = [new PivotValue("Sales")], Layout = PivotLayout.Outline },
            new("M", "Data!A1:F25") { Rows = ["Region"], Columns = ["Year", "Product"], Values = [new PivotValue("Sales"), new PivotValue("Units")] },
            new("N", "Data!A1:F25") { Filters = ["Region", "Product"], Rows = ["Month"], Values = [new PivotValue("Sales")], GrandTotalRow = false },
            new("O", "Data!A1:F25") { Rows = ["Region", "Product", "Month"], Values = [new PivotValue("Units", PivotAggregation.Count)], Layout = PivotLayout.Outline, GrandTotalRow = false },
            new("Q", "Data!A1:F25") { Columns = ["Region", "Product"], Values = [new PivotValue("Units", PivotAggregation.Max)], GrandTotalColumn = false },
            new("R", "Data!A1:F25") { Rows = ["Month", "Region"], Columns = ["Product"], Values = [new PivotValue("Sales", PivotAggregation.Average)], Layout = PivotLayout.Tabular, GrandTotalColumn = false },
        ];

        for (int index = 0; index < tables.Length; index++)
        {
            Add(sheet, tables[index], new CellAddress(index * 40, 0).ToString());
        }

        return book;
    }

    [Fact]
    public async Task More_layouts_have_the_locations_and_items_Excel_saves()
    {
        // What Excel 365 saved after it had opened and refreshed the file LayoutWorkbook writes;
        // the cells it showed were the ones the package wrote.
        Dictionary<string, (string Location, string Items)> excel = new()
        {
        ["H"] = ("<location ref=\"A1:B2\" firstHeaderRow=\"0\" firstDataRow=\"1\" firstDataCol=\"0\" />", "<rowItems count=\"1\"><i /></rowItems><colFields count=\"1\"><field x=\"-2\" /></colFields><colItems count=\"2\"><i><x /></i><i i=\"1\"><x v=\"1\" /></i></colItems>"),
        ["I"] = ("<location ref=\"A41:D45\" firstHeaderRow=\"1\" firstDataRow=\"2\" firstDataCol=\"1\" />", "<rowItems count=\"3\"><i><x /></i><i><x v=\"1\" /></i><i t=\"grand\"><x /></i></rowItems><colFields count=\"1\"><field x=\"0\" /></colFields><colItems count=\"3\"><i><x /></i><i><x v=\"1\" /></i><i t=\"grand\"><x /></i></colItems>"),
        ["K"] = ("<location ref=\"A81:D88\" firstHeaderRow=\"0\" firstDataRow=\"1\" firstDataCol=\"2\" />", "<rowItems count=\"7\"><i><x /><x /></i><i r=\"1\"><x v=\"1\" /></i><i t=\"default\"><x /></i><i><x v=\"1\" /><x /></i><i r=\"1\"><x v=\"1\" /></i><i t=\"default\"><x v=\"1\" /></i><i t=\"grand\"><x /></i></rowItems><colFields count=\"1\"><field x=\"-2\" /></colFields><colItems count=\"2\"><i><x /></i><i i=\"1\"><x v=\"1\" /></i></colItems>"),
        ["L"] = ("<location ref=\"A121:H126\" firstHeaderRow=\"1\" firstDataRow=\"3\" firstDataCol=\"1\" />", "<rowItems count=\"3\"><i><x /></i><i><x v=\"1\" /></i><i t=\"grand\"><x /></i></rowItems><colFields count=\"2\"><field x=\"0\" /><field x=\"3\" /></colFields><colItems count=\"7\"><i><x /><x /></i><i r=\"1\"><x v=\"1\" /></i><i t=\"default\"><x /></i><i><x v=\"1\" /><x /></i><i r=\"1\"><x v=\"1\" /></i><i t=\"default\"><x v=\"1\" /></i><i t=\"grand\"><x /></i></colItems>"),
        ["M"] = ("<location ref=\"A161:O167\" firstHeaderRow=\"1\" firstDataRow=\"4\" firstDataCol=\"1\" />", "<rowItems count=\"3\"><i><x /></i><i><x v=\"1\" /></i><i t=\"grand\"><x /></i></rowItems><colFields count=\"3\"><field x=\"0\" /><field x=\"3\" /><field x=\"-2\" /></colFields><colItems count=\"14\"><i><x /><x /><x /></i><i r=\"2\" i=\"1\"><x v=\"1\" /></i><i r=\"1\"><x v=\"1\" /><x /></i><i r=\"2\" i=\"1\"><x v=\"1\" /></i><i t=\"default\"><x /></i><i t=\"default\" i=\"1\"><x /></i><i><x v=\"1\" /><x /><x /></i><i r=\"2\" i=\"1\"><x v=\"1\" /></i><i r=\"1\"><x v=\"1\" /><x /></i><i r=\"2\" i=\"1\"><x v=\"1\" /></i><i t=\"default\"><x v=\"1\" /></i><i t=\"default\" i=\"1\"><x v=\"1\" /></i><i t=\"grand\"><x /></i><i t=\"grand\" i=\"1\"><x /></i></colItems>"),
        ["N"] = ("<location ref=\"A204:B207\" firstHeaderRow=\"1\" firstDataRow=\"1\" firstDataCol=\"1\" rowPageCount=\"2\" colPageCount=\"1\" />", "<rowItems count=\"3\"><i><x /></i><i><x v=\"1\" /></i><i><x v=\"2\" /></i></rowItems><colItems count=\"1\"><i /></colItems>"),
        ["O"] = ("<location ref=\"A241:D259\" firstHeaderRow=\"1\" firstDataRow=\"1\" firstDataCol=\"3\" />", "<rowItems count=\"18\"><i><x /></i><i r=\"1\"><x /></i><i r=\"2\"><x /></i><i r=\"2\"><x v=\"1\" /></i><i r=\"2\"><x v=\"2\" /></i><i r=\"1\"><x v=\"1\" /></i><i r=\"2\"><x /></i><i r=\"2\"><x v=\"1\" /></i><i r=\"2\"><x v=\"2\" /></i><i><x v=\"1\" /></i><i r=\"1\"><x /></i><i r=\"2\"><x /></i><i r=\"2\"><x v=\"1\" /></i><i r=\"2\"><x v=\"2\" /></i><i r=\"1\"><x v=\"1\" /></i><i r=\"2\"><x /></i><i r=\"2\"><x v=\"1\" /></i><i r=\"2\"><x v=\"2\" /></i></rowItems><colItems count=\"1\"><i /></colItems>"),
        ["Q"] = ("<location ref=\"A281:G284\" firstHeaderRow=\"1\" firstDataRow=\"3\" firstDataCol=\"1\" />", "<rowItems count=\"1\"><i /></rowItems><colFields count=\"2\"><field x=\"2\" /><field x=\"3\" /></colFields><colItems count=\"6\"><i><x /><x /></i><i r=\"1\"><x v=\"1\" /></i><i t=\"default\"><x /></i><i><x v=\"1\" /><x /></i><i r=\"1\"><x v=\"1\" /></i><i t=\"default\"><x v=\"1\" /></i></colItems>"),
        ["R"] = ("<location ref=\"A321:D332\" firstHeaderRow=\"1\" firstDataRow=\"2\" firstDataCol=\"2\" />", "<rowItems count=\"10\"><i><x /><x /></i><i r=\"1\"><x v=\"1\" /></i><i t=\"default\"><x /></i><i><x v=\"1\" /><x /></i><i r=\"1\"><x v=\"1\" /></i><i t=\"default\"><x v=\"1\" /></i><i><x v=\"2\" /><x /></i><i r=\"1\"><x v=\"1\" /></i><i t=\"default\"><x v=\"2\" /></i><i t=\"grand\"><x /></i></rowItems><colFields count=\"1\"><field x=\"3\" /></colFields><colItems count=\"2\"><i><x /></i><i><x v=\"1\" /></i></colItems>"),
        };

        Workbook book = LayoutWorkbook();
        AssertCells(Sheet(book, "P"), "A1=Sum of Sales B1=Sum of Units A2=3402 B2=72");
        await using MemoryStream file = await SaveAsync(book);
        Dictionary<string, string> tables = PivotTableParts(file);

        foreach ((string name, (string location, string items)) in excel)
        {
            Assert.Contains(location, tables[name], StringComparison.Ordinal);
            Assert.Contains(items, tables[name], StringComparison.Ordinal);
        }

        Assert.Empty(TestFiles.Validate(file));
    }

    [Fact]
    public async Task The_cache_holds_the_source_as_Excel_stores_it()
    {
        await using MemoryStream file = await SaveAsync(ExcelLikeWorkbook());
        string definition = TestFiles.ReadPart(file, "xl/pivotCache/pivotCacheDefinition1.xml");
        string records = TestFiles.ReadPart(file, "xl/pivotCache/pivotCacheRecords1.xml");

        Assert.Contains("refreshOnLoad=\"1\"", definition, StringComparison.Ordinal);
        Assert.Contains("recordCount=\"24\"", definition, StringComparison.Ordinal);
        Assert.Contains("<cacheSource type=\"worksheet\"><worksheetSource ref=\"A1:F25\" sheet=\"Data\" /></cacheSource>", definition, StringComparison.Ordinal);

        // The cache fields of the COM file, item for item.
        Assert.Contains("<cacheField name=\"Year\" numFmtId=\"0\"><sharedItems containsSemiMixedTypes=\"0\" containsString=\"0\" containsNumber=\"1\" containsInteger=\"1\" minValue=\"2025\" maxValue=\"2026\" count=\"2\"><n v=\"2025\" /><n v=\"2026\" /></sharedItems></cacheField>", definition, StringComparison.Ordinal);
        Assert.Contains("<cacheField name=\"Month\" numFmtId=\"0\"><sharedItems count=\"3\"><s v=\"Jan\" /><s v=\"Feb\" /><s v=\"Mar\" /></sharedItems></cacheField>", definition, StringComparison.Ordinal);
        Assert.Contains("<cacheField name=\"Sales\" numFmtId=\"0\"><sharedItems containsSemiMixedTypes=\"0\" containsString=\"0\" containsNumber=\"1\" minValue=\"21\" maxValue=\"262.5\" /></cacheField>", definition, StringComparison.Ordinal);
        Assert.Contains("<cacheField name=\"Units\" numFmtId=\"0\"><sharedItems containsSemiMixedTypes=\"0\" containsString=\"0\" containsNumber=\"1\" containsInteger=\"1\" minValue=\"0\" maxValue=\"6\" /></cacheField>", definition, StringComparison.Ordinal);
        Assert.Contains("<x14:pivotCacheDefinition pivotCacheId=", definition, StringComparison.Ordinal);

        // The first records of the COM file.
        Assert.StartsWith("<?xml version=\"1.0\" encoding=\"utf-8\" standalone=\"yes\"?>", records, StringComparison.Ordinal);
        Assert.Contains("count=\"24\"><r><x v=\"0\" /><x v=\"0\" /><x v=\"0\" /><x v=\"0\" /><n v=\"21\" /><n v=\"2\" /></r><r><x v=\"0\" /><x v=\"0\" /><x v=\"0\" /><x v=\"1\" /><n v=\"31.5\" /><n v=\"3\" /></r>", records, StringComparison.Ordinal);
        Assert.Equal(24, Regex.Count(records, "<r>"));
    }

    [Fact]
    public async Task The_cache_records_equal_the_source_cells()
    {
        Workbook book = SalesWorkbook(out Worksheet data, out Worksheet pivots);
        Add(pivots, new PivotTable("Sales", "Data!A1:F25") { Rows = ["Region"], Values = [new PivotValue("Sales")] }, "A1");
        await using MemoryStream file = await SaveAsync(book);
        string definition = TestFiles.ReadPart(file, "xl/pivotCache/pivotCacheDefinition1.xml");
        string records = TestFiles.ReadPart(file, "xl/pivotCache/pivotCacheRecords1.xml");
        List<List<string>> shared = [.. Regex.Matches(definition, "<sharedItems[^>]*?(/>|>(.*?)</sharedItems>)").Select(match => Regex.Matches(match.Groups[2].Value, "v=\"([^\"]*)\"").Select(item => item.Groups[1].Value).ToList())];
        List<string> rows = [.. Regex.Matches(records, "<r>(.*?)</r>").Select(match => match.Groups[1].Value)];

        Assert.Equal(24, rows.Count);

        for (int row = 0; row < rows.Count; row++)
        {
            List<Match> values = Regex.Matches(rows[row], "<(x|n|s) v=\"([^\"]*)\" />").ToList();
            Assert.Equal(6, values.Count);

            for (int column = 0; column < 6; column++)
            {
                string text = values[column].Groups[2].Value;

                if (values[column].Groups[1].Value == "x")
                {
                    text = shared[column][int.Parse(text, CultureInfo.InvariantCulture)];
                }

                Assert.Equal(data[row + 1, column].Value.ToString(), text);
            }
        }
    }

    [Fact]
    public async Task Slicers_are_written_as_Excel_writes_them()
    {
        await using MemoryStream file = await SaveAsync(ExcelLikeWorkbook());
        string workbook = TestFiles.ReadPart(file, "xl/workbook.xml");
        string year = TestFiles.ReadPart(file, "xl/slicerCaches/slicerCache1.xml");
        string month = TestFiles.ReadPart(file, "xl/slicerCaches/slicerCache2.xml");
        string slicers = TestFiles.ReadPart(file, "xl/slicers/slicer1.xml");
        string dash = TestFiles.ReadPart(file, "xl/worksheets/sheet3.xml");
        string dashRelationships = TestFiles.ReadPart(file, "xl/worksheets/_rels/sheet3.xml.rels");
        string drawing = TestFiles.ReadPart(file, "xl/drawings/drawing1.xml");
        string types = TestFiles.ReadPart(file, "[Content_Types].xml");

        Assert.Contains("<definedName name=\"Slicer_Year\">#N/A</definedName><definedName name=\"Slicer_Month\">#N/A</definedName>", workbook, StringComparison.Ordinal);
        Assert.Matches("<pivotCaches><pivotCache cacheId=\"0\" r:id=\"rId[0-9]+\" /></pivotCaches><extLst><ext uri=\"\\{BBE1A952-AA13-448e-AADC-164F8A28A991\\}\" xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\"><x14:slicerCaches><x14:slicerCache r:id=\"rId[0-9]+\" /><x14:slicerCache r:id=\"rId[0-9]+\" /></x14:slicerCaches></ext></extLst>", workbook);

        Assert.Contains("name=\"Slicer_Year\" sourceName=\"Year\"><pivotTables><pivotTable tabId=\"2\" name=\"PA\" /><pivotTable tabId=\"2\" name=\"PB\" /></pivotTables><data><tabular pivotCacheId=\"1000000\"><items count=\"2\"><i x=\"0\" /><i x=\"1\" s=\"1\" /></items></tabular></data></slicerCacheDefinition>", year, StringComparison.Ordinal);
        Assert.Contains("<items count=\"3\"><i x=\"0\" s=\"1\" /><i x=\"1\" s=\"1\" /><i x=\"2\" s=\"1\" /></items>", month, StringComparison.Ordinal);
        Assert.Contains("<slicer name=\"YearSlicer\" cache=\"Slicer_Year\" caption=\"Year\" columnCount=\"2\" style=\"SlicerStyleDark2\" rowHeight=\"241300\" /><slicer name=\"MonthSlicer\" cache=\"Slicer_Month\" caption=\"Month\" rowHeight=\"241300\" />", slicers, StringComparison.Ordinal);

        Assert.Matches("<drawing r:id=\"rId1\" /><extLst><ext uri=\"\\{A8765BA9-456A-4dab-B4F3-ACF838C121DE\\}\" xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\"><x14:slicerList><x14:slicer r:id=\"rId2\" /></x14:slicerList></ext></extLst></worksheet>", dash);
        Assert.Contains("Id=\"rId2\" Type=\"http://schemas.microsoft.com/office/2007/relationships/slicer\" Target=\"/xl/slicers/slicer1.xml\"", dashRelationships, StringComparison.Ordinal);
        Assert.Contains("<a:graphicData uri=\"http://schemas.microsoft.com/office/drawing/2010/slicer\"><sle:slicer name=\"YearSlicer\" xmlns:sle=\"http://schemas.microsoft.com/office/drawing/2010/slicer\" /></a:graphicData>", drawing, StringComparison.Ordinal);
        Assert.Contains("<mc:Choice xmlns:a14=\"http://schemas.microsoft.com/office/drawing/2010/main\" Requires=\"a14\">", drawing, StringComparison.Ordinal);
        Assert.Contains("<Override PartName=\"/xl/slicerCaches/slicerCache1.xml\" ContentType=\"application/vnd.ms-excel.slicerCache+xml\" />", types, StringComparison.Ordinal);
        Assert.Contains("<Override PartName=\"/xl/slicers/slicer1.xml\" ContentType=\"application/vnd.ms-excel.slicer+xml\" />", types, StringComparison.Ordinal);
        Assert.Contains("<Override PartName=\"/xl/pivotTables/pivotTable1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml\" />", types, StringComparison.Ordinal);
    }

    [Fact]
    public async Task A_workbook_with_pivot_tables_and_slicers_passes_schema_validation()
    {
        await using MemoryStream file = await SaveAsync(ExcelLikeWorkbook());

        List<string> problems = TestFiles.Validate(file);

        Assert.True(problems.Count == 0, string.Join(Environment.NewLine, problems));
    }

    [Fact]
    public async Task A_saved_workbook_opens_with_its_pivot_tables_and_GETPIVOTDATA_reads_them()
    {
        await using MemoryStream file = await SaveAsync(ExcelLikeWorkbook());
        Result<Workbook> opened = await Workbook.OpenAsync(file);
        Assert.True(opened.IsSuccess(out Workbook? book), opened.Error);

        Assert.Equal(["PG", "PF", "PE", "PD", "PC", "PB", "PA"], book.PivotTables.Select(table => table.Name).Order(StringComparer.Ordinal).Reverse());
        PivotTableInfo pb = book.PivotTables.Single(table => table.Name == "PB");
        Assert.Equal("P", pb.Sheet);
        Assert.Equal("E1:I10", pb.Location.ToString());
        PivotCacheInfo cache = Assert.Single(book.PivotCaches);
        Assert.Equal("'Data'!A1:F25", cache.Source);
        Assert.Equal(["Year", "Month", "Region", "Product", "Sales", "Units"], cache.Fields);
        Assert.Equal(24, cache.RecordCount);
        Assert.True(cache.RefreshOnLoad);

        // GETPIVOTDATA against the pivot tables as the file has them.
        Assert.Equal((CellValue)819, Evaluate(book, "GETPIVOTDATA(\"Sum of Sales\",P!$A$1,\"Month\",\"Feb\")"));
        Assert.Equal((CellValue)2457, Evaluate(book, "GETPIVOTDATA(\"Sales\",P!A3)"));
        Assert.Equal((CellValue)567, Evaluate(book, "GETPIVOTDATA(\"Sum of Sales\",P!E1,\"Region\",\"North\",\"Product\",\"A\",\"Year\",2026)"));
        Assert.Equal((CellValue)6, Evaluate(book, "GETPIVOTDATA(\"Count of Units\",P!E1,\"Region\",\"North\")"));
        Assert.Equal((CellValue)12, Evaluate(book, "GETPIVOTDATA(\"Count of Units\",P!E1)"));
        Assert.Equal((CellValue)409.5, Evaluate(book, "GETPIVOTDATA(\"Sum of Sales\",P!O30,\"Region\",\"North\",\"Year\",\"2025\")"));
        Assert.Equal((CellValue)141.75, Evaluate(book, "GETPIVOTDATA(\"Average of Sales\",P!A60,\"Month\",\"Feb\")"));
        Assert.Equal((CellValue)72, Evaluate(book, "GETPIVOTDATA(\"Sum of Units\",P!J60)"));
        Assert.Equal((CellValue)24, Evaluate(book, "GETPIVOTDATA(\"Sum of Units\",P!J70,\"Month\",\"Feb\")"));

        // What Excel answers with #REF!: no such item, a field off the axes, a total that is not shown, a cell outside every pivot table.
        Assert.Equal(CellValue.Error("#REF!"), Evaluate(book, "GETPIVOTDATA(\"Sum of Sales\",P!$A$1,\"Month\",\"Dec\")"));
        Assert.Equal(CellValue.Error("#REF!"), Evaluate(book, "GETPIVOTDATA(\"Sum of Sales\",P!$A$1,\"Region\",\"North\")"));
        Assert.Equal(CellValue.Error("#REF!"), Evaluate(book, "GETPIVOTDATA(\"Average of Sales\",P!A60)"));
        Assert.Equal(CellValue.Error("#REF!"), Evaluate(book, "GETPIVOTDATA(\"Sum of Sales\",P!Z99)"));
        Assert.Equal(CellValue.Error("#REF!"), Evaluate(book, "GETPIVOTDATA(\"Sum of Sales\",P!E1,\"Product\",\"A\")"));
    }

    [Fact]
    public async Task Opening_and_saving_keeps_every_pivot_part_byte_for_byte()
    {
        await using MemoryStream file = await SaveAsync(ExcelLikeWorkbook());
        Result<Workbook> opened = await Workbook.OpenAsync(file);
        Assert.True(opened.IsSuccess(out Workbook? book), opened.Error);
        await using MemoryStream saved = await SaveAsync(book);

        foreach (string part in new[] { "xl/pivotCache/pivotCacheDefinition1.xml", "xl/pivotCache/pivotCacheRecords1.xml", "xl/pivotTables/pivotTable1.xml", "xl/slicerCaches/slicerCache1.xml", "xl/slicers/slicer1.xml", "xl/drawings/drawing1.xml" })
        {
            Assert.Equal(TestFiles.ReadPart(file, part), TestFiles.ReadPart(saved, part));
        }

        Assert.Equal(TestFiles.ReadPart(file, "xl/worksheets/sheet3.xml").Contains("x14:slicerList", StringComparison.Ordinal), TestFiles.ReadPart(saved, "xl/worksheets/sheet3.xml").Contains("x14:slicerList", StringComparison.Ordinal));
        Assert.Contains("<pivotCaches>", TestFiles.ReadPart(saved, "xl/workbook.xml"), StringComparison.Ordinal);
        Assert.Empty(TestFiles.Validate(saved));
    }

    [Fact]
    public async Task RefreshPivotTablesOnOpen_marks_every_cache_of_an_opened_file()
    {
        // A template from elsewhere: a cache without refreshOnLoad.
        await using MemoryStream template = await SaveAsync(ExcelLikeWorkbook());
        string original = TestFiles.ReadPart(template, "xl/pivotCache/pivotCacheDefinition1.xml");
        await using MemoryStream withoutRefresh = Replace(template, "xl/pivotCache/pivotCacheDefinition1.xml", original.Replace(" refreshOnLoad=\"1\"", string.Empty, StringComparison.Ordinal));

        Result<Workbook> opened = await Workbook.OpenAsync(withoutRefresh);
        Assert.True(opened.IsSuccess(out Workbook? book), opened.Error);
        Assert.False(book.PivotCaches[0].RefreshOnLoad);

        book.RefreshPivotTablesOnOpen = true;
        Assert.True(book.PivotCaches[0].RefreshOnLoad);
        Sheet(book, "Data")[1, 4] = 1000;
        await using MemoryStream saved = await SaveAsync(book);
        string definition = TestFiles.ReadPart(saved, "xl/pivotCache/pivotCacheDefinition1.xml");

        Assert.Contains("refreshOnLoad=\"1\"", definition, StringComparison.Ordinal);
        Assert.Equal(original.Length, definition.Length);
        Assert.Empty(TestFiles.Validate(saved));
    }

    [Fact]
    public async Task Pivot_tables_and_slicers_can_be_added_to_an_opened_workbook()
    {
        await using MemoryStream first = await SaveAsync(ExcelLikeWorkbook());
        Result<Workbook> opened = await Workbook.OpenAsync(first);
        Assert.True(opened.IsSuccess(out Workbook? book), opened.Error);
        Worksheet dash = Sheet(book, "Dash");
        Assert.True(book.AddSheet("More").IsSuccess(out Worksheet? more));

        Add(dash, new PivotTable("ByProduct", "Data!A1:F25") { Rows = ["Product"], Values = [new PivotValue("Units")] }, "J20");
        Add(more, new PivotTable("ByRegion", "Data!A1:F25") { Rows = ["Region"], Values = [new PivotValue("Sales")] }, "A1");
        Result region = more.AddSlicer(new Slicer("Region", ["ByProduct", "ByRegion"]) { Selected = ["South"] }, "D1:F8");
        Assert.True(region.IsSuccess(out _, out _), region.Error);

        // The file's pivot tables keep their names: a slicer cannot filter them, and a new one cannot take a name on the same sheet.
        Assert.Equal(ResultStatus.NotFound, more.AddSlicer(new Slicer("Year", ["PA"]), "H1:J5").Status);
        Assert.Equal(ResultStatus.Conflict, Sheet(book, "P").AddPivotTable(new PivotTable("PA", "Data!A1:F25") { Values = [new PivotValue("Sales")] }, "AA1").Status);

        await using MemoryStream saved = await SaveAsync(book);
        Assert.Empty(TestFiles.Validate(saved));
        string workbook = TestFiles.ReadPart(saved, "xl/workbook.xml");

        // The new cache is numbered after the file's, and the slicer cache joins the file's two in the same list.
        Assert.Matches("<pivotCaches><pivotCache cacheId=\"0\" r:id=\"rId[0-9]+\" /><pivotCache cacheId=\"1\" r:id=\"rId[0-9]+\" /></pivotCaches>", workbook);
        Assert.Matches("<x14:slicerCaches><x14:slicerCache r:id=\"rId[0-9]+\" /><x14:slicerCache r:id=\"rId[0-9]+\" /><x14:slicerCache r:id=\"rId[0-9]+\" /></x14:slicerCaches>", workbook);
        Assert.Contains("<definedName name=\"Slicer_Region\">#N/A</definedName>", workbook, StringComparison.Ordinal);

        Result<Workbook> again = await Workbook.OpenAsync(saved);
        Assert.True(again.IsSuccess(out Workbook? reopened), again.Error);
        Assert.Equal(9, reopened.PivotTables.Count);
        Assert.Equal(2, reopened.PivotCaches.Count);
        Assert.Equal((CellValue)1827, Evaluate(reopened, "GETPIVOTDATA(\"Sum of Sales\",More!A1)"));
        Assert.Equal((CellValue)"South", Sheet(reopened, "More")["A2"].Value);
        Assert.Equal((CellValue)"Grand Total", Sheet(reopened, "More")["A3"].Value);
        Assert.True(Sheet(reopened, "More")["A4"].Value.IsEmpty);
    }

    [Fact]
    public async Task Recalculation_renders_the_pivot_tables_from_the_changed_source()
    {
        Workbook book = SalesWorkbook(out Worksheet data, out Worksheet pivots);
        Add(pivots, new PivotTable("ByMonth", "Data!A1:F25") { Rows = ["Month"], Values = [new PivotValue("Sales")] }, "A1");
        pivots["D1"] = Cell.FromFormula("GETPIVOTDATA(\"Sum of Sales\",$A$1,\"Month\",\"Jan\")");
        pivots["D2"] = Cell.FromFormula("Data!E2*2");
        await book.RecalculateAsync();
        Assert.Equal((CellValue)798, pivots["D1"].Value);

        data[1, 4] = 1021;
        data[25, 0] = 2027;
        data[25, 1] = "Apr";
        data[25, 2] = "North";
        data[25, 3] = "A";
        data[25, 4] = 5;
        await book.RecalculateAsync();

        // The source range does not grow by itself: row 26 is outside A1:F25.
        Assert.Equal((CellValue)1798, pivots["D1"].Value);
        Assert.Equal((CellValue)2042, pivots["D2"].Value);
        Assert.Equal((CellValue)"Grand Total", pivots["A5"].Value);
    }

    [Fact]
    public void A_slicer_selection_filters_every_pivot_table_it_is_connected_to()
    {
        Workbook book = SalesWorkbook(out _, out Worksheet pivots);
        Add(pivots, new PivotTable("Left", "Data!A1:F25") { Rows = ["Region"], Values = [new PivotValue("Sales")] }, "A1");
        Add(pivots, new PivotTable("Right", "Data!A1:F25") { Rows = ["Month"], Columns = ["Year"], Values = [new PivotValue("Units")] }, "E1");
        Add(pivots, new PivotTable("Alone", "Data!A1:F25") { Rows = ["Region"], Values = [new PivotValue("Sales")] }, "A10");

        Result added = pivots.AddSlicer(new Slicer("Month", ["Left", "Right"]) { Selected = ["Feb", "mar"] }, "M1:O10");
        Assert.True(added.IsSuccess(out _, out _), added.Error);

        // Feb and Mar only.
        AssertCells(pivots, "A2=North B2=1218 A4=Grand Total B4=2604");
        AssertCells(pivots, "E3=Feb E4=Mar E5=Grand Total");
        AssertCells(pivots, "A11=North B11=1575 A13=Grand Total B13=3402");

        // A second slicer on the same field shares the selection.
        Result conflict = pivots.AddSlicer(new Slicer("Month", ["Left"]) { Selected = ["Jan"] }, "Q1:S10");
        Assert.Equal(ResultStatus.Conflict, conflict.Status);
        Result shared = pivots.AddSlicer(new Slicer("Month", ["Alone"]), "Q1:S10");
        Assert.True(shared.IsSuccess(out _, out _), shared.Error);
        AssertCells(pivots, "B13=2604");
    }

    [Fact]
    public void Wrong_definitions_give_results_or_exceptions()
    {
        Workbook book = SalesWorkbook(out Worksheet data, out Worksheet pivots);

        Assert.Equal(ResultStatus.NotFound, pivots.AddPivotTable(new PivotTable("X", "Nowhere!A1:F25") { Values = [new PivotValue("Sales")] }, "A1").Status);
        Assert.Equal(ResultStatus.NotFound, pivots.AddPivotTable(new PivotTable("X", "Data!A1:F25") { Rows = ["Colour"], Values = [new PivotValue("Sales")] }, "A1").Status);
        Assert.Equal(ResultStatus.NotFound, pivots.AddPivotTable(new PivotTable("X", "NoSuchTable") { Values = [new PivotValue("Sales")] }, "A1").Status);
        Assert.Equal(ResultStatus.Conflict, data.AddPivotTable(new PivotTable("X", "Data!A1:F25") { Rows = ["Month"], Values = [new PivotValue("Sales")] }, "C3").Status);
        Assert.Throws<ArgumentException>(() => pivots.AddPivotTable(new PivotTable("X", "Data!A1:F25"), "A1"));
        Assert.Throws<ArgumentException>(() => pivots.AddPivotTable(new PivotTable("X", "Data!A1:F25") { Rows = ["Month"], Columns = ["month"], Values = [new PivotValue("Sales")] }, "A1"));
        Assert.Throws<ArgumentException>(() => new Slicer("Month", []));
        Assert.Throws<ArgumentException>(() => new Slicer("Month", ["X"]) { Selected = [] });

        data[0, 6] = "  ";
        data[1, 6] = 1;
        Assert.Equal(ResultStatus.Invalid, pivots.AddPivotTable(new PivotTable("X", "Data!A1:G25") { Values = [new PivotValue("Sales")] }, "A1").Status);

        Add(pivots, new PivotTable("X", "Data!A1:F25") { Rows = ["Month"], Values = [new PivotValue("Sales")] }, "A1");
        Assert.Equal(ResultStatus.Conflict, pivots.AddPivotTable(new PivotTable("x", "Data!A1:F25") { Values = [new PivotValue("Sales")] }, "H1").Status);
        Assert.Equal(ResultStatus.Conflict, pivots.AddPivotTable(new PivotTable("Y", "Data!A1:F25") { Values = [new PivotValue("Sales")] }, "B2").Status);
        Assert.Equal(ResultStatus.NotFound, pivots.AddSlicer(new Slicer("Colour", ["X"]), "H1:J5").Status);
        Assert.Equal(ResultStatus.NotFound, pivots.AddSlicer(new Slicer("Month", ["X"]) { Selected = ["Dec"] }, "H1:J5").Status);
        Assert.Equal(ResultStatus.NotFound, pivots.AddSlicer(new Slicer("Month", ["Nope"]), "H1:J5").Status);
    }

    [Fact]
    public async Task A_source_with_text_numbers_dates_blanks_and_errors_is_cached_and_sorted_as_Excel_does()
    {
        Workbook book = Workbook.Create();
        Assert.True(book.AddSheet("Data").IsSuccess(out Worksheet? data));
        data["A1"] = "Kind";
        data["B1"] = "When";
        data["C1"] = "Amount";
        object?[][] rows =
        [
            ["Zeta", new DateTime(2026, 3, 1), 1.5],
            ["alpha", new DateTime(2026, 1, 1), null],
            [10, new DateTime(2026, 3, 1), 2],
            ["Alpha", null, true],
            [null, new DateTime(2025, 12, 31), "x"],
        ];

        for (int row = 0; row < rows.Length; row++)
        {
            for (int column = 0; column < 3; column++)
            {
                data[row + 1, column] = rows[row][column] switch
                {
                    string text => text,
                    int number => number,
                    double number => number,
                    bool flag => flag,
                    DateTime date => date,
                    _ => Cell.Empty,
                };
            }
        }

        data["C7"] = CellValue.Error("#N/A");
        data["A7"] = "Zeta";
        Add(data, new PivotTable("Kinds", "A1:C7") { Rows = ["Kind"], Values = [new PivotValue("Amount", PivotAggregation.Count), new PivotValue("Amount") { Name = "Total" }] }, "F1");

        // Numbers first, then text without regard to case, blanks last; Alpha and alpha are one item.
        AssertCells(data, "F2=10 F3=alpha F4=Zeta F5=(blank) F6=Grand Total G2=1 G3=1 H3=0 G4=2 H4=#N/A G5=1 H5=0");

        await using MemoryStream file = await SaveAsync(book);
        string definition = TestFiles.ReadPart(file, "xl/pivotCache/pivotCacheDefinition1.xml");
        Assert.Contains("<sharedItems containsBlank=\"1\" containsMixedTypes=\"1\" containsNumber=\"1\" containsInteger=\"1\" minValue=\"10\" maxValue=\"10\" count=\"4\"><s v=\"Zeta\" /><s v=\"alpha\" /><n v=\"10\" /><m /></sharedItems>", definition, StringComparison.Ordinal);
        Assert.Contains("<sharedItems containsNonDate=\"0\" containsDate=\"1\" containsString=\"0\" containsBlank=\"1\"", definition, StringComparison.Ordinal);
        Assert.Contains("minDate=\"2025-12-31T00:00:00\" maxDate=\"2026-03-01T00:00:00\" />", definition, StringComparison.Ordinal);
        Assert.Empty(TestFiles.Validate(file));
    }

    [Fact]
    public async Task A_Strict_OOXML_file_gets_no_pivot_tables()
    {
        const string strictMain = "http://purl.oclc.org/ooxml/spreadsheetml/main";
        const string strictRelationships = "http://purl.oclc.org/ooxml/officeDocument/relationships";
        const string packageRelationships = "http://schemas.openxmlformats.org/package/2006/relationships";

        await using MemoryStream file = TestFiles.Package(
            ("_rels/.rels", $"""<Relationships xmlns="{packageRelationships}"><Relationship Id="rId1" Type="{strictRelationships}/officeDocument" Target="xl/workbook.xml"/></Relationships>"""),
            ("xl/workbook.xml", $"""<workbook xmlns="{strictMain}" xmlns:r="{strictRelationships}"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>"""),
            ("xl/_rels/workbook.xml.rels", $"""<Relationships xmlns="{packageRelationships}"><Relationship Id="rId1" Type="{strictRelationships}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"""),
            ("xl/worksheets/sheet1.xml", $"""<worksheet xmlns="{strictMain}"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>N</t></is></c></row><row r="2"><c r="A2"><v>1</v></c></row></sheetData></worksheet>"""));

        Result<Workbook> opened = await Workbook.OpenAsync(file);
        Assert.True(opened.IsSuccess(out Workbook? book), opened.Error);
        Assert.True(book.AddSheet("New").IsSuccess(out Worksheet? added));

        foreach (Worksheet sheet in new[] { book.Sheets[0], added })
        {
            Result result = sheet.AddPivotTable(new PivotTable("X", "Data!A1:A2") { Values = [new PivotValue("N")] }, "C1");
            Assert.True(result.IsFailure(out string? error, out ResultStatus status));
            Assert.Equal(ResultStatus.Invalid, status);
            Assert.Contains("Strict", error, StringComparison.Ordinal);
        }
    }

    [Fact]
    public async Task RemoveActiveContent_keeps_pivot_tables_on_worksheet_data()
    {
        await using MemoryStream file = await SaveAsync(ExcelLikeWorkbook());
        Result<Workbook> opened = await Workbook.OpenAsync(file);
        Assert.True(opened.IsSuccess(out Workbook? book), opened.Error);

        Result removed = book.RemoveActiveContent();
        Assert.True(removed.IsSuccess(out _, out _), removed.Error);
        await using MemoryStream cleaned = await SaveAsync(book);

        foreach (string part in new[] { "xl/pivotCache/pivotCacheDefinition1.xml", "xl/pivotCache/pivotCacheRecords1.xml", "xl/pivotTables/pivotTable7.xml", "xl/slicerCaches/slicerCache2.xml", "xl/slicers/slicer1.xml" })
        {
            Assert.Equal(TestFiles.ReadPart(file, part), TestFiles.ReadPart(cleaned, part));
        }

        Assert.Empty(TestFiles.Validate(cleaned));
        Result<Workbook> again = await Workbook.OpenAsync(cleaned);
        Assert.True(again.IsSuccess(out Workbook? reopened), again.Error);
        Assert.Equal(7, reopened.PivotTables.Count);
    }

    [Fact]
    public async Task A_pivot_part_that_is_not_XML_leaves_the_file_open_and_kept()
    {
        await using MemoryStream file = await SaveAsync(ExcelLikeWorkbook());
        await using MemoryStream broken = Replace(file, "xl/pivotTables/pivotTable1.xml", "<pivotTableDefinition><location ref=");
        await using MemoryStream brokenCache = Replace(file, "xl/pivotCache/pivotCacheDefinition1.xml", "<<<");

        Result<Workbook> opened = await Workbook.OpenAsync(broken);
        Assert.True(opened.IsSuccess(out Workbook? book), opened.Error);
        Assert.Equal(6, book.PivotTables.Count);
        await using MemoryStream saved = await SaveAsync(book);
        Assert.Equal("<pivotTableDefinition><location ref=", TestFiles.ReadPart(saved, "xl/pivotTables/pivotTable1.xml"));

        Result<Workbook> withoutCache = await Workbook.OpenAsync(brokenCache);
        Assert.True(withoutCache.IsSuccess(out Workbook? noCache), withoutCache.Error);
        Assert.Empty(noCache.PivotTables);
        Assert.Empty(noCache.PivotCaches);
        Assert.Equal(CellValue.Error("#REF!"), Evaluate(noCache, "GETPIVOTDATA(\"Sum of Sales\",P!A1)"));
    }

    [Fact]
    public async Task GETPIVOTDATA_is_a_known_function_and_follows_its_arguments()
    {
        Workbook book = ExcelLikeWorkbook();
        Worksheet dash = Sheet(book, "Dash");
        dash["A10"] = "Mar";
        dash["B10"] = Cell.FromFormula("GETPIVOTDATA(\"Sum of Sales\",P!$A$1,\"Month\",A10)");
        dash["C10"] = Cell.FromFormula("GETPIVOTDATA(\"Sum of Sales\",P!$A$1,\"Month\")");
        dash["D10"] = Cell.FromFormula("GETPIVOTDATA(\"Sum of Sales\",\"P!A1\")");
        dash["E10"] = Cell.FromFormula("GETPIVOTDATA(1/0,P!$A$1)");
        dash["F10"] = Cell.FromFormula("GETPIVOTDATA(\"sum of sales\",P!B3,\"MONTH\",\"feb\")");

        RecalculationReport report = await book.RecalculateAsync(RecalculationScope.All);

        Assert.Empty(report.UnknownFunctions);
        Assert.Equal((CellValue)987, dash["B10"].Value);
        Assert.Equal(CellValue.Error("#REF!"), dash["C10"].Value);
        Assert.Equal(CellValue.Error("#REF!"), dash["D10"].Value);
        Assert.Equal(CellValue.Error("#DIV/0!"), dash["E10"].Value);
        Assert.Equal((CellValue)819, dash["F10"].Value);

        // A slicer selection changes what the pivot table shows, and so what GETPIVOTDATA gives.
        Assert.Equal(CellValue.Error("#REF!"), Evaluate(book, "GETPIVOTDATA(\"Sum of Sales\",P!$E$1,\"Year\",2025)"));
    }

    private static void Add(Worksheet sheet, PivotTable table, string topLeft)
    {
        Result added = sheet.AddPivotTable(table, topLeft);
        Assert.True(added.IsSuccess(out _, out _), added.Error);
    }

    private static Worksheet Sheet(Workbook book, string name)
    {
        Assert.True(book.Sheet(name).IsSuccess(out Worksheet? sheet));

        return sheet;
    }

    private static CellValue Evaluate(Workbook book, string formula)
    {
        Result<CellValue> result = book.Evaluate(formula);
        Assert.True(result.IsSuccess(out CellValue value), result.Error);

        return value;
    }

    /// <summary>Checks cells written as "A1=text B2=12.5": a number when the text is one, an error when it starts with #.</summary>
    private static void AssertCells(Worksheet sheet, string expected)
    {
        foreach (Match match in CellPattern().Matches(expected))
        {
            string address = match.Groups[1].Value;
            string text = match.Groups[2].Value;
            CellValue actual = sheet[address].Value;
            CellValue wanted = text;

            if (double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out double number) == true)
            {
                wanted = number;
            }
            else if (text.StartsWith('#') == true)
            {
                wanted = CellValue.Error(text);
            }

            Assert.True(wanted == actual, $"{address}: expected {wanted}, got {actual}.");
        }
    }

    private static async Task<MemoryStream> SaveAsync(Workbook book)
    {
        MemoryStream file = new();
        await book.SaveAsync(file);
        file.Position = 0;

        return file;
    }

    /// <summary>The pivot table parts by the name of their table.</summary>
    private static Dictionary<string, string> PivotTableParts(MemoryStream file)
    {
        Dictionary<string, string> parts = [];

        for (int number = 1; number <= 20; number++)
        {
            try
            {
                string part = TestFiles.ReadPart(file, $"xl/pivotTables/pivotTable{number}.xml");
                parts[Regex.Match(part, "<pivotTableDefinition [^>]*?name=\"([^\"]+)\"").Groups[1].Value] = part;
            }
            catch (InvalidOperationException)
            {
                break;
            }
        }

        return parts;
    }

    private static MemoryStream Replace(MemoryStream file, string path, string content)
    {
        file.Position = 0;
        MemoryStream copy = new();

        using (System.IO.Compression.ZipArchive source = new(file, System.IO.Compression.ZipArchiveMode.Read, true))
        using (System.IO.Compression.ZipArchive target = new(copy, System.IO.Compression.ZipArchiveMode.Create, true))
        {
            foreach (System.IO.Compression.ZipArchiveEntry entry in source.Entries)
            {
                System.IO.Compression.ZipArchiveEntry created = target.CreateEntry(entry.FullName);
                using Stream output = created.Open();

                if (entry.FullName == path)
                {
                    byte[] bytes = System.Text.Encoding.UTF8.GetBytes(content);
                    output.Write(bytes, 0, bytes.Length);

                    continue;
                }

                using Stream input = entry.Open();
                input.CopyTo(output);
            }
        }

        copy.Position = 0;

        return copy;
    }

    [GeneratedRegex("([A-Z]+[0-9]+)=(.*?)(?= [A-Z]+[0-9]+=|$)")]
    private static partial Regex CellPattern();
}
