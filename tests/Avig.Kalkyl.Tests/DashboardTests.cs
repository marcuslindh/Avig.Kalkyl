using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using Excel = DocumentFormat.OpenXml.Spreadsheet;

namespace Avig.Kalkyl.Tests;

/// <summary>
/// Shapes, groups, SVG pictures, chart options, sheet views and hyperlinks: the parts a
/// dashboard is made of. The XML is checked against what Excel 365 writes (a dashboard template
/// saved by Excel was the model), against Microsoft's schema, and through a round trip in the
/// document model. Excel itself opened the dashboard without repair (ExcelCheck files 14 and 15).
/// </summary>
public class DashboardTests
{
    private const string Svg = """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 12"><rect width="24" height="12" fill="#2384F9"/></svg>""";

    [Fact]
    public async Task Every_shape_kind_from_both_writers_passes_schema_validation()
    {
        await using MemoryStream streamed = await TestFiles.WriteAsync(async writer =>
        {
            SheetWriter sheet = await TestFiles.AddSheetAsync(writer, "Shapes");

            foreach ((Shape shape, DrawingBounds bounds) in EveryShape())
            {
                sheet.AddShape(shape, bounds);
            }

            sheet.AddShape(new Shape(ShapeKind.Ellipse), "B30:D34");
            sheet.AddShape(new Shape(ShapeKind.Rectangle), new CellRange(35, 1, 38, 3));
        });

        Workbook book = Workbook.Create();
        Worksheet modelled = Added(book, "Shapes");

        foreach ((Shape shape, DrawingBounds bounds) in EveryShape())
        {
            Assert.True(modelled.AddShape(shape, bounds).IsSuccess(out _, out _));
        }

        await using MemoryStream saved = await SaveAsync(book);

        AssertValid(streamed);
        AssertValid(saved);
        string drawing = TestFiles.ReadPart(saved, "xl/drawings/drawing1.xml");
        Assert.Equal(TestFiles.ReadPart(streamed, "xl/drawings/drawing1.xml")[..(drawing.Length / 2)], drawing[..(drawing.Length / 2)]);

        foreach (string preset in (string[])["rect", "roundRect", "ellipse", "triangle", "diamond", "hexagon", "rightArrow", "leftArrow", "upArrow", "downArrow", "chevron", "homePlate", "line"])
        {
            Assert.Contains($"prst=\"{preset}\"", drawing, StringComparison.Ordinal);
        }

        Assert.Single(Regex.Matches(drawing, "<xdr:cxnSp "));
        Assert.Contains("<xdr:cNvSpPr txBox=\"1\" />", drawing.Replace("\"/>", "\" />", StringComparison.Ordinal), StringComparison.Ordinal);

        // A shape over cells stretches with them: no editAs, unlike a chart or a picture.
        string cells = TestFiles.ReadPart(streamed, "xl/drawings/drawing1.xml");
        Assert.Contains("<xdr:twoCellAnchor><xdr:from><xdr:col>1</xdr:col>", cells, StringComparison.Ordinal);
    }

    [Fact]
    public async Task A_shape_is_written_as_Excel_writes_a_card()
    {
        Shape card = new(ShapeKind.RoundedRectangle)
        {
            CornerRadius = 0.0945,
            Fill = Fill.Gradient(45, new GradientStop(1, "#605776"), new GradientStop(0.17, "#2384f9")),
            Outline = Outline.None,
            Shadow = new Shadow(),
            InnerShadow = new Shadow { Blur = 10, Distance = 0, Transparency = 0.76 },
            SoftEdges = 2,
            Rotation = -90,
            FlipVertical = true,
            Name = "Card",
            Description = "Projects card",
            Text = new ShapeText([new TextRun("PROJECT\n", new TextFont { Bold = true, Size = 12, Color = "#FFFFFF" }), new TextRun("DASHBOARD", new TextFont { Bold = true, Size = 12, Color = "#8BFFF7", Name = "Segoe UI" })])
            {
                Horizontal = HorizontalAlignment.Right,
                Vertical = VerticalAlignment.Bottom,
                Wrap = false,
                MarginLeft = 10,
                MarginTop = 0,
            },
        };

        string drawing = await DrawingOfAsync(sheet => sheet.AddShape(card, DrawingBounds.FromPoints(10, 20, 160, 90)));

        Assert.Contains("<xdr:absoluteAnchor><xdr:pos x=\"127000\" y=\"254000\" /><xdr:ext cx=\"2032000\" cy=\"1143000\" />", Spaced(drawing), StringComparison.Ordinal);
        Assert.Contains("<xdr:cNvPr id=\"2\" name=\"Card\" descr=\"Projects card\" />", Spaced(drawing), StringComparison.Ordinal);
        Assert.Contains("<a:xfrm rot=\"16200000\" flipV=\"1\"><a:off x=\"127000\" y=\"254000\" />", Spaced(drawing), StringComparison.Ordinal);
        Assert.Contains("<a:prstGeom prst=\"roundRect\"><a:avLst><a:gd name=\"adj\" fmla=\"val 9450\" /></a:avLst></a:prstGeom>", Spaced(drawing), StringComparison.Ordinal);

        // Stops in position order, the angle in 60 000ths of a degree, as in Excel's own cards.
        Assert.Contains("<a:gsLst><a:gs pos=\"17000\"><a:srgbClr val=\"2384F9\" /></a:gs><a:gs pos=\"100000\"><a:srgbClr val=\"605776\" /></a:gs></a:gsLst><a:lin ang=\"2700000\" scaled=\"0\" />", Spaced(drawing), StringComparison.Ordinal);
        Assert.Contains("<a:ln><a:noFill /></a:ln>", Spaced(drawing), StringComparison.Ordinal);
        Assert.Contains("<a:effectLst><a:innerShdw blurRad=\"127000\" dist=\"0\" dir=\"2700000\"><a:srgbClr val=\"000000\"><a:alpha val=\"24000\" /></a:srgbClr></a:innerShdw>"
            + "<a:outerShdw blurRad=\"50800\" dist=\"38100\" dir=\"2700000\" algn=\"tl\" rotWithShape=\"0\"><a:srgbClr val=\"000000\"><a:alpha val=\"40000\" /></a:srgbClr></a:outerShdw>"
            + "<a:softEdge rad=\"25400\" /></a:effectLst>", Spaced(drawing), StringComparison.Ordinal);
        Assert.Contains("<a:bodyPr vertOverflow=\"clip\" horzOverflow=\"clip\" wrap=\"none\" lIns=\"127000\" tIns=\"0\" rtlCol=\"0\" anchor=\"b\" />", Spaced(drawing), StringComparison.Ordinal);

        // \n starts a paragraph; each run keeps its font.
        Assert.Equal(2, Regex.Count(drawing, "<a:p>"));
        Assert.Contains("<a:r><a:rPr lang=\"en-US\" sz=\"1200\" b=\"1\" i=\"0\"><a:solidFill><a:srgbClr val=\"8BFFF7\" /></a:solidFill><a:latin typeface=\"Segoe UI\" /><a:cs typeface=\"Segoe UI\" /></a:rPr><a:t>DASHBOARD</a:t></a:r>", Spaced(drawing), StringComparison.Ordinal);
        Assert.Contains("<a:pPr algn=\"r\" />", Spaced(drawing), StringComparison.Ordinal);
    }

    [Fact]
    public async Task Outlines_and_solid_fills_take_widths_dashes_and_transparency()
    {
        Shape line = new(ShapeKind.Line) { Outline = new Outline("#3a4a6b", 1.5) { Dash = LineDash.LongDashDot, Transparency = 0.25 }, FlipHorizontal = true };

        Shape box = new(ShapeKind.Rectangle) { Fill = Fill.Solid("#FFFFFF", 0.85), Outline = new Outline("#000000") };

        string drawing = await DrawingOfAsync(sheet =>
        {
            sheet.AddShape(line, DrawingBounds.FromPoints(0, 0, 100, 0));
            sheet.AddShape(box, DrawingBounds.FromPoints(0, 10, 100, 50));
        });

        Assert.Contains("<a:ln w=\"19050\"><a:solidFill><a:srgbClr val=\"3A4A6B\"><a:alpha val=\"75000\" /></a:srgbClr></a:solidFill><a:prstDash val=\"lgDashDot\" /></a:ln>", Spaced(drawing), StringComparison.Ordinal);
        Assert.Contains("<a:xfrm flipH=\"1\">", Spaced(drawing), StringComparison.Ordinal);
        Assert.Contains("<a:solidFill><a:srgbClr val=\"FFFFFF\"><a:alpha val=\"15000\" /></a:srgbClr></a:solidFill><a:ln w=\"9525\">", Spaced(drawing), StringComparison.Ordinal);

        // A line has no inside and no text.
        string connector = Regex.Match(drawing, "<xdr:cxnSp .*?</xdr:cxnSp>").Value;
        Assert.DoesNotContain("txBody", connector, StringComparison.Ordinal);
        Assert.Contains("<a:lnRef idx=\"1\">", Spaced(connector), StringComparison.Ordinal);
    }

    [Fact]
    public async Task Links_inside_the_workbook_are_fragments_and_web_links_are_external()
    {
        Hyperlink toSheet = Hyperlink.ToCell("Sales 2026", "B3") with { Tooltip = "Go to sales" };

        Hyperlink toWeb = Hyperlink.ToUrl("https://example.com/report?id=1&x=2");

        await using MemoryStream file = await TestFiles.WriteAsync(async writer =>
        {
            SheetWriter sheet = await TestFiles.AddSheetAsync(writer, "Dashboard");
            sheet.AddShape(new Shape(ShapeKind.RoundedRectangle) { Link = toSheet }, DrawingBounds.FromPoints(0, 0, 100, 30));
            sheet.AddShape(new Shape(ShapeKind.RoundedRectangle) { Link = toSheet }, DrawingBounds.FromPoints(0, 40, 100, 30));
            sheet.AddPicture(DrawingTests.PngPicture() with { Link = toWeb }, DrawingBounds.FromPoints(0, 80, 30, 30));
            await TestFiles.AddSheetAsync(writer, "Sales 2026");
        });

        AssertValid(file);
        string drawing = Spaced(TestFiles.ReadPart(file, "xl/drawings/drawing1.xml"));
        string relationships = TestFiles.ReadPart(file, "xl/drawings/_rels/drawing1.xml.rels");

        // One relationship per link, shared by the shapes that use it.
        Assert.Equal(2, Regex.Count(drawing, "<a:hlinkClick r:id=\"rId1\" tooltip=\"Go to sales\" />"));
        Assert.Contains("Target=\"#'Sales 2026'!B3\"", relationships, StringComparison.Ordinal);
        Assert.Contains("Target=\"https://example.com/report?id=1&amp;x=2\" TargetMode=\"External\"", relationships, StringComparison.Ordinal);
        Assert.DoesNotContain("#'Sales 2026'!B3\" TargetMode", relationships, StringComparison.Ordinal);

        Workbook book = await OpenAsync(file);
        Assert.Empty(book.ExternalReferences);
    }

    [Fact]
    public void Mistakes_in_shapes_and_links_are_programming_errors()
    {
        Assert.Throws<ArgumentException>(() => Fill.Solid("blue"));
        Assert.Throws<ArgumentOutOfRangeException>(() => Fill.Solid("#000000", 1.5));
        Assert.Throws<ArgumentException>(() => Fill.Gradient(0, new GradientStop(0, "#000000")));
        Assert.Throws<ArgumentOutOfRangeException>(() => new GradientStop(1.1, "#000000"));
        Assert.Throws<ArgumentOutOfRangeException>(() => new Outline("#000000", -1));
        Assert.Throws<ArgumentOutOfRangeException>(() => new Shadow { Blur = 101 });
        Assert.Throws<ArgumentOutOfRangeException>(() => new Shape(ShapeKind.RoundedRectangle) { CornerRadius = 0.6 });
        Assert.Throws<ArgumentOutOfRangeException>(() => new Shape((ShapeKind)99));
        Assert.Throws<ArgumentOutOfRangeException>(() => new TextFont { Size = 0.5 });
        Assert.Throws<ArgumentException>(() => new ShapeText([]));
        Assert.Throws<ArgumentException>(() => Hyperlink.ToUrl("javascript:alert(1)"));
        Assert.Throws<ArgumentException>(() => Hyperlink.ToUrl("report.xlsx"));
        Assert.Throws<ArgumentException>(() => Hyperlink.ToCell("Bad[name]", "A1"));
        Assert.Throws<ArgumentException>(() => Hyperlink.ToCell("Sales", "A0"));
        Assert.Throws<ArgumentOutOfRangeException>(() => DrawingBounds.FromPoints(-1, 0, 10, 10));
        Assert.Throws<ArgumentOutOfRangeException>(() => DrawingBounds.FromPoints(0, 0, double.NaN, 10));
        Assert.Throws<ArgumentException>(() => new ShapeGroup([]));
        Assert.Throws<ArgumentOutOfRangeException>(() => new SheetView { Zoom = 5 });
        Assert.Throws<ArgumentException>(() => new SheetView { TabColor = "red" });
    }

    [Fact]
    public async Task A_group_maps_its_items_onto_its_place_and_links_as_one()
    {
        Picture icon = SvgPicture();
        ShapeGroup group = new(
        [
            GroupItem.Of(new Shape(ShapeKind.RoundedRectangle) { Fill = Fill.Solid("#232D45") }, DrawingBounds.FromPoints(10, 10, 145, 40)),
            GroupItem.Of(icon, DrawingBounds.FromPoints(20, 19, 22, 22)),
            GroupItem.Of(new Shape(ShapeKind.TextBox) { Text = new ShapeText("Projects") }, DrawingBounds.FromPoints(52, 10, 100, 40)),
        ])
        {
            Link = Hyperlink.ToCell("Projects", "A1"),
            Name = "Go to Projects",
        };

        Workbook book = Workbook.Create();
        Worksheet sheet = Added(book, "Dashboard");
        Assert.True(sheet.AddGroup(group, "B2:D4").IsSuccess(out _, out _));
        Assert.True(sheet.AddGroup(group, DrawingBounds.FromPoints(20, 100, 145, 40)).IsSuccess(out _, out _));
        await using MemoryStream file = await SaveAsync(book);

        AssertValid(file);
        string drawing = Spaced(TestFiles.ReadPart(file, "xl/drawings/drawing1.xml"));

        // The items' box (10, 10)-(155, 50) in points is what the group's place shows.
        Assert.Equal(2, Regex.Count(drawing, "<a:chOff x=\"127000\" y=\"127000\" /><a:chExt cx=\"1841500\" cy=\"508000\" />"));
        Assert.Contains("<xdr:cNvPr id=\"2\" name=\"Go to Projects\"><a:hlinkClick r:id=\"rId1\" /></xdr:cNvPr>", drawing, StringComparison.Ordinal);

        // Every item has its own id; the second group continues the count.
        List<int> ids = Regex.Matches(drawing, "<xdr:cNvPr id=\"(\\d+)\"").Select(match => int.Parse(match.Groups[1].Value, System.Globalization.CultureInfo.InvariantCulture)).ToList();
        Assert.Equal([2, 3, 4, 5, 6, 7, 8, 9], ids);

        // The SVG and its fallback are stored once for both groups.
        Assert.Equal(2, Regex.Count(TestFiles.ReadPart(file, "xl/drawings/_rels/drawing1.xml.rels"), "relationships/image"));
    }

    [Fact]
    public async Task An_svg_is_stored_as_Excel_stores_it_with_a_raster_fallback()
    {
        Picture fallback = DrawingTests.PngPicture();
        Assert.True(Picture.FromSvg(Encoding.UTF8.GetBytes(Svg), fallback).IsSuccess(out Picture? svg));
        Assert.Equal((PictureFormat.Svg, 24, 12), (svg!.Format, svg.Width, svg.Height));

        await using MemoryStream file = await TestFiles.WriteAsync(async writer =>
        {
            SheetWriter sheet = await TestFiles.AddSheetAsync(writer, "Icons");
            sheet.AddPicture(svg with { Description = "Bar", Transparency = 0.25, Shadow = new Shadow() }, "B2");
        });

        AssertValid(file);
        string drawing = Spaced(TestFiles.ReadPart(file, "xl/drawings/drawing1.xml"));
        Assert.Contains("<a:blip r:embed=\"rId1\"><a:alphaModFix amt=\"75000\" /><a:extLst><a:ext uri=\"{96DAC541-7B7A-43D3-8B79-37D633B846F1}\">"
            + "<asvg:svgBlip r:embed=\"rId2\" xmlns:asvg=\"http://schemas.microsoft.com/office/drawing/2016/SVG/main\" /></a:ext></a:extLst></a:blip>", drawing, StringComparison.Ordinal);
        Assert.Contains("<a:outerShdw", drawing, StringComparison.Ordinal);

        // At its own size: 24 × 12 pixels of 9525 EMU.
        Assert.Contains("<xdr:ext cx=\"228600\" cy=\"114300\" />", drawing, StringComparison.Ordinal);
        Assert.Contains("Extension=\"svg\" ContentType=\"image/svg+xml\"", TestFiles.ReadPart(file, "[Content_Types].xml"), StringComparison.Ordinal);
        Assert.Equal(Svg, TestFiles.ReadPart(file, "xl/media/image2.svg"));

        using ZipArchive archive = new(file, ZipArchiveMode.Read, true);
        Assert.Equal(DrawingTests.Png.Length, archive.GetEntry("xl/media/image1.png")!.Length);
    }

    [Theory]
    [InlineData("<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"48px\" height=\"36pt\"/>", 48, 48)]
    [InlineData("<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"1in\" height=\"2.54cm\"/>", 96, 96)]
    [InlineData("<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"100%\" viewBox=\"0, 0, 640, 480\"/>", 640, 480)]
    [InlineData("<?xml version=\"1.0\"?><!DOCTYPE svg PUBLIC \"-//W3C//DTD SVG 1.1//EN\" \"http://www.w3.org/Graphics/SVG/1.1/DTD/svg11.dtd\"><svg xmlns=\"http://www.w3.org/2000/svg\"/>", 96, 96)]
    public void The_size_of_an_svg_comes_from_its_width_and_height_or_its_view_box(string svg, int width, int height)
    {
        Assert.True(Picture.FromSvg(Encoding.UTF8.GetBytes(svg)).IsSuccess(out Picture? picture));
        Assert.Equal((width, height), (picture!.Width, picture.Height));
    }

    [Theory]
    [InlineData("")]
    [InlineData("<html><body/></html>")]
    [InlineData("<svg")]
    [InlineData("\u0089PNG")]
    [InlineData("<!DOCTYPE svg [<!ENTITY a \"aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa\"><!ENTITY b \"&a;&a;&a;&a;&a;&a;&a;&a;&a;&a;&a;&a;&a;&a;&a;&a;&a;&a;&a;&a;&a;&a;&a;\"><!ENTITY c \"&b;&b;&b;&b;&b;&b;&b;&b;&b;&b;&b;&b;&b;&b;&b;&b;&b;&b;&b;&b;&b;&b;&b;&b;&b;\">]><svg width=\"&c;\"/>")]
    public void Bytes_that_are_not_an_svg_are_Invalid(string text)
    {
        Assert.True(Picture.FromSvg(Encoding.UTF8.GetBytes(text)).IsFailure(out _, out ResultStatus status));
        Assert.Equal(ResultStatus.Invalid, status);
        Assert.Throws<ArgumentException>(() => Picture.FromSvg(Encoding.UTF8.GetBytes(Svg), SvgPicture()));
    }

    [Fact]
    public async Task Chart_options_are_written_where_the_schema_puts_them()
    {
        Chart chart = new(ChartType.Column)
        {
            Title = "Tasks",
            TitleFont = new TextFont { Size = 12, Bold = true, Color = "#FFFFFF" },
            Series =
            [
                new ChartSeries("Done", "B2:B7", "A2:A7", "#3A4A6B")
                {
                    PointColors = [null, null, "#00C5AE"],
                    Outline = new Outline("#FFFFFF", 0.5),
                    DataLabels = new DataLabels { Position = DataLabelPosition.InsideEnd, NumberFormat = "0.0", Font = new TextFont { Size = 8, Color = "#FFFFFF" } },
                },
                new ChartSeries("Plan", "C2:C7", "A2:A7") { Type = ChartType.Line, Smooth = true, Marker = MarkerStyle.Diamond, MarkerSize = 7, Outline = new Outline("#E1541C", 2.25), Shadow = new Shadow { Angle = 90 } },
                new ChartSeries("Share", "D2:D7", "A2:A7") { Type = ChartType.Area, SecondaryAxis = true, Fill = Fill.Gradient(90, new GradientStop(0, "#E1541C", 0.8), new GradientStop(1, "#FFFFFF", 1)) },
            ],
            GapWidth = 60,
            Overlap = -10,
            VaryColors = false,
            RoundedCorners = true,
            Background = Fill.Solid("#1F2940"),
            Border = Outline.None,
            Font = new TextFont { Size = 9, Color = "#AEB8CC" },
            LegendFont = new TextFont { Italic = true },
            ValueAxis = new ChartAxis { Minimum = 0, Maximum = 100, MajorUnit = 25, GridlineOutline = new Outline("#2E3850", 0.5), Line = Outline.None, NumberFormat = "0\" t\"" },
            CategoryAxis = new ChartAxis { Visible = false },
            SecondaryValueAxis = new ChartAxis { Gridlines = false, Maximum = 1 },
        };

        string xml = Spaced(await ChartOfAsync(chart));

        Assert.Contains("<c:roundedCorners val=\"1\" />", xml, StringComparison.Ordinal);

        // Areas behind columns behind lines; the secondary group last, on its own axes.
        Assert.Matches("<c:barChart>.*</c:barChart><c:lineChart>.*</c:lineChart><c:areaChart>.*<c:axId val=\"500000003\" /><c:axId val=\"500000004\" /></c:areaChart>", xml);
        Assert.Contains("<c:varyColors val=\"0\" />", xml, StringComparison.Ordinal);
        Assert.Contains("<c:gapWidth val=\"60\" /><c:overlap val=\"-10\" />", xml, StringComparison.Ordinal);
        Assert.Contains("<c:dPt><c:idx val=\"2\" /><c:invertIfNegative val=\"0\" /><c:bubble3D val=\"0\" /><c:spPr><a:solidFill><a:srgbClr val=\"00C5AE\" /></a:solidFill><a:ln w=\"6350\">", xml, StringComparison.Ordinal);
        Assert.Contains("<c:dLbls><c:numFmt formatCode=\"0.0\" sourceLinked=\"0\" />", xml, StringComparison.Ordinal);
        Assert.Contains("<c:dLblPos val=\"inEnd\" /><c:showLegendKey val=\"0\" /><c:showVal val=\"1\" />", xml, StringComparison.Ordinal);
        Assert.Contains("<c:marker><c:symbol val=\"diamond\" /><c:size val=\"7\" /><c:spPr><a:solidFill><a:srgbClr val=\"E1541C\" />", xml, StringComparison.Ordinal);
        Assert.Contains("<c:smooth val=\"1\" /></c:ser>", xml, StringComparison.Ordinal);
        Assert.Contains("<a:outerShdw blurRad=\"50800\" dist=\"38100\" dir=\"5400000\"", xml, StringComparison.Ordinal);
        Assert.Contains("<a:gs pos=\"0\"><a:srgbClr val=\"E1541C\"><a:alpha val=\"20000\" /></a:srgbClr></a:gs>", xml, StringComparison.Ordinal);

        // Axes: the category axis hidden, the value axis with a range and a step, the secondary on the right.
        Assert.Matches("<c:catAx><c:axId val=\"500000001\" /><c:scaling><c:orientation val=\"minMax\" /></c:scaling><c:delete val=\"1\" />", xml);
        Assert.Contains("<c:scaling><c:orientation val=\"minMax\" /><c:max val=\"100\" /><c:min val=\"0\" /></c:scaling><c:delete val=\"0\" /><c:axPos val=\"l\" /><c:majorGridlines><c:spPr><a:ln w=\"6350\">", xml, StringComparison.Ordinal);
        Assert.Contains("<c:numFmt formatCode=\"0&quot; t&quot;\" sourceLinked=\"0\" />", xml, StringComparison.Ordinal);
        Assert.Contains("<c:crossBetween val=\"between\" /><c:majorUnit val=\"25\" /></c:valAx>", xml, StringComparison.Ordinal);
        Assert.Contains("<c:axId val=\"500000004\" /><c:scaling><c:orientation val=\"minMax\" /><c:max val=\"1\" /></c:scaling><c:delete val=\"0\" /><c:axPos val=\"r\" />", xml, StringComparison.Ordinal);
        Assert.Contains("<c:crosses val=\"max\" />", xml, StringComparison.Ordinal);

        // The plot area shows the background; the chart area has it and no border.
        Assert.Contains("<c:spPr><a:noFill /></c:spPr></c:plotArea>", xml, StringComparison.Ordinal);
        Assert.Contains("</c:chart><c:spPr><a:solidFill><a:srgbClr val=\"1F2940\" /></a:solidFill><a:ln><a:noFill /></a:ln></c:spPr><c:txPr>", xml, StringComparison.Ordinal);
        Assert.Contains("<a:defRPr sz=\"900\" b=\"0\" i=\"0\"><a:solidFill><a:srgbClr val=\"AEB8CC\" />", xml, StringComparison.Ordinal);
        Assert.Contains("<c:legend><c:legendPos val=\"r\" /><c:overlay val=\"0\" /><c:txPr>", xml, StringComparison.Ordinal);
        Assert.Contains("<a:rPr lang=\"en-US\" sz=\"1200\" b=\"1\" i=\"0\">", xml, StringComparison.Ordinal);
    }

    [Fact]
    public async Task A_doughnut_gauge_has_its_hole_start_and_slice_colours()
    {
        Chart gauge = new(ChartType.Doughnut)
        {
            Series = [new ChartSeries("Share", "B2:B3", "A2:A3", "#FF0000") { PointColors = ["#00C5AE", "#2E3850"], Outline = Outline.None }],
            HoleSize = 75,
            FirstSliceAngle = 270,
            Legend = LegendPosition.None,
            DataLabels = new DataLabels { ShowValue = false, ShowPercent = true },
        };

        string xml = Spaced(await ChartOfAsync(gauge));

        Assert.Contains("<c:varyColors val=\"1\" />", xml, StringComparison.Ordinal);
        Assert.Contains("<c:firstSliceAng val=\"270\" /><c:holeSize val=\"75\" /></c:doughnutChart>", xml, StringComparison.Ordinal);
        Assert.Contains("<c:dPt><c:idx val=\"0\" /><c:bubble3D val=\"0\" /><c:spPr><a:solidFill><a:srgbClr val=\"00C5AE\" /></a:solidFill><a:ln><a:noFill /></a:ln></c:spPr></c:dPt>", xml, StringComparison.Ordinal);
        Assert.Contains("<c:showVal val=\"0\" /><c:showCatName val=\"0\" /><c:showSerName val=\"0\" /><c:showPercent val=\"1\" />", xml, StringComparison.Ordinal);

        // A pie colours its slices, so the series' own colour is not written.
        Assert.DoesNotContain("FF0000", xml, StringComparison.Ordinal);
        Assert.DoesNotContain("dLblPos", xml, StringComparison.Ordinal);
        Assert.DoesNotContain("<c:legend>", xml, StringComparison.Ordinal);
    }

    [Fact]
    public async Task Default_charts_are_written_as_before()
    {
        string xml = Spaced(await ChartOfAsync(new Chart(ChartType.Column) { Series = [new ChartSeries("A", "B2:B7")] }));

        Assert.Contains("<c:roundedCorners val=\"0\" />", xml, StringComparison.Ordinal);
        Assert.Contains("<c:gapWidth val=\"150\" />", xml, StringComparison.Ordinal);
        Assert.DoesNotContain("overlap", xml, StringComparison.Ordinal);
        Assert.DoesNotContain("<c:spPr>", xml, StringComparison.Ordinal);
        Assert.DoesNotContain("txPr", xml, StringComparison.Ordinal);
        Assert.DoesNotContain("dLbls", xml, StringComparison.Ordinal);
        Assert.Contains("<c:valAx><c:axId val=\"500000002\" /><c:scaling><c:orientation val=\"minMax\" /></c:scaling><c:delete val=\"0\" /><c:axPos val=\"l\" /><c:majorGridlines />", xml, StringComparison.Ordinal);
    }

    [Fact]
    public void Chart_options_Excel_would_repair_are_programming_errors()
    {
        Workbook book = Workbook.Create();
        Worksheet sheet = Added(book, "Data");
        ChartSeries series = new("A", "B2:B7");

        Assert.Throws<ArgumentException>(() => sheet.AddChart(new Chart(ChartType.Column) { Series = [series, new ChartSeries("B", "C2:C7") { Type = ChartType.Pie }] }, "D2:J20"));
        Assert.Throws<ArgumentException>(() => sheet.AddChart(new Chart(ChartType.Bar) { Series = [series, new ChartSeries("B", "C2:C7") { Type = ChartType.Line }] }, "D2:J20"));
        Assert.Throws<ArgumentException>(() => sheet.AddChart(new Chart(ChartType.Line) { Series = [new ChartSeries("A", "B2:B7") { SecondaryAxis = true }] }, "D2:J20"));
        Assert.Throws<ArgumentException>(() => sheet.AddChart(new Chart(ChartType.Column) { Stacked = true, Series = [series], DataLabels = new DataLabels { Position = DataLabelPosition.OutsideEnd } }, "D2:J20"));
        Assert.Throws<ArgumentException>(() => sheet.AddChart(new Chart(ChartType.Doughnut) { Series = [series], DataLabels = new DataLabels { Position = DataLabelPosition.Center } }, "D2:J20"));
        Assert.Throws<ArgumentException>(() => sheet.AddChart(new Chart(ChartType.Line) { Series = [series], DataLabels = new DataLabels { Position = DataLabelPosition.InsideEnd } }, "D2:J20"));
        Assert.Throws<ArgumentException>(() => sheet.AddChart(new Chart(ChartType.Column) { Series = [series], ValueAxis = new ChartAxis { Minimum = 10, Maximum = 10 } }, "D2:J20"));
        Assert.Throws<ArgumentOutOfRangeException>(() => new Chart(ChartType.Doughnut) { HoleSize = 95 });
        Assert.Throws<ArgumentOutOfRangeException>(() => new Chart(ChartType.Pie) { FirstSliceAngle = 361 });
        Assert.Throws<ArgumentOutOfRangeException>(() => new Chart(ChartType.Column) { GapWidth = 501 });
        Assert.Throws<ArgumentOutOfRangeException>(() => new Chart(ChartType.Column) { Overlap = -101 });
        Assert.Throws<ArgumentOutOfRangeException>(() => new ChartSeries("A", "B2:B7") { MarkerSize = 1 });
        Assert.Throws<ArgumentException>(() => new ChartSeries("A", "B2:B7") { PointColors = ["green"] });
        Assert.Throws<ArgumentOutOfRangeException>(() => new ChartAxis { MajorUnit = 0 });
        Assert.Throws<ArgumentOutOfRangeException>(() => new ChartAxis { Minimum = double.PositiveInfinity });
    }

    [Fact]
    public async Task The_view_is_written_by_both_writers_and_read_back()
    {
        SheetView view = new() { ShowGridLines = false, ShowHeaders = false, Zoom = 85, TabColor = "#1f4e79", ActiveCell = new CellAddress(4, 2), RightToLeft = true };

        await using MemoryStream streamed = await TestFiles.WriteAsync(async writer =>
        {
            SheetWriter sheet = await TestFiles.AddSheetAsync(writer, "View");
            sheet.SetView(view);
            sheet.FreezePanes(1);
            await sheet.WriteRowAsync(["x"]);
            Assert.Throws<InvalidOperationException>(() => sheet.SetView(SheetView.Default));
        });

        Workbook book = Workbook.Create();
        Worksheet modelled = Added(book, "View");
        modelled.View = view;
        modelled.FreezePanes(1);
        await using MemoryStream saved = await SaveAsync(book);

        foreach (MemoryStream file in (MemoryStream[])[streamed, saved])
        {
            AssertValid(file);
            string sheetXml = Spaced(TestFiles.ReadPart(file, "xl/worksheets/sheet1.xml"));
            Assert.Contains("<sheetPr><tabColor rgb=\"FF1F4E79\" /></sheetPr><sheetViews><sheetView tabSelected=\"1\" showGridLines=\"0\" showRowColHeaders=\"0\" rightToLeft=\"1\" zoomScale=\"85\" zoomScaleNormal=\"85\" workbookViewId=\"0\">", sheetXml, StringComparison.Ordinal);
            Assert.Contains("<selection pane=\"bottomLeft\" activeCell=\"C5\" sqref=\"C5\" />", sheetXml, StringComparison.Ordinal);

            Workbook opened = await OpenAsync(file);
            Assert.Equal(view, opened.Sheets[0].View);
            Assert.Equal(1, opened.Sheets[0].FrozenRows);
        }
    }

    [Fact]
    public async Task Changing_the_view_of_an_opened_file_keeps_what_the_model_does_not_know()
    {
        const string main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        await using MemoryStream original = TestFiles.Package(
            ("[Content_Types].xml", """<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>"""),
            ("_rels/.rels", """<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>"""),
            ("xl/workbook.xml", $"""<workbook xmlns="{main}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>"""),
            ("xl/_rels/workbook.xml.rels", """<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"""),
            ("xl/worksheets/sheet1.xml", $"""<worksheet xmlns="{main}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetPr codeName="Blad1"><tabColor theme="4"/><pageSetUpPr fitToPage="1"/></sheetPr><sheetViews><sheetView showGridLines="0" zoomScale="120" zoomScaleNormal="120" view="pageLayout" workbookViewId="0"><selection activeCell="B2" sqref="B2"/></sheetView></sheetViews><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData><hyperlinks><hyperlink ref="A1" location="'Data'!B5"/><hyperlink ref="A2" location="'Data'!B6"/></hyperlinks><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>"""));

        Workbook book = await OpenAsync(original);
        Worksheet sheet = book.Sheets[0];

        // A theme colour is no RGB, so it reads as none, and is kept until the tab colour is set.
        Assert.Equal(new SheetView { ShowGridLines = false, Zoom = 120, ActiveCell = new CellAddress(1, 1) }, sheet.View);
        sheet.View = sheet.View with { ShowHeaders = false };

        Assert.True(sheet.AddHyperlink("A2", Hyperlink.ToUrl("https://example.com/")).IsSuccess(out _, out _));
        Assert.True(sheet.AddHyperlink("A3", Hyperlink.ToCell("Data", "C1")).IsSuccess(out _, out _));
        await using MemoryStream saved = await SaveAsync(book);

        AssertValid(saved);
        string xml = Spaced(TestFiles.ReadPart(saved, "xl/worksheets/sheet1.xml"));
        Assert.Contains("<sheetPr codeName=\"Blad1\"><tabColor theme=\"4\" /><pageSetUpPr fitToPage=\"1\" /></sheetPr>", xml, StringComparison.Ordinal);
        Assert.Contains("<sheetView showGridLines=\"0\" showRowColHeaders=\"0\" zoomScale=\"120\" zoomScaleNormal=\"120\" view=\"pageLayout\" workbookViewId=\"0\"><selection activeCell=\"B2\" sqref=\"B2\" /></sheetView>", xml, StringComparison.Ordinal);

        // The file's link on A1 stays; the one on A2 is replaced; A3 is new, and only the web link needs a relationship.
        Assert.Contains("<hyperlinks><hyperlink ref=\"A1\" location=\"'Data'!B5\" /><hyperlink ref=\"A2\" r:id=\"rId1\" /><hyperlink ref=\"A3\" location=\"'Data'!C1\" /></hyperlinks>", xml, StringComparison.Ordinal);
        Assert.Contains("Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink\" Target=\"https://example.com/\" TargetMode=\"External\"", TestFiles.ReadPart(saved, "xl/worksheets/_rels/sheet1.xml.rels"), StringComparison.Ordinal);

        Workbook again = await OpenAsync(saved);
        again.Sheets[0].View = again.Sheets[0].View with { TabColor = "#FF0000" };

        await using MemoryStream recoloured = await SaveAsync(again);
        Assert.Contains("<sheetPr codeName=\"Blad1\"><tabColor rgb=\"FFFF0000\" /><pageSetUpPr fitToPage=\"1\" /></sheetPr>", Spaced(TestFiles.ReadPart(recoloured, "xl/worksheets/sheet1.xml")), StringComparison.Ordinal);
    }

    [Fact]
    public async Task Cells_link_inside_the_workbook_and_out_from_both_writers()
    {
        await using MemoryStream streamed = await TestFiles.WriteAsync(async writer =>
        {
            SheetWriter sheet = await TestFiles.AddSheetAsync(writer, "Links");
            sheet.AddHyperlink("A1", Hyperlink.ToUrl("https://example.com/"));
            sheet.AddHyperlink(new CellAddress(1, 0), Hyperlink.ToCell("Other sheet", "B2") with { Tooltip = "Other" });
            sheet.AddHyperlink("A3", Hyperlink.ToUrl("mailto:someone@example.com"));
            sheet.AddHyperlink("A3", Hyperlink.ToUrl("https://example.org/"));
            await sheet.WriteRowAsync(["Web"]);
            await sheet.WriteRowAsync(["Sheet"]);
            await sheet.WriteRowAsync(["Replaced"]);
            await TestFiles.AddSheetAsync(writer, "Other sheet");
        });

        AssertValid(streamed);
        string xml = Spaced(TestFiles.ReadPart(streamed, "xl/worksheets/sheet1.xml"));
        Assert.Contains("<hyperlinks><hyperlink ref=\"A1\" r:id=\"rId2\" /><hyperlink ref=\"A2\" location=\"'Other sheet'!B2\" tooltip=\"Other\" /><hyperlink ref=\"A3\" r:id=\"rId3\" /></hyperlinks>", xml, StringComparison.Ordinal);
        string relationships = TestFiles.ReadPart(streamed, "xl/worksheets/_rels/sheet1.xml.rels");
        Assert.Contains("Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink\" Target=\"https://example.org/\" TargetMode=\"External\"", relationships, StringComparison.Ordinal);
        Assert.DoesNotContain("mailto", relationships, StringComparison.Ordinal);

        using SpreadsheetDocument document = SpreadsheetDocument.Open(streamed, false);
        WorksheetPart part = document.WorkbookPart!.WorksheetParts.First();
        Assert.Equal(2, part.HyperlinkRelationships.Count());
        Assert.Equal(3, part.Worksheet!.Descendants<Excel.Hyperlink>().Count());
    }

    [Fact]
    public async Task The_dashboard_from_both_writers_passes_schema_validation_and_survives_a_round_trip()
    {
        await using MemoryStream document = await SaveAsync(DashboardWorkbook.Build());
        await using MemoryStream streamed = await TestFiles.WriteAsync(DashboardWorkbook.WriteAsync);

        foreach (MemoryStream file in (MemoryStream[])[document, streamed])
        {
            AssertValid(file);

            using (SpreadsheetDocument opened = SpreadsheetDocument.Open(file, false))
            {
                WorkbookPart workbook = opened.WorkbookPart!;
                Excel.Sheet sheet = workbook.Workbook!.Descendants<Excel.Sheet>().Single(candidate => candidate.Name == DashboardWorkbook.Dashboard);
                DrawingsPart drawing = ((WorksheetPart)workbook.GetPartById(sheet.Id!)).DrawingsPart!;
                Assert.Equal(4, drawing.ChartParts.Count());
                Assert.Equal(6, drawing.ImageParts.Count());
                Assert.Equal(5, drawing.HyperlinkRelationships.Count());
            }

            file.Position = 0;
        }

        // Opened and saved again, unchanged: every part is still there, the drawings byte for byte.
        Workbook book = await OpenAsync(document);
        Assert.False(book.Sheets[0].View.ShowGridLines);
        await using MemoryStream again = await SaveAsync(book);
        AssertValid(again);
        Dictionary<string, uint> before = Parts(document);
        Dictionary<string, uint> after = Parts(again);
        Assert.Equal(before.Keys.Order(StringComparer.Ordinal), after.Keys.Order(StringComparer.Ordinal));

        foreach (string name in before.Keys.Where(name => name.StartsWith("xl/drawings/", StringComparison.Ordinal) || name.StartsWith("xl/charts/", StringComparison.Ordinal) || name.StartsWith("xl/media/", StringComparison.Ordinal)))
        {
            Assert.True(before[name] == after[name], name + " changed in the round trip.");
        }
    }

    [Fact]
    public async Task Shapes_added_to_an_opened_drawing_go_after_its_own_with_new_ids_and_links()
    {
        await using MemoryStream original = await SaveAsync(DashboardWorkbook.Build());
        Workbook book = await OpenAsync(original);
        Worksheet dashboard = book.Sheets[0];
        Assert.True(dashboard.AddShape(new Shape(ShapeKind.Chevron) { Link = Hyperlink.ToCell("Tasks", "A1") }, DrawingBounds.FromPoints(400, 600, 60, 30)).IsSuccess(out _, out _));
        Assert.True(dashboard.AddChart(new Chart(ChartType.Pie) { Series = [new ChartSeries("Share", "Data!J2:J3")], FirstSliceAngle = 90 }, DrawingBounds.FromPoints(500, 600, 100, 100)).IsSuccess(out _, out _));
        await using MemoryStream saved = await SaveAsync(book);

        AssertValid(saved);
        string drawing = TestFiles.ReadPart(saved, "xl/drawings/drawing1.xml");
        List<string> ids = Regex.Matches(drawing, "<xdr:cNvPr id=\"(\\d+)\"").Select(match => match.Groups[1].Value).ToList();
        Assert.Equal(ids.Count, ids.Distinct().Count());
        string relationships = TestFiles.ReadPart(saved, "xl/drawings/_rels/drawing1.xml.rels");
        List<string> relationshipIds = Regex.Matches(relationships, "Id=\"(rId\\d+)\"").Select(match => match.Groups[1].Value).ToList();
        Assert.Equal(relationshipIds.Count, relationshipIds.Distinct().Count());
        Assert.Contains("Target=\"#'Tasks'!A1\"", relationships, StringComparison.Ordinal);
        Assert.Equal(5, Regex.Count(relationships, "relationships/chart\""));
    }

    [Fact]
    public async Task A_Strict_OOXML_file_gets_no_shapes_or_links()
    {
        const string strictMain = "http://purl.oclc.org/ooxml/spreadsheetml/main";
        const string strictRelationships = "http://purl.oclc.org/ooxml/officeDocument/relationships";
        const string packageRelationships = "http://schemas.openxmlformats.org/package/2006/relationships";

        await using MemoryStream file = TestFiles.Package(
            ("_rels/.rels", $"""<Relationships xmlns="{packageRelationships}"><Relationship Id="rId1" Type="{strictRelationships}/officeDocument" Target="xl/workbook.xml"/></Relationships>"""),
            ("xl/workbook.xml", $"""<workbook xmlns="{strictMain}" xmlns:r="{strictRelationships}"><sheets><sheet name="Strikt" sheetId="1" r:id="rId1"/></sheets></workbook>"""),
            ("xl/_rels/workbook.xml.rels", $"""<Relationships xmlns="{packageRelationships}"><Relationship Id="rId1" Type="{strictRelationships}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"""),
            ("xl/worksheets/sheet1.xml", $"""<worksheet xmlns="{strictMain}"><sheetData/></worksheet>"""));

        Workbook book = await OpenAsync(file);
        Worksheet sheet = book.Sheets[0];

        foreach (Result result in (Result[])[sheet.AddShape(new Shape(ShapeKind.Rectangle), "A1:B2"), sheet.AddGroup(new ShapeGroup([GroupItem.Of(new Shape(ShapeKind.Rectangle), default)]), "A1:B2"), sheet.AddHyperlink("A1", Hyperlink.ToUrl("https://example.com/")), sheet.AddPicture(DrawingTests.PngPicture(), DrawingBounds.FromPoints(0, 0, 10, 10))])
        {
            Assert.True(result.IsFailure(out _, out ResultStatus status));
            Assert.Equal(ResultStatus.Invalid, status);
        }
    }

    /// <summary>One shape of every kind, each with something set, in a column of places.</summary>
    private static List<(Shape Shape, DrawingBounds Bounds)> EveryShape()
    {
        List<(Shape, DrawingBounds)> shapes = [];
        int row = 0;

        foreach (ShapeKind kind in Enum.GetValues<ShapeKind>())
        {
            Shape shape = new(kind)
            {
                Fill = Fill.Solid("#2384F9"),
                Outline = new Outline("#1F2940", 1),
                Text = new ShapeText(kind.ToString()),
                Shadow = new Shadow(),
            };

            if (kind == ShapeKind.Line)
            {
                shape = new Shape(kind) { Outline = new Outline("#1F2940", 2) { Dash = LineDash.Dash } };
            }

            shapes.Add((shape, DrawingBounds.FromPoints(20, 10 + row * 40, 120, 30)));
            row++;
        }

        return shapes;
    }

    private static Picture SvgPicture()
    {
        Assert.True(Picture.FromSvg(Encoding.UTF8.GetBytes(Svg)).IsSuccess(out Picture? picture));

        return picture!;
    }

    /// <summary>The drawing part of a sheet built by the streaming writer.</summary>
    private static async Task<string> DrawingOfAsync(Action<SheetWriter> add)
    {
        await using MemoryStream file = await TestFiles.WriteAsync(async writer =>
        {
            SheetWriter sheet = await TestFiles.AddSheetAsync(writer, "Data");
            add(sheet);
        });

        AssertValid(file);

        return TestFiles.ReadPart(file, "xl/drawings/drawing1.xml");
    }

    /// <summary>The chart part of a chart over four columns of numbers, written by the streaming writer and checked against the schema.</summary>
    private static async Task<string> ChartOfAsync(Chart chart)
    {
        await using MemoryStream file = await TestFiles.WriteAsync(async writer =>
        {
            SheetWriter sheet = await TestFiles.AddSheetAsync(writer, "Data");
            await sheet.WriteRowAsync(["Month", "Done", "Plan", "Share"]);

            for (int row = 0; row < 6; row++)
            {
                await sheet.WriteRowAsync(["M" + row.ToString(System.Globalization.CultureInfo.InvariantCulture), 10 + row, 12 + row, 0.1 * row]);
            }

            Assert.True(sheet.AddChart(chart, "F2:M18").IsSuccess(out _, out _));
        });

        AssertValid(file);

        return TestFiles.ReadPart(file, "xl/charts/chart1.xml");
    }

    /// <summary>The XML with a space before every "/&gt;", whichever writer wrote it, so that expected strings can be written one way.</summary>
    private static string Spaced(string xml)
    {
        return Regex.Replace(xml, "(?<! )/>", " />");
    }

    private static Dictionary<string, uint> Parts(MemoryStream file)
    {
        file.Position = 0;
        using ZipArchive archive = new(file, ZipArchiveMode.Read, true);
        Dictionary<string, uint> parts = archive.Entries.ToDictionary(entry => entry.FullName, entry => entry.Crc32, StringComparer.Ordinal);
        file.Position = 0;

        return parts;
    }

    private static async Task<Workbook> OpenAsync(Stream file)
    {
        file.Position = 0;
        Result<Workbook> opened = await Workbook.OpenAsync(file);
        Assert.True(opened.IsSuccess(out Workbook? book), opened.Error);

        return book;
    }

    private static async Task<MemoryStream> SaveAsync(Workbook book)
    {
        MemoryStream saved = new();
        await book.SaveAsync(saved);
        saved.Position = 0;

        return saved;
    }

    private static Worksheet Added(Workbook book, string name)
    {
        Result<Worksheet> added = book.AddSheet(name);
        Assert.True(added.IsSuccess(out Worksheet? sheet), added.Error);

        return sheet;
    }

    private static void AssertValid(MemoryStream file)
    {
        List<string> problems = TestFiles.Validate(file);
        Assert.True(problems.Count == 0, string.Join(Environment.NewLine, problems));
        file.Position = 0;
    }
}
