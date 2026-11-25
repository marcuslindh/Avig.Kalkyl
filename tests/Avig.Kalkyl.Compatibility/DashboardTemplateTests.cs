using System.IO.Compression;

namespace Avig.Kalkyl.Compatibility;

/// <summary>
/// Opens a dashboard template that Excel saved (dark cards, grouped navigation buttons with SVG
/// icons and links, a hundred charts, pivot tables and slicers), saves it unchanged and with
/// drawings added, and checks that every part is still there and the drawings, charts and media
/// are byte for byte. The template is a third-party file and is not committed: set
/// KALKYL_DASHBOARD_TEMPLATE to its path, or the tests are skipped. Open the saved files in Excel
/// with tools/Avig.Kalkyl.ExcelCheck/open-in-excel.ps1; KALKYL_DASHBOARD_OUTPUT keeps them.
/// </summary>
public class DashboardTemplateTests
{
    private const string Variable = "KALKYL_DASHBOARD_TEMPLATE";

    [SkippableFact]
    public async Task An_unchanged_template_keeps_every_part()
    {
        string template = Template();
        Result<Workbook> opened = await Workbook.OpenAsync(template);
        Assert.True(opened.IsSuccess(out Workbook? book), opened.Error);

        byte[] saved = await SaveAsync(book, "dashboard-template-saved.xlsx");

        Dictionary<string, uint> before = Parts(await File.ReadAllBytesAsync(template));
        Dictionary<string, uint> after = Parts(saved);

        // calcChain is dropped on purpose (Excel rebuilds it); nothing else may go.
        Assert.DoesNotContain(before.Keys.Except(after.Keys), name => name.EndsWith("calcChain.xml", StringComparison.Ordinal) == false);
        Assert.Empty(after.Keys.Except(before.Keys));
        AssertDrawingsUnchanged(before, after, []);
    }

    [SkippableFact]
    public async Task Drawings_added_to_a_template_go_after_its_own()
    {
        string template = Template();
        Result<Workbook> opened = await Workbook.OpenAsync(template);
        Assert.True(opened.IsSuccess(out Workbook? book), opened.Error);
        Worksheet sheet = book.Sheets[0];
        string data = book.Sheets[^1].Name;
        sheet.View = sheet.View with { Zoom = 80 };

        Assert.True(sheet.AddShape(new Shape(ShapeKind.RoundedRectangle) { Fill = Fill.Gradient("#2384F9", "#605776"), Shadow = new Shadow(), Text = new ShapeText("Added"), Link = Hyperlink.ToCell(data, "A1") }, DrawingBounds.FromPoints(10, 600, 200, 50)).IsSuccess(out _, out _));
        Assert.True(sheet.AddChart(new Chart(ChartType.Doughnut) { Series = [new ChartSeries("Added", "'" + data.Replace("'", "''", StringComparison.Ordinal) + "'!B2:B4")], HoleSize = 70, Background = Fill.None }, DrawingBounds.FromPoints(220, 600, 150, 150)).IsSuccess(out _, out _));
        Assert.True(sheet.AddHyperlink(new CellAddress(200, 1), Hyperlink.ToUrl("https://example.com/")).IsSuccess(out _, out _));

        byte[] saved = await SaveAsync(book, "dashboard-template-added.xlsx");

        Dictionary<string, uint> before = Parts(await File.ReadAllBytesAsync(template));
        Dictionary<string, uint> after = Parts(saved);
        Assert.DoesNotContain(before.Keys.Except(after.Keys), name => name.EndsWith("calcChain.xml", StringComparison.Ordinal) == false);
        Assert.Single(after.Keys.Except(before.Keys));

        // The first sheet's drawing and its .rels have the additions; every other drawing part is as it was.
        AssertDrawingsUnchanged(before, after, ["xl/drawings/drawing1.xml", "xl/drawings/_rels/drawing1.xml.rels"]);
    }

    private static void AssertDrawingsUnchanged(Dictionary<string, uint> before, Dictionary<string, uint> after, HashSet<string> changed)
    {
        foreach (string name in before.Keys)
        {
            bool drawing = name.StartsWith("xl/drawings/", StringComparison.Ordinal) || name.StartsWith("xl/charts/", StringComparison.Ordinal)
                || name.StartsWith("xl/media/", StringComparison.Ordinal) || name.StartsWith("xl/slicers/", StringComparison.Ordinal)
                || name.StartsWith("xl/pivot", StringComparison.Ordinal);

            if (drawing == true && changed.Contains(name) == false)
            {
                Assert.True(before[name] == after[name], name + " changed.");
            }
        }
    }

    private static string Template()
    {
        string? path = Environment.GetEnvironmentVariable(Variable);
        Skip.If(string.IsNullOrWhiteSpace(path) == true || File.Exists(path) == false, $"Set {Variable} to a dashboard template saved by Excel to run this test.");

        return path!;
    }

    private static async Task<byte[]> SaveAsync(Workbook book, string name)
    {
        using MemoryStream stream = new();
        await book.SaveAsync(stream);
        byte[] bytes = stream.ToArray();
        string? output = Environment.GetEnvironmentVariable("KALKYL_DASHBOARD_OUTPUT");

        if (string.IsNullOrWhiteSpace(output) == false)
        {
            Directory.CreateDirectory(output);
            await File.WriteAllBytesAsync(Path.Combine(output, name), bytes);
        }

        return bytes;
    }

    private static Dictionary<string, uint> Parts(byte[] file)
    {
        using ZipArchive archive = new(new MemoryStream(file, false), ZipArchiveMode.Read);

        return archive.Entries.ToDictionary(entry => entry.FullName, entry => entry.Crc32, StringComparer.Ordinal);
    }
}
