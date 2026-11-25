namespace Avig.Kalkyl.Samples;

/// <summary>One workbook of the gallery and the part of it the README shows.</summary>
/// <param name="Name">The file name without extension, also the picture's name.</param>
/// <param name="Title">What it is, in a few words.</param>
/// <param name="Sheet">The sheet the picture is taken of.</param>
/// <param name="Range">The cells the picture covers.</param>
/// <param name="WriteAsync">Writes the workbook to a path.</param>
public sealed record GallerySample(string Name, string Title, string Sheet, string Range, Func<string, Task> WriteAsync);

/// <summary>The workbooks of the gallery, in the order the README shows them.</summary>
public static class Gallery
{
    /// <summary>Every sample.</summary>
    public static IReadOnlyList<GallerySample> All { get; } =
    [
        new("project-dashboard", "Project dashboard", ProjectDashboard.Dashboard, ProjectDashboard.PictureRange, path => SaveAsync(ProjectDashboard.Build(), path)),
        new("sales-dashboard", "Sales dashboard on pivot tables", SalesDashboard.Dashboard, SalesDashboard.PictureRange, path => SaveAsync(SalesDashboard.Build(), path)),
        new("profit-and-loss", "Profit and loss statement", ProfitAndLoss.Report, ProfitAndLoss.PictureRange, path => SaveAsync(ProfitAndLoss.Build(), path)),
        new("budget-vs-actual", "Budget against actual", BudgetVsActual.Report, BudgetVsActual.PictureRange, path => SaveAsync(BudgetVsActual.Build(), path)),
        new("invoice", "Invoice", Invoice.Sheet, Invoice.PictureRange, path => SaveAsync(Invoice.Build(), path)),
        new("project-plan", "Project plan with a Gantt chart", ProjectPlan.Plan, ProjectPlan.PictureRange, path => SaveAsync(ProjectPlan.Build(), path)),
        new("scorecard", "Customer and staff scorecard", Scorecard.Sheet, Scorecard.PictureRange, path => SaveAsync(Scorecard.Build(), path)),
        new("large-export", "100 000 orders from the streaming writer", LargeExport.Sheet, LargeExport.PictureRange, path => LargeExport.WriteAsync(path, LargeExport.DefaultRows)),
    ];

    private static async Task SaveAsync(Workbook book, string path)
    {
        SampleKit.Check(await book.SaveAsync(path));
    }
}
