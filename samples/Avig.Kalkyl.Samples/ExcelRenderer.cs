using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;

namespace Avig.Kalkyl.Samples;

/// <summary>One picture to take: a block of cells on a sheet of a saved workbook.</summary>
/// <param name="Workbook">The .xlsx file.</param>
/// <param name="Sheet">The sheet to take the picture of.</param>
/// <param name="Range">The cells, in Excel's notation, with every drawing over them.</param>
/// <param name="Png">Where the picture goes.</param>
public sealed record RenderJob(string Workbook, string Sheet, string Range, string Png);

/// <summary>
/// Takes pictures of workbooks as Excel 365 draws them, through COM with late binding, so the
/// project needs no interop assembly. Each block of cells is copied as a bitmap
/// (<c>Range.CopyPicture</c>), pasted into a temporary chart of the same size and exported as
/// PNG. Excel runs hidden with alerts off and macros disabled, on an STA thread of its own; if
/// it has not finished in time, the Excel processes it started are killed, since Excel can hang
/// on a prompt nobody sees.
/// </summary>
[SupportedOSPlatform("windows")]
public static class ExcelRenderer
{
    private const int ScreenAppearance = 1;   // xlScreen
    private const int Bitmap = 2;             // xlBitmap
    private const int ForceDisableMacros = 3; // msoAutomationSecurityForceDisable

    /// <summary>Renders every job and returns what went wrong, one line per problem.</summary>
    public static IReadOnlyList<string> Render(IReadOnlyList<RenderJob> jobs, TimeSpan timeout)
    {
        List<string> problems = [];

        HashSet<int> before = ExcelProcesses();
        Thread thread = new(() => RenderAll(jobs, problems));
        thread.SetApartmentState(ApartmentState.STA);
        thread.IsBackground = true;
        thread.Start();

        if (thread.Join(timeout) == false)
        {
            foreach (Process process in Process.GetProcessesByName("EXCEL"))
            {
                if (before.Contains(process.Id) == false)
                {
                    process.Kill();
                }
            }

            lock (problems)
            {
                problems.Add($"Excel did not finish within {timeout.TotalSeconds} seconds and was stopped.");
            }
        }

        return problems;
    }

    private static HashSet<int> ExcelProcesses()
    {
        HashSet<int> ids = [];

        foreach (Process process in Process.GetProcessesByName("EXCEL"))
        {
            ids.Add(process.Id);
        }

        return ids;
    }

    private static void RenderAll(IReadOnlyList<RenderJob> jobs, List<string> problems)
    {
        Type? type = Type.GetTypeFromProgID("Excel.Application");

        if (type is null)
        {
            lock (problems)
            {
                problems.Add("Excel is not installed.");
            }

            return;
        }

        MessageFilter.Register();
        dynamic excel = Activator.CreateInstance(type) ?? throw new InvalidOperationException("Excel did not start.");

        try
        {
            excel.DisplayAlerts = false;
            excel.Visible = false;
            excel.ScreenUpdating = true;
            excel.AutomationSecurity = ForceDisableMacros;

            foreach (RenderJob job in jobs)
            {
                try
                {
                    RenderOne(excel, job);
                    Console.WriteLine($"rendered  {Path.GetFileName(job.Png)}  {new FileInfo(job.Png).Length / 1024} KB");
                }
                catch (Exception exception) when (exception is not OutOfMemoryException)
                {
                    lock (problems)
                    {
                        problems.Add($"{Path.GetFileName(job.Workbook)}: {exception.Message}");
                    }
                }
            }
        }
        finally
        {
            excel.Quit();
            Marshal.FinalReleaseComObject(excel);
            MessageFilter.Revoke();
        }
    }

    private static void RenderOne(dynamic excel, RenderJob job)
    {
        dynamic book = excel.Workbooks.Open(Path.GetFullPath(job.Workbook), 0, true);

        try
        {
            string caption = excel.ActiveWindow.Caption;

            if (caption.Contains("Repaired", StringComparison.OrdinalIgnoreCase) == true)
            {
                throw new InvalidOperationException("Excel repaired the file.");
            }

            dynamic sheet = book.Worksheets[job.Sheet];
            sheet.Activate();
            excel.ActiveWindow.Zoom = 100;
            dynamic range = sheet.Range[job.Range];
            double width = range.Width;
            double height = range.Height;
            Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(job.Png)) ?? ".");

            // The clipboard and a hidden Excel sometimes give an empty picture the first time.
            for (int attempt = 0; attempt < 4; attempt++)
            {
                range.CopyPicture(ScreenAppearance, Bitmap);
                dynamic holder = sheet.ChartObjects().Add(0, 0, width, height);

                try
                {
                    holder.Chart.ChartArea.Format.Line.Visible = 0;
                    holder.Activate();
                    holder.Chart.Paste();
                    holder.Chart.Export(Path.GetFullPath(job.Png), "PNG");
                }
                finally
                {
                    holder.Delete();
                }

                if (new FileInfo(job.Png).Length > 4096)
                {
                    PngOptimizer.Optimize(job.Png);

                    return;
                }

                Thread.Sleep(500);
            }

            throw new InvalidOperationException("Excel gave an empty picture.");
        }
        finally
        {
            book.Close(false);
        }
    }
}
