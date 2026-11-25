namespace Avig.Kalkyl.Samples;

/// <summary>
/// Writes the gallery's workbooks and, with Excel on the machine, the pictures of them that the
/// README shows.
/// </summary>
/// <example>
/// dotnet run --project samples/Avig.Kalkyl.Samples                          (write and render)
/// dotnet run --project samples/Avig.Kalkyl.Samples -- --no-render           (only the .xlsx files)
/// dotnet run --project samples/Avig.Kalkyl.Samples -- --only invoice --out C:\temp\gallery
/// </example>
internal static class Program
{
    public static async Task<int> Main(string[] args)
    {
        string root = RepositoryRoot();
        string output = Path.GetFullPath(Option(args, "--out", Path.Combine(root, "artifacts", "gallery")));
        string images = Path.GetFullPath(Option(args, "--images", Path.Combine(root, "docs", "images", "gallery")));
        string only = Option(args, "--only", "");
        bool render = args.Contains("--no-render") == false;
        Directory.CreateDirectory(output);
        List<RenderJob> jobs = [];

        foreach (GallerySample sample in Gallery.All)
        {
            if (only.Length > 0 && sample.Name != only)
            {
                continue;
            }

            string path = Path.Combine(output, sample.Name + ".xlsx");
            await sample.WriteAsync(path);
            Console.WriteLine($"wrote     {path}");
            jobs.Add(new RenderJob(path, sample.Sheet, sample.Range, Path.Combine(images, sample.Name + ".png")));
        }

        if (render == false)
        {
            return 0;
        }

        IReadOnlyList<string> problems = ["The pictures need Excel on Windows; run with --no-render elsewhere."];

        if (OperatingSystem.IsWindows() == true)
        {
            // Guarded by the check above; the platform analyzer does not follow "== true".
#pragma warning disable CA1416
            problems = ExcelRenderer.Render(jobs, TimeSpan.FromMinutes(3));
#pragma warning restore CA1416
        }

        foreach (string problem in problems)
        {
            Console.Error.WriteLine(problem);
        }

        return problems.Count;
    }

    private static string Option(string[] args, string name, string fallback)
    {
        int index = Array.IndexOf(args, name);

        if (index >= 0 && index + 1 < args.Length)
        {
            return args[index + 1];
        }

        return fallback;
    }

    /// <summary>The folder with Avig.Kalkyl.slnx, found from the program's own folder.</summary>
    private static string RepositoryRoot()
    {
        DirectoryInfo? folder = new(AppContext.BaseDirectory);

        while (folder is not null)
        {
            if (File.Exists(Path.Combine(folder.FullName, "Avig.Kalkyl.slnx")) == true)
            {
                return folder.FullName;
            }

            folder = folder.Parent;
        }

        return Environment.CurrentDirectory;
    }
}
