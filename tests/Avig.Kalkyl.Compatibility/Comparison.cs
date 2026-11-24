using System.Globalization;
using System.Text;

namespace Avig.Kalkyl.Compatibility;

/// <summary>One cell where the library under test disagrees with the answer it is checked against.</summary>
internal sealed record Difference(string Sheet, int Row, int Column, string Actual, string Expected, string Detail)
{
    public string Address => new CellAddress(Row, Column).ToString();

    public override string ToString()
    {
        return $"{Sheet}!{Address}: {Actual}, expected {Expected}{Detail}";
    }
}

/// <summary>The outcome of comparing one reading with others.</summary>
internal sealed class ComparisonResult
{
    public int Compared { get; set; }

    /// <summary>Cells where the references did not agree among themselves, so there was nothing to check against.</summary>
    public int Unclear { get; set; }

    /// <summary>
    /// Dates before 1900-03-01 that differ by a day. Excel counts 1900 as a leap year, so
    /// serial 1 is 1900-01-01; libraries that use DateTime.FromOADate make it 1899-12-31.
    /// Avig.Kalkyl follows Excel, so this is a known difference, not a failure.
    /// </summary>
    public int EarlyDates { get; set; }

    /// <summary>Formulas with structured references (Table1[Column]), which the engine does not support.</summary>
    public int Unsupported { get; set; }

    public List<Difference> Differences { get; } = [];

    public List<string> Notes { get; } = [];
}

internal static class Comparison
{
    private const string Empty = "(empty)";

    /// <summary>
    /// Compares <paramref name="actual"/> with the majority of <paramref name="references"/>
    /// cell by cell. A cell counts only when at least two references agree and no other answer
    /// has as many votes. A reference that cannot read the file, or does not list a sheet
    /// (a chart sheet, say), takes no part for that file or sheet.
    /// </summary>
    public static ComparisonResult WithConsensus(WorkbookValues actual, IReadOnlyList<WorkbookValues> references)
    {
        ComparisonResult result = new();
        List<WorkbookValues> usable = [.. references.Where(reference => reference.Failure.Length == 0)];

        foreach (WorkbookValues reference in references.Where(reference => reference.Failure.Length > 0))
        {
            result.Notes.Add($"{reference.Library} could not read the file: {reference.Failure}");
        }

        foreach ((string sheet, Dictionary<(int Row, int Column), Value> cells) in actual.Sheets)
        {
            List<WorkbookValues> voters = [.. usable.Where(reference => reference.Sheets.ContainsKey(sheet))];

            if (voters.Count < 2)
            {
                result.Notes.Add($"Sheet '{sheet}': fewer than two other libraries read it.");

                continue;
            }

            HashSet<(int Row, int Column)> keys = [.. cells.Keys];

            foreach (WorkbookValues voter in voters)
            {
                keys.UnionWith(voter.Sheets[sheet].Keys);
            }

            foreach ((int row, int column) in keys)
            {
                CompareCell(result, sheet, row, column, actual, voters);
            }
        }

        return result;
    }

    private static void CompareCell(ComparisonResult result, string sheet, int row, int column, WorkbookValues actual, List<WorkbookValues> voters)
    {
        List<(Value? Value, List<string> Libraries)> groups = [];

        foreach (WorkbookValues voter in voters)
        {
            Value? vote = null;

            if (voter.TryGet(sheet, row, column, out Value value) == true)
            {
                vote = value;
            }

            int group = groups.FindIndex(existing => Same(existing.Value, vote));

            if (group < 0)
            {
                groups.Add((vote, [voter.Library]));
            }
            else
            {
                groups[group].Libraries.Add(voter.Library);
            }
        }

        List<(Value? Value, List<string> Libraries)> ranked = [.. groups.OrderByDescending(group => group.Libraries.Count)];
        bool clear = ranked[0].Libraries.Count >= 2 && (ranked.Count == 1 || ranked[1].Libraries.Count < ranked[0].Libraries.Count);

        if (clear == false)
        {
            result.Unclear++;

            return;
        }

        Value? mine = null;

        if (actual.TryGet(sheet, row, column, out Value own) == true)
        {
            mine = own;
        }

        result.Compared++;
        Value? expected = ranked[0].Value;

        if (Same(mine, expected) == true)
        {
            return;
        }

        if (IsEarlyDate(mine, expected) == true)
        {
            result.EarlyDates++;

            return;
        }

        StringBuilder detail = new();

        foreach ((Value? value, List<string> libraries) in groups)
        {
            detail.Append(CultureInfo.InvariantCulture, $"; {string.Join("+", libraries)}: {Describe(value)}");
        }

        result.Differences.Add(new Difference(sheet, row, column, Describe(mine), Describe(expected), detail.ToString()));
    }

    /// <summary>Compares one library's reading of two files cell by cell, such as the original and a copy saved by Avig.Kalkyl.</summary>
    public static ComparisonResult Exactly(WorkbookValues actual, WorkbookValues expected)
    {
        ComparisonResult result = new();

        if (actual.Failure.Length > 0)
        {
            result.Differences.Add(new Difference("(file)", 0, 0, actual.Failure, "a file it can read", string.Empty));

            return result;
        }

        foreach ((string sheet, Dictionary<(int Row, int Column), Value> cells) in expected.Sheets)
        {
            if (actual.Sheets.TryGetValue(sheet, out Dictionary<(int Row, int Column), Value>? actualCells) == false)
            {
                result.Differences.Add(new Difference(sheet, 0, 0, "the sheet is missing", "the sheet", string.Empty));

                continue;
            }

            HashSet<(int Row, int Column)> keys = [.. cells.Keys, .. actualCells.Keys];

            foreach ((int row, int column) in keys)
            {
                result.Compared++;
                Value? want = null;
                Value? have = null;

                if (cells.TryGetValue((row, column), out Value wanted) == true)
                {
                    want = wanted;
                }

                if (actualCells.TryGetValue((row, column), out Value had) == true)
                {
                    have = had;
                }

                if (IsEarlyDate(want, have) == true)
                {
                    result.EarlyDates++;
                }
                else if (Same(have, want) == false)
                {
                    result.Differences.Add(new Difference(sheet, row, column, Describe(have), Describe(want), string.Empty));
                }
            }
        }

        return result;
    }

    private static bool Same(Value? left, Value? right)
    {
        if (left is null || right is null)
        {
            return left is null && right is null;
        }

        return left.Value.Matches(right.Value);
    }

    private static bool IsEarlyDate(Value? mine, Value? expected)
    {
        if (mine is null || expected is null || mine.Value.Kind != ValueKind.Date || expected.Value.Kind != ValueKind.Date)
        {
            return false;
        }

        DateTime firstMarch = new(1900, 3, 1, 0, 0, 0, DateTimeKind.Unspecified);
        double days = Math.Abs((mine.Value.Date - expected.Value.Date).TotalDays);

        return mine.Value.Date < firstMarch && Math.Abs(days - 1) < 1e-6;
    }

    private static string Describe(Value? value)
    {
        if (value is null)
        {
            return Empty;
        }

        return $"{value.Value.Kind} {value.Value}";
    }

    /// <summary>A short summary for an assertion message, with the first differences.</summary>
    public static string Summary(string title, ComparisonResult result, int shown = 25)
    {
        StringBuilder builder = new();
        builder.AppendLine(CultureInfo.InvariantCulture, $"{title}: {result.Differences.Count} differences in {result.Compared} cells compared ({result.Unclear} unclear, {result.EarlyDates} early 1900 dates, {result.Unsupported} structured references).");

        foreach (string note in result.Notes)
        {
            builder.AppendLine("  note: " + note);
        }

        foreach (Difference difference in result.Differences.Take(shown))
        {
            builder.AppendLine("  " + difference);
        }

        if (result.Differences.Count > shown)
        {
            builder.AppendLine(CultureInfo.InvariantCulture, $"  ... and {result.Differences.Count - shown} more (see report/).");
        }

        return builder.ToString();
    }
}
