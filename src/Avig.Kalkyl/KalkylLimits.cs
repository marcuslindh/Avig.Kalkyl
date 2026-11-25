namespace Avig.Kalkyl;

/// <summary>
/// Limits on what a file may make the package do, so that a file from outside cannot run the
/// process out of memory, and how strictly the zip container is checked (<see cref="StrictZip"/>). An .xlsx is a zip file, and a few kilobytes of it can unpack to
/// gigabytes. A file that goes past a limit gives Invalid when it is opened. A formula that goes
/// past one gives #NUM!. Time is limited with the CancellationToken that every call takes, for
/// example <c>new CancellationTokenSource(TimeSpan.FromSeconds(10)).Token</c>.
/// </summary>
/// <example>
/// <code>
/// Result&lt;Workbook&gt; opened = await Workbook.OpenAsync(upload, KalkylLimits.Default with { MaxCells = 1_000_000 });
/// </code>
/// </example>
public sealed record KalkylLimits
{
    /// <summary>
    /// The limits used when none are given: well above any workbook people make by hand, and
    /// well below what takes a server down.
    /// </summary>
    public static KalkylLimits Default { get; } = new();

    /// <summary>No limits beyond Excel's own. Only for files you trust.</summary>
    public static KalkylLimits Unlimited { get; } = new()
    {
        MaxUncompressedBytes = long.MaxValue,
        MaxCells = long.MaxValue,
        MaxSharedStrings = long.MaxValue,
        MaxArrayCells = long.MaxValue,
    };

    /// <summary>
    /// How many shared strings the file may have. 10 million by default. They are all read into
    /// memory when the file is opened, by the streaming reader too, so 1 GiB of short strings
    /// could otherwise take several gigabytes.
    /// </summary>
    public long MaxSharedStrings { get; init; } = 10_000_000;

    /// <summary>
    /// How many bytes the parts of the file may unpack to, together. 1 GiB by default. Each part
    /// is also held to the size its zip header gives, so a header cannot lie its way past this.
    /// </summary>
    public long MaxUncompressedBytes { get; init; } = 1L << 30;

    /// <summary>
    /// How many cells the document model may hold when a file is opened. 20 million by
    /// default, which is a few gigabytes of memory. The streaming reader keeps one row at a
    /// time and is not affected.
    /// </summary>
    public long MaxCells { get; init; } = 20_000_000;

    /// <summary>
    /// How many values one array in a formula may hold, for example the result of
    /// A1:A100000*TRANSPOSE(A1:A100000). 10 million by default. A larger one gives #NUM!.
    /// </summary>
    public long MaxArrayCells { get; init; } = 10_000_000;

    /// <summary>
    /// Whether the zip file must say one thing only: every entry's local header as the zip
    /// directory has it, no two entries sharing bytes, and no bytes before, between or after the
    /// entries. True by default, also in <see cref="Unlimited"/>. Excel and every library that
    /// writes .xlsx files leave no such room, and a file that has it can show a virus scanner one
    /// workbook and Excel another. A file that fails is damaged. Turn it off only to read a file
    /// you trust that some tool has appended to.
    /// </summary>
    public bool StrictZip { get; init; } = true;
}
