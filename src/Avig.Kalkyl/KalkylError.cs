namespace Avig.Kalkyl;

/// <summary>
/// What kind of failure a result or exception from the package is, for an application to show
/// to its users. The error message is meant for the log: it can name parts, cells and sheets of
/// the file, and quote text from it. The kind says no more than what went wrong in general, so
/// it is safe to show, and an application can word it in its own language.
/// </summary>
public enum KalkylErrorKind
{
    /// <summary>Not a failure the package recognizes: a success, or an error that did not come from opening or reading a file.</summary>
    Unknown,

    /// <summary>The file or its folder does not exist, or the sheet that was asked for is not in the workbook.</summary>
    NotFound,

    /// <summary>The file is not an .xlsx workbook: an old .xls, a password-protected .xlsx, another kind of file, or not a zip file at all.</summary>
    NotXlsx,

    /// <summary>The file is an .xlsx, or claims to be one, but is damaged or holds something no spreadsheet program writes.</summary>
    Damaged,

    /// <summary>The file goes past one of the <see cref="KalkylLimits"/>. A trusted file can be opened with larger limits.</summary>
    TooLarge,

    /// <summary>The file exists but could not be opened or written: another program has it open, or the process may not use it.</summary>
    Inaccessible,
}

/// <summary>Finds the <see cref="KalkylErrorKind"/> of a failed result or of an exception thrown while a file is read.</summary>
/// <example>
/// <code>
/// Result&lt;Workbook&gt; opened = await Workbook.OpenAsync(upload);
///
/// if (opened.IsFailure(out string? error, out _) == true)
/// {
///     logger.LogWarning("Upload rejected: {Error}", error);
///
///     return opened.ErrorKind() switch
///     {
///         KalkylErrorKind.TooLarge => "The file is too large.",
///         KalkylErrorKind.NotXlsx => "Upload an .xlsx file.",
///         _ => "The file could not be read.",
///     };
/// }
/// </code>
/// </example>
public static class KalkylError
{
    /// <summary>How every message about a file that goes past a limit names the limit.</summary>
    internal const string LimitMarker = "the limit set in KalkylLimits.";

    /// <summary>How every message about a damaged file starts.</summary>
    internal const string DamagedPrefix = "The file is damaged";

    /// <summary>How the messages about a broken zip container or broken XML start; they mean a damaged file too.</summary>
    internal const string InvalidWorkbookPrefix = "The file is not a valid .xlsx workbook";

    /// <summary>How every message about a file that is not an .xlsx starts.</summary>
    internal const string NotXlsxPrefix = "The file is not an .xlsx workbook";

    /// <summary>The kind of a failed result; <see cref="KalkylErrorKind.Unknown"/> for a success.</summary>
    public static KalkylErrorKind ErrorKind(this Result result)
    {
        return Classify(result.Status, result.Error);
    }

    /// <summary>The kind of a failed result; <see cref="KalkylErrorKind.Unknown"/> for a success.</summary>
    public static KalkylErrorKind ErrorKind<T>(this Result<T> result)
    {
        return Classify(result.Status, result.Error);
    }

    /// <summary>
    /// The kind of an exception thrown while a file is read, as <see cref="XlsxReader"/> does
    /// with <see cref="InvalidDataException"/> for a damaged row. Any other exception, a
    /// cancellation among them, is <see cref="KalkylErrorKind.Unknown"/>.
    /// </summary>
    public static KalkylErrorKind ErrorKind(this Exception exception)
    {
        if (exception is FileNotFoundException or DirectoryNotFoundException)
        {
            return KalkylErrorKind.NotFound;
        }

        if (exception is InvalidDataException)
        {
            KalkylErrorKind kind = OfMessage(exception.Message);

            if (kind == KalkylErrorKind.Unknown)
            {
                return KalkylErrorKind.Damaged;
            }

            return kind;
        }

        return KalkylErrorKind.Unknown;
    }

    /// <summary>
    /// The message for a result made from an <see cref="InvalidDataException"/>: the package's own
    /// messages as they are, and one from .NET (a broken deflate stream) marked as a damaged file.
    /// </summary>
    internal static string Describe(InvalidDataException exception)
    {
        if (OfMessage(exception.Message) != KalkylErrorKind.Unknown)
        {
            return exception.Message;
        }

        return $"{DamagedPrefix}: {exception.Message}";
    }

    private static KalkylErrorKind Classify(ResultStatus status, string? error)
    {
        if (status == ResultStatus.Ok || error is null)
        {
            return KalkylErrorKind.Unknown;
        }

        if (status == ResultStatus.NotFound)
        {
            return KalkylErrorKind.NotFound;
        }

        if (status == ResultStatus.Conflict)
        {
            return KalkylErrorKind.Inaccessible;
        }

        return OfMessage(error);
    }

    private static KalkylErrorKind OfMessage(string message)
    {
        if (message.Contains(LimitMarker, StringComparison.Ordinal) == true)
        {
            return KalkylErrorKind.TooLarge;
        }

        if (message.StartsWith(DamagedPrefix, StringComparison.Ordinal) == true || message.StartsWith(InvalidWorkbookPrefix, StringComparison.Ordinal) == true)
        {
            return KalkylErrorKind.Damaged;
        }

        if (message.StartsWith(NotXlsxPrefix, StringComparison.Ordinal) == true)
        {
            return KalkylErrorKind.NotXlsx;
        }

        return KalkylErrorKind.Unknown;
    }
}
