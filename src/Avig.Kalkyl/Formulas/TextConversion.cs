using System.Globalization;

namespace Avig.Kalkyl.Formulas;

/// <summary>
/// What it takes to read text as a number, a date or a criterion the way Excel does: the date
/// system of the workbook and the culture whose separators and date order apply.
/// </summary>
/// <param name="Uses1904">True for the 1904 date system.</param>
/// <param name="Culture">The culture text is read in; Excel uses the calculating machine's.</param>
internal readonly record struct TextConversion(bool Uses1904, CultureInfo Culture);
