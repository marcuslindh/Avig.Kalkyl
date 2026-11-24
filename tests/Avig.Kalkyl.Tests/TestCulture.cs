using System.Globalization;
using System.Runtime.CompilerServices;

namespace Avig.Kalkyl.Tests;

/// <summary>
/// Workbook.Culture starts as the machine's culture, as in Excel, so the tests fix it: they
/// are written for en-US ("1,000" is a thousand) and must give the same answer on every
/// machine. Tests about other cultures set Workbook.Culture themselves.
/// </summary>
internal static class TestCulture
{
#pragma warning disable CA2255 // A test assembly, not a library: the culture must be set before any test runs.
    [ModuleInitializer]
    internal static void Initialize()
    {
        CultureInfo english = CultureInfo.GetCultureInfo("en-US");
        CultureInfo.DefaultThreadCurrentCulture = english;
        CultureInfo.CurrentCulture = english;
    }
#pragma warning restore CA2255
}
