namespace Avig.Kalkyl;

/// <summary>Whether a sheet's tab is shown.</summary>
public enum SheetVisibility
{
    /// <summary>The tab is shown.</summary>
    Visible,

    /// <summary>The tab is hidden but can be shown from Excel's menu.</summary>
    Hidden,

    /// <summary>The tab is hidden and can only be shown with VBA.</summary>
    VeryHidden,
}

/// <summary>A worksheet in a workbook. Chart sheets are not listed.</summary>
/// <param name="Name">The name on the tab.</param>
/// <param name="Index">The position among the worksheets, 0-based.</param>
/// <param name="Visibility">Whether the tab is shown.</param>
public sealed record SheetInfo(string Name, int Index, SheetVisibility Visibility);
