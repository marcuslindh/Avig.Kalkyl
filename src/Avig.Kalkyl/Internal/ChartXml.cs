using System.Xml;

namespace Avig.Kalkyl.Internal;

/// <summary>
/// Writes a chart part. The part itself (c:chartSpace) is DrawingML and is written by
/// Avig.OpenXml's <see cref="ChartWriter"/>; what is Excel's is the series' references, which
/// <see cref="Chart.ToChartSpace(string)"/> has already turned into formulas on the sheet.
/// </summary>
internal static class ChartXml
{
    /// <summary>A complete chart part.</summary>
    public static Task WriteChartAsync(XmlWriter xml, DrawingItem item)
    {
        return ChartWriter.WriteChartSpaceAsync(xml, item.ChartSpace!);
    }
}
