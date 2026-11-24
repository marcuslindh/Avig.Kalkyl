using System.Globalization;
using System.Xml;
using Avig.Kalkyl.Writing;

namespace Avig.Kalkyl.Internal;

/// <summary>Writes &lt;conditionalFormatting&gt; elements. Shared by the streaming writer and the document model.</summary>
internal static class ConditionalFormatXml
{
    /// <summary>
    /// Writes one element per rule. <paramref name="firstPriority"/> is one more than the highest
    /// priority already on the sheet: Excel wants every rule on a sheet to have its own, and a
    /// lower number wins, so rules added later lose to the ones the file had.
    /// </summary>
    public static async Task WriteAsync(XmlWriter xml, IReadOnlyList<(CellRange Range, ConditionalFormat Format)> rules, StyleRegistry styles, int firstPriority)
    {
        int priority = firstPriority;

        foreach ((CellRange range, ConditionalFormat format) in rules)
        {
            await xml.WriteStartElementAsync(null, "conditionalFormatting", null);
            await xml.WriteAttributeStringAsync(null, "sqref", null, range.ToString());
            await WriteRuleAsync(xml, range, format, styles, priority);
            await xml.WriteEndElementAsync();
            priority++;
        }
    }

    private static async Task WriteRuleAsync(XmlWriter xml, CellRange range, ConditionalFormat format, StyleRegistry styles, int priority)
    {
        await xml.WriteStartElementAsync(null, "cfRule", null);
        await xml.WriteAttributeStringAsync(null, "type", null, TypeName(format));

        if (format.Style is not null)
        {
            await xml.WriteAttributeStringAsync(null, "dxfId", null, Text(styles.GetDifferentialIndex(format.Style)));
        }

        await xml.WriteAttributeStringAsync(null, "priority", null, Text(priority));

        switch (format.Kind)
        {
            case ConditionalFormatKind.CellIs:
                await xml.WriteAttributeStringAsync(null, "operator", null, OperatorName(format));
                await WriteFormulasAsync(xml, format.Arguments);

                break;

            case ConditionalFormatKind.Expression:
                await WriteFormulasAsync(xml, format.Arguments);

                break;

            case ConditionalFormatKind.ContainsText:
                string text = format.Text!;
                string quoted = "\"" + text.Replace("\"", "\"\"", StringComparison.Ordinal) + "\"";
                await xml.WriteAttributeStringAsync(null, "operator", null, "containsText");
                await xml.WriteAttributeStringAsync(null, "text", null, text);

                // Excel evaluates the formula, not the text attribute, so it must be the formula
                // Excel would write: relative to the range's top-left cell.
                await WriteFormulasAsync(xml, ["NOT(ISERROR(SEARCH(" + quoted + "," + range.First.ToString() + ")))"]);

                break;

            case ConditionalFormatKind.Top:
                await xml.WriteAttributeStringAsync(null, "rank", null, Text(format.Rank));

                if (format.Percent == true)
                {
                    await xml.WriteAttributeStringAsync(null, "percent", null, "1");
                }

                if (format.FromBottom == true)
                {
                    await xml.WriteAttributeStringAsync(null, "bottom", null, "1");
                }

                break;

            case ConditionalFormatKind.AboveAverage when format.FromBottom == true:
                await xml.WriteAttributeStringAsync(null, "aboveAverage", null, "0");

                break;

            case ConditionalFormatKind.ColorScale:
                await xml.WriteStartElementAsync(null, "colorScale", null);
                await WriteValueAsync(xml, "min", null);

                if (format.Colors.Count == 3)
                {
                    await WriteValueAsync(xml, "percentile", "50");
                }

                await WriteValueAsync(xml, "max", null);

                foreach (string color in format.Colors)
                {
                    await WriteColorAsync(xml, color);
                }

                await xml.WriteEndElementAsync();

                break;

            case ConditionalFormatKind.DataBar:
                await xml.WriteStartElementAsync(null, "dataBar", null);
                await WriteValueAsync(xml, "min", null);
                await WriteValueAsync(xml, "max", null);
                await WriteColorAsync(xml, format.Colors[0]);
                await xml.WriteEndElementAsync();

                break;

            case ConditionalFormatKind.IconSet:
                (string name, int count) = IconSetName(format.Icons);
                await xml.WriteStartElementAsync(null, "iconSet", null);
                await xml.WriteAttributeStringAsync(null, "iconSet", null, name);

                if (format.ShowValue == false)
                {
                    await xml.WriteAttributeStringAsync(null, "showValue", null, "0");
                }

                if (format.Reverse == true)
                {
                    await xml.WriteAttributeStringAsync(null, "reverse", null, "1");
                }

                // Equal parts of the spread, as Excel sets them by default: 0, 33, 67 for three icons.
                for (int icon = 0; icon < count; icon++)
                {
                    int percent = (int)Math.Round(100.0 * icon / count);
                    await WriteValueAsync(xml, "percent", Text(percent));
                }

                await xml.WriteEndElementAsync();

                break;
        }

        await xml.WriteEndElementAsync();
    }

    private static async Task WriteFormulasAsync(XmlWriter xml, IReadOnlyList<string> formulas)
    {
        foreach (string formula in formulas)
        {
            await xml.WriteElementStringAsync(null, "formula", null, formula);
        }
    }

    private static async Task WriteValueAsync(XmlWriter xml, string type, string? value)
    {
        await xml.WriteStartElementAsync(null, "cfvo", null);
        await xml.WriteAttributeStringAsync(null, "type", null, type);

        if (value is not null)
        {
            await xml.WriteAttributeStringAsync(null, "val", null, value);
        }

        await xml.WriteEndElementAsync();
    }

    private static async Task WriteColorAsync(XmlWriter xml, string color)
    {
        await xml.WriteStartElementAsync(null, "color", null);
        await xml.WriteAttributeStringAsync(null, "rgb", null, "FF" + color[1..]);
        await xml.WriteEndElementAsync();
    }

    private static string TypeName(ConditionalFormat format)
    {
        return format.Kind switch
        {
            ConditionalFormatKind.CellIs => "cellIs",
            ConditionalFormatKind.Expression => "expression",
            ConditionalFormatKind.ContainsText => "containsText",
            ConditionalFormatKind.Top => "top10",
            ConditionalFormatKind.AboveAverage => "aboveAverage",
            ConditionalFormatKind.Duplicates => "duplicateValues",
            ConditionalFormatKind.Unique => "uniqueValues",
            ConditionalFormatKind.ColorScale => "colorScale",
            ConditionalFormatKind.DataBar => "dataBar",
            _ => "iconSet",
        };
    }

    private static string OperatorName(ConditionalFormat format)
    {
        if (format.Arguments.Count == 2)
        {
            if (format.Negated == true)
            {
                return "notBetween";
            }

            return "between";
        }

        return format.Operator switch
        {
            ComparisonOperator.NotEqual => "notEqual",
            ComparisonOperator.GreaterThan => "greaterThan",
            ComparisonOperator.GreaterThanOrEqual => "greaterThanOrEqual",
            ComparisonOperator.LessThan => "lessThan",
            ComparisonOperator.LessThanOrEqual => "lessThanOrEqual",
            _ => "equal",
        };
    }

    private static (string Name, int Count) IconSetName(IconSetKind icons)
    {
        return icons switch
        {
            IconSetKind.ThreeArrows => ("3Arrows", 3),
            IconSetKind.ThreeArrowsGray => ("3ArrowsGray", 3),
            IconSetKind.ThreeFlags => ("3Flags", 3),
            IconSetKind.ThreeTrafficLightsRimmed => ("3TrafficLights2", 3),
            IconSetKind.ThreeSigns => ("3Signs", 3),
            IconSetKind.ThreeSymbols => ("3Symbols", 3),
            IconSetKind.ThreeSymbolsUncircled => ("3Symbols2", 3),
            IconSetKind.FourArrows => ("4Arrows", 4),
            IconSetKind.FourArrowsGray => ("4ArrowsGray", 4),
            IconSetKind.FourRedToBlack => ("4RedToBlack", 4),
            IconSetKind.FourRating => ("4Rating", 4),
            IconSetKind.FourTrafficLights => ("4TrafficLights", 4),
            IconSetKind.FiveArrows => ("5Arrows", 5),
            IconSetKind.FiveArrowsGray => ("5ArrowsGray", 5),
            IconSetKind.FiveRating => ("5Rating", 5),
            IconSetKind.FiveQuarters => ("5Quarters", 5),
            _ => ("3TrafficLights1", 3),
        };
    }

    private static string Text(int number)
    {
        return number.ToString(CultureInfo.InvariantCulture);
    }
}
