using System.Text;
using System.Text.RegularExpressions;

namespace Avig.Kalkyl.Fuzz;

/// <summary>
/// <c>attribute-sweep</c> — every number attribute in the template workbook (style and font
/// indexes, row and column numbers, shared-string and shared-formula indexes, counts, sizes,
/// sheet ids) gets every edge value in turn, one at a time. Random mutation reaches a given
/// attribute with a given value too rarely; this visits each pair once, which is about 2 000
/// inputs (run with --iter 3000). The harness is the one of <c>xlsx-part</c>.
/// </summary>
internal static partial class AttributeSweepTarget
{
    private static readonly string[] EdgeValues =
    [
        "0", "-1", "1", "65535", "65536", "1048576", "1048577", "16384", "16385", "2147483647", "2147483648",
        "-2147483649", "99999999999999999999", "1e9", "0.5", "", " 1 ", "x",
    ];

    /// <summary>Every (part, attribute occurrence) in the template whose value is a number.</summary>
    private static readonly (int Part, int Start, int Length)[] Places = FindPlaces();

    private static int position;

    public static FuzzTarget Create()
    {
        return new FuzzTarget("attribute-sweep", "every number attribute of the template with every edge value", PackageTarget.RunAsync, Generate);
    }

    [GeneratedRegex("\\b[A-Za-z:]+=\"(-?[0-9][0-9.E+]*)\"")]
    private static partial Regex NumberAttribute();

    private static (int Part, int Start, int Length)[] FindPlaces()
    {
        List<(int, int, int)> places = [];

        for (int part = 0; part < PackageTarget.Parts.Length; part++)
        {
            foreach (Match match in NumberAttribute().Matches(PackageTarget.Parts[part].Content))
            {
                places.Add((part, match.Groups[1].Index, match.Groups[1].Length));
            }
        }

        return [.. places];
    }

    private static byte[] Generate(Random random)
    {
        int index = position;
        position++;
        (int part, int start, int length) = Places[(index / EdgeValues.Length) % Places.Length];
        string value = EdgeValues[index % EdgeValues.Length];
        string content = PackageTarget.Parts[part].Content;
        string changed = string.Concat(content.AsSpan(0, start), value, content.AsSpan(start + length));
        byte[] text = Encoding.UTF8.GetBytes(changed);
        byte[] input = new byte[text.Length + 1];
        input[0] = (byte)part;
        text.CopyTo(input, 1);

        return input;
    }
}
