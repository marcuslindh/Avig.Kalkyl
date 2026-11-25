namespace Avig.Kalkyl;

/// <summary>
/// The Excel side of <see cref="Hyperlink"/>, which lives in Avig.OpenXml.Drawing with the other
/// DrawingML types: a link to a cell in the workbook, stored as the location <c>'Sheet'!A1</c>.
/// Put one on a shape, a group, a picture or a cell.
/// </summary>
public static class HyperlinkExtensions
{
    extension(Hyperlink)
    {
        /// <summary>A link to a cell in the workbook. The sheet does not have to exist yet.</summary>
        /// <param name="sheet">The sheet's name.</param>
        /// <param name="cell">The cell.</param>
        /// <returns>The link.</returns>
        public static Hyperlink ToCell(string sheet, CellAddress cell)
        {
            ArgumentNullException.ThrowIfNull(sheet);

            if (XlsxWriter.ValidateName(sheet).IsFailure(out string? error, out ResultStatus _) == true)
            {
                throw new ArgumentException(error, nameof(sheet));
            }

            return Hyperlink.ToLocation("'" + sheet.Replace("'", "''", StringComparison.Ordinal) + "'!" + cell.ToString());
        }

        /// <summary>A link to a cell in the workbook, such as <c>Hyperlink.ToCell("Sales", "A1")</c>.</summary>
        /// <param name="sheet">The sheet's name.</param>
        /// <param name="cell">The cell in A1 notation.</param>
        /// <returns>The link.</returns>
        public static Hyperlink ToCell(string sheet, string cell = "A1")
        {
            ArgumentNullException.ThrowIfNull(cell);

            if (CellAddress.TryParse(cell, out CellAddress address) == false)
            {
                throw new ArgumentException($"'{cell}' is not a cell address. Expected a reference such as A1.", nameof(cell));
            }

            return Hyperlink.ToCell(sheet, address);
        }
    }

    extension(Hyperlink link)
    {
        /// <summary>The sheet a link inside the workbook goes to, or null for a web address.</summary>
        public string? Sheet
        {
            get
            {
                string? sheet = null;

                if (TrySplit(link.Location, out string? name, out _) == true)
                {
                    sheet = name;
                }

                return sheet;
            }
        }

        /// <summary>The cell a link inside the workbook goes to.</summary>
        public CellAddress Cell
        {
            get
            {
                TrySplit(link.Location, out _, out CellAddress cell);

                return cell;
            }
        }
    }

    /// <summary>Splits a location as <see cref="ToCell(string, CellAddress)"/> writes it into the sheet and the cell.</summary>
    private static bool TrySplit(string? location, [System.Diagnostics.CodeAnalysis.NotNullWhen(true)] out string? sheet, out CellAddress cell)
    {
        sheet = null;
        cell = default;

        if (location is null)
        {
            return false;
        }

        int bang = location.LastIndexOf('!');

        if (bang < 0 || CellAddress.TryParse(location[(bang + 1)..], out cell) == false)
        {
            return false;
        }

        string name = location[..bang];

        if (name.Length >= 2 && name.StartsWith('\'') == true && name.EndsWith('\'') == true)
        {
            name = name[1..^1].Replace("''", "'", StringComparison.Ordinal);
        }

        sheet = name;

        return true;
    }
}
