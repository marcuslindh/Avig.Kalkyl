using Avig.Kalkyl.Reading;

namespace Avig.Kalkyl.Document;

/// <summary>
/// The sheet list a save writes, in tab order. The model holds only worksheets, but the list in
/// the file also has chart sheets, macro sheets and dialog sheets, and defined names and the
/// active tab count positions in that whole list (localSheetId, activeTab). So the sheets of the
/// opened file are written where they were, the model's worksheets in their places and the
/// others as they were, and new worksheets go last. A sheet whose relationship is missing, or
/// that <see cref="Workbook.RemoveActiveContent"/> takes out, is left out, and positions after it
/// move up.
/// </summary>
internal sealed class SheetList
{
    private readonly List<SheetSlot> slots;
    private readonly int[] writtenIndexOfOriginal;

    private SheetList(List<SheetSlot> slots, int[] writtenIndexOfOriginal, List<string> droppedNames)
    {
        this.slots = slots;
        this.writtenIndexOfOriginal = writtenIndexOfOriginal;
        DroppedNames = droppedNames;
    }

    /// <summary>The sheets in the order they are written.</summary>
    public IReadOnlyList<SheetSlot> Slots => slots;

    /// <summary>The names of the opened file's sheets that are left out.</summary>
    public IReadOnlyList<string> DroppedNames { get; }

    /// <summary>True when a sheet of the opened file is left out, so that positions move.</summary>
    public bool HasDropped => DroppedNames.Count > 0;

    /// <summary>Works out the list for a save; <paramref name="removedRelationships"/> are the workbook's relationships that are left out.</summary>
    public static SheetList Create(Workbook book, IReadOnlySet<string> removedRelationships)
    {
        List<SheetSlot> slots = [];
        IReadOnlyList<DeclaredSheet> declared = book.DeclaredSheets;
        int[] map = new int[declared.Count];
        List<string> dropped = [];

        for (int index = 0; index < declared.Count; index++)
        {
            DeclaredSheet sheet = declared[index];
            Worksheet? model = book.Sheets.FirstOrDefault(candidate => candidate.Source is not null && string.Equals(candidate.Source.RelationshipId, sheet.RelationshipId, StringComparison.Ordinal) == true);

            if (model is not null)
            {
                map[index] = slots.Count;
                slots.Add(new SheetSlot(model, null));

                continue;
            }

            if (sheet.WorksheetIndex < 0 && sheet.Kind.Length > 0 && removedRelationships.Contains(sheet.RelationshipId) == false)
            {
                map[index] = slots.Count;
                slots.Add(new SheetSlot(null, sheet));

                continue;
            }

            map[index] = -1;
            dropped.Add(sheet.Name);
        }

        foreach (Worksheet sheet in book.Sheets.Where(sheet => sheet.Source is null))
        {
            slots.Add(new SheetSlot(sheet, null));
        }

        return new SheetList(slots, map, dropped);
    }

    /// <summary>The position a worksheet of the model is written at.</summary>
    public int WrittenIndexOf(Worksheet sheet)
    {
        return slots.FindIndex(slot => ReferenceEquals(slot.Model, sheet) == true);
    }

    /// <summary>
    /// The position a sheet of the opened file, given by its localSheetId, is written at. False
    /// when it is left out; a position the file did not have stays as it was.
    /// </summary>
    public bool TryMapOriginal(int localSheetId, out int written)
    {
        written = localSheetId;

        if (localSheetId < 0 || localSheetId >= writtenIndexOfOriginal.Length)
        {
            return true;
        }

        written = writtenIndexOfOriginal[localSheetId];

        return written >= 0;
    }

    /// <summary>The position for an activeTab or firstSheet: the sheet's own, or the next one kept when it is left out.</summary>
    public int MapNearest(int localSheetId)
    {
        for (int index = Math.Max(0, localSheetId); index < writtenIndexOfOriginal.Length; index++)
        {
            if (writtenIndexOfOriginal[index] >= 0)
            {
                return writtenIndexOfOriginal[index];
            }
        }

        return Math.Max(0, Math.Min(localSheetId, slots.Count - 1));
    }
}

/// <summary>A sheet in the list: a worksheet of the model, or a sheet of the opened file the model does not hold.</summary>
internal sealed record SheetSlot(Worksheet? Model, DeclaredSheet? Original);
