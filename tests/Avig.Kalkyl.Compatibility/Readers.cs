using ClosedXML.Excel;
using MiniExcelLibs;
using EdrReader = ExcelDataReader.IExcelDataReader;
using EdrFactory = ExcelDataReader.ExcelReaderFactory;
using SylvanReader = Sylvan.Data.Excel.ExcelDataReader;

namespace Avig.Kalkyl.Compatibility;

/// <summary>Reads a workbook with each library and brings the answers to <see cref="Value"/>.</summary>
internal static class Readers
{
    /// <summary>Rows read per sheet. The largest tutorial workbooks have several hundred thousand.</summary>
    public const int MaxRows = 20_000;

    static Readers()
    {
        // ExcelDataReader decodes some parts with Windows-1252, which .NET only has once the code pages are registered.
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
    }

    public static async Task<WorkbookValues> KalkylStreamingAsync(string path)
    {
        WorkbookValues values = new("Avig.Kalkyl (XlsxReader)");
        Result<XlsxReader> opened = await XlsxReader.OpenAsync(path);

        if (opened.IsSuccess(out XlsxReader? reader) == false)
        {
            values.Failure = opened.Error ?? "failed";

            return values;
        }

        await using (reader)
        {
            for (int index = 0; index < reader.Sheets.Count; index++)
            {
                string name = reader.Sheets[index].Name;
                values.AddSheet(name);

                if (reader.Sheet(index).IsSuccess(out SheetReader? sheet) == false)
                {
                    continue;
                }

                await foreach (Row row in sheet.ReadRowsAsync())
                {
                    if (row.Index >= MaxRows)
                    {
                        break;
                    }

                    for (int column = 0; column < row.ColumnCount; column++)
                    {
                        AddKalkyl(values, name, row.Index, column, row[column].Value);
                    }
                }
            }
        }

        return values;
    }

    public static WorkbookValues KalkylDocument(Workbook book)
    {
        WorkbookValues values = new("Avig.Kalkyl (Workbook)");

        foreach (Worksheet sheet in book.Sheets)
        {
            values.AddSheet(sheet.Name);

            foreach (Row row in sheet.Rows)
            {
                if (row.Index >= MaxRows)
                {
                    break;
                }

                for (int column = 0; column < row.ColumnCount; column++)
                {
                    AddKalkyl(values, sheet.Name, row.Index, column, row[column].Value);
                }
            }
        }

        return values;
    }

    public static void AddKalkyl(WorkbookValues values, string sheet, int row, int column, CellValue value)
    {
        if (value.TryGetNumber(out double number) == true)
        {
            values.Add(sheet, row, column, Value.Of(number));
        }
        else if (value.TryGetText(out string? text) == true)
        {
            values.Add(sheet, row, column, Value.Of(text));
        }
        else if (value.TryGetBool(out bool flag) == true)
        {
            values.Add(sheet, row, column, Value.Of(flag));
        }
        else if (value.TryGetDate(out DateTime date) == true)
        {
            values.Add(sheet, row, column, Value.Of(date));
        }
        else if (value.TryGetError(out string? code) == true)
        {
            values.Add(sheet, row, column, Value.Error(code));
        }
    }

    public static WorkbookValues WithExcelDataReader(string path)
    {
        WorkbookValues values = new("ExcelDataReader");

        try
        {
            using FileStream stream = File.OpenRead(path);
            using EdrReader reader = EdrFactory.CreateReader(stream);

            do
            {
                string name = reader.Name;
                values.AddSheet(name);
                int row = 0;

                while (reader.Read() == true && row < MaxRows)
                {
                    for (int column = 0; column < reader.FieldCount; column++)
                    {
                        ExcelDataReader.CellError? error = reader.GetCellError(column);

                        if (error is not null)
                        {
                            values.Add(name, row, column, Value.Error(ErrorCode(error.Value)));

                            continue;
                        }

                        AddObject(values, name, row, column, reader.GetValue(column));
                    }

                    row++;
                }
            }
            while (reader.NextResult() == true);
        }
        catch (Exception exception) when (exception is not OutOfMemoryException)
        {
            values.Failure = exception.GetType().Name + ": " + exception.Message;
        }

        return values;
    }

    public static WorkbookValues WithClosedXml(string path)
    {
        WorkbookValues values = new("ClosedXML");

        try
        {
            using XLWorkbook book = new(path);

            foreach (IXLWorksheet sheet in book.Worksheets)
            {
                values.AddSheet(sheet.Name);

                foreach (IXLCell cell in sheet.CellsUsed(XLCellsUsedOptions.Contents))
                {
                    int row = cell.Address.RowNumber - 1;
                    int column = cell.Address.ColumnNumber - 1;
                    XLCellValue value = cell.CachedValue;

                    switch (value.Type)
                    {
                        case XLDataType.Number:
                            values.Add(sheet.Name, row, column, Value.Of(value.GetNumber()));
                            break;
                        case XLDataType.Text:
                            // ClosedXML turns \n in an inline string into \r\n (the file holds \n,
                            // and the other three libraries read \n), so it is turned back here.
                            values.Add(sheet.Name, row, column, Value.Of(value.GetText().Replace("\r\n", "\n", StringComparison.Ordinal)));
                            break;
                        case XLDataType.Boolean:
                            values.Add(sheet.Name, row, column, Value.Of(value.GetBoolean()));
                            break;
                        case XLDataType.DateTime:
                            values.Add(sheet.Name, row, column, Value.Of(value.GetDateTime()));
                            break;
                        case XLDataType.TimeSpan:
                            // A time without a date; Excel stores it as a fraction of day zero.
                            values.Add(sheet.Name, row, column, Value.Of(new DateTime(1899, 12, 31, 0, 0, 0, DateTimeKind.Unspecified).Add(value.GetTimeSpan())));
                            break;
                        case XLDataType.Error:
                            values.Add(sheet.Name, row, column, Value.Error(ErrorCode(value.GetError())));
                            break;
                        default:
                            break;
                    }
                }
            }
        }
        catch (Exception exception) when (exception is not OutOfMemoryException)
        {
            values.Failure = exception.GetType().Name + ": " + exception.Message;
        }

        return values;
    }

    public static WorkbookValues WithSylvan(string path)
    {
        WorkbookValues values = new("Sylvan.Data.Excel");

        try
        {
            Sylvan.Data.Excel.ExcelDataReaderOptions options = new()
            {
                Schema = Sylvan.Data.Excel.ExcelSchema.NoHeaders,
                ReadHiddenWorksheets = true,
                ReadHiddenRows = true,
            };

            using SylvanReader reader = SylvanReader.Create(path, options);

            do
            {
                string name = reader.WorksheetName ?? string.Empty;
                values.AddSheet(name);

                while (reader.Read() == true)
                {
                    int row = reader.RowNumber - 1;

                    if (row >= MaxRows)
                    {
                        break;
                    }

                    for (int column = 0; column < reader.RowFieldCount; column++)
                    {
                        switch (reader.GetExcelDataType(column))
                        {
                            // Sylvan types a cell by how it is stored, so a date is Numeric; the
                            // number format says it is a date.
                            case Sylvan.Data.Excel.ExcelDataType.Numeric when IsDateFormat(reader.GetFormat(column)) == true:
                                values.Add(name, row, column, Value.Of(SylvanDate(reader, column)));
                                break;
                            case Sylvan.Data.Excel.ExcelDataType.Numeric:
                                values.Add(name, row, column, Value.Of(reader.GetDouble(column)));
                                break;
                            case Sylvan.Data.Excel.ExcelDataType.DateTime:
                                values.Add(name, row, column, Value.Of(reader.GetDateTime(column)));
                                break;
                            case Sylvan.Data.Excel.ExcelDataType.String:
                                values.Add(name, row, column, Value.Of(reader.GetString(column)));
                                break;
                            case Sylvan.Data.Excel.ExcelDataType.Boolean:
                                values.Add(name, row, column, Value.Of(reader.GetBoolean(column)));
                                break;
                            case Sylvan.Data.Excel.ExcelDataType.Error:
                                values.Add(name, row, column, Value.Error(ErrorCode(reader.GetFormulaError(column))));
                                break;
                            default:
                                break;
                        }
                    }
                }
            }
            while (reader.NextResult() == true);
        }
        catch (Exception exception) when (exception is not OutOfMemoryException)
        {
            values.Failure = exception.GetType().Name + ": " + exception.Message;
        }

        return values;
    }

    public static WorkbookValues WithMiniExcel(string path)
    {
        WorkbookValues values = new("MiniExcel");

        try
        {
            foreach (string name in MiniExcelLibs.MiniExcel.GetSheetNames(path))
            {
                values.AddSheet(name);
                int row = 0;

                foreach (object item in MiniExcelLibs.MiniExcel.Query(path, false, name, ExcelType.XLSX, "A1", null))
                {
                    if (row >= MaxRows)
                    {
                        break;
                    }

                    if (item is IDictionary<string, object?> cells)
                    {
                        foreach (KeyValuePair<string, object?> cell in cells)
                        {
                            if (CellAddress.TryParseColumn(cell.Key, out int column) == true)
                            {
                                AddObject(values, name, row, column, cell.Value);
                            }
                        }
                    }

                    row++;
                }
            }
        }
        catch (Exception exception) when (exception is not OutOfMemoryException)
        {
            values.Failure = exception.GetType().Name + ": " + exception.Message;
        }

        return values;
    }

    /// <summary>Sylvan's GetDateTime refuses serial numbers below 1, which a time without a date is; those are counted from day zero here.</summary>
    private static DateTime SylvanDate(SylvanReader reader, int column)
    {
        double serial = reader.GetDouble(column);

        if (serial < 1)
        {
            return new DateTime(1899, 12, 31, 0, 0, 0, DateTimeKind.Unspecified).AddMilliseconds(Math.Round(serial * 86_400_000));
        }

        return reader.GetDateTime(column);
    }

    private static bool IsDateFormat(Sylvan.Data.Excel.ExcelFormat? format)
    {
        return format is not null && format.Kind is Sylvan.Data.Excel.FormatKind.Date or Sylvan.Data.Excel.FormatKind.Time;
    }

    private static void AddObject(WorkbookValues values, string sheet, int row, int column, object? value)
    {
        switch (value)
        {
            case null:
                break;
            case double number:
                values.Add(sheet, row, column, Value.Of(number));
                break;
            case decimal number:
                values.Add(sheet, row, column, Value.Of((double)number));
                break;
            case int number:
                values.Add(sheet, row, column, Value.Of(number));
                break;
            case long number:
                values.Add(sheet, row, column, Value.Of(number));
                break;
            case bool flag:
                values.Add(sheet, row, column, Value.Of(flag));
                break;
            case DateTime date:
                values.Add(sheet, row, column, Value.Of(date));
                break;
            case TimeSpan time:
                values.Add(sheet, row, column, Value.Of(new DateTime(1899, 12, 31, 0, 0, 0, DateTimeKind.Unspecified).Add(time)));
                break;
            case string text:
                values.Add(sheet, row, column, Value.Of(text));
                break;
            default:
                values.Add(sheet, row, column, Value.Of(Convert.ToString(value, System.Globalization.CultureInfo.InvariantCulture) ?? string.Empty));
                break;
        }
    }

    private static string ErrorCode(ExcelDataReader.CellError error)
    {
        return error switch
        {
            ExcelDataReader.CellError.NULL => "#NULL!",
            ExcelDataReader.CellError.DIV0 => "#DIV/0!",
            ExcelDataReader.CellError.VALUE => "#VALUE!",
            ExcelDataReader.CellError.REF => "#REF!",
            ExcelDataReader.CellError.NAME => "#NAME?",
            ExcelDataReader.CellError.NUM => "#NUM!",
            ExcelDataReader.CellError.NA => "#N/A",
            _ => "#" + error,
        };
    }

    private static string ErrorCode(XLError error)
    {
        return error switch
        {
            XLError.NullValue => "#NULL!",
            XLError.DivisionByZero => "#DIV/0!",
            XLError.IncompatibleValue => "#VALUE!",
            XLError.CellReference => "#REF!",
            XLError.NameNotRecognized => "#NAME?",
            XLError.NumberInvalid => "#NUM!",
            XLError.NoValueAvailable => "#N/A",
            _ => "#" + error,
        };
    }

    private static string ErrorCode(Sylvan.Data.Excel.ExcelErrorCode error)
    {
        return error switch
        {
            Sylvan.Data.Excel.ExcelErrorCode.Null => "#NULL!",
            Sylvan.Data.Excel.ExcelErrorCode.DivideByZero => "#DIV/0!",
            Sylvan.Data.Excel.ExcelErrorCode.Value => "#VALUE!",
            Sylvan.Data.Excel.ExcelErrorCode.Reference => "#REF!",
            Sylvan.Data.Excel.ExcelErrorCode.Name => "#NAME?",
            Sylvan.Data.Excel.ExcelErrorCode.Number => "#NUM!",
            Sylvan.Data.Excel.ExcelErrorCode.NotAvailable => "#N/A",
            _ => "#" + error,
        };
    }
}
