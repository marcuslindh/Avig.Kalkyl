using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;

namespace Avig.Kalkyl.Tests;

/// <summary>
/// Hostile and malformed files from the security test plan (docs/testplan.md), built in memory.
/// Each must give a Result failure or InvalidDataException, never another exception, and never
/// touch anything outside the file.
/// </summary>
public class HostileFilesTests
{
    private const string Main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private const string Relationships = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private const string PackageRelationships = "http://schemas.openxmlformats.org/package/2006/relationships";
    private const string OneCell = "<row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t>x</t></is></c></row>";

    private static string ContentTypes(string main = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml", string extra = "")
    {
        return $"""<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/><Override PartName="/xl/workbook.xml" ContentType="{main}"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>{extra}</Types>""";
    }

    private static string WorkbookRelationships(string extra = "", string sheetTarget = "worksheets/sheet1.xml")
    {
        return $"""<Relationships xmlns="{PackageRelationships}"><Relationship Id="rId1" Type="{Relationships}/worksheet" Target="{sheetTarget}"/>{extra}</Relationships>""";
    }

    private static string Sheet(string data, string after = "")
    {
        return $"""<worksheet xmlns="{Main}"><sheetData>{data}</sheetData>{after}</worksheet>""";
    }

    private static List<(string Path, string Content)> Parts(string sheet, string relationships = "", string workbook = "", string contentTypes = "")
    {
        if (relationships.Length == 0)
        {
            relationships = WorkbookRelationships();
        }

        if (workbook.Length == 0)
        {
            workbook = $"""<workbook xmlns="{Main}" xmlns:r="{Relationships}"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>""";
        }

        if (contentTypes.Length == 0)
        {
            contentTypes = ContentTypes();
        }

        return
        [
            ("[Content_Types].xml", contentTypes),
            ("_rels/.rels", $"""<Relationships xmlns="{PackageRelationships}"><Relationship Id="rId1" Type="{Relationships}/officeDocument" Target="xl/workbook.xml"/></Relationships>"""),
            ("xl/workbook.xml", workbook),
            ("xl/_rels/workbook.xml.rels", relationships),
            ("xl/worksheets/sheet1.xml", sheet),
        ];
    }

    private static MemoryStream File(List<(string Path, string Content)> parts)
    {
        return TestFiles.Package([.. parts]);
    }

    [Theory]
    [InlineData("inte XML alls")]
    [InlineData("<wsDr><a><a><a>")]
    [InlineData("deep")]
    public async Task A_damaged_drawing_that_a_picture_is_added_to_is_a_damaged_file(string drawing)
    {
        if (drawing == "deep")
        {
            drawing = "<wsDr>" + string.Concat(Enumerable.Repeat("<x>", 500)) + string.Concat(Enumerable.Repeat("</x>", 500)) + "</wsDr>";
        }

        List<(string Path, string Content)> parts = Parts(Sheet(OneCell, "<drawing xmlns:r=\"" + Relationships + "\" r:id=\"rId1\"/>"));
        parts.Add(("xl/worksheets/_rels/sheet1.xml.rels", $"""<Relationships xmlns="{PackageRelationships}"><Relationship Id="rId1" Type="{Relationships}/drawing" Target="../drawings/drawing1.xml"/></Relationships>"""));
        parts.Add(("xl/drawings/drawing1.xml", drawing));
        await using MemoryStream file = File(parts);

        // The drawing is not read when the file is opened, only copied when it is saved.
        Result<Workbook> opened = await Workbook.OpenAsync(file);
        Assert.True(opened.IsSuccess(out Workbook? book), opened.Error);
        Assert.True(book.Sheets[0].AddPicture(DrawingTests.PngPicture(), "B2").IsSuccess(out _, out _));

        await using MemoryStream saved = new();
        InvalidDataException damaged = await Assert.ThrowsAsync<InvalidDataException>(() => book.SaveAsync(saved));
        Assert.Equal(KalkylErrorKind.Damaged, damaged.ErrorKind());

        // Saved to a path, nothing is left behind: the file is built in memory first.
        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
        await Assert.ThrowsAsync<InvalidDataException>(() => book.SaveAsync(path));
        Assert.False(System.IO.File.Exists(path));
    }

    /// <summary>The kinds a refused hostile file may have: every refusal must be one an application can show.</summary>
    private static readonly KalkylErrorKind[] Refusals = [KalkylErrorKind.Damaged, KalkylErrorKind.TooLarge, KalkylErrorKind.NotXlsx];

    private static async Task<string> RefusedAsync(MemoryStream file)
    {
        Result<XlsxReader> reader = await XlsxReader.OpenAsync(file, true);
        string streaming = string.Empty;

        if (reader.IsSuccess(out XlsxReader? opened) == true)
        {
            await using (opened)
            {
                InvalidDataException damaged = await Assert.ThrowsAsync<InvalidDataException>(async () =>
                {
                    for (int index = 0; index < opened.Sheets.Count; index++)
                    {
                        if (opened.Sheet(index).IsSuccess(out SheetReader? sheet) == true)
                        {
                            await foreach (Row row in sheet.ReadRowsAsync())
                            {
                                _ = row.Cells.Count;
                            }
                        }
                    }
                });
                streaming = damaged.Message;
                Assert.Contains(damaged.ErrorKind(), Refusals);
            }
        }
        else
        {
            streaming = reader.Error!;
            Assert.Contains(reader.ErrorKind(), Refusals);
        }

        file.Position = 0;
        Result<Workbook> book = await Workbook.OpenAsync(file);
        Assert.True(book.IsFailure(out string? error, out ResultStatus status), "The document model opened the file.");
        Assert.Equal(ResultStatus.Invalid, status);
        Assert.True(Refusals.Contains(book.ErrorKind()), $"No error kind for: {error}");

        return streaming + " | " + error;
    }

    [Fact]
    public async Task Two_entries_with_the_same_name_are_refused()
    {
        List<(string Path, string Content)> parts = Parts(Sheet(OneCell));
        parts.Add(("XL/Worksheets/Sheet1.xml", Sheet("<row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t>SECOND</t></is></c></row>")));
        await using MemoryStream file = File(parts);

        Assert.Contains("twice", await RefusedAsync(file), StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("../../evil.txt")]
    [InlineData("xl/../../evil.txt")]
    [InlineData("/abs-evil.txt")]
    [InlineData("C:/evil.txt")]
    [InlineData("..\\..\\evil.txt")]
    [InlineData("xl/evil\u0001.txt")]
    public async Task An_entry_name_outside_the_package_is_refused(string name)
    {
        List<(string Path, string Content)> parts = Parts(Sheet(OneCell));
        parts.Add((name, "pwned"));
        await using MemoryStream file = File(parts);

        Assert.Contains("outside the package", await RefusedAsync(file), StringComparison.Ordinal);
    }

    /// <summary>Found by the fuzzer: two sheets on one part opened, and the save wrote the part twice, which no program can read.</summary>
    [Fact]
    public async Task Two_sheets_on_the_same_part_are_refused()
    {
        string workbook = $"""<workbook xmlns="{Main}" xmlns:r="{Relationships}"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/><sheet name="Copy" sheetId="2" r:id="rId1"/></sheets></workbook>""";
        await using MemoryStream file = File(Parts(Sheet(OneCell), workbook: workbook));

        Result<XlsxReader> reader = await XlsxReader.OpenAsync(file, true);

        Assert.True(reader.IsFailure(out string? error, out _));
        Assert.Contains("point to the same part", error, StringComparison.Ordinal);
    }

    /// <summary>
    /// Found by libFuzzer: in .NET 10, XmlTextReaderImpl throws ArgumentOutOfRangeException from
    /// Buffer.BlockCopy for this malformed declaration read from a deflate stream. It is broken
    /// XML, and must be refused as such.
    /// </summary>
    [Fact]
    public async Task A_declaration_that_breaks_the_XML_reader_is_a_damaged_file()
    {
        byte[] declaration = [.. Encoding.ASCII.GetBytes("<?xml version=\"1.0zzzzzzzz"), .. Enumerable.Repeat((byte)0xFF, 94), .. Encoding.ASCII.GetBytes(new string('z', 28) + "-8\" standalone=\"yes\"?>\rmats.org/>5.5")];
        await using MemoryStream file = new();

        using (ZipArchive archive = new(file, ZipArchiveMode.Create, true))
        {
            foreach ((string path, string content) in Parts(Sheet(OneCell)))
            {
                byte[] bytes = Encoding.UTF8.GetBytes(content);

                if (path == "_rels/.rels")
                {
                    bytes = declaration;
                }

                await using Stream entry = await archive.CreateEntry(path, CompressionLevel.Fastest).OpenAsync();
                await entry.WriteAsync(bytes);
            }
        }

        file.Position = 0;

        Assert.Contains("broken XML", await RefusedAsync(file), StringComparison.Ordinal);
    }

    /// <summary>
    /// Found by libFuzzer: sheets and names outside &lt;sheets&gt; and &lt;definedNames&gt; were read
    /// as well. The saver rewrote &lt;sheets&gt; and copied the unknown element, so every sheet came
    /// back twice and the saved file could not be opened.
    /// </summary>
    [Fact]
    public async Task Sheets_and_names_outside_their_sections_are_not_read()
    {
        string workbook = $"""<workbook xmlns="{Main}" xmlns:r="{Relationships}"><sheetz><sheet name="Ghost" sheetId="9" r:id="rId1"/></sheetz><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets><definedNamez><definedName name="Ghost">Data!$A$1</definedName></definedNamez></workbook>""";
        await using MemoryStream file = File(Parts(Sheet(OneCell), workbook: workbook));

        Assert.True((await Workbook.OpenAsync(file)).IsSuccess(out Workbook? book));
        Assert.Equal(["Data"], book.Sheets.Select(sheet => sheet.Name));
        Assert.True(book.Evaluate("Ghost").IsSuccess(out CellValue ghost));
        Assert.Equal(CellValue.Error("#NAME?"), ghost);

        book.Sheets[0][0, 0] = "changed";
        await using MemoryStream saved = new();
        await book.SaveAsync(saved);
        saved.Position = 0;

        Assert.True((await Workbook.OpenAsync(saved)).IsSuccess(out Workbook? again));
        Assert.Equal(["Data"], again.Sheets.Select(sheet => sheet.Name));
    }

    /// <summary>Found by the zip-raw fuzz target: an entry without a name opened, and the save failed when it wrote the entry again.</summary>
    [Fact]
    public async Task An_entry_without_a_name_is_refused()
    {
        List<(string Path, string Content)> parts = Parts(Sheet(OneCell));
        parts.Add(("xl/zzzz.xml", "<x/>"));
        await using MemoryStream original = File(parts);
        byte[] bytes = original.ToArray();
        SetNameLength(bytes, "xl/zzzz.xml");
        await using MemoryStream file = new(bytes);

        // The local header still has the name, so the strict zip check is the first to see it.
        Assert.Contains("does not say one thing only", await RefusedAsync(file), StringComparison.Ordinal);

        file.Position = 0;
        Result<XlsxReader> lax = await XlsxReader.OpenAsync(file, KalkylLimits.Default with { StrictZip = false }, true);
        Assert.True(lax.IsFailure(out string? error, out _));
        Assert.Contains("without a name", error, StringComparison.Ordinal);
    }

    /// <summary>Sets the name length of an entry to 0 in its central directory record, as a damaged file can have it.</summary>
    private static void SetNameLength(byte[] zip, string name)
    {
        byte[] wanted = Encoding.UTF8.GetBytes(name);

        for (int position = 0; position < zip.Length - 46; position++)
        {
            if (zip[position] == 0x50 && zip[position + 1] == 0x4B && zip[position + 2] == 1 && zip[position + 3] == 2 && zip.AsSpan(position + 46).StartsWith(wanted) == true)
            {
                // The name's bytes stay, now read as the extra field, so the record keeps its length.
                int extraLength = zip[position + 30] | (zip[position + 31] << 8);
                int newExtra = extraLength + wanted.Length;
                zip[position + 28] = 0;
                zip[position + 29] = 0;
                zip[position + 30] = (byte)newExtra;
                zip[position + 31] = (byte)(newExtra >> 8);
            }
        }
    }

    [Fact]
    public async Task Backslashes_in_entry_names_are_still_read()
    {
        List<(string Path, string Content)> parts = Parts(Sheet(OneCell));
        parts[4] = ("xl\\worksheets\\sheet1.xml", parts[4].Content);
        await using MemoryStream file = File(parts);

        Assert.True((await Workbook.OpenAsync(file)).IsSuccess(out Workbook? book));
        Assert.Equal((CellValue)"x", book.Sheets[0]["A1"].Value);
    }

    [Fact]
    public async Task An_entry_with_a_compression_method_dotnet_lacks_is_refused()
    {
        await using MemoryStream original = File(Parts(Sheet(OneCell)));
        byte[] bytes = original.ToArray();
        SetCompressionMethod(bytes, "xl/worksheets/sheet1.xml", 14);
        await using MemoryStream file = new(bytes);

        Assert.Contains("cannot be unpacked", await RefusedAsync(file), StringComparison.Ordinal);
    }

    /// <summary>Sets the compression method of an entry in both its local header and the central directory, as a file made with LZMA would have it.</summary>
    private static void SetCompressionMethod(byte[] zip, string name, ushort method)
    {
        byte[] wanted = Encoding.UTF8.GetBytes(name);

        for (int position = 0; position < zip.Length - 46; position++)
        {
            bool local = zip[position] == 0x50 && zip[position + 1] == 0x4B && zip[position + 2] == 3 && zip[position + 3] == 4;
            bool central = zip[position] == 0x50 && zip[position + 1] == 0x4B && zip[position + 2] == 1 && zip[position + 3] == 2;
            int nameAt = 46;
            int methodAt = 10;

            if (local == true)
            {
                nameAt = 30;
                methodAt = 8;
            }

            if ((local == true || central == true) && zip.AsSpan(position + nameAt).StartsWith(wanted) == true)
            {
                zip[position + methodAt] = (byte)method;
                zip[position + methodAt + 1] = (byte)(method >> 8);
            }
        }
    }

    [Theory]
    [InlineData("<row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t>{0}</t></is></c></row>")]
    [InlineData("<row r=\"1\"><c r=\"A1\" t=\"str\"><f>A2</f><v>{0}</v></c></row>")]
    public async Task Text_longer_than_a_cell_holds_is_a_damaged_file(string cell)
    {
        await using MemoryStream file = File(Parts(Sheet(string.Format(System.Globalization.CultureInfo.InvariantCulture, cell, new string('a', 32_768)))));

        Assert.Contains("at most 32", await RefusedAsync(file), StringComparison.Ordinal);
    }

    [Fact]
    public async Task A_long_shared_string_is_a_damaged_file()
    {
        List<(string Path, string Content)> parts = Parts(Sheet("<row r=\"1\"><c r=\"A1\" t=\"s\"><v>0</v></c></row>"), WorkbookRelationships($"""<Relationship Id="rId2" Type="{Relationships}/sharedStrings" Target="sharedStrings.xml"/>"""));
        parts.Add(("xl/sharedStrings.xml", $"""<sst xmlns="{Main}"><si><t>{new string('a', 40_000)}</t></si></sst>"""));
        await using MemoryStream file = File(parts);

        Assert.Contains("at most 32", await RefusedAsync(file), StringComparison.Ordinal);
    }

    [Fact]
    public async Task More_shared_strings_than_the_limit_is_Invalid()
    {
        List<(string Path, string Content)> parts = Parts(Sheet("<row r=\"1\"><c r=\"A1\" t=\"s\"><v>0</v></c></row>"), WorkbookRelationships($"""<Relationship Id="rId2" Type="{Relationships}/sharedStrings" Target="sharedStrings.xml"/>"""));
        parts.Add(("xl/sharedStrings.xml", $"<sst xmlns=\"{Main}\">" + string.Concat(Enumerable.Repeat("<si><t>a</t></si>", 11)) + "</sst>"));
        await using MemoryStream file = File(parts);

        Result<XlsxReader> reader = await XlsxReader.OpenAsync(file, KalkylLimits.Default with { MaxSharedStrings = 10 }, true);
        file.Position = 0;
        Result<XlsxReader> enough = await XlsxReader.OpenAsync(file, KalkylLimits.Default with { MaxSharedStrings = 11 }, true);

        Assert.True(reader.IsFailure(out string? error, out _));
        Assert.Contains("MaxSharedStrings", error, StringComparison.Ordinal);
        Assert.Equal(KalkylErrorKind.TooLarge, reader.ErrorKind());
        Assert.True(enough.IsSuccess(out XlsxReader? opened));
        await opened.DisposeAsync();
    }

    /// <summary>200 000 nested &lt;r&gt; in one cell took 28 s and 107 GB of allocations before the depth limit.</summary>
    [Theory]
    [InlineData("cell")]
    [InlineData("extension")]
    public async Task Elements_nested_past_the_limit_are_a_damaged_file_and_quick(string where)
    {
        string nested = string.Concat(Enumerable.Repeat("<r>", 200_000)) + "<t>x</t>" + string.Concat(Enumerable.Repeat("</r>", 200_000));
        string sheet = Sheet($"<row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is>{nested}</is></c></row>");

        if (where == "extension")
        {
            sheet = Sheet(OneCell, $"<extLst>{nested}</extLst>");
        }

        await using MemoryStream file = File(Parts(sheet));
        System.Diagnostics.Stopwatch watch = System.Diagnostics.Stopwatch.StartNew();

        Assert.Contains("levels deep", await RefusedAsync(file), StringComparison.Ordinal);
        Assert.True(watch.Elapsed < TimeSpan.FromSeconds(5), $"Took {watch.Elapsed}.");
    }

    [Fact]
    public async Task Rich_text_nested_as_deep_as_Excel_writes_is_read()
    {
        string nested = string.Concat(Enumerable.Repeat("<r>", 20)) + "<t>deep</t>" + string.Concat(Enumerable.Repeat("</r>", 20));
        await using MemoryStream file = File(Parts(Sheet($"<row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is>{nested}</is></c></row>")));

        Assert.True((await Workbook.OpenAsync(file)).IsSuccess(out Workbook? book));
        Assert.Equal((CellValue)"deep", book.Sheets[0]["A1"].Value);
    }

    [Fact]
    public async Task An_external_entity_is_refused()
    {
        await using MemoryStream file = File(Parts("<?xml version=\"1.0\"?><!DOCTYPE x [<!ENTITY e SYSTEM \"file:///C:/Windows/win.ini\">]>" + Sheet("<row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t>&e;</t></is></c></row>")));

        Assert.Contains("DTD", await RefusedAsync(file), StringComparison.Ordinal);
    }

    [Fact]
    public async Task A_sheet_pointing_outside_the_package_is_refused()
    {
        await using MemoryStream file = File(Parts(Sheet(OneCell), WorkbookRelationships(sheetTarget: "../../../../Windows/win.ini")));

        Result<Workbook> opened = await Workbook.OpenAsync(file);

        Assert.True(opened.IsFailure(out string? error, out _));
        Assert.Contains("not in the file", error, StringComparison.Ordinal);
    }

    /// <summary>Nothing in a formula reaches outside the workbook: no DDE, no web, no file, no information about the machine.</summary>
    [Theory]
    [InlineData("cmd|' /C calc'!A0", "#NAME?")]
    [InlineData("[1]Sheet1!A1", "#NAME?")]
    [InlineData("WEBSERVICE(\"http://attacker.invalid\")", "#NAME?")]
    [InlineData("FILTERXML(\"<a/>\",\"//a\")", "#NAME?")]
    [InlineData("CALL(\"kernel32\",\"WinExec\",\"JC\",\"calc\",0)", "#NAME?")]
    [InlineData("CELL(\"filename\")", "#NAME?")]
    [InlineData("INFO(\"directory\")", "#NAME?")]
    [InlineData("INDIRECT(\"'C:\\secret\\[x.xlsx]S'!A1\")", "#REF!")]
    [InlineData("HYPERLINK(\"http://attacker.invalid/x\",\"click\")", "click")]
    public async Task Formulas_that_would_reach_outside_give_errors_or_text(string formula, string expected)
    {
        await using MemoryStream file = File(Parts(Sheet($"<row r=\"1\"><c r=\"A1\"><f>{System.Security.SecurityElement.Escape(formula)}</f><v>0</v></c></row>")));
        Assert.True((await Workbook.OpenAsync(file)).IsSuccess(out Workbook? book));

        await book.RecalculateAsync(RecalculationScope.All);

        Assert.Equal(expected, book.Sheets[0]["A1"].Value.ToString());
    }

    [Fact]
    public async Task Names_that_refer_to_themselves_end()
    {
        string workbook = $"""<workbook xmlns="{Main}" xmlns:r="{Relationships}"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets><definedNames><definedName name="Loop">Loop+1</definedName><definedName name="A_">B_</definedName><definedName name="B_">A_</definedName></definedNames></workbook>""";
        await using MemoryStream file = File(Parts(Sheet("<row r=\"1\"><c r=\"A1\"><f>Loop</f></c><c r=\"B1\"><f>A_</f></c></row>"), workbook: workbook));
        Assert.True((await Workbook.OpenAsync(file)).IsSuccess(out Workbook? book));

        await book.RecalculateAsync(RecalculationScope.All);

        Assert.True(book.Sheets[0]["A1"].Value.TryGetError(out _));
        Assert.True(book.Sheets[0]["B1"].Value.TryGetError(out _));
    }

    [Fact]
    public async Task A_shared_formula_over_the_whole_sheet_is_read_cell_by_cell()
    {
        await using MemoryStream file = File(Parts(Sheet("<row r=\"1\"><c r=\"A1\"><f t=\"shared\" ref=\"A1:XFD1048576\" si=\"0\">ROW()</f><v>1</v></c></row><row r=\"2\"><c r=\"B2\"><f t=\"shared\" si=\"0\"/></c></row>")));
        Assert.True((await Workbook.OpenAsync(file)).IsSuccess(out Workbook? book));

        await book.RecalculateAsync(RecalculationScope.All);

        Assert.Equal((CellValue)2, book.Sheets[0]["B2"].Value);
    }

    [Fact]
    public async Task A_file_without_content_types_is_saved_with_them()
    {
        List<(string Path, string Content)> parts = Parts(Sheet(OneCell));
        parts.RemoveAt(0);
        await using MemoryStream file = File(parts);
        Assert.True((await Workbook.OpenAsync(file)).IsSuccess(out Workbook? book));
        await using MemoryStream saved = new();

        await book.SaveAsync(saved);

        string types = TestFiles.ReadPart(saved, "[Content_Types].xml");
        Assert.Contains("PartName=\"/xl/workbook.xml\"", types, StringComparison.Ordinal);
        Assert.Contains("PartName=\"/xl/worksheets/sheet1.xml\"", types, StringComparison.Ordinal);
        Assert.Contains("Extension=\"rels\"", types, StringComparison.Ordinal);
    }

    private static List<(string Path, string Content)> MacroWorkbook(string extra = "")
    {
        List<(string Path, string Content)> parts = Parts(
            Sheet(OneCell),
            WorkbookRelationships("""<Relationship Id="rId9" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/>"""),
            contentTypes: ContentTypes("application/vnd.ms-excel.sheet.macroEnabled.main+xml"));
        parts.Add(("xl/vbaProject.bin", "MACRO-BYTES"));
        parts.Add(("xl/_rels/vbaProject.bin.rels", $"""<Relationships xmlns="{PackageRelationships}"><Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProjectSignature" Target="vbaProjectSignature.bin"/></Relationships>"""));
        parts.Add(("xl/vbaProjectSignature.bin", "SIGNATURE"));

        if (extra.Length > 0)
        {
            parts.Add((extra, "<x/>"));
        }

        return parts;
    }

    [Fact]
    public async Task An_xlsm_keeps_its_macros_unless_they_are_removed()
    {
        await using MemoryStream file = File(MacroWorkbook());
        Assert.True((await Workbook.OpenAsync(file)).IsSuccess(out Workbook? book));
        Assert.True(book.HasMacros);
        await using MemoryStream kept = new();
        await book.SaveAsync(kept);
        Assert.Contains("xl/vbaProject.bin", Entries(kept));

        Assert.True(book.RemoveMacros().IsSuccess(out string? error, out _), error);
        await using MemoryStream removed = new();
        await book.SaveAsync(removed);

        string[] entries = Entries(removed);
        Assert.DoesNotContain(entries, entry => entry.Contains("vbaProject", StringComparison.OrdinalIgnoreCase) == true);
        string types = TestFiles.ReadPart(removed, "[Content_Types].xml");
        Assert.DoesNotContain("macroEnabled", types, StringComparison.Ordinal);
        Assert.DoesNotContain("vbaProject", types, StringComparison.Ordinal);
        Assert.Contains("spreadsheetml.sheet.main+xml", types, StringComparison.Ordinal);
        Assert.DoesNotContain("vbaProject", TestFiles.ReadPart(removed, "xl/_rels/workbook.xml.rels"), StringComparison.Ordinal);
        removed.Position = 0;
        Assert.True((await Workbook.OpenAsync(removed)).IsSuccess(out Workbook? clean));
        Assert.False(clean.HasMacros);
        Assert.Empty(TestFiles.Validate(removed));
    }

    [Theory]
    [InlineData("xl/macrosheets/sheet1.xml")]
    [InlineData("xl/activeX/activeX1.xml")]
    public async Task Macros_woven_into_the_sheets_cannot_be_removed(string part)
    {
        await using MemoryStream file = File(MacroWorkbook(part));
        Assert.True((await Workbook.OpenAsync(file)).IsSuccess(out Workbook? book));

        Assert.True(book.HasMacros);
        Assert.True(book.RemoveMacros().IsFailure(out string? error, out _));
        Assert.Contains("cannot remove", error, StringComparison.Ordinal);
    }

    private static string[] Entries(MemoryStream zip)
    {
        zip.Position = 0;
        using ZipArchive archive = new(zip, ZipArchiveMode.Read, true);

        return [.. archive.Entries.Select(entry => entry.FullName)];
    }

    [Theory]
    [InlineData("=1+1")]
    [InlineData("@SUM(A1)")]
    [InlineData("+cmd|' /C calc'!A0")]
    [InlineData("-2+3")]
    public async Task Text_that_looks_like_a_formula_is_written_as_text(string text)
    {
        await using MemoryStream file = await TestFiles.WriteAsync(async writer =>
        {
            SheetWriter sheet = await TestFiles.AddSheetAsync(writer, "Data");
            await sheet.WriteRowAsync([text]);
        });

        string xml = TestFiles.ReadPart(file, "xl/worksheets/sheet1.xml");

        Assert.Contains("t=\"inlineStr\"", xml, StringComparison.Ordinal);
        Assert.DoesNotContain("<f>", xml, StringComparison.Ordinal);
    }

    /// <summary>
    /// One value in XFD1048576 makes the used area of A:XFD 17 billion cells. The aggregates read
    /// only the cells that are there and give Excel's answers at once (it used to be #NUM!).
    /// </summary>
    [Theory]
    [InlineData("SUM(A:XFD)", "8")]
    [InlineData("SUM(1:1048576)", "8")]
    [InlineData("COUNT(A:XFD)", "3")]
    [InlineData("COUNTA(A:XFD)", "4")]
    [InlineData("AVERAGE(A:XFD)", "2.6666666666666665")]
    [InlineData("MAX(A:XFD)", "5")]
    [InlineData("MIN(A:XFD)", "1")]
    [InlineData("MEDIAN(A:XFD)", "2")]
    [InlineData("PRODUCT(A:XFD)", "10")]
    [InlineData("COUNTIF(A:XFD,\">0\")", "3")]
    [InlineData("COUNTIFS(A:XFD,\"x\")", "1")]
    [InlineData("SUMIF(A:XFD,\">1\")", "7")]
    [InlineData("SUMIFS(A:XFD,A:XFD,\">1\")", "7")]
    [InlineData("AVERAGEIF(A:XFD,\">1\")", "3.5")]
    [InlineData("COUNTBLANK(A:XFD)", "17179869180")]
    [InlineData("COUNTIF(A:XFD,\"\")", "17179869180")]
    public void Whole_sheet_aggregates_with_one_value_far_away_give_Excels_answer(string formula, string expected)
    {
        Workbook book = Workbook.Create();
        Assert.True(book.AddSheet("Data").IsSuccess(out Worksheet? sheet));
        sheet["A1"] = 1;
        sheet["B2"] = 5;
        sheet["C3"] = "x";
        sheet["XFD1048576"] = 2;
        using CancellationTokenSource timeout = new(TimeSpan.FromSeconds(10));

        Assert.True(book.Evaluate(formula, sheet, timeout.Token).IsSuccess(out CellValue value));

        Assert.Equal(expected, value.ToString());
    }

    [Fact]
    public async Task A_whole_sheet_criteria_walk_can_be_cancelled()
    {
        Workbook book = Workbook.Create();
        Assert.True(book.AddSheet("Data").IsSuccess(out Worksheet? sheet));
        sheet["A1"] = 1;
        sheet["XFD1048576"] = 2;
        using CancellationTokenSource timeout = new(TimeSpan.FromMilliseconds(300));
        System.Diagnostics.Stopwatch watch = System.Diagnostics.Stopwatch.StartNew();

        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => Task.Run(() => book.Evaluate("COUNTIFS(A:XFD,\"\",A:XFD,\"\")", sheet, timeout.Token)));

        Assert.True(watch.Elapsed < TimeSpan.FromSeconds(5), $"Stopped after {watch.Elapsed}.");
    }

    [Fact]
    public async Task Parallel_recalculations_keep_their_own_limits()
    {
        Workbook template = Workbook.Create();
        Assert.True(template.AddSheet("Data").IsSuccess(out Worksheet? sheet));

        for (int row = 0; row < 200; row++)
        {
            sheet[row, 0] = row + 1;
            sheet[row, 1] = Cell.FromFormula($"SUMPRODUCT(A1:A200*1)+A{row + 1}");
        }

        await using MemoryStream file = new();
        await template.SaveAsync(file);
        byte[] bytes = file.ToArray();

        CellValue[] answers = await Task.WhenAll(Enumerable.Range(0, 16).Select(async index =>
        {
            Assert.True((await Workbook.OpenAsync(new MemoryStream(bytes))).IsSuccess(out Workbook? book));

            if (index % 2 == 0)
            {
                book.Limits = KalkylLimits.Default with { MaxArrayCells = 10 };
            }

            await book.RecalculateAsync(RecalculationScope.All);

            return book.Sheets[0]["B5"].Value;
        }));

        for (int index = 0; index < answers.Length; index++)
        {
            if (index % 2 == 0)
            {
                Assert.Equal(CellValue.Error("#NUM!"), answers[index]);
            }
            else
            {
                Assert.Equal((CellValue)(20_100 + 5), answers[index]);
            }
        }
    }

    /// <summary>
    /// The package must not reach the network, start processes, load code, deserialize objects
    /// or use unsafe code (docs/testplan.md, S-15, S-18, S-24). Checked in the source, so that a
    /// change that brings one in fails here rather than in a review.
    /// </summary>
    [Fact]
    public void The_package_uses_no_network_process_deserialization_or_unsafe_code()
    {
        string source = Path.Combine(RepositoryRoot(), "src", "Avig.Kalkyl");
        Regex forbidden = new(@"System\.Net\b|HttpClient|WebRequest|Socket|Process\.Start|ProcessStartInfo|Assembly\.Load|Activator\.CreateInstance|BinaryFormatter|XmlSerializer|DataContractSerializer|JsonSerializer|TypeNameHandling|\bunsafe\b|DllImport|LibraryImport|Marshal\.|XmlResolver\s*=\s*new|DtdProcessing\.Parse");
        List<string> found = [];

        foreach (string path in Directory.GetFiles(source, "*.cs", SearchOption.AllDirectories).Where(path => path.Contains($"{Path.DirectorySeparatorChar}obj{Path.DirectorySeparatorChar}", StringComparison.Ordinal) == false))
        {
            string[] lines = System.IO.File.ReadAllLines(path);

            for (int index = 0; index < lines.Length; index++)
            {
                string line = lines[index].TrimStart();

                if (line.StartsWith("//", StringComparison.Ordinal) == false && forbidden.IsMatch(line) == true)
                {
                    found.Add($"{Path.GetFileName(path)}:{index + 1}: {line}");
                }
            }
        }

        Assert.True(found.Count == 0, string.Join(Environment.NewLine, found));
    }

    private static string RepositoryRoot()
    {
        DirectoryInfo? directory = new(AppContext.BaseDirectory);

        while (directory is not null && System.IO.File.Exists(Path.Combine(directory.FullName, "Avig.Kalkyl.slnx")) == false)
        {
            directory = directory.Parent;
        }

        Assert.NotNull(directory);

        return directory.FullName;
    }
}
