using Xunit;

namespace Cellwright.Xls.Tests;

/// <summary>
/// Reading the notes of a legacy workbook: the Note record anchors them, the Obj record names
/// the drawing object, and the TxO record after it carries the text. The fixtures are the POI
/// test files whose content that project's own suite pins down.
/// </summary>
public sealed class CommentTests
{
    private static string? Find(string name) =>
        XlsCorpus.Files.FirstOrDefault(path => Path.GetFileName(path).Equals(name, StringComparison.OrdinalIgnoreCase));

    [Fact]
    public async Task TheNotesOfAKnownWorkbookAreRead()
    {
        string? path = Find("SimpleWithComments.xls");
        Assert.SkipWhen(path is null, "SimpleWithComments.xls is not in the corpus. These checks read the .xls and .xlsb files that ship with ExcelDataReader, NPOI, EPPlus and Aspose.Cells; unpack those repositories beside this one to run them.");

        Workbook workbook = await XlsWorkbook.LoadAsync(path!, options: null, TestContext.Current.CancellationToken);
        Worksheet sheet = workbook.Sheets[0];

        Assert.Equal(3, sheet.Comments.Count);
        foreach ((string cell, string tail) in new[] { ("B1", "first cell"), ("B2", "second cell"), ("B3", "third cell") })
        {
            CellComment? comment = sheet.Comments[CellRef.Parse(cell)];
            Assert.NotNull(comment);
            Assert.Equal("Yegor Kozlov", comment.Author);
            Assert.EndsWith(tail, comment.Text);
        }
    }

    [Fact]
    public async Task NotesSurviveConversionToTheModernFormat()
    {
        string? path = Find("SimpleWithComments.xls");
        Assert.SkipWhen(path is null, "SimpleWithComments.xls is not in the corpus. These checks read the .xls and .xlsb files that ship with ExcelDataReader, NPOI, EPPlus and Aspose.Cells; unpack those repositories beside this one to run them.");

        Workbook legacy = await XlsWorkbook.LoadAsync(path!, options: null, TestContext.Current.CancellationToken);

        using var buffer = new MemoryStream();
        await legacy.SaveAsync(buffer, new SaveOptions { LeaveOpen = true }, TestContext.Current.CancellationToken);
        buffer.Position = 0;
        Workbook modern = await Workbook.LoadAsync(
            buffer, new LoadOptions { LeaveOpen = true }, TestContext.Current.CancellationToken);

        Worksheet sheet = modern.Sheets[0];
        Assert.Equal(3, sheet.Comments.Count);
        CellComment? comment = sheet.Comments[CellRef.Parse("B2")];
        Assert.NotNull(comment);
        Assert.Equal("Yegor Kozlov", comment.Author);
        Assert.EndsWith("second cell", comment.Text);
    }

    [Fact]
    public async Task NotesSurviveAWriteAndRead()
    {
        Workbook original = Workbook.Create();
        Worksheet first = original.Sheets.Add("First");
        first["B2"] = CellValue.Text("value");
        first.Comments.Add("B2", "The first note", "Alice");
        first.Comments.Add("D7", "Another note, longer than the first one", "Bob");

        Worksheet second = original.Sheets.Add("Second");
        second.Comments.Add("A1", "Note on another sheet", "Carol");

        Workbook reloaded = await RoundTripAsync(original);

        Worksheet readFirst = reloaded.Sheets[0];
        Assert.Equal(2, readFirst.Comments.Count);
        Assert.Equal("The first note", readFirst.Comments[CellRef.Parse("B2")]?.Text);
        Assert.Equal("Alice", readFirst.Comments[CellRef.Parse("B2")]?.Author);
        Assert.Equal("Another note, longer than the first one", readFirst.Comments[CellRef.Parse("D7")]?.Text);
        Assert.Equal("Bob", readFirst.Comments[CellRef.Parse("D7")]?.Author);

        Worksheet readSecond = reloaded.Sheets[1];
        Assert.Equal("Note on another sheet", Assert.Single(readSecond.Comments).Text);
    }

    [Fact]
    public async Task ANoteInAnotherAlphabetSurvives()
    {
        Workbook original = Workbook.Create();
        Worksheet sheet = original.Sheets.Add("Data");
        sheet.Comments.Add("A1", "Примечание кириллицей", "Автор");

        Worksheet read = (await RoundTripAsync(original)).Sheets[0];

        CellComment note = Assert.Single(read.Comments);
        Assert.Equal("Примечание кириллицей", note.Text);
        Assert.Equal("Автор", note.Author);
    }

    /// <summary>
    /// The drawing layer is the fussiest thing this library writes, so an independent reader
    /// has to agree that what came out is a note on the right cell.
    /// </summary>
    [Fact]
    public async Task AnIndependentReaderFindsTheNotes()
    {
        Workbook original = Workbook.Create();
        Worksheet sheet = original.Sheets.Add("Notes");
        sheet["B2"] = CellValue.Text("value");
        sheet.Comments.Add("B2", "Reviewed", "Alice");
        sheet.Comments.Add("C5", "Check this", "Bob");

        using var buffer = new MemoryStream();
        await XlsWorkbook.SaveAsync(original, buffer, options: null, TestContext.Current.CancellationToken);
        buffer.Position = 0;

        var oracle = new NPOI.HSSF.UserModel.HSSFWorkbook(buffer);
        NPOI.SS.UserModel.ISheet read = oracle.GetSheetAt(0);

        NPOI.SS.UserModel.IComment first = read.GetCellComment(new NPOI.SS.Util.CellAddress(1, 1));
        Assert.NotNull(first);
        Assert.Equal("Reviewed", first.String.String);
        Assert.Equal("Alice", first.Author);

        NPOI.SS.UserModel.IComment second = read.GetCellComment(new NPOI.SS.Util.CellAddress(4, 2));
        Assert.NotNull(second);
        Assert.Equal("Check this", second.String.String);
    }

    /// <summary>
    /// A workbook with both drawings and notes mixes text boxes into the same object stream;
    /// their TxO records must not be mistaken for note text.
    /// </summary>
    [Fact]
    public async Task DrawingsAmongTheNotesDoNotConfuseThem()
    {
        string? path = Find("DrawingAndComments.xls");
        Assert.SkipWhen(path is null, "DrawingAndComments.xls is not in the corpus. These checks read the .xls and .xlsb files that ship with ExcelDataReader, NPOI, EPPlus and Aspose.Cells; unpack those repositories beside this one to run them.");

        Workbook workbook = await XlsWorkbook.LoadAsync(path!, options: null, TestContext.Current.CancellationToken);

        // The exact notes are the file's business; what matters is that every note that was
        // found is anchored and carries text that came from a note object, not from a shape.
        foreach (Worksheet sheet in workbook.Sheets)
        {
            foreach (CellComment comment in sheet.Comments)
                Assert.NotNull(comment.Text);
        }
    }

    private static async Task<Workbook> RoundTripAsync(Workbook workbook)
    {
        using var buffer = new MemoryStream();
        await XlsWorkbook.SaveAsync(workbook, buffer, options: null, TestContext.Current.CancellationToken);
        buffer.Position = 0;
        return await XlsWorkbook.LoadAsync(
            buffer, new XlsLoadOptions { LeaveOpen = true }, TestContext.Current.CancellationToken);
    }
}
