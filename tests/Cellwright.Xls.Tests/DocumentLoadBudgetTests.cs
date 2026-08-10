using System.Buffers.Binary;
using System.IO.Compression;
using Cellwright.Model;
using Cellwright.Xls;
using Cellwright.Xls.Biff;
using Cellwright.Xls.Formulas;
using Cellwright.Xls.Reading;
using Cellwright.Xls.Vba;
using Cellwright.Xls.Xlsb;
using Xunit;

namespace Cellwright.Xls.Tests;

public sealed class DocumentLoadBudgetTests
{
    [Fact]
    public void XlsOptionsExposeTheEffectiveBudgetAndTheLastAssignmentWins()
    {
        var legacy = new PackageLimits { MaxPartSize = 17, MaxTotalSize = 29 };
        var custom = DocumentLoadBudget.Default with { MaxCells = 13, MaxPartBytes = 41 };

        var legacyLast = new XlsLoadOptions { Budget = custom, Limits = legacy };
        Assert.Equal(17, legacyLast.Budget.MaxPartBytes);
        Assert.Equal(29, legacyLast.Budget.MaxInputBytes);
        Assert.Equal(13, legacyLast.Budget.MaxCells);

        var budgetLast = new XlsLoadOptions { Limits = legacy, Budget = custom };
        Assert.Same(custom, budgetLast.Budget);
        Assert.Same(
            DocumentLoadBudget.Unlimited,
            new XlsLoadOptions { Limits = PackageLimits.None }.Budget);
    }

    [Fact]
    public async Task DefaultBudgetAcceptsAndRepairsTheKnownOverlongOpenOfficeString()
    {
        string? path = XlsCorpus.Files.FirstOrDefault(static path =>
            Path.GetFileName(path).Equals("46368.xls", StringComparison.OrdinalIgnoreCase));
        Assert.SkipWhen(path is null, "46368.xls is not in the sibling corpus.");

        var warnings = new List<WorkbookWarning>();
        Workbook workbook = await XlsWorkbook.LoadAsync(
            path!,
            new XlsLoadOptions { OnWarning = warnings.Add },
            TestContext.Current.CancellationToken);

        Assert.Contains(warnings, static warning => warning.Code == "TextTooLong");
        Assert.Contains(
            workbook.Sheets.SelectMany(static sheet => sheet.Values.EnumerateOccupied()),
            static item => item.Value.IsText && item.Value.GetText().Length == SheetLimits.MaxCellTextLength);
    }

    [Fact]
    public async Task CustomExcelLengthBudgetRejectsTheKnownOverlongOpenOfficeString()
    {
        string? path = XlsCorpus.Files.FirstOrDefault(static path =>
            Path.GetFileName(path).Equals("46368.xls", StringComparison.OrdinalIgnoreCase));
        Assert.SkipWhen(path is null, "46368.xls is not in the sibling corpus.");

        DocumentLoadLimitException error = await Assert.ThrowsAsync<DocumentLoadLimitException>(async () =>
            await XlsWorkbook.LoadAsync(
                path!,
                new XlsLoadOptions
                {
                    Budget = DocumentLoadBudget.Default with
                    {
                        MaxCellTextCharacters = SheetLimits.MaxCellTextLength,
                    },
                },
                TestContext.Current.CancellationToken));

        Assert.Equal(nameof(DocumentLoadBudget.MaxCellTextCharacters), error.LimitName);
        Assert.Equal(SheetLimits.MaxCellTextLength, error.Limit);
        Assert.Equal(32_770, error.Observed);
    }

    [Fact]
    public async Task LegacyWorkbookUsesSharedInputContainerCellAndStringCeilings()
    {
        byte[] content = await LegacyWorkbookAsync(includePicture: false);
        (string Name, DocumentLoadBudget Budget)[] cases =
        [
            (nameof(DocumentLoadBudget.MaxInputBytes),
                DocumentLoadBudget.Default with { MaxInputBytes = content.LongLength - 1 }),
            (nameof(DocumentLoadBudget.MaxPackageParts),
                DocumentLoadBudget.Default with { MaxPackageParts = 1 }),
            (nameof(DocumentLoadBudget.MaxInflatedBytes),
                DocumentLoadBudget.Default with { MaxInflatedBytes = 1 }),
            (nameof(DocumentLoadBudget.MaxPartBytes),
                DocumentLoadBudget.Default with { MaxPartBytes = 1 }),
            (nameof(DocumentLoadBudget.MaxCells),
                DocumentLoadBudget.Default with { MaxCells = 1 }),
            (nameof(DocumentLoadBudget.MaxSharedStrings),
                DocumentLoadBudget.Default with { MaxSharedStrings = 1 }),
            (nameof(DocumentLoadBudget.MaxCellTextCharacters),
                DocumentLoadBudget.Default with { MaxCellTextCharacters = 5 }),
            (nameof(DocumentLoadBudget.MaxTotalCellTextCharacters),
                DocumentLoadBudget.Default with { MaxTotalCellTextCharacters = 10 }),
        ];

        foreach ((string name, DocumentLoadBudget budget) in cases)
        {
            DocumentLoadLimitException error = await Assert.ThrowsAsync<DocumentLoadLimitException>(async () =>
                await XlsWorkbook.LoadAsync(
                    new MemoryStream(content, writable: false),
                    new XlsLoadOptions { Budget = budget },
                    TestContext.Current.CancellationToken));
            Assert.Equal(name, error.LimitName);
            Assert.True(error.Observed > error.Limit);
        }
    }

    [Fact]
    public async Task LegacyPictureInflationUsesTheSameMediaBudget()
    {
        byte[] content = await LegacyWorkbookAsync(includePicture: true);
        int pngLength = Png().Length;
        DocumentLoadLimitException error = await Assert.ThrowsAsync<DocumentLoadLimitException>(async () =>
            await XlsWorkbook.LoadAsync(
                new MemoryStream(content, writable: false),
                new XlsLoadOptions
                {
                    Budget = DocumentLoadBudget.Default with { MaxMediaBytes = pngLength - 1 },
                },
                TestContext.Current.CancellationToken));

        Assert.Equal(nameof(DocumentLoadBudget.MaxMediaBytes), error.LimitName);
        Assert.Equal(pngLength, error.Observed);
    }

    [Theory]
    [InlineData(nameof(DocumentLoadBudget.MaxCellTextCharacters), 1, 2)]
    [InlineData(nameof(DocumentLoadBudget.MaxTotalCellTextCharacters), 3, 4)]
    public void LegacySharedFormulaExpansionUsesTheCellTextBudget(
        string limitName, int limit, int observed)
    {
        DocumentLoadLimitException error = Assert.Throws<DocumentLoadLimitException>(() =>
            ParseLegacySharedFormula(FormulaBudget(limitName, limit)));

        Assert.Equal(limitName, error.LimitName);
        Assert.Equal(limit, error.Limit);
        Assert.Equal(observed, error.Observed);
    }

    [Theory]
    [InlineData(nameof(DocumentLoadBudget.MaxCellTextCharacters), 1, 2)]
    [InlineData(nameof(DocumentLoadBudget.MaxTotalCellTextCharacters), 3, 4)]
    public void XlsbSharedFormulaExpansionUsesTheCellTextBudget(
        string limitName, int limit, int observed)
    {
        DocumentLoadLimitException error = Assert.Throws<DocumentLoadLimitException>(() =>
            ParseXlsbSharedFormula(FormulaBudget(limitName, limit)));

        Assert.Equal(limitName, error.LimitName);
        Assert.Equal(limit, error.Limit);
        Assert.Equal(observed, error.Observed);
    }

    [Fact]
    public async Task FormatDetectionHandlesBoundedNonSeekableXlsxInput()
    {
        Workbook source = Workbook.Create();
        source.Sheets.Add("Data")["A1"] = "xlsx";
        using var package = new MemoryStream();
        await source.SaveAsync(package, cancellationToken: TestContext.Current.CancellationToken);
        byte[] packageBytes = package.ToArray();

        await using var input = new NonSeekableReadStream(packageBytes);
        Workbook loaded = await Spreadsheet.LoadAsync(
            input,
            new XlsLoadOptions
            {
                LeaveOpen = true,
                Budget = DocumentLoadBudget.Default with { MaxInputBytes = packageBytes.Length },
            },
            TestContext.Current.CancellationToken);

        Assert.Equal("xlsx", loaded.Sheets[0]["A1"].GetText());
    }

    [Fact]
    public async Task EncryptedContainerAndDecryptedPackageShareOneCumulativeState()
    {
        const string password = "load-budget";
        Workbook source = Workbook.Create();
        source.Sheets.Add("Data")["A1"] = "encrypted";

        using var plain = new MemoryStream();
        await source.SaveAsync(plain, cancellationToken: TestContext.Current.CancellationToken);
        int packageParts;
        using (var archive = new ZipArchive(
                   new MemoryStream(plain.ToArray(), writable: false),
                   ZipArchiveMode.Read))
        {
            packageParts = archive.Entries.Count;
        }

        using var encrypted = new MemoryStream();
        await source.SaveEncryptedAsync(
            encrypted,
            password,
            cancellationToken: TestContext.Current.CancellationToken);

        DocumentLoadLimitException error = await Assert.ThrowsAsync<DocumentLoadLimitException>(async () =>
            await Spreadsheet.LoadAsync(
                new MemoryStream(encrypted.ToArray(), writable: false),
                new XlsLoadOptions
                {
                    Password = password,
                    Budget = DocumentLoadBudget.Default with
                    {
                        MaxPackageParts = packageParts + 1,
                    },
                },
                TestContext.Current.CancellationToken));

        Assert.Equal(nameof(DocumentLoadBudget.MaxPackageParts), error.LimitName);
        Assert.True(error.Observed > packageParts + 1);
    }

    [Theory]
    [InlineData(nameof(DocumentLoadBudget.MaxPartBytes))]
    [InlineData(nameof(DocumentLoadBudget.MaxInflatedBytes))]
    public void VbaExpansionUsesTheOuterMachineReadableBudget(string limitName)
    {
        byte[] compressed = [0x01, 0x03, 0xB0, 0x02, (byte)'a', 0x04, 0x00];
        DocumentLoadBudget budget = limitName == nameof(DocumentLoadBudget.MaxPartBytes)
            ? DocumentLoadBudget.Default with { MaxPartBytes = 7 }
            : DocumentLoadBudget.Default with { MaxInflatedBytes = 7 };

        DocumentLoadLimitException error = Assert.Throws<DocumentLoadLimitException>(() =>
            VbaCompression.Decompress(
                compressed,
                new DocumentLoadBudgetState(budget),
                TestContext.Current.CancellationToken));

        Assert.Equal(limitName, error.LimitName);
        Assert.Equal(8, error.Observed);
    }

    [Fact]
    public void VbaExpansionHonorsCancellationInsideTheCodec()
    {
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        Assert.ThrowsAny<OperationCanceledException>(() =>
            VbaCompression.Decompress(
                [0x01, 0x03, 0xB0, 0x02, (byte)'a', 0x04, 0x00],
                loadBudget: null,
                cancellation.Token));
    }

    [Fact]
    public async Task BinaryReadersHonorCancellationAfterTheInputWasBuffered()
    {
        byte[] content = await LegacyWorkbookAsync(includePicture: false);
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();

        await Assert.ThrowsAnyAsync<OperationCanceledException>(async () =>
            await XlsWorkbook.LoadAsync(
                new MemoryStream(content, writable: false),
                cancellationToken: cancellation.Token));
    }

    private static async Task<byte[]> LegacyWorkbookAsync(bool includePicture)
    {
        Workbook workbook = Workbook.Create();
        Worksheet sheet = workbook.Sheets.Add("Data");
        sheet["A1"] = "first";
        sheet["B1"] = "second";
        if (includePicture)
        {
            sheet.Pictures.Add(
                Png(),
                CellRef.Parse("A2"),
                CellRef.Parse("B3"));
        }

        using var stream = new MemoryStream();
        await XlsWorkbook.SaveAsync(
            workbook, stream, cancellationToken: TestContext.Current.CancellationToken);
        return stream.ToArray();
    }

    private static DocumentLoadBudget FormulaBudget(string limitName, int limit) =>
        limitName == nameof(DocumentLoadBudget.MaxCellTextCharacters)
            ? DocumentLoadBudget.Default with { MaxCellTextCharacters = limit }
            : DocumentLoadBudget.Default with { MaxTotalCellTextCharacters = limit };

    private static void ParseLegacySharedFormula(DocumentLoadBudget budget)
    {
        using var writer = new BiffWriter(128);
        writer.WriteRecord(BiffRecordType.Formula, LegacyFormulaMember(row: 0));
        writer.WriteRecord(BiffRecordType.Formula, LegacyFormulaMember(row: 1));
        writer.WriteRecord(BiffRecordType.ShrFmla, LegacySharedFormula());
        writer.WriteRecord(BiffRecordType.Eof);

        using var reader = new BiffReader(writer.ToArray());
        var globals = new XlsGlobals { Version = BiffVersion.Biff8 };
        var decoder = new PtgDecoder(globals, ["Sheet1"]);
        Worksheet sheet = Workbook.Create().Sheets.Add("Sheet1");
        var loadBudget = new DocumentLoadBudgetState(budget);
        var sink = new WorksheetContentSink(
            sheet, styleMap: [], relationships: [], preserveMarkup: false, _ => { }, loadBudget);

        new SheetParser(
                reader,
                globals,
                decoder,
                sheetIndex: 0,
                substreamEnd: int.MaxValue,
                _ => { },
                loadBudget)
            .Parse(sink);
    }

    private static byte[] LegacyFormulaMember(int row)
    {
        var body = new byte[27];
        BinaryPrimitives.WriteUInt16LittleEndian(body, (ushort)row);
        BinaryPrimitives.WriteUInt16LittleEndian(body.AsSpan(20), 5);
        body[22] = 0x01; // PtgExp naming the A1 shared-formula master.
        return body;
    }

    private static byte[] LegacySharedFormula()
    {
        var body = new byte[13];
        BinaryPrimitives.WriteUInt16LittleEndian(body.AsSpan(8), 3);
        body[10] = 0x1E; // PtgInt 42.
        body[11] = 42;
        return body;
    }

    private static void ParseXlsbSharedFormula(DocumentLoadBudget budget)
    {
        byte[] part = new Biff12Builder()
            .Record(Biff12RecordType.RowHdr, XlsbRow(row: 0))
            .Record(Biff12RecordType.FmlaNum, XlsbFormulaMember())
            .Record(Biff12RecordType.RowHdr, XlsbRow(row: 1))
            .Record(Biff12RecordType.FmlaNum, XlsbFormulaMember())
            .Record(Biff12RecordType.ShrFmla, XlsbSharedFormula())
            .ToArray();

        var globals = new XlsGlobals { Version = BiffVersion.Biff12 };
        var decoder = new PtgDecoder(globals, ["Sheet1"]);
        Worksheet sheet = Workbook.Create().Sheets.Add("Sheet1");
        var loadBudget = new DocumentLoadBudgetState(budget);
        var sink = new WorksheetContentSink(
            sheet, styleMap: [], relationships: [], preserveMarkup: false, _ => { }, loadBudget);

        new XlsbSheetParser(globals, decoder, sheetIndex: 0, _ => { }, loadBudget).Parse(part, sink);
    }

    private static byte[] XlsbRow(int row)
    {
        var body = new byte[13];
        BinaryPrimitives.WriteUInt32LittleEndian(body, (uint)row);
        return body;
    }

    private static byte[] XlsbFormulaMember()
    {
        var body = new byte[27];
        BinaryPrimitives.WriteUInt32LittleEndian(body.AsSpan(18), 5);
        body[22] = 0x01; // PtgExp naming the A1 shared-formula master.
        return body;
    }

    private static byte[] XlsbSharedFormula()
    {
        var body = new byte[23];
        BinaryPrimitives.WriteUInt32LittleEndian(body.AsSpan(4), 1); // A1:A2.
        BinaryPrimitives.WriteUInt32LittleEndian(body.AsSpan(16), 3);
        body[20] = 0x1E; // PtgInt 42.
        body[21] = 42;
        return body;
    }

    private static byte[] Png() => Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==");

    private sealed class NonSeekableReadStream(byte[] content) : Stream
    {
        private readonly MemoryStream _inner = new(content, writable: false);

        public override bool CanRead => true;
        public override bool CanSeek => false;
        public override bool CanWrite => false;
        public override long Length => throw new NotSupportedException();
        public override long Position
        {
            get => throw new NotSupportedException();
            set => throw new NotSupportedException();
        }

        public override void Flush() { }
        public override int Read(byte[] buffer, int offset, int count) => _inner.Read(buffer, offset, count);
        public override ValueTask<int> ReadAsync(
            Memory<byte> buffer, CancellationToken cancellationToken = default) =>
            _inner.ReadAsync(buffer, cancellationToken);
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();

        protected override void Dispose(bool disposing)
        {
            if (disposing)
                _inner.Dispose();
            base.Dispose(disposing);
        }
    }
}
