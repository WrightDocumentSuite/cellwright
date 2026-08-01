using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Xunit;

namespace Cellwright.Tests.Model;

/// <summary>
/// Sheet and workbook protection, and the two password verifiers that go with them.
/// </summary>
public sealed class ProtectionTests
{
    private static async Task<Workbook> SaveAndReloadAsync(Workbook workbook)
    {
        using var stream = new MemoryStream();
        await workbook.SaveAsync(stream, new SaveOptions { LeaveOpen = true }, TestContext.Current.CancellationToken);

        stream.Position = 0;
        using (SpreadsheetDocument document = SpreadsheetDocument.Open(stream, isEditable: false))
        {
            ValidationErrorInfo? error = new OpenXmlValidator()
                .Validate(document, TestContext.Current.CancellationToken)
                .FirstOrDefault();
            Assert.True(error is null, $"{error?.Path?.XPath} {error?.Description}");
        }

        stream.Position = 0;
        return await Workbook.LoadAsync(stream, new LoadOptions { LeaveOpen = true }, TestContext.Current.CancellationToken);
    }

    private static Workbook Sample()
    {
        var book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("Locked");
        sheet["A1"] = "Total";
        return book;
    }

    [Fact]
    public async Task SheetProtectionFlagsRoundTrip()
    {
        Workbook book = Sample();
        SheetProtection protection = book.Sheets[0].Protection;
        protection.Protect();
        protection.AllowFormatCells = true;
        protection.AllowInsertRows = true;
        protection.AllowSort = true;
        protection.AllowSelectLockedCells = false;

        SheetProtection loaded = (await SaveAndReloadAsync(book)).Sheets[0].Protection;

        Assert.True(loaded.IsEnabled);
        Assert.True(loaded.AllowFormatCells);
        Assert.True(loaded.AllowInsertRows);
        Assert.True(loaded.AllowSort);
        Assert.False(loaded.AllowSelectLockedCells);
        Assert.True(loaded.AllowSelectUnlockedCells);
        Assert.False(loaded.AllowDeleteRows);
        Assert.Null(loaded.Password);
    }

    [Fact]
    public async Task SheetPasswordSurvivesAndStillVerifies()
    {
        Workbook book = Sample();
        book.Sheets[0].Protection.Protect("hunter2");

        SheetProtection loaded = (await SaveAndReloadAsync(book)).Sheets[0].Protection;

        PasswordHash password = Assert.IsType<PasswordHash>(loaded.Password);
        Assert.False(password.IsLegacy);
        Assert.Equal("SHA-512", password.Algorithm);
        Assert.True(password.Matches("hunter2"));
        Assert.False(password.Matches("hunter3"));
    }

    [Fact]
    public async Task EditableRangesRoundTrip()
    {
        Workbook book = Sample();
        SheetProtection protection = book.Sheets[0].Protection;
        protection.Protect("secret");
        protection.EditableRanges.Add(ProtectedRange.For("Input", "B2:D10"));
        protection.EditableRanges.Add(ProtectedRange.For("Notes", "F2:F10 H2:H10", "other"));

        SheetProtection loaded = (await SaveAndReloadAsync(book)).Sheets[0].Protection;

        Assert.Equal(2, loaded.EditableRanges.Count);
        Assert.Equal("Input", loaded.EditableRanges[0].Name);
        Assert.Equal(RangeRef.Parse("B2:D10"), Assert.Single(loaded.EditableRanges[0].Ranges));
        Assert.Equal(2, loaded.EditableRanges[1].Ranges.Count);
        Assert.True(loaded.EditableRanges[1].Password?.Matches("other"));
    }

    [Fact]
    public async Task WorkbookProtectionRoundTrips()
    {
        Workbook book = Sample();
        book.Protection.Protect("structure");
        book.Protection.LockWindows = true;

        WorkbookProtection loaded = (await SaveAndReloadAsync(book)).Protection;

        Assert.True(loaded.LockStructure);
        Assert.True(loaded.LockWindows);
        Assert.True(loaded.Password?.Matches("structure"));
    }

    [Fact]
    public async Task ClosedXmlSeesTheSheetAsProtected()
    {
        Workbook book = Sample();
        book.Sheets[0].Protection.Protect("hunter2");

        using var stream = new MemoryStream();
        await book.SaveAsync(stream, new SaveOptions { LeaveOpen = true }, TestContext.Current.CancellationToken);
        stream.Position = 0;

        using var interop = new XLWorkbook(stream);
        Assert.True(interop.Worksheet("Locked").Protection.IsProtected);
    }

    [Theory]
    [InlineData("test", "CBEB")]
    [InlineData("test", "cbeb")]
    [InlineData("secret", "DAA7")]
    public void LegacyChecksumIsRecognisedAndVerified(string password, string stored)
    {
        // Pre-2013 files carry this 16-bit checksum, written with whatever padding and case
        // their producer felt like; ClosedXML computes the same values.
        var legacy = new PasswordHash { Value = stored };

        Assert.True(legacy.IsLegacy);
        Assert.True(legacy.Matches(password));
        Assert.False(legacy.Matches("something else"));
    }

    [Fact]
    public void UnprotectingForgetsThePassword()
    {
        var protection = new SheetProtection();
        protection.Protect("hunter2");
        protection.Unprotect();

        Assert.False(protection.IsEnabled);
        Assert.Null(protection.Password);
    }
}
