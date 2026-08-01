using Xunit;

namespace Cellwright.Formulas.Tests;

/// <summary>
/// Exact lookups switch to a hash index once a range is long enough to be cached. The index
/// is only worth anything if it answers exactly what the scan answered, so every case here is
/// run twice: once on a short table that scans, once on a long one that indexes.
/// </summary>
public sealed class LookupIndexTests
{
    /// <summary>Long enough to be cached and indexed; the short variant stays on the scan path.</summary>
    private const int IndexedRows = 300;

    private const int ScannedRows = 20;

    [Theory]
    [InlineData(ScannedRows)]
    [InlineData(IndexedRows)]
    public void ExactLookupsFindTheSameRow(int rows)
    {
        Worksheet sheet = BuildTable(rows);

        Assert.Equal(70.0, sheet.Evaluate($"VLOOKUP(7,$A$1:$B${rows},2,FALSE)", "E1").GetNumber());
        Assert.Equal(7.0, sheet.Evaluate($"MATCH(7,$A$1:$A${rows},0)", "E1").GetNumber());
        Assert.Equal(70.0, sheet.Evaluate($"XLOOKUP(7,$A$1:$A${rows},$B$1:$B${rows})", "E1").GetNumber());
        Assert.Equal(7.0, sheet.Evaluate($"XMATCH(7,$A$1:$A${rows})", "E1").GetNumber());
    }

    [Theory]
    [InlineData(ScannedRows)]
    [InlineData(IndexedRows)]
    public void MissesReturnNotAvailable(int rows)
    {
        Worksheet sheet = BuildTable(rows);

        Assert.Equal(ExcelError.NotAvailable, sheet.Evaluate($"VLOOKUP(-1,$A$1:$B${rows},2,FALSE)", "E1").GetError());
        Assert.Equal(ExcelError.NotAvailable, sheet.Evaluate($"MATCH(-1,$A$1:$A${rows},0)", "E1").GetError());
        Assert.Equal("none", sheet.Evaluate($"XLOOKUP(-1,$A$1:$A${rows},$B$1:$B${rows},\"none\")", "E1").GetText());
    }

    [Theory]
    [InlineData(ScannedRows)]
    [InlineData(IndexedRows)]
    public void TextKeysIgnoreCase(int rows)
    {
        Worksheet sheet = BuildTable(rows, textKeys: true);

        Assert.Equal(70.0, sheet.Evaluate($"VLOOKUP(\"KEY7\",$A$1:$B${rows},2,FALSE)", "E1").GetNumber());
        Assert.Equal(7.0, sheet.Evaluate($"MATCH(\"key7\",$A$1:$A${rows},0)", "E1").GetNumber());
    }

    [Theory]
    [InlineData(ScannedRows)]
    [InlineData(IndexedRows)]
    public void DuplicateKeysResolveToTheFirstRow(int rows)
    {
        Worksheet sheet = BuildTable(rows);
        sheet[rows, 1] = 7;      // A second row carrying key 7, far below the first.
        sheet[rows, 2] = -1;

        Assert.Equal(70.0, sheet.Evaluate($"VLOOKUP(7,$A$1:$B${rows},2,FALSE)", "E1").GetNumber());
        Assert.Equal(7.0, sheet.Evaluate($"MATCH(7,$A$1:$A${rows},0)", "E1").GetNumber());
    }

    [Theory]
    [InlineData(ScannedRows)]
    [InlineData(IndexedRows)]
    public void DateKeysMatchTheirSerialNumber(int rows)
    {
        Worksheet sheet = BuildTable(rows);
        var date = new DateTime(2026, 7, 25);
        sheet[3, 1] = date;

        double serial = ExcelSerialDate.FromDateTime(date);
        Assert.Equal(30.0, sheet.Evaluate($"VLOOKUP({serial},$A$1:$B${rows},2,FALSE)", "E1").GetNumber());
    }

    [Theory]
    [InlineData(ScannedRows)]
    [InlineData(IndexedRows)]
    public void BlanksAndErrorsInTheKeyColumnFallBackToScanning(int rows)
    {
        Worksheet sheet = BuildTable(rows);
        sheet.Cell(5, 1).Clear();
        sheet.Cell(6, 1).SetValue(CellValue.Error(ExcelError.DivisionByZero));

        // A blank compares equal to zero in a scan, and that is what the engine has always
        // reported; the index must not quietly change the answer.
        Assert.Equal(50.0, sheet.Evaluate($"VLOOKUP(0,$A$1:$B${rows},2,FALSE)", "E1").GetNumber());
        Assert.Equal(70.0, sheet.Evaluate($"VLOOKUP(7,$A$1:$B${rows},2,FALSE)", "E1").GetNumber());
    }

    [Theory]
    [InlineData(ScannedRows)]
    [InlineData(IndexedRows)]
    public void WildcardKeysKeepScanning(int rows)
    {
        Worksheet sheet = BuildTable(rows, textKeys: true);

        Assert.Equal(70.0, sheet.Evaluate($"VLOOKUP(\"key7\",$A$1:$B${rows},2,FALSE)", "E1").GetNumber());

        // "key1?" matches key10 first; an index lookup would have missed the literal outright.
        Assert.Equal(100.0, sheet.Evaluate($"VLOOKUP(\"key1?\",$A$1:$B${rows},2,FALSE)", "E1").GetNumber());
    }

    [Theory]
    [InlineData(ScannedRows)]
    [InlineData(IndexedRows)]
    public void ApproximateLookupIsUnaffected(int rows)
    {
        Worksheet sheet = BuildTable(rows);

        // 7.5 falls between keys, so an approximate match takes the row below it.
        Assert.Equal(70.0, sheet.Evaluate($"VLOOKUP(7.5,$A$1:$B${rows},2,TRUE)", "E1").GetNumber());
    }

    [Fact]
    public void EditingTheTableIsVisibleToIndexedLookups()
    {
        Worksheet sheet = BuildTable(IndexedRows);
        Workbook book = sheet.Workbook;
        sheet.Cell("E1").SetFormula($"VLOOKUP(7,$A$1:$B${IndexedRows},2,FALSE)");
        book.Calculate();
        Assert.Equal(70.0, sheet["E1"].GetNumber());

        // The row moves to a different position: a stale index would still report the old one.
        sheet[7, 1] = 999;
        sheet[12, 1] = 7;
        sheet[12, 2] = 1_234;
        book.Calculate();

        Assert.Equal(1_234.0, sheet["E1"].GetNumber());
    }

    private static Worksheet BuildTable(int rows, bool textKeys = false)
    {
        var book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("S");
        for (int row = 1; row <= rows; row++)
        {
            if (textKeys)
                sheet[row, 1] = "key" + row;
            else
                sheet[row, 1] = row;
            sheet[row, 2] = row * 10;
        }

        return sheet;
    }
}
