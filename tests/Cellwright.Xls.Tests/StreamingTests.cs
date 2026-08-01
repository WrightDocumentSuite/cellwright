using Xunit;

namespace Cellwright.Xls.Tests;

/// <summary>
/// Reading a legacy workbook row by row. What the streaming reader yields has to agree with
/// what the model reader builds, since the two share nothing but the file.
/// </summary>
public sealed class StreamingTests
{
    [Fact]
    public async Task RowsComeBackInOrderWithTheirValues()
    {
        Workbook original = Workbook.Create();
        Worksheet sheet = original.Sheets.Add("Data");
        sheet["A1"] = CellValue.Text("Name");
        sheet["B1"] = CellValue.Text("Amount");
        sheet["A2"] = CellValue.Text("Widget");
        sheet["B2"] = CellValue.Number(12.5);
        sheet["A4"] = CellValue.Text("Gadget");      // Row 3 is empty and must be skipped.
        sheet["C4"] = CellValue.Boolean(true);

        using var buffer = new MemoryStream();
        await XlsWorkbook.SaveAsync(original, buffer, options: null, TestContext.Current.CancellationToken);
        buffer.Position = 0;

        await using XlsReader reader = await XlsReader.OpenAsync(
            buffer, new XlsLoadOptions { LeaveOpen = true }, TestContext.Current.CancellationToken);

        Assert.Equal("Data", Assert.Single(reader.Sheets).Name);

        var rows = new List<(int Number, CellValue[] Cells)>();
        await foreach (RowData row in reader.ReadRowsAsync(0, TestContext.Current.CancellationToken))
            rows.Add((row.RowNumber, row.ToArray()));

        Assert.Equal(3, rows.Count);
        Assert.Equal([1, 2, 4], rows.Select(static r => r.Number));
        Assert.Equal("Name", rows[0].Cells[0].GetText());
        Assert.Equal(12.5, rows[1].Cells[1].GetNumber());
        Assert.Equal("Gadget", rows[2].Cells[0].GetText());
        Assert.True(rows[2].Cells[2].GetBoolean());

        // Nothing of the previous row may linger in the reused buffer.
        Assert.True(rows[2].Cells[1].IsBlank);
    }

    [Fact]
    public async Task ASheetCanBeFoundByName()
    {
        Workbook original = Workbook.Create();
        original.Sheets.Add("First")["A1"] = CellValue.Number(1);
        original.Sheets.Add("Second")["A1"] = CellValue.Number(2);

        using var buffer = new MemoryStream();
        await XlsWorkbook.SaveAsync(original, buffer, options: null, TestContext.Current.CancellationToken);
        buffer.Position = 0;

        await using XlsReader reader = await XlsReader.OpenAsync(
            buffer, new XlsLoadOptions { LeaveOpen = true }, TestContext.Current.CancellationToken);

        RowData row = await FirstRowAsync(reader, "Second");
        Assert.Equal(2, row.GetValueOrBlank(0).GetNumber());
        await Assert.ThrowsAsync<ArgumentException>(async () => await FirstRowAsync(reader, "Third"));
    }

    /// <summary>
    /// The streaming reader and the model reader see the same file from different ends, so
    /// running both over the corpus is the check that neither drifted.
    /// </summary>
    [Fact]
    public async Task StreamedRowsAgreeWithTheModel()
    {
        Assert.SkipWhen(XlsCorpus.Files.Count == 0, "No corpus available. These checks read the .xls and .xlsb files that ship with ExcelDataReader, NPOI, EPPlus and Aspose.Cells; unpack those repositories beside this one to run them.");

        int compared = 0;
        int files = 0;

        foreach (string path in XlsCorpus.Files.Take(80))
        {
            Workbook workbook;
            XlsReader reader;
            try
            {
                workbook = await XlsWorkbook.LoadAsync(path, options: null, TestContext.Current.CancellationToken);
                reader = await XlsReader.OpenAsync(path, options: null, TestContext.Current.CancellationToken);
            }
            catch (Exception exception) when (XlsCorpus.IsUnreadable(exception))
            {
                continue;
            }

            await using (reader)
            {
                files++;
                foreach (SheetInfo info in reader.Sheets)
                {
                    Worksheet? sheet = workbook.Sheets.Find(info.Name);
                    if (sheet is null)
                        continue;

                    await foreach (RowData row in reader.ReadRowsAsync(info.Index, TestContext.Current.CancellationToken))
                    {
                        for (int column = 1; column <= row.Count; column++)
                        {
                            CellValue streamed = row.GetValueOrBlank(column - 1);
                            CellValue modelled = sheet.Cell(row.RowNumber, column).Value;
                            if (streamed.IsBlank && modelled.IsBlank)
                                continue;

                            compared++;
                            Assert.True(
                                Render(modelled) == Render(streamed),
                                $"{Path.GetFileName(path)} [{info.Name}] R{row.RowNumber}C{column}: " +
                                $"model '{Render(modelled)}' vs streamed '{Render(streamed)}'");
                        }
                    }
                }
            }
        }

        Assert.True(files > 20, $"Only {files} workbooks were readable both ways.");
        Assert.True(compared > 5_000, $"Only {compared} cells were compared.");
    }

    /// <summary>Renders a value as text so two readings can be compared without arguing about types.</summary>
    private static string Render(in CellValue value) => value.Kind switch
    {
        CellValueKind.Blank => "",
        CellValueKind.Text => value.GetText(),
        CellValueKind.Number => value.GetNumber().ToString("G12", System.Globalization.CultureInfo.InvariantCulture),
        CellValueKind.Boolean => value.GetBoolean() ? "TRUE" : "FALSE",
        CellValueKind.Error => value.GetError().ToText(),
        CellValueKind.DateTime => value.GetDateTime().ToString("O"),
        CellValueKind.TimeSpan => value.GetTimeSpan().ToString(),
        _ => "",
    };

    private static async Task<RowData> FirstRowAsync(XlsReader reader, string sheetName)
    {
        await foreach (RowData row in reader.ReadRowsAsync(sheetName, TestContext.Current.CancellationToken))
            return row;

        throw new InvalidOperationException($"'{sheetName}' has no rows.");
    }
}
