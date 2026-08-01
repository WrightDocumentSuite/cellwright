using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Xunit;

namespace Cellwright.Tests.Styles;

public sealed class StyleRoundTripTests
{
    private static readonly CellStyle HeaderStyle = new()
    {
        Font = new FontStyle { Bold = true, Size = 14, Color = ExcelColor.FromRgb(255, 255, 255), Name = "Arial" },
        Fill = FillStyle.Solid(ExcelColor.FromArgb(0xFF2278D4)),
        Border = BorderStyle.All(new BorderEdge(BorderLineStyle.Thin, ExcelColor.FromRgb(0, 0, 0))),
        Alignment = new Alignment { Horizontal = HorizontalAlignment.Center, Vertical = VerticalAlignment.Center, WrapText = true },
    };

    private static async Task<MemoryStream> WriteStyledWorkbookAsync()
    {
        var stream = new MemoryStream();
        await using (XlsxWriter writer = await XlsxWriter.CreateAsync(
            stream, new XlsxWriterOptions { LeaveOpen = true }, TestContext.Current.CancellationToken))
        {
            StyleId header = writer.RegisterStyle(HeaderStyle);
            StyleId money = writer.RegisterStyle(new CellStyle { NumberFormat = "#,##0.00" });
            StyleId percent = writer.RegisterStyle(new CellStyle { NumberFormat = "0.00%" });
            StyleId isoDate = writer.RegisterStyle(new CellStyle { NumberFormat = "yyyy-mm-dd" });

            XlsxSheetWriter sheet = await writer.BeginSheetAsync("Styled", cancellationToken: TestContext.Current.CancellationToken);
            sheet.WriteRow(["Products", "Value"], header);
            sheet.WriteRow(
                ["Total", 1234.5, 0.25, new DateTime(2026, 7, 24)],
                [StyleId.Default, money, percent, isoDate]);
            sheet.WriteRow([CellValue.Blank], header); // Styled blank keeps the fill.
        }

        stream.Position = 0;
        return stream;
    }

    [Fact]
    public async Task StyledWorkbook_PassesOpenXmlValidation()
    {
        using MemoryStream stream = await WriteStyledWorkbookAsync();
        using SpreadsheetDocument document = SpreadsheetDocument.Open(stream, isEditable: false);
        var errors = new OpenXmlValidator().Validate(document, TestContext.Current.CancellationToken).ToList();
        Assert.Empty(errors);
    }

    [Fact]
    public async Task StyledWorkbook_ReadableByClosedXml()
    {
        using MemoryStream stream = await WriteStyledWorkbookAsync();
        using var workbook = new XLWorkbook(stream);
        IXLWorksheet sheet = workbook.Worksheet("Styled");

        IXLCell header = sheet.Cell(1, 1);
        Assert.True(header.Style.Font.Bold);
        Assert.Equal(14, header.Style.Font.FontSize);
        Assert.Equal("Arial", header.Style.Font.FontName);
        Assert.Equal(XLColor.FromArgb(0x22, 0x78, 0xD4), header.Style.Fill.BackgroundColor);
        Assert.Equal(XLAlignmentHorizontalValues.Center, header.Style.Alignment.Horizontal);
        Assert.Equal(XLAlignmentVerticalValues.Center, header.Style.Alignment.Vertical);
        Assert.True(header.Style.Alignment.WrapText);
        Assert.Equal(XLBorderStyleValues.Thin, header.Style.Border.LeftBorder);

        Assert.Equal(4, sheet.Cell(2, 2).Style.NumberFormat.NumberFormatId); // "#,##0.00" is builtin id 4.
        Assert.Equal(1234.5, sheet.Cell(2, 2).GetDouble());
        Assert.Equal(10, sheet.Cell(2, 3).Style.NumberFormat.NumberFormatId); // "0.00%" is builtin id 10.
        Assert.Equal(new DateTime(2026, 7, 24), sheet.Cell(2, 4).GetDateTime());

        // The styled blank cell keeps its fill.
        Assert.Equal(XLColor.FromArgb(0x22, 0x78, 0xD4), sheet.Cell(3, 1).Style.Fill.BackgroundColor);
    }

    [Fact]
    public async Task RegisterStyle_InternsEqualRecords()
    {
        using var stream = new MemoryStream();
        await using XlsxWriter writer = await XlsxWriter.CreateAsync(
            stream, new XlsxWriterOptions { LeaveOpen = true }, TestContext.Current.CancellationToken);

        StyleId first = writer.RegisterStyle(new CellStyle { Font = new FontStyle { Bold = true } });
        StyleId second = writer.RegisterStyle(new CellStyle { Font = new FontStyle { Bold = true } });
        StyleId different = writer.RegisterStyle(new CellStyle { Font = new FontStyle { Italic = true } });

        Assert.Equal(first, second);
        Assert.NotEqual(first, different);
    }

    [Fact]
    public async Task StylesWrittenByClosedXml_ParseIntoCellStyles()
    {
        using var stream = new MemoryStream();
        using (var workbook = new XLWorkbook())
        {
            IXLWorksheet sheet = workbook.AddWorksheet("S");
            IXLCell cell = sheet.Cell(1, 1);
            cell.Value = "styled";
            cell.Style.Font.SetBold(true).Font.SetFontSize(16).Font.SetFontColor(XLColor.FromArgb(0xCC, 0x00, 0x00));
            cell.Style.Fill.SetBackgroundColor(XLColor.FromArgb(0x11, 0x22, 0x33));
            cell.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Right);
            cell.Style.NumberFormat.SetFormat("0.000");
            workbook.SaveAs(stream);
        }

        stream.Position = 0;
        await using XlsxReader reader = await XlsxReader.OpenAsync(
            stream, new XlsxReaderOptions { LeaveOpen = true }, TestContext.Current.CancellationToken);

        // Force styles parsing through a row read, then inspect the parsed part.
        await foreach (RowData _ in reader.ReadRowsAsync(0, TestContext.Current.CancellationToken))
        {
        }

        Cellwright.Styles.StylesPart styles = await reader.GetStylesAsync(TestContext.Current.CancellationToken);
        Assert.Contains(styles.CellFormats, s =>
            s.Font.Bold &&
            s.Font.Size == 16 &&
            !s.Font.Color.IsAuto && s.Font.Color.Argb == 0xFFCC0000 &&
            s.Fill.IsSolid && s.Fill.Color.Argb == 0xFF112233 &&
            s.Alignment.Horizontal == HorizontalAlignment.Right &&
            s.NumberFormat == "0.000");
    }
}
