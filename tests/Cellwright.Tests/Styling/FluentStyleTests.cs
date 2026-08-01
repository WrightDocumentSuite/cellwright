using Xunit;

namespace Cellwright.Tests.Styling;

public sealed class FluentStyleTests
{
    [Fact]
    public void CellStyleBuilders_ComposeImmutably()
    {
        CellStyle style = CellStyle.Default
            .Bold()
            .Italic()
            .FontSize(16)
            .FontName("Consolas")
            .Fill(0xFF112233)
            .Border()
            .NumberFormat("#,##0.00")
            .Center()
            .Wrap();

        Assert.True(style.Font.Bold);
        Assert.True(style.Font.Italic);
        Assert.Equal(16, style.Font.Size);
        Assert.Equal("Consolas", style.Font.Name);
        Assert.Equal(0xFF112233, style.Fill.Color.Argb);
        Assert.Equal(BorderLineStyle.Thin, style.Border.Left.Line);
        Assert.Equal("#,##0.00", style.NumberFormat);
        Assert.Equal(HorizontalAlignment.Center, style.Alignment.Horizontal);
        Assert.Equal(VerticalAlignment.Center, style.Alignment.Vertical);
        Assert.True(style.Alignment.WrapText);

        Assert.True(CellStyle.Default.IsDefault);
        Assert.False(style.IsDefault);
    }

    [Fact]
    public void CellFluent_WritesThroughToSheet()
    {
        var book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("S");

        sheet.Cell("A1").SetValue(1.5).Bold().NumberFormat("0.00");

        CellStyle applied = sheet.Cell("A1").StyleValue;
        Assert.True(applied.Font.Bold);
        Assert.Equal("0.00", applied.NumberFormat);
        Assert.True(sheet.Cell("A1").HasFormula == false);
    }

    [Fact]
    public void RangeStyle_MemoizesPerSourceStyle()
    {
        var book = Workbook.Create();
        Worksheet sheet = book.Sheets.Add("S");
        sheet.Cell("A1").SetValue(1).Italic();
        sheet.Cell("B1").SetValue(2);

        sheet.Range("A1:B1").Bold();

        Assert.True(sheet.Cell("A1").StyleValue.Font.Bold);
        Assert.True(sheet.Cell("A1").StyleValue.Font.Italic);
        Assert.True(sheet.Cell("B1").StyleValue.Font.Bold);
        Assert.False(sheet.Cell("B1").StyleValue.Font.Italic);

        // Two distinct source styles produced exactly two distinct result styles.
        Assert.NotEqual(sheet.Cell("A1").Style, sheet.Cell("B1").Style);
    }
}
