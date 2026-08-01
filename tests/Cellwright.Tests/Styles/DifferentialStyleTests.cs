using Cellwright.Styles;
using Xunit;

namespace Cellwright.Tests.Styles;

/// <summary>
/// Laying a differential format over a cell's own style — what a conditional format does to the
/// cell it paints.
/// </summary>
public sealed class DifferentialStyleTests
{
    private static readonly CellStyle Base = CellStyle.Default with
    {
        Font = FontStyle.Default with { Name = "Georgia", Size = 14, Italic = true },
        Fill = FillStyle.Patterned(FillPattern.LightGrid, ExcelColor.FromRgb(0x11, 0x22, 0x33)),
        Border = BorderStyle.None,
        Alignment = Alignment.Default with { WrapText = true, Indent = 3 },
        NumberFormat = "#,##0.00",
    };

    [Fact]
    public void AnEmptyFormatChangesNothing()
    {
        Assert.Same(Base, DifferentialStyle.Empty.ApplyTo(Base));
    }

    [Fact]
    public void WhatTheFormatDoesNotSetIsKept()
    {
        var format = new DifferentialStyle { Bold = true };

        CellStyle result = format.ApplyTo(Base);

        Assert.True(result.Font.Bold);
        Assert.Equal("Georgia", result.Font.Name);
        Assert.Equal(14, result.Font.Size);
        Assert.True(result.Font.Italic);
        Assert.Equal("#,##0.00", result.NumberFormat);
        Assert.Equal(Base.Fill, result.Fill);
        Assert.Equal(Base.Alignment, result.Alignment);
    }

    [Fact]
    public void AFormatCanTurnAPropertyOffAsWellAsOn()
    {
        var format = new DifferentialStyle { Italic = false };

        Assert.False(format.ApplyTo(Base).Font.Italic);
    }

    [Fact]
    public void AColourReplacesTheCellsOwn()
    {
        var format = new DifferentialStyle
        {
            FontColor = ExcelColor.FromRgb(0x9C, 0x00, 0x06),
            FillColor = ExcelColor.FromRgb(0xFF, 0xC7, 0xCE),
        };

        CellStyle result = format.ApplyTo(Base);

        Assert.Equal(ExcelColor.FromRgb(0x9C, 0x00, 0x06), result.Font.Color);

        // The format names a solid colour, so a patterned fill goes while the rule holds.
        Assert.Equal(FillPattern.Solid, result.Fill.Pattern);
        Assert.Equal(ExcelColor.FromRgb(0xFF, 0xC7, 0xCE), result.Fill.Color);
    }

    [Fact]
    public void AnAutomaticColourLeavesTheCellsAlone()
    {
        CellStyle coloured = Base with
        {
            Font = Base.Font with { Color = ExcelColor.FromRgb(0x00, 0x00, 0xFF) },
        };

        CellStyle result = new DifferentialStyle { Bold = true }.ApplyTo(coloured);

        Assert.Equal(ExcelColor.FromRgb(0x00, 0x00, 0xFF), result.Font.Color);
        Assert.Equal(coloured.Fill, result.Fill);
    }

    [Fact]
    public void UnderlineIsSingleWhenSetAndNoneWhenCleared()
    {
        CellStyle underlined = Base with
        {
            Font = Base.Font with { Underline = UnderlineStyle.Double },
        };

        Assert.Equal(UnderlineStyle.Single, new DifferentialStyle { Underline = true }.ApplyTo(Base).Font.Underline);
        Assert.Equal(UnderlineStyle.None, new DifferentialStyle { Underline = false }.ApplyTo(underlined).Font.Underline);
    }

    [Fact]
    public void ANumberFormatReplacesTheCellsOwn()
    {
        var format = new DifferentialStyle { NumberFormat = "0.00%" };

        Assert.Equal("0.00%", format.ApplyTo(Base).NumberFormat);
    }

    [Fact]
    public void TheBuiltInFormatsPaintAsExcelPaintsThem()
    {
        CellStyle bad = DifferentialStyle.Bad.ApplyTo(CellStyle.Default);

        Assert.Equal(ExcelColor.FromRgb(0x9C, 0x00, 0x06), bad.Font.Color);
        Assert.Equal(ExcelColor.FromRgb(0xFF, 0xC7, 0xCE), bad.Fill.Color);
        Assert.Equal(FillPattern.Solid, bad.Fill.Pattern);
    }

    [Fact]
    public void AlignmentAndProtectionAreNeverTouched()
    {
        // A dxf can carry both, but the model does not express them, so a format must not
        // pretend to reset what it cannot describe.
        var format = new DifferentialStyle { Bold = true, FillColor = ExcelColor.FromRgb(1, 2, 3) };

        CellStyle result = format.ApplyTo(Base);

        Assert.True(result.Alignment.WrapText);
        Assert.Equal(3, result.Alignment.Indent);
        Assert.Equal(Base.Protection, result.Protection);
    }
}
