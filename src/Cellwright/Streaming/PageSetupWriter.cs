using Cellwright.IO;
using Cellwright.Xml;

namespace Cellwright.Streaming;

/// <summary>
/// Generates the print settings of a worksheet from the model, in the order CT_Worksheet
/// prescribes. Used only once the settings have been touched; an untouched sheet gets its
/// original elements replayed.
/// </summary>
internal static class PageSetupWriter
{
    /// <summary>Writes <c>printOptions</c> through <c>colBreaks</c>.</summary>
    public static void Write(Utf8XmlWriter xml, PageSetup setup, WorksheetMarkup? markup)
    {
        WritePrintOptions(xml, setup);
        WriteMargins(xml, setup);
        WritePageSetup(xml, setup, markup);
        WriteHeaderFooter(xml, setup);
        WriteBreaks(xml, "rowBreaks"u8, setup.RowBreaks, horizontal: true);
        WriteBreaks(xml, "colBreaks"u8, setup.ColumnBreaks, horizontal: false);
    }

    /// <summary>Writes the <c>pageSetUpPr</c> child of <c>sheetPr</c>, or nothing when it is all defaults.</summary>
    public static void WritePageSetupProperties(Utf8XmlWriter xml, PageSetup setup)
    {
        if (!setup.FitToPage && setup.AutoPageBreaks)
            return;

        xml.WriteRaw("<pageSetUpPr"u8);
        if (!setup.AutoPageBreaks)
            xml.WriteRaw(" autoPageBreaks=\"0\""u8);
        if (setup.FitToPage)
            xml.WriteRaw(" fitToPage=\"1\""u8);
        xml.WriteRaw("/>"u8);
    }

    private static void WritePrintOptions(Utf8XmlWriter xml, PageSetup setup)
    {
        if (!setup.CenterHorizontally && !setup.CenterVertically && !setup.PrintGridLines && !setup.PrintHeadings)
            return;

        xml.WriteRaw("<printOptions"u8);
        WriteFlag(xml, " horizontalCentered=\"1\""u8, setup.CenterHorizontally);
        WriteFlag(xml, " verticalCentered=\"1\""u8, setup.CenterVertically);
        WriteFlag(xml, " gridLines=\"1\""u8, setup.PrintGridLines);
        WriteFlag(xml, " headings=\"1\""u8, setup.PrintHeadings);
        xml.WriteRaw("/>"u8);
    }

    private static void WriteMargins(Utf8XmlWriter xml, PageSetup setup)
    {
        // A pageSetup element without margins in front of it fails validation, so the normal
        // margins stand in whenever the sheet has anything else to say about printing.
        PageMargins margins = setup.Margins ?? PageMargins.Normal;
        xml.WriteRaw("<pageMargins left=\""u8);
        xml.WriteDouble(margins.Left);
        xml.WriteRaw("\" right=\""u8);
        xml.WriteDouble(margins.Right);
        xml.WriteRaw("\" top=\""u8);
        xml.WriteDouble(margins.Top);
        xml.WriteRaw("\" bottom=\""u8);
        xml.WriteDouble(margins.Bottom);
        xml.WriteRaw("\" header=\""u8);
        xml.WriteDouble(margins.Header);
        xml.WriteRaw("\" footer=\""u8);
        xml.WriteDouble(margins.Footer);
        xml.WriteRaw("\"/>"u8);
    }

    private static void WritePageSetup(Utf8XmlWriter xml, PageSetup setup, WorksheetMarkup? markup)
    {
        xml.WriteRaw("<pageSetup"u8);
        if (setup.PaperSize > 0)
        {
            xml.WriteRaw(" paperSize=\""u8);
            xml.WriteInt32(setup.PaperSize);
            xml.WriteRaw("\""u8);
        }

        if (setup.FitToPage)
        {
            xml.WriteRaw(" fitToWidth=\""u8);
            xml.WriteInt32(setup.FitToWidth ?? 0);
            xml.WriteRaw("\" fitToHeight=\""u8);
            xml.WriteInt32(setup.FitToHeight ?? 0);
            xml.WriteRaw("\""u8);
        }
        else if (setup.Scale != 100)
        {
            xml.WriteRaw(" scale=\""u8);
            xml.WriteInt32(setup.Scale);
            xml.WriteRaw("\""u8);
        }

        if (setup.Order == PageOrder.OverThenDown)
            xml.WriteRaw(" pageOrder=\"overThenDown\""u8);
        if (setup.Orientation != PageOrientation.Default)
        {
            xml.WriteRaw(" orientation=\""u8);
            xml.WriteRaw(setup.Orientation == PageOrientation.Portrait ? "portrait"u8 : "landscape"u8);
            xml.WriteRaw("\""u8);
        }

        WriteFlag(xml, " blackAndWhite=\"1\""u8, setup.BlackAndWhite);
        WriteFlag(xml, " draft=\"1\""u8, setup.Draft);
        if (setup.FirstPageNumber > 0)
        {
            xml.WriteRaw(" firstPageNumber=\""u8);
            xml.WriteInt32(setup.FirstPageNumber);
            xml.WriteRaw("\" useFirstPageNumber=\"1\""u8);
        }

        RawXml.WriteAttributes(xml, markup?.PageSetupAttributes);
        xml.WriteRaw("/>"u8);
    }

    private static void WriteHeaderFooter(Utf8XmlWriter xml, PageSetup setup)
    {
        bool hasHeader = setup.Header is { IsEmpty: false };
        bool hasFooter = setup.Footer is { IsEmpty: false };
        if (!hasHeader && !hasFooter)
            return;

        xml.WriteRaw("<headerFooter"u8);
        WriteFlag(xml, " differentOddEven=\"1\""u8, setup.DifferentOddEven);
        WriteFlag(xml, " differentFirst=\"1\""u8, setup.DifferentFirstPage);
        WriteFlag(xml, " scaleWithDoc=\"0\""u8, !setup.ScaleHeaderWithDocument);
        WriteFlag(xml, " alignWithMargins=\"0\""u8, !setup.AlignHeaderWithMargins);
        xml.WriteRaw(">"u8);

        WriteLine(xml, "oddHeader"u8, setup.Header?.Default);
        WriteLine(xml, "oddFooter"u8, setup.Footer?.Default);
        if (setup.DifferentOddEven)
        {
            WriteLine(xml, "evenHeader"u8, setup.Header?.Even);
            WriteLine(xml, "evenFooter"u8, setup.Footer?.Even);
        }

        if (setup.DifferentFirstPage)
        {
            WriteLine(xml, "firstHeader"u8, setup.Header?.First);
            WriteLine(xml, "firstFooter"u8, setup.Footer?.First);
        }

        xml.WriteRaw("</headerFooter>"u8);
    }

    private static void WriteLine(Utf8XmlWriter xml, scoped ReadOnlySpan<byte> name, string? text)
    {
        if (string.IsNullOrEmpty(text))
            return;

        xml.WriteRaw("<"u8);
        xml.WriteRaw(name);
        xml.WriteRaw(">"u8);
        xml.WriteText(text);
        xml.WriteRaw("</"u8);
        xml.WriteRaw(name);
        xml.WriteRaw(">"u8);
    }

    private static void WriteBreaks(
        Utf8XmlWriter xml, scoped ReadOnlySpan<byte> name, IReadOnlyList<int> breaks, bool horizontal)
    {
        if (breaks.Count == 0)
            return;

        xml.WriteRaw("<"u8);
        xml.WriteRaw(name);
        xml.WriteRaw(" count=\""u8);
        xml.WriteInt32(breaks.Count);
        xml.WriteRaw("\" manualBreakCount=\""u8);
        xml.WriteInt32(breaks.Count);
        xml.WriteRaw("\">"u8);
        foreach (int at in breaks)
        {
            // The file names the last row or column before the break and spans the other axis.
            xml.WriteRaw("<brk id=\""u8);
            xml.WriteInt32(at - 1);
            xml.WriteRaw("\" max=\""u8);
            xml.WriteInt32(horizontal ? SheetLimits.MaxColumn - 1 : SheetLimits.MaxRow - 1);
            xml.WriteRaw("\" man=\"1\"/>"u8);
        }

        xml.WriteRaw("</"u8);
        xml.WriteRaw(name);
        xml.WriteRaw(">"u8);
    }

    private static void WriteFlag(Utf8XmlWriter xml, scoped ReadOnlySpan<byte> attribute, bool value)
    {
        if (value)
            xml.WriteRaw(attribute);
    }
}
