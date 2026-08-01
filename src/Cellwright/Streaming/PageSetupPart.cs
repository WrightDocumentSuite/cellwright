using System.Globalization;
using System.Xml;
using Cellwright.IO;
using Cellwright.Xml;

namespace Cellwright.Streaming;

/// <summary>
/// Reading and writing the print settings of a worksheet: <c>printOptions</c>,
/// <c>pageMargins</c>, <c>pageSetup</c>, <c>headerFooter</c> and the manual page breaks. The
/// attributes of <c>pageSetup</c> the model does not represent — chiefly the link to the
/// printer settings part — travel in the markup so that regenerating keeps them.
/// </summary>
internal static class PageSetupPart
{
    /// <summary>Reads one captured element into the setup, leaving the model unmodified.</summary>
    public static void Read(string element, PageSetup setup, WorksheetMarkup markup)
    {
        using var xml = XmlReader.Create(new StringReader(element), XmlDefaults.ReaderSettings);
        while (xml.Read())
        {
            if (xml.NodeType != XmlNodeType.Element)
                continue;

            switch (xml.LocalName)
            {
                case "printOptions":
                    ReadPrintOptions(xml, setup);
                    return;
                case "pageMargins":
                    ReadMargins(xml, setup);
                    return;
                case "pageSetup":
                    ReadPageSetup(xml, setup, markup);
                    return;
                case "headerFooter":
                    ReadHeaderFooter(xml, setup);
                    return;
                case "rowBreaks":
                    setup.SetBreaksFromFile(ReadBreaks(xml, "rowBreaks"), [.. setup.ColumnBreaks]);
                    return;
                case "colBreaks":
                    setup.SetBreaksFromFile([.. setup.RowBreaks], ReadBreaks(xml, "colBreaks"));
                    return;
                default:
                    return;
            }
        }
    }

    /// <summary>Reads the <c>pageSetUpPr</c> child of <c>sheetPr</c>.</summary>
    public static void ReadPageSetupProperties(XmlReader xml, PageSetup setup)
    {
        bool fitToPage = IsTrue(xml.GetAttribute("fitToPage"));
        bool autoBreaks = xml.GetAttribute("autoPageBreaks") is not { } value || IsTrue(value);
        setup.SetFromFile(target =>
        {
            target.AutoPageBreaks = autoBreaks;

            // Excel writes the two counts whether or not they are in force; without the flag
            // they mean nothing, so the model only learns them when fitting is on.
            if (fitToPage)
                target.FitToPageRequested = true;
        });
    }

    private static void ReadPrintOptions(XmlReader xml, PageSetup setup) => setup.SetFromFile(target =>
    {
        target.CenterHorizontally = IsTrue(xml.GetAttribute("horizontalCentered"));
        target.CenterVertically = IsTrue(xml.GetAttribute("verticalCentered"));
        target.PrintGridLines = IsTrue(xml.GetAttribute("gridLines"));
        target.PrintHeadings = IsTrue(xml.GetAttribute("headings"));
    });

    private static void ReadMargins(XmlReader xml, PageSetup setup)
    {
        var margins = new PageMargins(
            Number(xml.GetAttribute("left"), 0.7),
            Number(xml.GetAttribute("right"), 0.7),
            Number(xml.GetAttribute("top"), 0.75),
            Number(xml.GetAttribute("bottom"), 0.75),
            Number(xml.GetAttribute("header"), 0.3),
            Number(xml.GetAttribute("footer"), 0.3));
        setup.SetFromFile(target => target.Margins = margins);
    }

    private static void ReadPageSetup(XmlReader xml, PageSetup setup, WorksheetMarkup markup)
    {
        PageOrientation orientation = xml.GetAttribute("orientation") switch
        {
            "portrait" => PageOrientation.Portrait,
            "landscape" => PageOrientation.Landscape,
            _ => PageOrientation.Default,
        };
        int paper = Integer(xml.GetAttribute("paperSize"), 0);
        int scale = Integer(xml.GetAttribute("scale"), 100);
        int fitWidth = Integer(xml.GetAttribute("fitToWidth"), 1);
        int fitHeight = Integer(xml.GetAttribute("fitToHeight"), 1);
        bool useFirstPage = IsTrue(xml.GetAttribute("useFirstPageNumber"));
        int firstPage = Integer(xml.GetAttribute("firstPageNumber"), 1);
        bool blackAndWhite = IsTrue(xml.GetAttribute("blackAndWhite"));
        bool draft = IsTrue(xml.GetAttribute("draft"));
        PageOrder order = xml.GetAttribute("pageOrder") == "overThenDown" ? PageOrder.OverThenDown : PageOrder.DownThenOver;

        markup.PageSetupAttributes = RawXml.ReadAttributes(
            xml, "orientation", "paperSize", "scale", "fitToWidth", "fitToHeight",
            "useFirstPageNumber", "firstPageNumber", "blackAndWhite", "draft", "pageOrder");

        setup.SetFromFile(target =>
        {
            target.Orientation = orientation;
            target.PaperSize = paper;
            target.Scale = scale;
            target.BlackAndWhite = blackAndWhite;
            target.Draft = draft;
            target.Order = order;
            target.FirstPageNumber = useFirstPage ? firstPage : 0;
            if (target.FitToPageRequested)
            {
                target.FitToWidth = fitWidth == 0 ? null : fitWidth;
                target.FitToHeight = fitHeight == 0 ? null : fitHeight;
            }
        });
    }

    private static void ReadHeaderFooter(XmlReader xml, PageSetup setup)
    {
        bool differentOddEven = IsTrue(xml.GetAttribute("differentOddEven"));
        bool differentFirst = IsTrue(xml.GetAttribute("differentFirst"));
        bool scaleWithDocument = xml.GetAttribute("scaleWithDoc") is not { } scale || IsTrue(scale);
        bool alignWithMargins = xml.GetAttribute("alignWithMargins") is not { } align || IsTrue(align);

        string? oddHeader = null, evenHeader = null, firstHeader = null;
        string? oddFooter = null, evenFooter = null, firstFooter = null;
        if (!xml.IsEmptyElement)
        {
            bool advance = true;
            while (true)
            {
                if (advance && !xml.Read())
                    break;
                advance = true;

                if (xml.NodeType == XmlNodeType.EndElement && xml.LocalName == "headerFooter")
                    break;
                if (xml.NodeType != XmlNodeType.Element)
                    continue;

                string name = xml.LocalName;

                // Reading the content lands on the node after the element, so the next line is
                // already current; advancing again would swallow it.
                string content = xml.ReadElementContentAsString();
                advance = false;
                switch (name)
                {
                    case "oddHeader":
                        oddHeader = content;
                        break;
                    case "evenHeader":
                        evenHeader = content;
                        break;
                    case "firstHeader":
                        firstHeader = content;
                        break;
                    case "oddFooter":
                        oddFooter = content;
                        break;
                    case "evenFooter":
                        evenFooter = content;
                        break;
                    case "firstFooter":
                        firstFooter = content;
                        break;
                    default:
                        break;
                }
            }
        }

        setup.SetFromFile(target =>
        {
            target.DifferentOddEven = differentOddEven;
            target.DifferentFirstPage = differentFirst;
            target.ScaleHeaderWithDocument = scaleWithDocument;
            target.AlignHeaderWithMargins = alignWithMargins;
            target.Header = Lines(oddHeader, evenHeader, firstHeader);
            target.Footer = Lines(oddFooter, evenFooter, firstFooter);
        });
    }

    private static HeaderFooterText? Lines(string? standard, string? even, string? first)
    {
        var text = new HeaderFooterText { Default = standard, Even = even, First = first };
        return text.IsEmpty ? null : text;
    }

    private static List<int>? ReadBreaks(XmlReader xml, string container)
    {
        if (xml.IsEmptyElement)
            return null;

        List<int> breaks = [];
        while (xml.Read())
        {
            if (xml.NodeType == XmlNodeType.EndElement && xml.LocalName == container)
                break;
            if (xml.NodeType == XmlNodeType.Element && xml.LocalName == "brk" &&
                int.TryParse(xml.GetAttribute("id"), NumberStyles.None, CultureInfo.InvariantCulture, out int id) && id > 0)
            {
                // The file counts from zero and names the last row before the break.
                breaks.Add(id + 1);
            }
        }

        breaks.Sort();
        return breaks.Count == 0 ? null : breaks;
    }

    private static double Number(string? text, double fallback) =>
        double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out double value) ? value : fallback;

    private static int Integer(string? text, int fallback) =>
        int.TryParse(text, NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out int value) ? value : fallback;

    private static bool IsTrue(string? value) => value is "1" or "true" or "TRUE";
}
