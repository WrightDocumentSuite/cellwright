using System.Globalization;
using System.Xml;
using Cellwright.IO;
using Cellwright.Xml;

namespace Cellwright.Streaming;

/// <summary>
/// Reads the pictures out of a drawing part.
/// </summary>
/// <remarks>
/// Both package formats share this. A binary workbook keeps its cells in BIFF12 records but its
/// drawings in the very same DrawingML the modern format uses, so there is one reader rather
/// than two — the only difference is which reader hands it the bytes.
/// <para>
/// A drawing part holds more than pictures: shapes, charts, group shapes and SmartArt all live
/// in the same anchors. Everything that is not a picture is counted and left alone, so the
/// caller can say plainly what a regenerated part would cost.
/// </para>
/// </remarks>
internal static class DrawingPart
{
    /// <summary>What one drawing part turned out to hold.</summary>
    /// <param name="Pictures">The pictures, in the order they are stacked.</param>
    /// <param name="ChartFrames">The graphic frames holding charts: each an anchor and the relationship naming the chart part.</param>
    /// <param name="OtherShapes">
    /// How many anchors held something else. A part with any of these cannot be regenerated from
    /// the model without loss.
    /// </param>
    public readonly record struct Result(List<Picture> Pictures, List<ChartFrame> ChartFrames, int OtherShapes);

    /// <summary>A chart's place on the sheet and the relationship its frame names it by.</summary>
    /// <param name="Anchor">Where the frame sits.</param>
    /// <param name="RelationshipId">The <c>r:id</c> of the chart part.</param>
    /// <param name="Name">The frame's name, e.g. <c>Chart 1</c>.</param>
    public readonly record struct ChartFrame(PictureAnchor Anchor, string RelationshipId, string? Name);

    /// <summary>
    /// Parses a drawing part.
    /// </summary>
    /// <param name="stream">The part's bytes.</param>
    /// <param name="images">
    /// Image bytes by the relationship id an anchor names them with. A blip pointing at an id
    /// that is not here is an external or missing image, and its anchor is counted as another
    /// shape rather than dropped silently.
    /// </param>
    public static Result Read(Stream stream, IReadOnlyDictionary<string, ReadOnlyMemory<byte>> images)
    {
        using var xml = XmlReader.Create(stream, XmlDefaults.ReaderSettings);
        var pictures = new List<Picture>();
        var charts = new List<ChartFrame>();
        int others = 0;

        while (xml.Read())
        {
            if (xml.NodeType != XmlNodeType.Element)
                continue;

            PictureAnchorKind kind = xml.LocalName switch
            {
                "twoCellAnchor" => PictureAnchorKind.TwoCell,
                "oneCellAnchor" => PictureAnchorKind.OneCell,
                "absoluteAnchor" => PictureAnchorKind.Absolute,
                _ => (PictureAnchorKind)(-1),
            };

            if (kind == (PictureAnchorKind)(-1) || xml.NamespaceURI != XlsxSchema.NsSpreadsheetDrawing)
                continue;

            string? editAs = xml.GetAttribute("editAs");
            switch (ReadAnchor(xml, kind, editAs, images))
            {
                case { Picture: { } picture }:
                    pictures.Add(picture);
                    break;
                case { Chart: { } frame }:
                    charts.Add(frame);
                    break;
                default:
                    others++;
                    break;
            }
        }

        return new Result(pictures, charts, others);
    }

    /// <summary>What one anchor turned out to hold: a picture, a chart frame, or neither.</summary>
    private readonly record struct AnchorContent(Picture? Picture, ChartFrame? Chart);

    /// <summary>
    /// Reads one anchor: where it sits, then what it holds. An anchor holding neither a picture
    /// this reader can produce bytes for nor a chart frame reads as empty content.
    /// </summary>
    private static AnchorContent ReadAnchor(
        XmlReader xml, PictureAnchorKind kind, string? editAs, IReadOnlyDictionary<string, ReadOnlyMemory<byte>> images)
    {
        var anchor = new PictureAnchor { Kind = kind, MoveBehavior = MoveBehavior(editAs, kind) };
        string? embed = null;
        string? name = null;
        string? description = null;
        string? chartId = null;
        bool hidden = false;
        bool isPicture = false;
        bool isFrame = false;
        int depth = xml.Depth;

        // The extent of a one-cell or absolute anchor sits directly under it, but a picture's
        // own shape properties carry an a:ext of their own further down. Only the first, at the
        // anchor's own depth, says how big the picture is on the sheet.
        while (xml.Read() && xml.Depth > depth)
        {
            if (xml.NodeType != XmlNodeType.Element)
                continue;

            switch (xml.LocalName)
            {
                case "from" when xml.Depth == depth + 1:
                    (anchor.From, anchor.FromOffset) = ReadCellPosition(xml);
                    break;

                case "to" when xml.Depth == depth + 1:
                    (anchor.To, anchor.ToOffset) = ReadCellPosition(xml);
                    break;

                case "ext" when xml.Depth == depth + 1:
                    anchor.Extent = (Number(xml.GetAttribute("cx")), Number(xml.GetAttribute("cy")));
                    break;

                case "pos" when xml.Depth == depth + 1:
                    anchor.Position = (Number(xml.GetAttribute("x")), Number(xml.GetAttribute("y")));
                    break;

                case "pic":
                    isPicture = true;
                    break;

                case "graphicFrame":
                    isFrame = true;
                    break;

                case "cNvPr" when isPicture || isFrame:
                    name = xml.GetAttribute("name");
                    description = xml.GetAttribute("descr");
                    hidden = xml.GetAttribute("hidden") is "1" or "true";
                    break;

                case "blip" when isPicture:
                    embed = xml.GetAttribute("embed", XlsxSchema.NsDocumentRelationships)
                        ?? xml.GetAttribute("embed", XlsxSchema.NsStrictDocumentRelationships);
                    break;

                // What a frame holds is named by its chart element's r:id; a frame holding
                // something else — SmartArt, say — has no chart element and stays unmodelled.
                case "chart" when isFrame:
                    chartId = xml.GetAttribute("id", XlsxSchema.NsDocumentRelationships)
                        ?? xml.GetAttribute("id", XlsxSchema.NsStrictDocumentRelationships);
                    break;

                default:
                    break;
            }
        }

        if (isFrame && chartId is not null)
            return new AnchorContent(null, new ChartFrame(anchor, chartId, name));

        if (!isPicture || embed is null || !images.TryGetValue(embed, out ReadOnlyMemory<byte> content) || content.IsEmpty)
            return default;

        return new AnchorContent(new Picture
        {
            Content = content,
            Format = PictureCollection.DetectFormat(content.Span),
            Anchor = anchor,
            Name = name,
            Description = description,
            IsHidden = hidden,
        }, null);
    }

    /// <summary>Reads a <c>from</c> or <c>to</c>: a cell, zero-based, and an offset into it.</summary>
    private static (CellRef Cell, (long X, long Y) Offset) ReadCellPosition(XmlReader xml)
    {
        int column = 0, row = 0;
        long columnOffset = 0, rowOffset = 0;
        int depth = xml.Depth;

        if (xml.IsEmptyElement)
            return (new CellRef(1, 1), (0, 0));

        while (xml.Read() && xml.Depth > depth)
        {
            if (xml.NodeType != XmlNodeType.Element)
                continue;

            switch (xml.LocalName)
            {
                case "col":
                    column = (int)Number(xml.ReadElementContentAsString());
                    break;
                case "colOff":
                    columnOffset = Number(xml.ReadElementContentAsString());
                    break;
                case "row":
                    row = (int)Number(xml.ReadElementContentAsString());
                    break;
                case "rowOff":
                    rowOffset = Number(xml.ReadElementContentAsString());
                    break;
                default:
                    break;
            }

            // ReadElementContentAsString consumed the end tag, so the loop's own Read would
            // step past the next sibling; checking the depth again keeps it honest.
            if (xml.Depth <= depth)
                break;
        }

        return (Clamp(row + 1, column + 1), (columnOffset, rowOffset));
    }

    /// <summary>
    /// A file may anchor a picture outside the grid — a negative index, or one past the last
    /// row. Excel shows such a picture at the edge rather than refusing the file, and a cell
    /// reference the rest of the library can hold has to stay inside the sheet.
    /// </summary>
    private static CellRef Clamp(int row, int column) => new(
        Math.Clamp(row, 1, SheetLimits.MaxRow),
        Math.Clamp(column, 1, SheetLimits.MaxColumn));

    /// <summary>
    /// What <c>editAs</c> means, and what the anchor means when it says nothing. Only a
    /// two-cell anchor carries the attribute; a one-cell anchor moves without resizing because
    /// it has a size of its own, and an absolute anchor is pinned to the sheet, not the grid.
    /// </summary>
    private static PictureMoveBehavior MoveBehavior(string? editAs, PictureAnchorKind kind) => editAs switch
    {
        "absolute" => PictureMoveBehavior.Fixed,
        "oneCell" => PictureMoveBehavior.MoveOnly,
        "twoCell" => PictureMoveBehavior.MoveAndResize,
        _ => kind switch
        {
            PictureAnchorKind.TwoCell => PictureMoveBehavior.MoveAndResize,
            PictureAnchorKind.OneCell => PictureMoveBehavior.MoveOnly,
            _ => PictureMoveBehavior.Fixed,
        },
    };

    private static long Number(string? text) =>
        long.TryParse(text, NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out long value) ? value : 0;
}
