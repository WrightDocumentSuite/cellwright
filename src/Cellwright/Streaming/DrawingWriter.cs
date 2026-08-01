using Cellwright.Xml;

namespace Cellwright.Streaming;

/// <summary>
/// Writes a drawing part: one anchor per picture, each pointing at an image part by
/// relationship (ECMA-376 Part 1 §20.5).
/// </summary>
internal static class DrawingWriter
{
    /// <summary>A chart to place: its anchor and the relationship naming its part.</summary>
    /// <param name="Anchor">Where the frame sits.</param>
    /// <param name="RelationshipId">The <c>r:id</c> of the chart part.</param>
    /// <param name="Name">The frame's name, e.g. <c>Chart 1</c>.</param>
    public readonly record struct ChartFrame(PictureAnchor Anchor, string RelationshipId, string? Name);

    /// <summary>Writes the part.</summary>
    /// <param name="xml">Where to write.</param>
    /// <param name="pictures">The sheet's pictures, back to front.</param>
    /// <param name="relationshipIds">Relationship id of each picture's image part, in the same order.</param>
    /// <param name="charts">The sheet's charts, drawn after the pictures.</param>
    public static void Write(
        Utf8XmlWriter xml, IReadOnlyList<Picture> pictures, IReadOnlyList<string> relationshipIds,
        IReadOnlyList<ChartFrame>? charts = null)
    {
        xml.WriteDeclaration();
        xml.WriteRaw(
            "<xdr:wsDr xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\""u8);
        xml.WriteRaw(" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\""u8);
        xml.WriteRaw(" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"u8);

        for (int i = 0; i < pictures.Count && i < relationshipIds.Count; i++)
            WriteAnchor(xml, pictures[i], relationshipIds[i], i + 1);

        for (int i = 0; i < charts?.Count; i++)
            WriteChartFrame(xml, charts[i], pictures.Count + i + 1);

        xml.WriteRaw("</xdr:wsDr>"u8);
    }

    /// <summary>Writes one chart's frame inside a two-cell anchor.</summary>
    private static void WriteChartFrame(Utf8XmlWriter xml, in ChartFrame frame, int id)
    {
        PictureAnchor anchor = frame.Anchor;
        xml.WriteRaw("<xdr:twoCellAnchor>"u8);
        WriteCellPosition(xml, "from"u8, anchor.From, anchor.FromOffset);
        WriteCellPosition(xml, "to"u8, anchor.To, anchor.ToOffset);

        xml.WriteRaw("<xdr:graphicFrame macro=\"\"><xdr:nvGraphicFramePr><xdr:cNvPr id=\""u8);
        xml.WriteInt32(id);
        xml.WriteRaw("\" name=\""u8);
        xml.WriteAttributeText(frame.Name ?? $"Chart {id}");
        xml.WriteRaw("\"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>"u8);
        xml.WriteRaw("<xdr:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"0\" cy=\"0\"/></xdr:xfrm>"u8);
        xml.WriteRaw("<a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/chart\">"u8);
        xml.WriteRaw("<c:chart xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" r:id=\""u8);
        xml.WriteAttributeText(frame.RelationshipId);
        xml.WriteRaw("\"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>"u8);
    }

    private static void WriteAnchor(Utf8XmlWriter xml, Picture picture, string relationshipId, int id)
    {
        PictureAnchor anchor = picture.Anchor;
        switch (anchor.Kind)
        {
            case PictureAnchorKind.TwoCell:
                xml.WriteRaw("<xdr:twoCellAnchor editAs=\""u8);
                xml.WriteRaw(EditAs(anchor.MoveBehavior));
                xml.WriteRaw("\">"u8);
                WriteCellPosition(xml, "from"u8, anchor.From, anchor.FromOffset);
                WriteCellPosition(xml, "to"u8, anchor.To, anchor.ToOffset);
                WritePicture(xml, picture, relationshipId, id);
                xml.WriteRaw("</xdr:twoCellAnchor>"u8);
                break;

            case PictureAnchorKind.OneCell:
                xml.WriteRaw("<xdr:oneCellAnchor>"u8);
                WriteCellPosition(xml, "from"u8, anchor.From, anchor.FromOffset);
                WriteExtent(xml, anchor.Extent);
                WritePicture(xml, picture, relationshipId, id);
                xml.WriteRaw("</xdr:oneCellAnchor>"u8);
                break;

            default:
                xml.WriteRaw("<xdr:absoluteAnchor><xdr:pos x=\""u8);
                xml.WriteInt64(anchor.Position.X);
                xml.WriteRaw("\" y=\""u8);
                xml.WriteInt64(anchor.Position.Y);
                xml.WriteRaw("\"/>"u8);
                WriteExtent(xml, anchor.Extent);
                WritePicture(xml, picture, relationshipId, id);
                xml.WriteRaw("</xdr:absoluteAnchor>"u8);
                break;
        }
    }

    private static void WriteCellPosition(Utf8XmlWriter xml, ReadOnlySpan<byte> element, CellRef cell, (long X, long Y) offset)
    {
        xml.WriteRaw("<xdr:"u8);
        xml.WriteRaw(element);
        xml.WriteRaw("><xdr:col>"u8);
        xml.WriteInt32(cell.Column - 1);
        xml.WriteRaw("</xdr:col><xdr:colOff>"u8);
        xml.WriteInt64(offset.X);
        xml.WriteRaw("</xdr:colOff><xdr:row>"u8);
        xml.WriteInt32(cell.Row - 1);
        xml.WriteRaw("</xdr:row><xdr:rowOff>"u8);
        xml.WriteInt64(offset.Y);
        xml.WriteRaw("</xdr:rowOff></xdr:"u8);
        xml.WriteRaw(element);
        xml.WriteRaw(">"u8);
    }

    private static void WriteExtent(Utf8XmlWriter xml, (long Width, long Height) extent)
    {
        xml.WriteRaw("<xdr:ext cx=\""u8);
        xml.WriteInt64(extent.Width);
        xml.WriteRaw("\" cy=\""u8);
        xml.WriteInt64(extent.Height);
        xml.WriteRaw("\"/>"u8);
    }

    /// <summary>
    /// The picture itself. The shape transform is written at the origin with the anchor's
    /// extent: for a two-cell anchor the cells decide the size and Excel recomputes it anyway,
    /// and for the other two the extent beside the anchor is the same number.
    /// </summary>
    private static void WritePicture(Utf8XmlWriter xml, Picture picture, string relationshipId, int id)
    {
        xml.WriteRaw("<xdr:pic><xdr:nvPicPr><xdr:cNvPr id=\""u8);
        xml.WriteInt32(id);
        xml.WriteRaw("\" name=\""u8);
        xml.WriteAttributeText(picture.Name ?? $"Picture {id}");
        if (picture.Description is { Length: > 0 } description)
        {
            xml.WriteRaw("\" descr=\""u8);
            xml.WriteAttributeText(description);
        }

        if (picture.IsHidden)
            xml.WriteRaw("\" hidden=\"1"u8);

        xml.WriteRaw("\"/><xdr:cNvPicPr><a:picLocks noChangeAspect=\"1\"/></xdr:cNvPicPr></xdr:nvPicPr>"u8);

        xml.WriteRaw("<xdr:blipFill><a:blip r:embed=\""u8);
        xml.WriteAttributeText(relationshipId);
        xml.WriteRaw("\"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>"u8);

        xml.WriteRaw("<xdr:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/>"u8);
        xml.WriteRaw("<a:ext cx=\""u8);
        xml.WriteInt64(picture.Anchor.Extent.Width);
        xml.WriteRaw("\" cy=\""u8);
        xml.WriteInt64(picture.Anchor.Extent.Height);
        xml.WriteRaw("\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></xdr:spPr>"u8);

        xml.WriteRaw("</xdr:pic><xdr:clientData/>"u8);
    }

    private static ReadOnlySpan<byte> EditAs(PictureMoveBehavior behavior) => behavior switch
    {
        PictureMoveBehavior.Fixed => "absolute"u8,
        PictureMoveBehavior.MoveOnly => "oneCell"u8,
        _ => "twoCell"u8,
    };

    /// <summary>Content type of an image part, which the package declares by extension.</summary>
    public static string ContentType(PictureFormat format) => format switch
    {
        PictureFormat.Png => "image/png",
        PictureFormat.Jpeg => "image/jpeg",
        PictureFormat.Gif => "image/gif",
        PictureFormat.Bmp => "image/bmp",
        PictureFormat.Tiff => "image/tiff",
        PictureFormat.Emf => "image/x-emf",
        PictureFormat.Wmf => "image/x-wmf",
        _ => "application/octet-stream",
    };

    /// <summary>File extension an image part is given, which is how its content type is found.</summary>
    public static string Extension(PictureFormat format) => format switch
    {
        PictureFormat.Png => "png",
        PictureFormat.Jpeg => "jpeg",
        PictureFormat.Gif => "gif",
        PictureFormat.Bmp => "bmp",
        PictureFormat.Tiff => "tiff",
        PictureFormat.Emf => "emf",
        PictureFormat.Wmf => "wmf",
        _ => "bin",
    };
}
