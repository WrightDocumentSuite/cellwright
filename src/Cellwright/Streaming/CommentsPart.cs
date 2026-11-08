using System.Globalization;
using System.Xml;
using Cellwright.Xml;

namespace Cellwright.Streaming;

/// <summary>
/// Reading and writing the notes part of a worksheet, plus the legacy VML drawing Excel
/// insists on: without a shape for every note, the file opens but the notes are invisible.
/// </summary>
internal static class CommentsPart
{
    private const string RevisionNamespace = "http://schemas.microsoft.com/office/spreadsheetml/2014/revision";

    /// <summary>
    /// Parses a comments part into flat notes; rich-text runs are concatenated. Placeholder
    /// candidates stay in the result until the loader can reconcile them with live threads.
    /// </summary>
    public static List<CellComment> Read(Stream stream)
    {
        var authors = new List<string>();
        var comments = new List<CellComment>();
        using var xml = XmlReader.Create(stream, XmlDefaults.ReaderSettings);
        bool advance = true;
        while (true)
        {
            if (advance && !xml.Read())
                break;
            advance = true;

            if (xml.NodeType != XmlNodeType.Element)
                continue;

            if (xml.LocalName == "author")
            {
                authors.Add(xml.ReadElementContentAsString());
                advance = false;
            }
            else if (xml.LocalName == "comment")
            {
                if (ReadComment(xml, authors) is { } comment)
                    comments.Add(comment);
                advance = false;
            }
        }

        return comments;
    }

    private static CellComment? ReadComment(XmlReader xml, List<string> authors)
    {
        string? reference = xml.GetAttribute("ref");
        string? revisionUid = xml.GetAttribute("uid", RevisionNamespace);
        int authorId = int.TryParse(xml.GetAttribute("authorId"), NumberStyles.None, CultureInfo.InvariantCulture, out int id)
            ? id
            : -1;

        using XmlReader subtree = xml.ReadSubtree();
        subtree.Read();
        var text = new System.Text.StringBuilder();
        bool advance = true;
        while (true)
        {
            if (advance && !subtree.Read())
                break;
            advance = true;

            if (subtree.NodeType == XmlNodeType.Element && subtree.LocalName == "t")
            {
                text.Append(subtree.ReadElementContentAsString());
                advance = false;
            }
        }

        if (!CellRef.TryParse(reference, null, out CellRef cell))
            return null;

        string? author = (uint)authorId < (uint)authors.Count ? authors[authorId] : null;

        return new CellComment
        {
            Cell = cell,
            Text = OoxmlText.Decode(text.ToString()),
            Author = author,
            RevisionUid = revisionUid,
        };
    }

    /// <summary>Writes the comments part for a set of notes.</summary>
    public static void Write(Utf8XmlWriter xml, IReadOnlyList<CellComment> comments)
    {
        List<string> authors = [];
        foreach (CellComment comment in comments)
        {
            string author = comment.Author ?? "";
            if (!authors.Contains(author, StringComparer.Ordinal))
                authors.Add(author);
        }

        bool hasRevisionUids = comments.Any(static comment => comment.RevisionUid is { Length: > 0 });
        xml.WriteDeclaration();
        xml.WriteRaw("<comments xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\""u8);
        if (hasRevisionUids)
        {
            xml.WriteRaw(" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\""u8);
            xml.WriteRaw(" xmlns:xr=\"http://schemas.microsoft.com/office/spreadsheetml/2014/revision\""u8);
            xml.WriteRaw(" mc:Ignorable=\"xr\""u8);
        }
        xml.WriteRaw("><authors>"u8);
        foreach (string author in authors)
        {
            xml.WriteRaw("<author>"u8);
            xml.WriteText(author);
            xml.WriteRaw("</author>"u8);
        }

        xml.WriteRaw("</authors><commentList>"u8);
        foreach (CellComment comment in comments)
        {
            xml.WriteRaw("<comment ref=\""u8);
            xml.WriteFormattable(comment.Cell);
            xml.WriteRaw("\" authorId=\""u8);
            xml.WriteInt32(authors.IndexOf(comment.Author ?? ""));
            if (comment.RevisionUid is { Length: > 0 } revisionUid)
            {
                xml.WriteRaw("\" xr:uid=\""u8);
                xml.WriteAttributeText(revisionUid);
            }
            // A single run rather than a bare t element: the schema allows both, but readers
            // that expect Excel's own output only look at runs.
            xml.WriteRaw("\"><text><r><t xml:space=\"preserve\">"u8);
            xml.WriteText(comment.Text);
            xml.WriteRaw("</t></r></text></comment>"u8);
        }

        xml.WriteRaw("</commentList></comments>"u8);
    }

    /// <summary>
    /// Writes the legacy VML that gives each note a box: without a shape per note the file
    /// opens but the notes are invisible. A file whose notes were never touched keeps its own
    /// VML instead, geometry and all.
    /// </summary>
    /// <param name="xml">Part being written.</param>
    /// <param name="comments">Notes of one sheet.</param>
    /// <param name="drawingIndex">
    /// One-based position of this drawing among the workbook's. Shape identifiers are handed
    /// out in blocks of 1024, one block per drawing, and the block is what <c>o:idmap</c>
    /// declares — two sheets claiming the same block is a file Excel repairs.
    /// </param>
    /// <param name="leftovers">
    /// Non-note shapes of the original part — form controls, mostly — written back verbatim
    /// so that editing a comment does not cost the sheet its check boxes.
    /// </param>
    public static void WriteVml(
        Utf8XmlWriter xml, IReadOnlyList<CellComment> comments, int drawingIndex, VmlLeftovers? leftovers = null)
    {
        // Note ids start after whatever the kept shapes already use; a collision would make
        // Excel repair the file.
        int shapeId = (drawingIndex * 1024) + 1;
        if (leftovers is not null && leftovers.MaxShapeId >= shapeId)
            shapeId = leftovers.MaxShapeId + 1;

        xml.WriteRaw("<xml xmlns:v=\"urn:schemas-microsoft-com:vml\""u8);
        xml.WriteRaw(" xmlns:o=\"urn:schemas-microsoft-com:office:office\""u8);
        xml.WriteRaw(" xmlns:x=\"urn:schemas-microsoft-com:office:excel\">"u8);
        xml.WriteRaw("<o:shapelayout v:ext=\"edit\"><o:idmap v:ext=\"edit\" data=\""u8);
        WriteIdMapBlocks(xml, drawingIndex, shapeId, comments.Count, leftovers);
        xml.WriteRaw("\"/></o:shapelayout>"u8);
        xml.WriteRaw("<v:shapetype id=\"_x0000_t202\" coordsize=\"21600,21600\" o:spt=\"202\""u8);
        xml.WriteRaw(" path=\"m,l,21600r21600,l21600,xe\"><v:stroke joinstyle=\"miter\"/>"u8);
        xml.WriteRaw("<v:path gradientshapeok=\"t\" o:connecttype=\"rect\"/></v:shapetype>"u8);

        if (leftovers is not null)
        {
            foreach (string shapeType in leftovers.ShapeTypes)
                RawXml.Write(xml, shapeType);
            foreach (string shape in leftovers.Shapes)
                RawXml.Write(xml, shape);
        }

        foreach (CellComment comment in comments)
        {
            xml.WriteRaw("<v:shape id=\"_x0000_s"u8);
            xml.WriteInt32(shapeId++);
            xml.WriteRaw("\" type=\"#_x0000_t202\" style=\"position:absolute;margin-left:60pt;margin-top:2pt;"u8);
            xml.WriteRaw("width:108pt;height:60pt;z-index:1;visibility:hidden\" fillcolor=\"#ffffe1\" o:insetmode=\"auto\">"u8);
            xml.WriteRaw("<v:fill color2=\"#ffffe1\"/><v:shadow on=\"t\" color=\"black\" obscured=\"t\"/>"u8);
            xml.WriteRaw("<v:path o:connecttype=\"none\"/><v:textbox style=\"mso-direction-alt:auto\">"u8);
            xml.WriteRaw("<div style=\"text-align:left\"/></v:textbox>"u8);
            xml.WriteRaw("<x:ClientData ObjectType=\"Note\"><x:MoveWithCells/><x:SizeWithCells/>"u8);
            WriteAnchor(xml, comment.Cell);
            xml.WriteRaw("<x:AutoFill>False</x:AutoFill><x:Row>"u8);
            xml.WriteInt32(comment.Cell.Row - 1);
            xml.WriteRaw("</x:Row><x:Column>"u8);
            xml.WriteInt32(comment.Cell.Column - 1);
            xml.WriteRaw("</x:Column></x:ClientData></v:shape>"u8);
        }

        xml.WriteRaw("</xml>"u8);
    }

    /// <summary>
    /// Writes the <c>o:idmap</c> block list: every block the note ids will span, plus every
    /// block the kept foreign shapes occupy. An id block used by a shape but missing from the
    /// map is one of the things that sends Excel into repair.
    /// </summary>
    private static void WriteIdMapBlocks(
        Utf8XmlWriter xml, int drawingIndex, int firstNoteId, int noteCount, VmlLeftovers? leftovers)
    {
        var blocks = new SortedSet<int>();
        if (noteCount > 0)
        {
            for (int block = firstNoteId / 1024; block <= (firstNoteId + noteCount - 1) / 1024; block++)
                blocks.Add(block);
        }

        if (leftovers is not null)
        {
            foreach (int block in leftovers.Blocks)
                blocks.Add(block);
        }

        if (blocks.Count == 0)
            blocks.Add(drawingIndex);

        bool first = true;
        foreach (int block in blocks)
        {
            if (!first)
                xml.WriteRaw(","u8);
            xml.WriteInt32(block);
            first = false;
        }
    }

    /// <summary>
    /// Writes where a note's box sits: the cells it spans and how far into each it starts.
    /// Excel places the box from this rather than from the shape's own style, so without it
    /// every note on a sheet lands in the same spot. Like Excel, the box goes up and to the
    /// right of its cell, two columns wide and four rows tall.
    /// </summary>
    private static void WriteAnchor(Utf8XmlWriter xml, CellRef cell)
    {
        // Offsets into the first and last cell, in the format's own units: 1/1024 of a column
        // width across, 1/256 of a row height down.
        const int LeftOffset = 15;
        const int TopOffset = 2;
        const int BottomOffset = 16;

        int firstColumn = Math.Min(cell.Column, SheetLimits.MaxColumn - 3);
        int firstRow = Math.Max(cell.Row - 2, 0);
        int lastRow = Math.Min(firstRow + 4, SheetLimits.MaxRow - 1);

        xml.WriteRaw("<x:Anchor>"u8);
        xml.WriteInt32(firstColumn);
        xml.WriteRaw(", "u8);
        xml.WriteInt32(LeftOffset);
        xml.WriteRaw(", "u8);
        xml.WriteInt32(firstRow);
        xml.WriteRaw(", "u8);
        xml.WriteInt32(TopOffset);
        xml.WriteRaw(", "u8);
        xml.WriteInt32(firstColumn + 2);
        xml.WriteRaw(", "u8);
        xml.WriteInt32(LeftOffset);
        xml.WriteRaw(", "u8);
        xml.WriteInt32(lastRow);
        xml.WriteRaw(", "u8);
        xml.WriteInt32(BottomOffset);
        xml.WriteRaw("</x:Anchor>"u8);
    }
}
