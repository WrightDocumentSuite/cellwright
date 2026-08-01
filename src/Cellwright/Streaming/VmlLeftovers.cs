using System.Xml;
using Cellwright.Xml;

namespace Cellwright.Streaming;

/// <summary>
/// The shapes of a sheet's legacy VML drawing that are not note boxes: check boxes, buttons
/// and the rest of the form controls that share the part with the comments. Regenerating the
/// part for edited notes must carry these through, or the first touched comment silently
/// strips the sheet of its controls.
/// </summary>
internal sealed class VmlLeftovers
{
    /// <summary>Shape type declarations other than the note box type, verbatim.</summary>
    public required IReadOnlyList<string> ShapeTypes { get; init; }

    /// <summary>Shapes that are not notes, verbatim and keeping their original ids.</summary>
    public required IReadOnlyList<string> Shapes { get; init; }

    /// <summary>
    /// Identifier blocks (<c>id / 1024</c>) the kept shapes occupy — <c>o:idmap</c> has to
    /// declare them alongside the block the regenerated note boxes use.
    /// </summary>
    public required IReadOnlyList<int> Blocks { get; init; }

    /// <summary>Highest shape id among the kept shapes, for collision-free numbering.</summary>
    public required int MaxShapeId { get; init; }

    /// <summary>
    /// Extracts the non-note content of a VML part. Returns <see langword="null"/> when there
    /// is nothing to keep or the part does not parse — the caller then regenerates from the
    /// model alone, which is what happened before leftovers were carried at all.
    /// </summary>
    public static VmlLeftovers? Parse(byte[] content)
    {
        List<string> shapeTypes = [];
        List<string> shapes = [];
        var blocks = new SortedSet<int>();
        int maxShapeId = 0;

        try
        {
            using var stream = new MemoryStream(content, writable: false);
            using var xml = XmlReader.Create(stream, XmlDefaults.ReaderSettings);

            // Skip, and ReadOuterXml both leave the reader on the following node, so the next
            // iteration must not advance over it.
            bool advance = true;
            while (true)
            {
                if (advance && !xml.Read())
                    break;
                advance = true;

                if (xml.NodeType != XmlNodeType.Element || xml.Depth != 1)
                    continue;

                switch (xml.LocalName)
                {
                    // The layout is regenerated, and the note box type is re-declared by the
                    // writer; everything else is somebody's and travels as it was.
                    case "shapelayout":
                    case "shapetype" when xml.GetAttribute("id") == "_x0000_t202":
                        xml.Skip();
                        advance = false;
                        break;
                    case "shapetype":
                        shapeTypes.Add(xml.ReadOuterXml());
                        advance = false;
                        break;
                    case "shape":
                        string? shapeId = xml.GetAttribute("id");
                        string outer = xml.ReadOuterXml();
                        advance = false;
                        if (IsNoteShape(outer))
                            break;

                        shapes.Add(outer);
                        if (TryParseShapeId(shapeId, out int numeric))
                        {
                            blocks.Add(numeric / 1024);
                            maxShapeId = Math.Max(maxShapeId, numeric);
                        }

                        break;
                    default:
                        shapes.Add(xml.ReadOuterXml());
                        advance = false;
                        break;
                }
            }
        }
        catch (XmlException)
        {
            return null;
        }

        return shapes.Count == 0 && shapeTypes.Count == 0
            ? null
            : new VmlLeftovers { ShapeTypes = shapeTypes, Shapes = shapes, Blocks = [.. blocks], MaxShapeId = maxShapeId };
    }

    /// <summary>Whether the shape's client data declares it a note box.</summary>
    private static bool IsNoteShape(string outerXml)
    {
        try
        {
            using var xml = XmlReader.Create(new StringReader(outerXml), XmlDefaults.ReaderSettings);
            while (xml.Read())
            {
                if (xml.NodeType == XmlNodeType.Element && xml.LocalName == "ClientData")
                    return xml.GetAttribute("ObjectType") == "Note";
            }
        }
        catch (XmlException)
        {
            // ReadOuterXml already parsed it once; being here means something exotic, and a
            // shape we cannot classify is safer kept than dropped.
        }

        return false;
    }

    /// <summary>Parses the numeric part of a VML shape id like <c>_x0000_s1026</c>.</summary>
    private static bool TryParseShapeId(string? id, out int numeric)
    {
        numeric = 0;
        if (id is null)
            return false;

        int start = id.Length;
        while (start > 0 && char.IsAsciiDigit(id[start - 1]))
            start--;

        return start < id.Length && int.TryParse(id.AsSpan(start), out numeric);
    }
}
