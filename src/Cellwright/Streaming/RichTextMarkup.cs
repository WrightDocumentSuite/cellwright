using System.Xml;
using Cellwright.Styles;
using Cellwright.Xml;

namespace Cellwright.Streaming;

/// <summary>
/// The bridge between a cell's formatted text and the runs the file stores it as.
/// </summary>
/// <remarks>
/// A shared string keeps its formatting in <c>r</c> children, each an optional <c>rPr</c> font
/// and a <c>t</c> of text (§18.4.4). The reader captures those children verbatim so an untouched
/// string is written back exactly as it arrived; this turns that markup into a model and back
/// when somebody actually wants to read or change the runs.
/// <para>
/// A captured fragment can hold more than runs: phonetic hints (<c>rPh</c>, <c>phoneticPr</c>)
/// sit beside them for East Asian text. Those are read past — they annotate the text rather than
/// being it — and they are not regenerated, so rewriting a string's runs drops the hints that
/// were attached to the old text. That is the honest outcome: a hint names a stretch of the
/// reading it annotates, and the stretch is gone.
/// </para>
/// </remarks>
internal static class RichTextMarkup
{
    /// <summary>Parses captured run markup into runs.</summary>
    /// <param name="markup">The <c>r</c> children of an <c>si</c>, as captured by the reader.</param>
    public static RichText Parse(string markup)
    {
        if (string.IsNullOrEmpty(markup))
            return RichText.Empty;

        var runs = new List<RichTextRun>();

        // The captured fragment is a run of siblings without a root, which no XML reader will
        // take on its own; a fragment-mode reader is exactly the one that will.
        var settings = XmlDefaults.ReaderSettings.Clone();
        settings.ConformanceLevel = ConformanceLevel.Fragment;

        using var xml = XmlReader.Create(new StringReader(markup), settings);
        while (xml.Read())
        {
            if (xml.NodeType != XmlNodeType.Element || xml.LocalName != "r")
                continue;

            if (ReadRun(xml) is { } run)
                runs.Add(run);
        }

        return runs.Count == 0 ? RichText.Empty : new RichText(runs);
    }

    /// <summary>Writes runs as the <c>r</c> children of an <c>si</c>.</summary>
    /// <param name="xml">Where to write.</param>
    /// <param name="text">The runs to write.</param>
    public static void Write(Utf8XmlWriter xml, RichText text)
    {
        foreach (RichTextRun run in text)
        {
            xml.WriteRaw("<r>"u8);

            if (run.Font is { } font)
            {
                xml.WriteRaw("<rPr>"u8);
                FontMarkup.WriteChildren(xml, font, runProperties: true);
                xml.WriteRaw("</rPr>"u8);
            }

            if (XlsxSheetWriter.NeedsSpacePreserve(run.Text))
                xml.WriteRaw("<t xml:space=\"preserve\">"u8);
            else
                xml.WriteRaw("<t>"u8);

            xml.WriteText(run.Text);
            xml.WriteRaw("</t>"u8);
            xml.WriteRaw("</r>"u8);
        }
    }

    /// <summary>Renders runs as the markup the model stores beside a cell.</summary>
    /// <param name="text">The runs to render.</param>
    public static string ToMarkup(RichText text) => Utf8XmlWriter.Render(xml => Write(xml, text));

    /// <summary>Reads one <c>r</c> element; the reader starts on it and ends on its end tag.</summary>
    private static RichTextRun? ReadRun(XmlReader xml)
    {
        if (xml.IsEmptyElement)
            return null;

        FontStyle? font = null;
        string? runText = null;

        // Reading an element's content leaves the reader on whatever follows it, so advancing
        // again would step over that node. For the last child of a run, the node that would be
        // stepped over is the run's own end tag — and past it, the next run.
        bool advance = true;
        while (true)
        {
            if (advance && !xml.Read())
                break;
            advance = true;

            if (xml.NodeType == XmlNodeType.EndElement && xml.LocalName == "r")
                break;
            if (xml.NodeType != XmlNodeType.Element)
                continue;

            switch (xml.LocalName)
            {
                case "rPr":
                    font = FontMarkup.Parse(xml, "rPr");
                    break;

                case "t":
                    runText = xml.ReadElementContentAsString();
                    advance = false;
                    break;
            }
        }

        return runText is null ? null : new RichTextRun(runText, font);
    }
}
