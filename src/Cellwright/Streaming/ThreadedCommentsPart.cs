using System.Globalization;
using System.Xml;
using Cellwright.Xml;

namespace Cellwright.Streaming;

/// <summary>
/// Reading and writing the threaded comments of a worksheet and the people who wrote them.
/// </summary>
/// <remarks>
/// Threaded comments are an extension rather than part of ISO/IEC 29500: the parts, their
/// content types and their relationships are defined by MS-XLSX 2.1.17 and 2.1.18, and the
/// elements by 2.6.203 through 2.6.207. Two parts are involved and they live at different
/// levels — the comments hang off a sheet, the people off the workbook — because a person who
/// comments on three sheets is one person.
/// <para>
/// A file that has threads also carries a legacy note for each of them, so that a reader which
/// predates threads shows something. Excel marks those notes by giving them an author of
/// <c>tc={id}</c> and hides them; they are not notes anybody wrote, and the model keeps them out
/// of <see cref="CellComment"/> for that reason. Writing has to put them back, which is what ties
/// this part to the ordinary comments part.
/// </para>
/// </remarks>
internal static class ThreadedCommentsPart
{
    /// <summary>Namespace of the threaded comment elements.</summary>
    private const string Namespace = "http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments";

    /// <summary>How the placeholder note names the thread it stands for.</summary>
    private const string PlaceholderPrefix = "tc=";

    /// <summary>Reads the people of a workbook.</summary>
    /// <param name="stream">The persons part.</param>
    /// <param name="people">Where to put them.</param>
    public static void ReadPeople(Stream stream, PersonCollection people)
    {
        using var xml = XmlReader.Create(stream, XmlDefaults.ReaderSettings);
        while (xml.Read())
        {
            if (xml.NodeType != XmlNodeType.Element || xml.LocalName != "person")
                continue;

            if (xml.GetAttribute("id") is not { Length: > 0 } id)
                continue;

            people.AddFromFile(new Person(
                id,
                xml.GetAttribute("displayName") ?? string.Empty,
                Empty(xml.GetAttribute("userId")),
                Empty(xml.GetAttribute("providerId"))));
        }
    }

    /// <summary>Reads the threaded comments of one sheet.</summary>
    /// <param name="stream">The threaded comments part.</param>
    /// <param name="comments">Where to put them.</param>
    public static void Read(Stream stream, ThreadedCommentCollection comments)
    {
        using var xml = XmlReader.Create(stream, XmlDefaults.ReaderSettings);
        bool advance = true;
        while (true)
        {
            if (advance && !xml.Read())
                break;
            advance = true;

            if (xml.NodeType != XmlNodeType.Element || xml.LocalName != "threadedComment")
                continue;

            if (ReadComment(xml) is { } comment)
                comments.AddFromFile(comment);

            advance = false;
        }
    }

    private static ThreadedComment? ReadComment(XmlReader xml)
    {
        string? reference = xml.GetAttribute("ref");
        string? personId = xml.GetAttribute("personId");
        string? id = xml.GetAttribute("id");
        string? parentId = Empty(xml.GetAttribute("parentId"));
        bool resolved = xml.GetAttribute("done") is "1" or "true";

        DateTimeOffset? created =
            DateTimeOffset.TryParse(
                xml.GetAttribute("dT"), CultureInfo.InvariantCulture, DateTimeStyles.AdjustToUniversal, out DateTimeOffset parsed)
                ? parsed
                : null;

        string text = string.Empty;
        if (!xml.IsEmptyElement)
        {
            using XmlReader subtree = xml.ReadSubtree();
            subtree.Read();

            bool advance = true;
            while (true)
            {
                if (advance && !subtree.Read())
                    break;
                advance = true;

                if (subtree.NodeType == XmlNodeType.Element && subtree.LocalName == "text")
                {
                    text = subtree.ReadElementContentAsString();
                    advance = false;
                }
            }
        }

        if (!CellRef.TryParse(reference, null, out CellRef cell) || personId is null || id is null)
            return null;

        return new ThreadedComment
        {
            Cell = cell,
            Text = OoxmlText.Decode(text),
            PersonId = personId,
            Id = id,
            ParentId = parentId,
            Created = created,
            Resolved = resolved,
        };
    }

    /// <summary>Writes the persons part.</summary>
    /// <param name="xml">Where to write.</param>
    /// <param name="people">The people to write.</param>
    public static void WritePeople(Utf8XmlWriter xml, IReadOnlyList<Person> people)
    {
        xml.WriteDeclaration();
        xml.WriteRaw("<personList xmlns=\""u8);
        xml.WriteRaw(System.Text.Encoding.UTF8.GetBytes(Namespace));
        xml.WriteRaw("\" xmlns:x=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"u8);

        foreach (Person person in people)
        {
            xml.WriteRaw("<person displayName=\""u8);
            xml.WriteAttributeText(person.DisplayName);
            xml.WriteRaw("\" id=\""u8);
            xml.WriteAttributeText(person.Id);

            if (person.UserId is { Length: > 0 } userId)
            {
                xml.WriteRaw("\" userId=\""u8);
                xml.WriteAttributeText(userId);
            }

            xml.WriteRaw("\" providerId=\""u8);
            xml.WriteAttributeText(person.ProviderId ?? "None");
            xml.WriteRaw("\"/>"u8);
        }

        xml.WriteRaw("</personList>"u8);
    }

    /// <summary>Writes the threaded comments of one sheet.</summary>
    /// <param name="xml">Where to write.</param>
    /// <param name="comments">The comments to write, in file order.</param>
    public static void Write(Utf8XmlWriter xml, IReadOnlyList<ThreadedComment> comments)
    {
        xml.WriteDeclaration();
        xml.WriteRaw("<ThreadedComments xmlns=\""u8);
        xml.WriteRaw(System.Text.Encoding.UTF8.GetBytes(Namespace));
        xml.WriteRaw("\" xmlns:x=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"u8);

        foreach (ThreadedComment comment in comments)
        {
            xml.WriteRaw("<threadedComment ref=\""u8);
            xml.WriteFormattable(comment.Cell);
            xml.WriteRaw("\" dT=\""u8);
            xml.WriteAttributeText(
                (comment.Created ?? DateTimeOffset.UtcNow).UtcDateTime.ToString("yyyy-MM-ddTHH:mm:ssZ", CultureInfo.InvariantCulture));
            xml.WriteRaw("\" personId=\""u8);
            xml.WriteAttributeText(comment.PersonId);
            xml.WriteRaw("\" id=\""u8);
            xml.WriteAttributeText(comment.Id);

            if (comment.ParentId is { Length: > 0 } parent)
            {
                xml.WriteRaw("\" parentId=\""u8);
                xml.WriteAttributeText(parent);
            }

            if (comment.Resolved)
                xml.WriteRaw("\" done=\"1"u8);

            xml.WriteRaw("\"><text>"u8);
            xml.WriteText(comment.Text);
            xml.WriteRaw("</text></threadedComment>"u8);
        }

        xml.WriteRaw("</ThreadedComments>"u8);
    }

    /// <summary>
    /// The legacy notes a set of threads needs beside them, so that a reader which does not know
    /// about threads still shows something on those cells.
    /// </summary>
    /// <param name="comments">The sheet's threaded comments.</param>
    /// <param name="people">The workbook's people, for the name to show.</param>
    /// <remarks>
    /// One note per thread, not per message: the note stands for the conversation. Its author is
    /// the marker Excel uses to recognise it as a placeholder and hide it, and its text is the
    /// first message prefixed with the author's name, which is what Excel writes.
    /// </remarks>
    public static List<CellComment> Placeholders(
        IReadOnlyList<ThreadedComment> comments, PersonCollection people)
    {
        var notes = new List<CellComment>();
        foreach (ThreadedComment comment in comments)
        {
            if (!comment.IsFirst)
                continue;

            string author = people.ById(comment.PersonId)?.DisplayName ?? string.Empty;
            notes.Add(new CellComment
            {
                Cell = comment.Cell,
                Text = author.Length > 0 ? $"Comment:\n    {comment.Text}" : comment.Text,
                Author = PlaceholderPrefix + comment.Id,
            });
        }

        return notes;
    }

    private static string? Empty(string? value) => string.IsNullOrEmpty(value) ? null : value;
}
