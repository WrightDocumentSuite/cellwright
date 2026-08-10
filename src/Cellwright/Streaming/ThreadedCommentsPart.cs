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
    private const string _namespace = "http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments";

    /// <summary>How the placeholder note names the thread it stands for.</summary>
    private const string _placeholderPrefix = "tc=";

    /// <summary>
    /// Extracts the threaded root identifier from the mandatory <c>tc={uid}</c> author token.
    /// The revision uid is only a SHOULD in MS-XLSX and is validated separately.
    /// </summary>
    public static bool TryGetPlaceholderId(CellComment note, out string id)
    {
        id = string.Empty;
        if (note.Author is not { } author || !author.StartsWith(_placeholderPrefix, StringComparison.Ordinal))
            return false;

        string token = author[_placeholderPrefix.Length..];
        if (!Guid.TryParse(token, out _))
            return false;

        id = token;
        return true;
    }

    /// <summary>Reads the people of a workbook.</summary>
    /// <param name="stream">The persons part.</param>
    /// <param name="people">Where to put them.</param>
    public static void ReadPeople(Stream stream, PersonCollection people)
    {
        using var xml = XmlReader.Create(stream, XmlDefaults.ReaderSettings);
        bool advance = true;
        while (true)
        {
            if (advance && !xml.Read())
                break;
            advance = true;

            if (xml.NodeType != XmlNodeType.Element)
                continue;

            if (xml.LocalName == "person" && xml.NamespaceURI == _namespace)
            {
                bool empty = xml.IsEmptyElement;
                if (ReadPerson(xml) is { } person)
                    people.AddFromFile(person);

                advance = empty;
                continue;
            }

            if (xml.LocalName == "extLst" && xml.Depth == 1)
            {
                people.ExtensionListXml = xml.ReadOuterXml();
                advance = false;
            }
        }
    }

    private static Person? ReadPerson(XmlReader xml)
    {
        string? id = xml.GetAttribute("id");
        string displayName = xml.GetAttribute("displayName") ?? string.Empty;
        string? userId = Empty(xml.GetAttribute("userId"));
        string? providerId = Empty(xml.GetAttribute("providerId"));
        string? extensionList = null;

        if (!xml.IsEmptyElement)
        {
            using XmlReader subtree = xml.ReadSubtree();
            subtree.Read();
            while (subtree.Read())
            {
                if (subtree.NodeType == XmlNodeType.Element &&
                    subtree.LocalName == "extLst" && subtree.Depth == 1)
                {
                    extensionList = subtree.ReadOuterXml();
                }
            }
        }

        return id is not { Length: > 0 } ? null : new Person(id, displayName, userId, providerId)
        {
            ExtensionListXml = extensionList,
        };
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
            {
                if (xml.NodeType == XmlNodeType.Element && xml.LocalName == "extLst" && xml.Depth == 1)
                {
                    comments.ExtensionListXml = xml.ReadOuterXml();
                    advance = false;
                }
                continue;
            }

            bool empty = xml.IsEmptyElement;
            if (ReadComment(xml) is { } comment)
                comments.AddFromFile(comment);

            advance = empty;
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
        string? mentions = null;
        string? extensionList = null;
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
                else if (subtree.NodeType == XmlNodeType.Element && subtree.Depth == 1 &&
                    subtree.LocalName == "mentions")
                {
                    mentions = subtree.ReadOuterXml();
                    advance = false;
                }
                else if (subtree.NodeType == XmlNodeType.Element && subtree.Depth == 1 &&
                    subtree.LocalName == "extLst")
                {
                    extensionList = subtree.ReadOuterXml();
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
            MentionsXml = mentions,
            ExtensionListXml = extensionList,
        };
    }

    /// <summary>Writes the persons part.</summary>
    /// <param name="xml">Where to write.</param>
    /// <param name="people">The people to write.</param>
    public static void WritePeople(Utf8XmlWriter xml, PersonCollection people)
    {
        xml.WriteDeclaration();
        xml.WriteRaw("<personList xmlns=\""u8);
        xml.WriteRaw(System.Text.Encoding.UTF8.GetBytes(_namespace));
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
            if (person.ExtensionListXml is { Length: > 0 } extensions)
            {
                xml.WriteRaw("\">"u8);
                xml.WriteRawXml(extensions);
                xml.WriteRaw("</person>"u8);
            }
            else
            {
                xml.WriteRaw("\"/>"u8);
            }
        }

        RawXml.Write(xml, people.ExtensionListXml);

        xml.WriteRaw("</personList>"u8);
    }

    /// <summary>Writes the threaded comments of one sheet.</summary>
    /// <param name="xml">Where to write.</param>
    /// <param name="comments">The comments to write, in file order.</param>
    public static void Write(Utf8XmlWriter xml, ThreadedCommentCollection comments)
    {
        xml.WriteDeclaration();
        xml.WriteRaw("<ThreadedComments xmlns=\""u8);
        xml.WriteRaw(System.Text.Encoding.UTF8.GetBytes(_namespace));
        xml.WriteRaw("\" xmlns:x=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"u8);

        foreach (ThreadedComment comment in comments)
        {
            xml.WriteRaw("<threadedComment ref=\""u8);
            xml.WriteFormattable(comment.Cell);
            if (comment.Created is { } created)
            {
                xml.WriteRaw("\" dT=\""u8);
                xml.WriteAttributeText(
                    created.UtcDateTime.ToString("yyyy-MM-ddTHH:mm:ssZ", CultureInfo.InvariantCulture));
            }
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
            xml.WriteRaw("</text>"u8);
            RawXml.Write(xml, comment.MentionsXml);
            RawXml.Write(xml, comment.ExtensionListXml);
            xml.WriteRaw("</threadedComment>"u8);
        }

        RawXml.Write(xml, comments.ExtensionListXml);

        xml.WriteRaw("</ThreadedComments>"u8);
    }

    /// <summary>
    /// The legacy notes a set of threads needs beside them, so that a reader which does not know
    /// about threads still shows something on those cells.
    /// </summary>
    /// <param name="roots">The sheet's projected thread roots.</param>
    /// <param name="people">The workbook's people, for the name to show.</param>
    /// <remarks>
    /// One note per thread, not per message: the note stands for the conversation. Its author is
    /// the marker Excel uses to recognise it as a placeholder and hide it, and its text is the
    /// first message prefixed with the author's name, which is what Excel writes.
    /// </remarks>
    public static List<CellComment> Placeholders(
        IReadOnlyList<ThreadedComment> roots, PersonCollection people)
    {
        var notes = new List<CellComment>();
        foreach (ThreadedComment comment in roots)
        {
            string author = people.ById(comment.PersonId)?.DisplayName ?? string.Empty;
            notes.Add(new CellComment
            {
                Cell = comment.Cell,
                Text = author.Length > 0 ? $"Comment:\n    {comment.Text}" : comment.Text,
                Author = _placeholderPrefix + comment.Id,
                RevisionUid = comment.Id,
            });
        }

        return notes;
    }

    private static string? Empty(string? value) => string.IsNullOrEmpty(value) ? null : value;
}
