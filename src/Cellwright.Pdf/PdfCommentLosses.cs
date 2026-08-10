namespace Cellwright.Pdf;

/// <summary>Aggregates comment metadata that interactive PDF annotations cannot represent.</summary>
internal sealed class PdfCommentLosses
{
    private Workbook? _peopleScannedFor;
    private int _mentions;
    private int _messageExtensions;
    private int _collectionExtensions;
    private int _personExtensions;
    private int _personListExtensions;

    internal void Collect(Workbook workbook, ThreadedCommentCollection comments)
    {
        if (!string.IsNullOrWhiteSpace(comments.ExtensionListXml))
            _collectionExtensions++;

        foreach (ThreadedComment comment in comments)
        {
            if (!string.IsNullOrWhiteSpace(comment.MentionsXml))
                _mentions++;
            if (!string.IsNullOrWhiteSpace(comment.ExtensionListXml))
                _messageExtensions++;
        }

        if (ReferenceEquals(_peopleScannedFor, workbook))
            return;

        _peopleScannedFor = workbook;
        if (!string.IsNullOrWhiteSpace(workbook.People.ExtensionListXml))
            _personListExtensions++;

        foreach (Person person in workbook.People)
        {
            if (!string.IsNullOrWhiteSpace(person.ExtensionListXml))
                _personExtensions++;
        }
    }

    internal void Report(PdfExportDiagnostics diagnostics)
    {
        int total = _mentions + _messageExtensions + _collectionExtensions +
            _personExtensions + _personListExtensions;
        if (total == 0)
            return;

        diagnostics.Add(
            PdfExportWarningKind.ContentSkipped,
            "Interactive PDF annotations have no equivalent for threaded-comment opaque metadata: " +
            $"{_mentions} mention block(s), {_messageExtensions} message extension list(s), " +
            $"{_collectionExtensions} collection extension list(s), {_personExtensions} person extension list(s), " +
            $"and {_personListExtensions} person-list extension list(s) were omitted. " +
            "The opaque XML itself is not exposed in diagnostics.",
            "comment-metadata");
    }
}
