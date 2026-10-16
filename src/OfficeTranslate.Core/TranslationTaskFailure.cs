using System;
using System.Threading;

namespace OfficeTranslate.Core
{
    // One policy for all three hosts. The service exposes its live summary
    // before the first request, so later exceptions cannot strand note text.
    public sealed class TranslationTaskFailure
    {
        public TranslationTaskFailure(Exception error, TranslationTaskSummary? summary, CancellationToken userToken)
        {
            Error = error ?? throw new ArgumentNullException(nameof(error));
            Summary = summary;
            IsUserCancellation = TranslationFailurePolicy.IsUserCancellation(error, userToken);
        }

        public Exception Error { get; }
        public TranslationTaskSummary? Summary { get; }
        public bool IsUserCancellation { get; }
        public bool HasRecoverableNotes => Summary != null && (Summary.UnplacedImageNotes.Count > 0 || Summary.TextBlockFailures.Count > 0);
        public bool ShowErrorDialog => !HasRecoverableNotes && !IsUserCancellation;
        public string HeadingKey => IsUserCancellation ? "SummaryCancelled" : "SummaryFailed";
    }
}
