using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeTranslate.Core;

namespace OfficeTranslate.Tests
{
    [TestClass]
    public class TranslationTaskFailureTests
    {
        [DataTestMethod]
        [DataRow(false)]
        [DataRow(true)]
        public void EarlierNoteFailure_ThenErrorOrCancel_KeepsFullTextAndEarlierSuccess(bool cancel)
        {
            var summary = new TranslationTaskSummary(2);
            var entries = new[] { "placed:" + new string('中', 1200), "unplaced:" + new string('译', 1200) };
            var placed = new List<string>();
            ImageOverlayNotes.Place(entries, 160, 0, (chunk, top) => {
                if (chunk.Fragments[0].EntryIndex == 1) throw new InvalidOperationException("layout unavailable");
                placed.Add(chunk.Fragments.Single().Text); return 300;
            }, (chunk, ex) => summary.RecordImageNoteFailure("image-1", chunk.Text, ex.Message));
            using (var token = new CancellationTokenSource())
            {
                if (cancel) token.Cancel();
                Exception error = cancel ? (Exception)new OperationCanceledException(token.Token)
                    : new ImageOcrTimeoutException(TimeSpan.FromSeconds(600), new TimeoutException("second image"));
                var failure = new TranslationTaskFailure(error, summary, token.Token);
                Assert.IsTrue(failure.HasRecoverableNotes); Assert.IsFalse(failure.ShowErrorDialog);
                Assert.AreEqual(cancel, failure.IsUserCancellation);
                Assert.AreEqual(cancel ? "SummaryCancelled" : "SummaryFailed", failure.HeadingKey);
                Assert.AreSame(error, failure.Error); Assert.AreSame(summary, failure.Summary);
                Assert.AreEqual(1, summary.UnplacedImageNotes.Count);
                StringAssert.Contains(summary.UnplacedImageNotes[0], entries[1]);
                CollectionAssert.AreEqual(new[] { entries[0] }, placed);
            }
        }

        [TestMethod]
        public void CancellationWithoutNotes_RemainsSilent()
        {
            using (var token = new CancellationTokenSource())
            {
                token.Cancel();
                var failure = new TranslationTaskFailure(new OperationCanceledException(token.Token), null, token.Token);
                Assert.IsTrue(failure.IsUserCancellation); Assert.IsFalse(failure.HasRecoverableNotes); Assert.IsFalse(failure.ShowErrorDialog);
            }
        }

        [TestMethod]
        public void ErrorBeforeSummaryExists_StillShowsOriginalError()
        {
            var error = new InvalidOperationException("settings invalid");
            var failure = new TranslationTaskFailure(error, null, CancellationToken.None);
            Assert.AreSame(error, failure.Error); Assert.IsTrue(failure.ShowErrorDialog); Assert.IsFalse(failure.HasRecoverableNotes);
        }

        [TestMethod]
        public void TransportCancellationWithoutUserToken_IsErrorNotSilentCancel()
        {
            var failure = new TranslationTaskFailure(new OperationCanceledException(), new TranslationTaskSummary(2), CancellationToken.None);
            Assert.IsFalse(failure.IsUserCancellation); Assert.IsTrue(failure.ShowErrorDialog);
        }

        [TestMethod]
        public void UnrelatedErrorEvenWhenTokenCancelled_IsNotHidden()
        {
            using (var token = new CancellationTokenSource())
            {
                token.Cancel();
                var failure = new TranslationTaskFailure(new InvalidOperationException("writeback failed"), null, token.Token);
                Assert.IsFalse(failure.IsUserCancellation); Assert.IsTrue(failure.ShowErrorDialog);
            }
        }
    }
}
