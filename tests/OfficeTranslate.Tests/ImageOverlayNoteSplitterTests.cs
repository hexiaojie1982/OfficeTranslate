using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeTranslate.Core;

namespace OfficeTranslate.Tests
{
    [TestClass]
    public sealed class ImageOverlayNoteSplitterTests
    {
        [TestMethod]
        public void Split_ManyEntries_KeepsWholeEntriesAndOrder()
        {
            var entries = Enumerable.Range(0, 40).Select(i => "条目" + i + "\r" + new string('中', 160)).ToArray();
            var chunks = ImageOverlayNotes.Split(entries, 160F);
            Assert.IsTrue(chunks.Count > 1);
            Assert.AreEqual(entries.Length, chunks.Sum(c => c.Fragments.Count));
            AssertLossless(entries, chunks);
        }

        [TestMethod]
        public void Split_SingleHugeMultilineEntry_PreservesLinesAndContinuation()
        {
            var entries = new[] { string.Concat(Enumerable.Repeat("原文与译文\r\n", 1000)) };
            var chunks = ImageOverlayNotes.Split(entries, 160F);
            Assert.IsTrue(chunks.Count > 1);
            Assert.IsTrue(chunks.Skip(1).All(c => c.Text.Contains("续接")));
            Assert.IsTrue(chunks.Take(chunks.Count - 1).All(c => c.Fragments.Last().Text.EndsWith("\r\n")));
            AssertLossless(entries, chunks);
        }

        [TestMethod]
        public void Split_SingleHugeLine_SplitsWithoutDroppingText()
        {
            var entries = new[] { new string('中', 30000) };
            var chunks = ImageOverlayNotes.Split(entries, 160F);
            Assert.IsTrue(chunks.Count > 1);
            AssertLossless(entries, chunks);
        }

        [DataTestMethod]
        [DataRow(160F)]
        [DataRow(300F)]
        [DataRow(420F)]
        public void Split_EveryChunkEstimate_IsUnderCap(float width)
        {
            var entries = new[] { new string('中', 30000), string.Concat(Enumerable.Repeat("abcd\r", 3000)), "末尾" };
            var chunks = ImageOverlayNotes.Split(entries, width);
            Assert.IsTrue(chunks.All(c => ImageOverlayNotes.EstimateHeight(width, c) <= ImageOverlayLayout.MaxNoteHeightPt));
            AssertLossless(entries, chunks);
        }

        [TestMethod]
        public void Split_UnicodeAndMixedLineEndings_KeepExactTextElements()
        {
            var text = string.Concat(Enumerable.Repeat("A\r\nB\rC\n😀e\u0301中", 2500));
            var chunks = ImageOverlayNotes.Split(new[] { text }, 160F);
            foreach (var part in chunks.SelectMany(c => c.Fragments))
            {
                Assert.IsFalse(char.IsLowSurrogate(part.Text[0]));
                Assert.IsFalse(char.IsHighSurrogate(part.Text[part.Text.Length - 1]));
                Assert.IsFalse(part.Offset > 0 && text[part.Offset - 1] == '\r' && text[part.Offset] == '\n');
                Assert.IsFalse(part.Text[0] == '\u0301');
            }
            AssertLossless(new[] { text }, chunks);
        }

        [TestMethod]
        public void Split_EmptyInput_HasNoNotes()
        {
            Assert.AreEqual(0, ImageOverlayNotes.Split(Array.Empty<string>(), 160F).Count);
        }

        [TestMethod]
        public void Split_EmptyEntries_AreRepresentedExactlyOnce()
        {
            var entries = new[] { "", "", "end" };
            var chunks = ImageOverlayNotes.Split(entries, 160F);
            Assert.AreEqual(3, chunks.Sum(c => c.Fragments.Count));
            AssertLossless(entries, chunks);
        }

        [TestMethod]
        public void Split_InvalidWidth_RejectsNonFiniteGeometry()
        {
            Assert.ThrowsException<ArgumentOutOfRangeException>(() => ImageOverlayNotes.Split(new[] { "a" }, float.NaN));
            Assert.ThrowsException<ArgumentOutOfRangeException>(() => ImageOverlayNotes.Split(new[] { "a" }, float.PositiveInfinity));
        }

        [TestMethod]
        public void SplitForRetry_MultipleEntries_PreservesOrder()
        {
            var entries = new[] { "one", "two", "three" };
            var chunk = ImageOverlayNotes.Split(entries, 160F).Single();
            var pieces = ImageOverlayNotes.SplitForRetry(chunk);
            Assert.AreEqual(2, pieces.Count);
            AssertLossless(entries, pieces);
        }

        [TestMethod]
        public void SplitForRetry_LongEntry_PreservesOffsetsAndText()
        {
            var entries = new[] { "首行\r\n次行\r\n末行" };
            var pieces = ImageOverlayNotes.SplitForRetry(ImageOverlayNotes.Split(entries, 160F).Single());
            Assert.AreEqual(2, pieces.Count);
            Assert.IsTrue(pieces[1].Text.Contains("续接"));
            AssertLossless(entries, pieces);
        }

        [TestMethod]
        public void Place_RealOverflow_SplitsAgainAndStacksActualHeights()
        {
            var placed = new List<ImageOverlayNoteChunk>();
            var tops = new List<float>();
            ImageOverlayNotes.Place(new[] { "abcd" }, 160F, 20F, (chunk, top) =>
            {
                if (chunk.Fragments.Sum(p => p.Text.Length) > 1) throw new ImageOverlayNoteOverflowException("overflow");
                placed.Add(chunk); tops.Add(top); return 50F;
            }, (chunk, error) => Assert.Fail(error.Message));
            AssertLossless(new[] { "abcd" }, placed);
            CollectionAssert.AreEqual(new[] { 20F, 76F, 132F, 188F }, tops);
        }

        [TestMethod]
        public void Place_OneFailure_PreservesSuccessAndContinues()
        {
            var entries = new[] { "first" + new string('中', 1200), "bad" + new string('中', 1200), "last" + new string('中', 1200) };
            var placed = new List<ImageOverlayNoteChunk>();
            var failed = new List<ImageOverlayNoteChunk>();
            ImageOverlayNotes.Place(entries, 160F, 0F, (chunk, top) =>
            {
                if (chunk.Fragments[0].EntryIndex == 1) throw new InvalidOperationException("COM unavailable");
                placed.Add(chunk); return 300F;
            }, (chunk, error) => failed.Add(chunk));
            Assert.AreEqual(2, placed.Count);
            Assert.AreEqual(1, failed.Count);
            Assert.AreEqual(2, placed.Last().Fragments[0].EntryIndex);
            Assert.AreEqual(entries[1], failed[0].Fragments[0].Text);
        }

        [TestMethod]
        public void Place_UnsplittableOverflow_ReportsOnlyThatFragment()
        {
            var placed = new List<ImageOverlayNoteChunk>();
            var failed = new List<ImageOverlayNoteChunk>();
            ImageOverlayNotes.Place(new[] { "abc" }, 160F, 0F, (chunk, top) =>
            {
                if (chunk.Fragments[0].Text.Contains("b")) throw new ImageOverlayNoteOverflowException("overflow");
                placed.Add(chunk); return 30F;
            }, (chunk, error) => failed.Add(chunk));
            Assert.AreEqual("b", failed.Single().Fragments.Single().Text);
            Assert.AreEqual("ac", string.Concat(placed.SelectMany(c => c.Fragments).Select(p => p.Text)));
        }

        [TestMethod]
        public void Place_Cancellation_IsNotSwallowedAsNoteFailure()
        {
            Assert.ThrowsException<OperationCanceledException>(() => ImageOverlayNotes.Place(new[] { "a" }, 160F, 0F,
                (chunk, top) => throw new OperationCanceledException(), (chunk, error) => Assert.Fail("cancel swallowed")));
        }

        [TestMethod]
        public void Summary_UnplacedNotes_KeepFullTextAndOwner()
        {
            var summary = new TranslationTaskSummary(1);
            var text = new string('中', 20000);
            summary.RecordImageNoteFailure("image-1", text, "layout failed");
            Assert.AreEqual(1, summary.UnplacedImageNotes.Count);
            StringAssert.Contains(summary.UnplacedImageNotes[0], "image-1");
            StringAssert.Contains(summary.UnplacedImageNotes[0], text);
        }

        [TestMethod]
        public void ChunkMarkers_AllBelongToSameImage_NotOtherImage()
        {
            foreach (var chunk in ImageOverlayNotes.Split(new[] { new string('中', 10000) }, 160F))
            {
                StringAssert.Contains(chunk.Text, ImageOverlayNotes.Header);
                Assert.IsTrue(ImageOverlayTags.WordExcelMarkerBelongsTo(ImageOverlayTags.NoteFor("one"), "one"));
                Assert.IsFalse(ImageOverlayTags.WordExcelMarkerBelongsTo(ImageOverlayTags.NoteFor("one"), "two"));
                Assert.IsTrue(ImageOverlayTags.PowerPointTagBelongsTo("Note:one", "one"));
                Assert.IsFalse(ImageOverlayTags.PowerPointTagBelongsTo("Note:one", "two"));
            }
        }

        private static void AssertLossless(IReadOnlyList<string> entries, IEnumerable<ImageOverlayNoteChunk> chunks)
        {
            var parts = chunks.SelectMany(c => c.Fragments).ToArray();
            var lastEntry = -1;
            foreach (var part in parts) { Assert.IsTrue(part.EntryIndex >= lastEntry); lastEntry = part.EntryIndex; }
            for (var entry = 0; entry < entries.Count; entry++)
            {
                var offset = 0;
                foreach (var part in parts.Where(p => p.EntryIndex == entry))
                { Assert.AreEqual(offset, part.Offset); offset += part.Text.Length; }
                Assert.AreEqual(entries[entry], string.Concat(parts.Where(p => p.EntryIndex == entry).Select(p => p.Text)));
            }
        }
    }
}
