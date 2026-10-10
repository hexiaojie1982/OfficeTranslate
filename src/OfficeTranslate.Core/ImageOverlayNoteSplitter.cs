using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

namespace OfficeTranslate.Core
{
    public sealed class ImageOverlayNoteFragment
    {
        internal ImageOverlayNoteFragment(int entryIndex, int offset, string text)
        { EntryIndex = entryIndex; Offset = offset; Text = text; }

        public int EntryIndex { get; }
        public int Offset { get; }
        public string Text { get; }
    }

    public sealed class ImageOverlayNoteChunk
    {
        internal ImageOverlayNoteChunk(IEnumerable<ImageOverlayNoteFragment> fragments)
        {
            Fragments = new List<ImageOverlayNoteFragment>(fragments).AsReadOnly();
            var builder = new StringBuilder(ImageOverlayNotes.Header);
            foreach (var part in Fragments)
            {
                builder.Append("\r\r【条目 ").Append(part.EntryIndex + 1);
                if (part.Offset > 0) builder.Append(" 续接");
                builder.Append("，字符 ").Append(part.Offset + 1).Append("–")
                    .Append(part.Offset + part.Text.Length).Append("】\r").Append(part.Text);
            }
            Text = builder.ToString();
        }

        public IReadOnlyList<ImageOverlayNoteFragment> Fragments { get; }
        public string Text { get; }
    }

    // Only a verified layout overflow is eligible for finer splitting. A COM
    // placement failure is reported for that chunk, not recursively retried.
    public sealed class ImageOverlayNoteOverflowException : InvalidOperationException
    {
        public ImageOverlayNoteOverflowException(string message) : base(message) { }
    }

    public static partial class ImageOverlayNotes
    {
        public const float NoteFontSize = 9F;
        public const float NoteGap = 6F;

        public static float NoteWidth(float imageWidth) => Math.Min(420F, Math.Max(160F, imageWidth));

        // Notes have 4pt margins on EACH side, unlike the overlay estimator's
        // 4pt total margin. Reserve additional height for host/font differences.
        public static float EstimateHeight(float width, ImageOverlayNoteChunk chunk) =>
            ImageOverlayLayout.EstimateNoteHeight(width - 4F, chunk.Text, NoteFontSize);

        public static IReadOnlyList<ImageOverlayNoteChunk> Split(IReadOnlyList<string> entries, float width)
        {
            if (entries == null) throw new ArgumentNullException(nameof(entries));
            if (float.IsNaN(width) || float.IsInfinity(width) || width < 160F || width > 420F)
                throw new ArgumentOutOfRangeException(nameof(width));
            var chunks = new List<ImageOverlayNoteChunk>();
            var pending = new List<ImageOverlayNoteFragment>();
            var budget = ImageOverlayLayout.MaxNoteHeightPt * 0.7F;
            for (var entry = 0; entry < entries.Count; entry++)
            {
                var text = entries[entry] ?? string.Empty;
                var whole = new ImageOverlayNoteFragment(entry, 0, text);
                var candidate = new ImageOverlayNoteChunk(pending.Concat(new[] { whole }));
                if (EstimateHeight(width, candidate) <= budget) { pending.Add(whole); continue; }
                if (pending.Count > 0) { chunks.Add(new ImageOverlayNoteChunk(pending)); pending.Clear(); }
                if (EstimateHeight(width, new ImageOverlayNoteChunk(new[] { whole })) <= budget)
                { pending.Add(whole); continue; }

                // Preserve exact substrings, including line endings and Unicode
                // text elements. Prefer whole lines; split a single long line too.
                var boundaries = StringInfo.ParseCombiningCharacters(text);
                var startElement = 0;
                while (startElement < boundaries.Length)
                {
                    var start = boundaries[startElement];
                    var low = startElement + 1;
                    var high = boundaries.Length;
                    var best = low;
                    while (low <= high)
                    {
                        var mid = low + (high - low) / 2;
                        var end = mid == boundaries.Length ? text.Length : boundaries[mid];
                        var piece = new ImageOverlayNoteChunk(new[] {
                            new ImageOverlayNoteFragment(entry, start, text.Substring(start, end - start)) });
                        if (EstimateHeight(width, piece) <= budget) { best = mid; low = mid + 1; }
                        else high = mid - 1;
                    }
                    var maxEnd = best == boundaries.Length ? text.Length : boundaries[best];
                    var cut = PreferLineBoundary(text, start, maxEnd);
                    if (cut < text.Length && cut > start && text[cut - 1] == '\r' && text[cut] == '\n') cut--;
                    if (cut <= start) cut = maxEnd;
                    var chunk = new ImageOverlayNoteChunk(new[] {
                        new ImageOverlayNoteFragment(entry, start, text.Substring(start, cut - start)) });
                    if (EstimateHeight(width, chunk) > ImageOverlayLayout.MaxNoteHeightPt)
                        throw new InvalidOperationException("单个文字元素已超出旁注估高上限。");
                    chunks.Add(chunk);
                    startElement = cut == text.Length ? boundaries.Length : Array.BinarySearch(boundaries, cut);
                    if (startElement < 0) throw new InvalidOperationException("旁注切分边界不完整。");
                }
            }
            if (pending.Count > 0) chunks.Add(new ImageOverlayNoteChunk(pending));
            return chunks.AsReadOnly();
        }

        // Returns smaller, lossless chunks for a REAL host overflow. No retry
        // can repeat identical content forever: every split reduces the piece.
        public static IReadOnlyList<ImageOverlayNoteChunk> SplitForRetry(ImageOverlayNoteChunk chunk)
        {
            if (chunk.Fragments.Count > 1)
            {
                var mid = chunk.Fragments.Count / 2;
                return new[] { new ImageOverlayNoteChunk(chunk.Fragments.Take(mid)),
                    new ImageOverlayNoteChunk(chunk.Fragments.Skip(mid)) };
            }
            var part = chunk.Fragments[0];
            var boundaries = StringInfo.ParseCombiningCharacters(part.Text);
            if (boundaries.Length < 2) return Array.Empty<ImageOverlayNoteChunk>();
            var cut = PreferLineBoundary(part.Text, 0, boundaries[boundaries.Length / 2]);
            if (cut > 0 && cut < part.Text.Length && part.Text[cut - 1] == '\r' && part.Text[cut] == '\n') cut--;
            if (cut <= 0) return Array.Empty<ImageOverlayNoteChunk>();
            return new[] {
                new ImageOverlayNoteChunk(new[] { new ImageOverlayNoteFragment(part.EntryIndex, part.Offset, part.Text.Substring(0, cut)) }),
                new ImageOverlayNoteChunk(new[] { new ImageOverlayNoteFragment(part.EntryIndex, part.Offset + cut, part.Text.Substring(cut)) }) };
        }

        private static int PreferLineBoundary(string text, int start, int maxEnd)
        {
            // Avoid tiny leading lines that would cause quadratic splitting.
            for (var i = maxEnd - 1; i >= start + (maxEnd - start) / 2; i--)
                if (text[i] == '\n' || (text[i] == '\r' && (i + 1 == text.Length || text[i + 1] != '\n')))
                    return i + 1;
            return maxEnd;
        }

        // Common orchestration: place, check real layout, then stack by the
        // actual final height. The callback must delete ONLY its failed shape.
        // Other errors retain earlier successes and report the complete chunk.
        public static void Place(IReadOnlyList<string> entries, float width, float top,
            Func<ImageOverlayNoteChunk, float, float> place,
            Action<ImageOverlayNoteChunk, Exception> reportFailure)
        {
            var pending = new LinkedList<ImageOverlayNoteChunk>(Split(entries, width));
            while (pending.First != null)
            {
                var chunk = pending.First.Value;
                pending.RemoveFirst();
                try
                {
                    var height = place(chunk, top);
                    top += height + NoteGap;
                }
                catch (OperationCanceledException) { throw; }
                catch (ImageOverlayNoteOverflowException ex)
                {
                    var pieces = SplitForRetry(chunk);
                    if (pieces.Count == 0) reportFailure(chunk, ex);
                    else for (var i = pieces.Count - 1; i >= 0; i--) pending.AddFirst(pieces[i]);
                }
                catch (Exception ex) { reportFailure(chunk, ex); }
            }
        }
    }
}
