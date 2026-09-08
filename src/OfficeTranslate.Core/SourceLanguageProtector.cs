using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;

namespace OfficeTranslate.Core
{
    internal sealed class ProtectedContentException : InvalidOperationException
    {
        public ProtectedContentException(string message) : base(message) { }
    }

    internal sealed class ProtectedTranslationPart
    {
        public ProtectedTranslationPart(string text, bool shouldTranslate)
        {
            Text = text;
            ShouldTranslate = shouldTranslate;
        }

        public string Text { get; }
        public bool ShouldTranslate { get; }
    }

    public sealed class ProtectedTranslationText
    {
        private readonly IReadOnlyList<KeyValuePair<string, string>> _segments;

        internal ProtectedTranslationText(string text, IReadOnlyList<KeyValuePair<string, string>> segments)
        {
            Text = text;
            _segments = segments;
        }

        public string Text { get; }
        public bool HasProtectedSegments => _segments.Count > 0;
        public bool IsFullyProtected => _segments.Count == 1 && Text == _segments[0].Key;

        internal IReadOnlyList<ProtectedTranslationPart> GetParts()
        {
            var parts = new List<ProtectedTranslationPart>();
            var position = 0;
            foreach (var segment in _segments)
            {
                var tokenPosition = Text.IndexOf(segment.Key, position, StringComparison.Ordinal);
                if (tokenPosition < 0)
                    throw new ProtectedContentException("无法重建受保护的混合语言文本。");
                if (tokenPosition > position)
                    parts.Add(new ProtectedTranslationPart(Text.Substring(position, tokenPosition - position), true));
                parts.Add(new ProtectedTranslationPart(segment.Value, false));
                position = tokenPosition + segment.Key.Length;
            }
            if (position < Text.Length)
                parts.Add(new ProtectedTranslationPart(Text.Substring(position), true));
            return parts;
        }

        public string Restore(string translatedText)
        {
            var restored = translatedText ?? string.Empty;
            foreach (var segment in _segments)
                restored = NormalizeTokenVariant(restored, segment.Key);

            EnsureNoUnknownTokenFragments(restored);

            foreach (var segment in _segments)
            {
                if (restored.IndexOf(segment.Key, StringComparison.Ordinal) < 0)
                {
                    var atStart = Text.StartsWith(segment.Key, StringComparison.Ordinal);
                    var atEnd = Text.EndsWith(segment.Key, StringComparison.Ordinal);
                    if (atStart && atEnd) return segment.Value;
                    if (atStart) restored = segment.Value + restored;
                    else if (atEnd) restored += segment.Value;
                    else throw new ProtectedContentException("翻译模型修改了受保护的非源语言内容。");
                    continue;
                }
                restored = restored.Replace(segment.Key, segment.Value);
            }
            return restored;
        }

        private static string NormalizeTokenVariant(string text, string canonicalToken)
        {
            var numberMatch = Regex.Match(canonicalToken, @"\d+");
            if (!numberMatch.Success) return text;
            var sequence = int.Parse(numberMatch.Value);
            var separator = @"[\s_\-:：.．＿－\u200B-\u200D\uFEFF]*";
            var pattern = @"[\p{Ps}\p{Pi}<＜]*\s*[OＯ]" + separator + @"[TＴ]" + separator +
                @"[KＫ]" + separator + @"[EＥ]" + separator + @"[EＥ]" + separator + @"[PＰ]" +
                separator + @"[0０]*" + BuildDigitPattern(sequence) + @"\s*[\p{Pe}\p{Pf}>＞]*";
            return Regex.Replace(text, pattern, canonicalToken, RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
        }

        private static string BuildDigitPattern(int value)
        {
            var result = new StringBuilder();
            foreach (var digit in value.ToString())
                result.Append('[').Append(digit).Append((char)('０' + digit - '0')).Append(']');
            return result.ToString();
        }

        private void EnsureNoUnknownTokenFragments(string text)
        {
            if (_segments.Count == 0) return;

            var withoutKnownTokens = text;
            foreach (var segment in _segments)
                withoutKnownTokens = withoutKnownTokens.Replace(segment.Key, string.Empty);

            var separator = @"[\s_\-:：.．＿－\u200B-\u200D\uFEFF]*";
            var markerName = @"[OＯ]" + separator + @"[TＴ]" + separator +
                @"[KＫ]" + separator + @"[EＥ]" + separator + @"[EＥ]" + separator + @"[PＰ]";
            var bracketedFragment = @"[\p{Ps}\p{Pi}<＜]\s*" + markerName;
            var numberedFragment = markerName + separator + @"\d+";
            if (Regex.IsMatch(withoutKnownTokens, bracketedFragment + "|" + numberedFragment,
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant))
                throw new ProtectedContentException("翻译模型返回了无法识别的保护标记。");
        }
    }

    public static class SourceLanguageProtector
    {
        public static ProtectedTranslationText Protect(string text, string sourceLanguage)
        {
            if (string.IsNullOrEmpty(text) || IsAutomatic(sourceLanguage))
                return new ProtectedTranslationText(text ?? string.Empty, new KeyValuePair<string, string>[0]);

            var output = new StringBuilder(text.Length);
            var protectedSegments = new List<KeyValuePair<string, string>>();
            var position = 0;
            while (position < text.Length)
            {
                if (IsSourceLetter(text[position], sourceLanguage))
                {
                    output.Append(text[position]);
                    position++;
                    continue;
                }

                var start = position;
                var containsForeignLetter = false;
                while (position < text.Length && !IsSourceLetter(text[position], sourceLanguage))
                {
                    if (char.IsLetter(text[position])) containsForeignLetter = true;
                    position++;
                }

                var segment = text.Substring(start, position - start);
                if (!containsForeignLetter)
                {
                    output.Append(segment);
                    continue;
                }

                var token = CreateUniqueToken(text, protectedSegments, protectedSegments.Count + 1);
                protectedSegments.Add(new KeyValuePair<string, string>(token, segment));
                output.Append(token);
            }

            return new ProtectedTranslationText(output.ToString(), protectedSegments);
        }

        public static bool IsAutomatic(string sourceLanguage)
        {
            return string.IsNullOrWhiteSpace(sourceLanguage) || sourceLanguage == "自动检测";
        }

        private static string CreateUniqueToken(string originalText, IReadOnlyList<KeyValuePair<string, string>> existing, int sequence)
        {
            var token = "⟦OT_KEEP_" + sequence.ToString("D4") + "⟧";
            while (originalText.IndexOf(token, StringComparison.Ordinal) >= 0 || ContainsToken(existing, token))
            {
                sequence++;
                token = "⟦OT_KEEP_" + sequence.ToString("D4") + "⟧";
            }
            return token;
        }

        private static bool ContainsToken(IReadOnlyList<KeyValuePair<string, string>> existing, string token)
        {
            for (var i = 0; i < existing.Count; i++)
                if (existing[i].Key == token) return true;
            return false;
        }

        private static bool IsSourceLetter(char character, string sourceLanguage)
        {
            if (!char.IsLetter(character)) return false;
            switch (sourceLanguage)
            {
                case "简体中文":
                case "繁體中文":
                    return IsHan(character) || IsBopomofo(character);
                case "日语":
                    return IsHan(character) || IsHiragana(character) || IsKatakana(character);
                case "韩语":
                    // Hanja can be part of Korean prose, so keep it available to
                    // the translator together with Hangul.
                    return IsHangul(character) || IsHan(character);
                case "俄语":
                    return IsCyrillic(character);
                case "阿拉伯语":
                    return IsArabic(character);
                case "英语":
                case "法语":
                case "德语":
                case "西班牙语":
                case "葡萄牙语":
                case "意大利语":
                    return IsLatin(character);
                default:
                    return true;
            }
        }

        private static bool IsLatin(char c) =>
            (c >= '\u0041' && c <= '\u024F') || (c >= '\u1E00' && c <= '\u1EFF') ||
            (c >= '\uFF21' && c <= '\uFF5A');

        private static bool IsHan(char c) =>
            (c >= '\u3400' && c <= '\u4DBF') || (c >= '\u4E00' && c <= '\u9FFF') ||
            (c >= '\uF900' && c <= '\uFAFF');

        private static bool IsBopomofo(char c) => c >= '\u3100' && c <= '\u312F';
        private static bool IsHiragana(char c) => c >= '\u3040' && c <= '\u309F';
        private static bool IsKatakana(char c) =>
            (c >= '\u30A0' && c <= '\u30FF') || (c >= '\u31F0' && c <= '\u31FF') ||
            (c >= '\uFF66' && c <= '\uFF9D');

        private static bool IsHangul(char c) =>
            (c >= '\u1100' && c <= '\u11FF') || (c >= '\u3130' && c <= '\u318F') ||
            (c >= '\uA960' && c <= '\uA97F') || (c >= '\uAC00' && c <= '\uD7AF') ||
            (c >= '\uD7B0' && c <= '\uD7FF');

        private static bool IsCyrillic(char c) =>
            (c >= '\u0400' && c <= '\u052F') || (c >= '\u2DE0' && c <= '\u2DFF') ||
            (c >= '\uA640' && c <= '\uA69F');

        private static bool IsArabic(char c) =>
            (c >= '\u0600' && c <= '\u06FF') || (c >= '\u0750' && c <= '\u077F') ||
            (c >= '\u08A0' && c <= '\u08FF') || (c >= '\uFB50' && c <= '\uFDFF') ||
            (c >= '\uFE70' && c <= '\uFEFF');
    }
}
