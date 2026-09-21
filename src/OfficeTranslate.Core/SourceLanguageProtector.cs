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
        private readonly string _originalText;
        private readonly IReadOnlyList<KeyValuePair<string, string>> _segments;

        internal ProtectedTranslationText(string originalText, string text, IReadOnlyList<KeyValuePair<string, string>> segments)
        {
            _originalText = originalText;
            Text = text;
            _segments = segments;
        }

        public string OriginalText => _originalText;
        public string Text { get; }
        public bool HasProtectedSegments => _segments.Count > 0;
        public bool IsFullyProtected => _segments.Count == 1 && Text == _segments[0].Key;

        internal bool PreservesProtectedSegmentsInOrder(string candidate)
        {
            if (candidate == null) return false;
            var position = 0;
            foreach (var segment in _segments)
            {
                var found = candidate.IndexOf(segment.Value, position, StringComparison.Ordinal);
                if (found < 0) return false;
                position = found + segment.Value.Length;
            }
            return true;
        }

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
            return Restore(translatedText, false);
        }

        public string Restore(string translatedText, bool separateLatinWordBoundaries)
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
                    if (atStart) restored = JoinWithBoundary(segment.Value, restored, separateLatinWordBoundaries);
                    else if (atEnd) restored = JoinWithBoundary(restored, segment.Value, separateLatinWordBoundaries);
                    else throw new ProtectedContentException("翻译模型修改了受保护的非源语言内容。");
                    continue;
                }
                restored = ReplaceToken(restored, segment.Key, segment.Value, separateLatinWordBoundaries);
            }
            return restored;
        }

        private static string ReplaceToken(string text, string token, string value, bool separateLatinWordBoundaries)
        {
            var position = 0;
            while ((position = text.IndexOf(token, position, StringComparison.Ordinal)) >= 0)
            {
                var replacement = value;
                if (separateLatinWordBoundaries && replacement.Length > 0)
                {
                    if (position > 0 && NeedsBoundarySpace(text[position - 1], replacement[0]))
                        replacement = " " + replacement;
                    var afterToken = position + token.Length;
                    if (afterToken < text.Length && NeedsBoundarySpace(replacement[replacement.Length - 1], text[afterToken]))
                        replacement += " ";
                }
                text = text.Remove(position, token.Length).Insert(position, replacement);
                position += replacement.Length;
            }
            return text;
        }

        private static string JoinWithBoundary(string left, string right, bool enabled)
        {
            if (!enabled || string.IsNullOrEmpty(left) || string.IsNullOrEmpty(right)) return left + right;
            return NeedsBoundarySpace(left[left.Length - 1], right[0]) ? left + " " + right : left + right;
        }

        internal static bool NeedsBoundarySpace(char left, char right)
        {
            return IsLatinOrDigit(left) && IsLatinOrDigit(right);
        }

        private static bool IsLatinOrDigit(char value)
        {
            return char.IsDigit(value) ||
                (value >= '\u0041' && value <= '\u024F') ||
                (value >= '\u1E00' && value <= '\u1EFF') ||
                (value >= '\uFF21' && value <= '\uFF5A');
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
            if (string.IsNullOrEmpty(text))
                return new ProtectedTranslationText(text ?? string.Empty, text ?? string.Empty, new KeyValuePair<string, string>[0]);

            if (IsAutomatic(sourceLanguage))
                return ProtectLayoutSeparators(text);

            if (!ContainsSourceLetter(text, sourceLanguage))
            {
                var wholeToken = CreateUniqueToken(text, new KeyValuePair<string, string>[0], 1);
                return new ProtectedTranslationText(text, wholeToken,
                    new[] { new KeyValuePair<string, string>(wholeToken, text) });
            }

            var output = new StringBuilder(text.Length);
            var protectedSegments = new List<KeyValuePair<string, string>>();
            var position = 0;
            while (position < text.Length)
            {
                if (TryReadLayoutSeparator(text, position, out var separator))
                {
                    var layoutToken = CreateUniqueToken(text, protectedSegments, protectedSegments.Count + 1);
                    protectedSegments.Add(new KeyValuePair<string, string>(layoutToken, separator));
                    output.Append(layoutToken);
                    position += separator.Length;
                    continue;
                }

                if (IsSourceLetter(text[position], sourceLanguage))
                {
                    output.Append(text[position]);
                    position++;
                    continue;
                }

                if (!char.IsLetterOrDigit(text[position]))
                {
                    output.Append(text[position]);
                    position++;
                    continue;
                }

                var start = position;
                var containsForeignLetter = false;
                while (position < text.Length && IsForeignTokenCharacter(text, position, sourceLanguage))
                {
                    if (char.IsLetter(text[position]) && !IsSourceLetter(text[position], sourceLanguage))
                        containsForeignLetter = true;
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

            return new ProtectedTranslationText(text, output.ToString(), protectedSegments);
        }

        private static ProtectedTranslationText ProtectLayoutSeparators(string text)
        {
            var output = new StringBuilder(text.Length);
            var protectedSegments = new List<KeyValuePair<string, string>>();
            var position = 0;
            while (position < text.Length)
            {
                if (!TryReadLayoutSeparator(text, position, out var separator))
                {
                    output.Append(text[position]);
                    position++;
                    continue;
                }

                var token = CreateUniqueToken(text, protectedSegments, protectedSegments.Count + 1);
                protectedSegments.Add(new KeyValuePair<string, string>(token, separator));
                output.Append(token);
                position += separator.Length;
            }
            return new ProtectedTranslationText(text, output.ToString(), protectedSegments);
        }

        private static bool TryReadLayoutSeparator(string text, int position, out string separator)
        {
            separator = string.Empty;
            if (position < 0 || position >= text.Length) return false;
            var character = text[position];
            if (character == '\r')
            {
                separator = position + 1 < text.Length && text[position + 1] == '\n' ? "\r\n" : "\r";
                return true;
            }
            if (character == '\n' || character == '\v' || character == '\f' ||
                character == '\u2028' || character == '\u2029')
            {
                separator = character.ToString();
                return true;
            }
            return false;
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

        private static bool ContainsSourceLetter(string text, string sourceLanguage)
        {
            foreach (var character in text)
                if (IsSourceLetter(character, sourceLanguage)) return true;
            return false;
        }

        private static bool IsForeignTokenCharacter(string text, int position, string sourceLanguage)
        {
            var character = text[position];
            if (IsSourceLetter(character, sourceLanguage) || char.IsWhiteSpace(character)) return false;
            if (char.IsLetterOrDigit(character) || IsCombiningMark(character)) return true;
            if (!IsWordConnector(character) || position + 1 >= text.Length) return false;
            var next = text[position + 1];
            return !IsSourceLetter(next, sourceLanguage) &&
                (char.IsLetterOrDigit(next) || IsCombiningMark(next));
        }

        private static bool IsWordConnector(char character)
        {
            switch (character)
            {
                case '-': case '_': case '/': case '\\': case '.': case '@': case '+':
                case '#': case '%': case '&': case '=': case ':':
                    return true;
                default:
                    return false;
            }
        }

        private static bool IsCombiningMark(char character)
        {
            var category = char.GetUnicodeCategory(character);
            return category == System.Globalization.UnicodeCategory.NonSpacingMark ||
                category == System.Globalization.UnicodeCategory.SpacingCombiningMark ||
                category == System.Globalization.UnicodeCategory.EnclosingMark;
        }

        internal static int CountDistinctiveSourceLetters(string text, string sourceLanguage)
        {
            if (string.IsNullOrEmpty(text)) return 0;
            var count = 0;
            foreach (var character in text)
            {
                switch (sourceLanguage)
                {
                    case "简体中文":
                    case "繁體中文":
                        if (IsHan(character) || IsBopomofo(character)) count++;
                        break;
                    case "韩语":
                        if (IsHangul(character)) count++;
                        break;
                    case "日语":
                        if (IsHiragana(character) || IsKatakana(character)) count++;
                        break;
                    case "俄语":
                        if (IsCyrillic(character)) count++;
                        break;
                    case "阿拉伯语":
                        if (IsArabic(character)) count++;
                        break;
                    case "英语":
                    case "法语":
                    case "德语":
                    case "西班牙语":
                    case "葡萄牙语":
                    case "意大利语":
                        if (IsLatin(character)) count++;
                        break;
                }
            }
            return count;
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
