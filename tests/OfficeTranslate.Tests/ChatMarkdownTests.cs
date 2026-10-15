using System;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeTranslate.Core;
using OfficeTranslate.WordAddIn;

namespace OfficeTranslate.Tests
{
    [TestClass]
    public class ChatMarkdownTests
    {
        [DataTestMethod]
        [DataRow("**加粗** 和 *斜体*", "加粗 和 斜体")]
        [DataRow("***both***", "both")]
        [DataRow("**bold *italic***", "bold italic")]
        [DataRow("*italic **bold***", "italic bold")]
        [DataRow("`**literal**` and ``a ` b``", "**literal** and a ` b")]
        [DataRow(@"\*literal\* \\ path C:\test\file", @"*literal* \ path C:\test\file")]
        [DataRow(@"\cf0 \cf3 {\rtf1}", @"\cf0 \cf3 {\rtf1}")]
        [DataRow("**unfinished", "**unfinished")]
        [DataRow("text * unfinished", "text * unfinished")]
        [DataRow("a\r\nb\r\n\r\nc\rd", "a\nb\n\nc\nd")]
        [DataRow("# 一级\n## 二级\n###### 六级 ###", "一级\n二级\n六级")]
        [DataRow("#hashtag\n####### literal", "#hashtag\n####### literal")]
        [DataRow("- one\n* two\n+ three\n1. first\n2) second", "•\tone\n•\ttwo\n•\tthree\n1.\tfirst\n2.\tsecond")]
        [DataRow("> **引用**\n>> nested\n> - item", "引用\nnested\n•\titem")]
        [DataRow("```csharp\n**literal**\n{\\rtf1} 😀\n```\nnormal", "**literal**\n{\\rtf1} 😀\nnormal")]
        [DataRow("```\npartial **code", "partial **code")]
        [DataRow("````text\n```\n# literal\n````\n*normal*", "```\n# literal\nnormal")]
        [DataRow("~~~\ncode\n~~~", "code\n")]
        [DataRow("<img src='https://example.com/x'> 中文😀 {\\b injected}", "<img src='https://example.com/x'> 中文😀 {\\b injected}")]
        [DataRow("", "")]
        [DataRow("a\n", "a\n")]
        [DataRow("\n", "\n")]
        [DataRow("a\n\n", "a\n\n")]
        public void Parser_RequiredSyntaxEscapesAndIncompleteStreams(string source, string expected)
        {
            var paragraphs = ChatMarkdown.Parse(source, 18);
            Assert.AreEqual(expected, string.Concat(paragraphs.Select(p => p.Text)));
            DocumentChatHostTests.Sta(() =>
            {
                using (var box = new RichTextBox())
                {
                    box.SelectedRtf = ChatMarkdown.Document(paragraphs);
                    Assert.AreEqual(expected, box.Text, "Native RTF text and parser offsets must agree, including Unicode and newlines.");
                }
            });
        }
        private static void Select(RichTextBox box, string value) => box.Select(box.Text.IndexOf(value, StringComparison.Ordinal), value.Length);

        [TestMethod]
        public void NativeRtf_HasRealFontStylesHeadingListQuoteAndGrayCode()
        {
            DocumentChatHostTests.Sta(() =>
            {
                using (var box = new RichTextBox())
                {
                    box.Rtf = ChatMarkdown.Document(ChatMarkdown.Parse("# 标题\n**bold *both*** *italic* plain\n`inline`\n```cs\ncode\n```\n- bullet\n1. numbered\n> quote\n正文", 18));
                    Select(box, "标题"); Assert.IsTrue(box.SelectionFont!.Bold); Assert.IsTrue(box.SelectionFont.SizeInPoints > 9F);
                    Select(box, "bold"); Assert.IsTrue(box.SelectionFont!.Bold); Assert.IsFalse(box.SelectionFont.Italic);
                    Select(box, "both"); Assert.IsTrue(box.SelectionFont!.Bold); Assert.IsTrue(box.SelectionFont.Italic);
                    Select(box, "italic"); Assert.IsTrue(box.SelectionFont!.Italic); Assert.IsFalse(box.SelectionFont.Bold);
                    foreach (var code in new[] { "inline", "code" })
                    { Select(box, code); Assert.AreEqual(Color.FromArgb(242, 244, 247), box.SelectionBackColor); }
                    foreach (var list in new[] { "bullet", "numbered" })
                    { Select(box, list); Assert.IsTrue(box.SelectionIndent > 0); Assert.IsTrue(box.SelectionHangingIndent > 0); }
                    Select(box, "quote"); Assert.IsTrue(box.SelectionIndent > 0); Assert.AreEqual(Color.FromArgb(90, 104, 122), box.SelectionColor);
                    Select(box, "正文"); Assert.IsFalse(box.SelectionFont!.Bold); Assert.IsFalse(box.SelectionFont.Italic);
                    Assert.AreEqual("Microsoft YaHei UI", box.SelectionFont.Name); Assert.AreEqual(9F, box.SelectionFont.SizeInPoints);
                    Assert.AreNotEqual(Color.FromArgb(242, 244, 247), box.SelectionBackColor); Assert.AreEqual(0, box.SelectionIndent);
                }
            });
        }
        [TestMethod]
        public void Message_UpdatesUnclosedMarkupInPlace_AndPreservesPlainQuestionsAndRawHistory()
        {
            DocumentChatHostTests.Sta(() =>
            {
                using (var box = new ChatTranscriptBox { Font = new Font("Microsoft YaHei UI", 9F) })
                {
                    var question = "## user **plain** `code`\n- literal";
                    box.ShowMessage(question, false);
                    Assert.AreEqual(question, box.Text);
                    Select(box, "user **plain**"); Assert.IsFalse(box.SelectionFont!.Bold);
                    foreach (var snapshot in new[] { "## Heading\n**bo", "## Heading\n**bold**", "## Heading\n**bold**\n```cs\ncode", "## Heading\n**bold**\n```cs\ncode\n```\nnormal" })
                    {
                        box.ShowMessage(snapshot, true);
                        Assert.AreEqual(string.Concat(ChatMarkdown.Parse(snapshot, 18).Select(p => p.Text)), box.Text);
                    }
                    Select(box, "bold"); Assert.IsTrue(box.SelectionFont!.Bold);
                    box.ShowMessage("recovered", true);
                    Select(box, "recovered"); Assert.IsFalse(box.SelectionFont!.Bold); Assert.AreEqual(0, box.SelectionIndent);
                    var changes = 0; box.TextChanged += (s, e) => changes++;
                    box.ShowMessage("recovered", true); Assert.AreEqual(0, changes);
                    box.ShowMessage("", true); Assert.AreEqual("", box.Text);
                    var message = new ChatMessage("assistant", "*answer*");
                    box.ShowMessage(message.Content, true, "\nerror **plain**");
                    Select(box, "answer"); Assert.IsTrue(box.SelectionFont!.Italic);
                    Select(box, "error **plain**"); Assert.IsFalse(box.SelectionFont!.Bold);
                    Assert.AreEqual("*answer*", message.Content);
                    box.ShowMessage("> last quote", true);
                    Select(box, "last quote"); Assert.IsTrue(box.SelectionIndent > 0);
                    box.ShowMessage("normal\nplain", true);
                    box.ShowMessage("normal\n> last quote", true);
                    Select(box, "last quote"); Assert.IsTrue(box.SelectionIndent > 0);
                }
            });
        }
        [TestMethod]
        public void Rtf_UntrustedControlWordsRemainLiteral_AndDoNotBecomeLinksOrImages()
        {
            DocumentChatHostTests.Sta(() =>
            {
                const string raw = @"{\rtf1\b attack} \cf0 \cf3 <script>alert(1)</script> ![x](https://example.com/x)";
                using (var box = new ChatTranscriptBox())
                {
                    box.ShowMessage(raw, true);
                    StringAssert.Contains(box.Text, raw); Select(box, "attack"); Assert.IsFalse(box.SelectionFont!.Bold); Assert.IsFalse(box.DetectUrls);
                }
            });
        }
    }
}
