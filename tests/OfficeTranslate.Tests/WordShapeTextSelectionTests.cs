using System;
using System.Collections;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.Remoting.Messaging;
using System.Runtime.Remoting.Proxies;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace OfficeTranslate.Tests
{
    [TestClass]
    public class WordShapeTextSelectionTests
    {
        // Strict Office interface doubles exercise the actual compiled Word
        // collector and the same writeback method used by TranslateAsync.
        private sealed class OfficeProxy : RealProxy
        {
            private readonly Func<IMethodCallMessage, object?> _call;
            public OfficeProxy(Type type, Func<IMethodCallMessage, object?> call) : base(type) => _call = call;
            public override IMessage Invoke(IMessage message)
            {
                var call = (IMethodCallMessage)message;
                try { return new ReturnMessage(_call(call), null, 0, call.LogicalCallContext, call); }
                catch (Exception ex) { return new ReturnMessage(ex, call); }
            }
        }

        private static readonly Type Service = LoadService();
        private static readonly Assembly Interop = Service.GetField("_word", BindingFlags.Instance | BindingFlags.NonPublic)!.FieldType.Assembly;
        private static Type OfficeType(string name) => Interop.GetType("Microsoft.Office.Interop.Word." + name, true)!;
        private static object Proxy(string type, Func<IMethodCallMessage, object?> call) => new OfficeProxy(OfficeType(type), call).GetTransparentProxy();
        private static object Unexpected(IMethodCallMessage call) => throw new AssertFailedException("Unexpected Office access: " + call.MethodName);
        private static object Collection(string type, params object[] items) => Proxy(type, call =>
            call.MethodName == "GetEnumerator" ? items.GetEnumerator() :
            call.MethodName == "get_Item" ? items[(int)call.Args[0] - 1] : Unexpected(call));

        private sealed class TextRange
        {
            public string Text;
            public int Start;
            public int End;
            public readonly object Value;
            public object Font;
            public object Paragraph;
            public bool SameStory;
            public bool IsTextFrame;
            public int Writes;
            public int Appends;
            public TextRange(string text)
            {
                Text = text; Start = 0; End = text.Length;
                object? font = null, paragraph = null;
                font = Proxy("Font", call => call.MethodName == "get_Duplicate" ? font : Unexpected(call));
                paragraph = Proxy("ParagraphFormat", call => call.MethodName == "get_Duplicate" ? paragraph : Unexpected(call));
                Font = font; Paragraph = paragraph;
                Value = Proxy("Range", call =>
                {
                    switch (call.MethodName)
                    {
                        case "get_Text": return Text;
                        case "set_Text":
                            Text = (string)call.Args[0]; Writes++;
                            // Model replacement resetting style: the host must restore it.
                            Font = null!; Paragraph = null!; return null;
                        case "get_Start": return Start;
                        case "get_End": return End;
                        case "set_End": End = (int)call.Args[0]; return null;
                        case "get_Duplicate": return Value;
                        case "get_Characters": return Collection("Characters", Value!);
                        case "get_Font": return Font;
                        case "set_Font": Font = call.Args[0]; return null;
                        case "get_ParagraphFormat": return Paragraph;
                        case "set_ParagraphFormat": Paragraph = call.Args[0]; return null;
                        case "get_StoryType": return Enum.Parse(OfficeType("WdStoryType"), IsTextFrame ? "wdTextFrameStory" : "wdMainTextStory");
                        case "InStory": return SameStory;
                        case "Collapse": return null;
                        case "InsertAfter": Text = Text.Substring(0, End) + (string)call.Args[0] + Text.Substring(End); Appends++; return null;
                        default: return Unexpected(call);
                    }
                });
            }
        }

        private static object Shape(TextRange? text, bool unsupported = false) => Proxy("Shape", call =>
        {
            if (unsupported) throw new COMException("No text frame");
            // Type and AlternativeText must never filter text targets, including OCR boxes.
            if (call.MethodName == "get_TextFrame") return Proxy("TextFrame", frame =>
                frame.MethodName == "get_HasText" ? (text == null ? 0 : -1) :
                frame.MethodName == "get_TextRange" ? text!.Value : Unexpected(frame));
            return Unexpected(call);
        });

        private static object CreateService(TextRange body, bool shapeSelection, object[]? shapes)
        {
            var selection = Proxy("Selection", call =>
            {
                if (call.MethodName == "get_Range") return body.Value;
                if (call.MethodName == "get_Type") return Enum.Parse(OfficeType("WdSelectionType"), shapeSelection ? "wdSelectionShape" : "wdSelectionNormal");
                if (call.MethodName == "get_ShapeRange")
                {
                    if (shapes == null) throw new COMException("No selected shapes");
                    return Collection("ShapeRange", shapes);
                }
                return Unexpected(call);
            });
            var application = Proxy("Application", call => call.MethodName == "get_Selection" ? selection : Unexpected(call));
            return Service.GetConstructor(new[] { OfficeType("Application") })!.Invoke(new[] { application });
        }
        private static IList Read(object service) => (IList)Service.GetMethod("ReadSelection", BindingFlags.Instance | BindingFlags.NonPublic)!.Invoke(service, null)!;
        private static string Source(object target) => (string)target.GetType().GetProperty("Text")!.GetValue(target)!;
        private static void Write(object target, string translated, bool bilingual = false) =>
            Service.GetMethod("WriteTarget", BindingFlags.Static | BindingFlags.NonPublic)!.Invoke(null, new[] { target, translated, (object)bilingual });
        private static Type LoadService()
        {
            var configuration = new DirectoryInfo(AppDomain.CurrentDomain.BaseDirectory).Parent!.Name;
            var path = Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory,
                "../../../../../src/OfficeTranslate.WordAddIn/bin/" + configuration + "/net48/OfficeTranslate.WordAddIn.dll"));
            return Assembly.LoadFrom(path).GetType("OfficeTranslate.WordAddIn.WordTranslationService", true)!;
        }

        [TestMethod]
        public void SelectedTextBox_WithEmptySelectionRange_ReadsAndWritesShapePreservingBaseFormat()
        {
            var body = new TextRange(""); var box = new TextRange("OCR source text.\r");
            var font = box.Font; var paragraph = box.Paragraph;
            var targets = Read(CreateService(body, true, new[] { Shape(box) }));
            Assert.AreEqual(1, targets.Count); Assert.AreEqual("OCR source text.", Source(targets[0]!));
            Write(targets[0]!, "识别文字的译文。");
            Assert.AreEqual("识别文字的译文。", box.Text); Assert.AreEqual(1, box.Writes);
            Assert.AreSame(font, box.Font); Assert.AreSame(paragraph, box.Paragraph);
            Assert.AreEqual(0, body.Writes);
        }

        [TestMethod]
        public void ShapeSelection_DoesNotTranslateUnselectedAnchorText()
        {
            var body = new TextRange("Unselected anchor paragraph.");
            var targets = Read(CreateService(body, true, new[] { Shape(new TextRange("Selected box.")) }));
            Assert.AreEqual(1, targets.Count); Assert.AreEqual("Selected box.", Source(targets[0]!));
            Write(targets[0]!, "文本框译文。"); Assert.AreEqual(0, body.Writes);
        }

        [TestMethod]
        public void MultipleSelectedTextShapes_IgnoreEmptyAndUnsupportedShapes_WriteToTheirOwnFrames()
        {
            var first = new TextRange("First shape."); var second = new TextRange("Second shape.");
            var targets = Read(CreateService(new TextRange(""), true, new[] {
                Shape(null), Shape(first), Shape(null, true), Shape(new TextRange(" \r")), Shape(second) }));
            Assert.AreEqual(2, targets.Count);
            Write(targets[0]!, "第一个。"); Write(targets[1]!, "第二个。");
            Assert.AreEqual("第一个。", first.Text); Assert.AreEqual("第二个。", second.Text);
        }

        [TestMethod]
        public void MixedBodyAndSelectedShape_BothTranslateAndWriteToCorrectTargets()
        {
            var body = new TextRange("Body sentence.\r"); var box = new TextRange("Shape sentence.\r");
            var targets = Read(CreateService(body, false, new[] { Shape(box) }));
            Assert.AreEqual(2, targets.Count);
            Assert.AreEqual("Body sentence.", Source(targets[0]!)); Assert.AreEqual("Shape sentence.", Source(targets[1]!));
            Write(targets[0]!, "正文译文。"); Write(targets[1]!, "文本框译文。");
            Assert.AreEqual("正文译文。", body.Text); Assert.AreEqual("文本框译文。", box.Text);
        }

        [TestMethod]
        public void OrdinaryBodySelection_WithoutShapeRange_KeepsExistingWriteback()
        {
            var body = new TextRange("Selected body.");
            var targets = Read(CreateService(body, false, null));
            Assert.AreEqual(1, targets.Count); Write(targets[0]!, "正文译文。");
            Assert.AreEqual("正文译文。", body.Text);
        }

        [DataTestMethod]
        [DataRow("")]
        [DataRow(" \r")]
        [DataRow("\uFFFC\r")]
        public void SelectedShapesWithoutTranslatableText_ReturnEmpty(string text)
        {
            Assert.AreEqual(0, Read(CreateService(new TextRange(""), true, new[] { Shape(new TextRange(text)), Shape(null) })).Count);
        }

        [TestMethod]
        public void SelectedTextInsideBox_DoesNotAlsoTranslateWholeFrame()
        {
            var selected = new TextRange("Selected words") { IsTextFrame = true, SameStory = true };
            var box = new TextRange("Selected words and unselected remainder.");
            var targets = Read(CreateService(selected, false, new[] { Shape(box) }));
            Assert.AreEqual(1, targets.Count); Assert.AreEqual("Selected words", Source(targets[0]!));
        }

        [DataTestMethod]
        [DataRow(true)]
        [DataRow(false)]
        public void BilingualMode_AppendsTranslationInsideCorrectTarget(bool shapeSelection)
        {
            var body = new TextRange(shapeSelection ? "" : "Original sentence.\r");
            var box = new TextRange("Original sentence.\r");
            var range = shapeSelection ? box : body; var font = range.Font;
            var targets = Read(CreateService(body, shapeSelection, shapeSelection ? new[] { Shape(box) } : null));
            Write(targets[0]!, "译文。", true);
            Assert.AreEqual("Original sentence.\r译文。\r", range.Text);
            Assert.AreEqual(0, range.Writes); Assert.AreEqual(1, range.Appends); Assert.AreSame(font, range.Font);
        }
    }
}
