using System;
using System.Collections;
using System.IO;
using System.Reflection;
using System.Runtime.Remoting;
using System.Runtime.Remoting.Messaging;
using System.Runtime.Remoting.Proxies;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace OfficeTranslate.Tests
{
    [TestClass]
    public class OcrShapeTextSelectionTests
    {
        // PowerPoint's production group access uses IDispatch.Item dynamically.
        // Provide that member for managed doubles, whose PIA exposes only an indexer.
        public interface IDispatchGroupItems
        {
            object Item(int index);
        }

        // Run the real host collectors and write callbacks with strict Office doubles.
        // Unexpected COM access fails, including text access on skipped overlays.
        private sealed class OfficeProxy : RealProxy, IRemotingTypeInfo
        {
            private readonly Type _type;
            private readonly Type _runtimeType;
            private readonly Func<IMethodCallMessage, object?> _call;
            public OfficeProxy(Type type, Func<IMethodCallMessage, object?> call, Type? runtimeType = null) : base(type)
            { _type = type; _runtimeType = runtimeType ?? type; _call = call; }
            public string TypeName { get; set; } = "Office test double";
            public bool CanCastTo(Type type, object obj) => type == _type || type == _runtimeType || type == typeof(object);
            public override IMessage Invoke(IMessage message)
            {
                var call = (IMethodCallMessage)message;
                try { return new ReturnMessage(call.MethodName == "GetType" ? _runtimeType : _call(call), null, 0, call.LogicalCallContext, call); }
                catch (Exception ex) { return new ReturnMessage(ex, call); }
            }
        }

        private sealed class Host
        {
            public readonly string Name;
            public readonly Type Service;
            private readonly Assembly _interop;
            public Host(string name)
            {
                Name = name;
                var configuration = new DirectoryInfo(AppDomain.CurrentDomain.BaseDirectory).Parent!.Name;
                var assemblyName = "OfficeTranslate." + name + "AddIn";
                var path = Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory,
                    "../../../../../src/" + assemblyName + "/bin/" + configuration + "/net48/" + assemblyName + ".dll"));
                Service = Assembly.LoadFrom(path).GetType(assemblyName + "." + name + "TranslationService", true)!;
                _interop = Service.GetField(name == "Excel" ? "_excel" : "_powerPoint", BindingFlags.Instance | BindingFlags.NonPublic)!.FieldType.Assembly;
            }
            public Type Type(string name) => _interop.GetType("Microsoft.Office.Interop." + Name + "." + name, true)!;
            public object Proxy(string type, Func<IMethodCallMessage, object?> call) => Proxy(Type(type), call);
            public static object Proxy(Type type, Func<IMethodCallMessage, object?> call) => new OfficeProxy(type, call).GetTransparentProxy();
            public static object Unexpected(IMethodCallMessage call) => throw new AssertFailedException("Unexpected Office access: " + call.MethodName);
            public static object EnumProperty(Type type, string property, string value) => Enum.Parse(type.GetProperty(property)!.PropertyType, value);
            public object Collection(string type, params object[] items) => Collection(Type(type), items);
            public static object Collection(Type type, params object[] items) => new OfficeProxy(type, call =>
            {
                if (call.MethodName == "GetEnumerator") return items.GetEnumerator();
                if (call.MethodName == "get_Count") return items.Length;
                if (call.MethodName == "get_Item" || call.MethodName == "Item") return items[Convert.ToInt32(call.Args[0]) - 1];
                return Unexpected(call);
            }, type.FullName == "Microsoft.Office.Interop.PowerPoint.GroupShapes" ? typeof(IDispatchGroupItems) : null).GetTransparentProxy();

            public object TextShape(string text, bool ocr, Action<string> write, bool forbidText = false)
            {
                var shapeType = Type("Shape");
                var frameType = shapeType.GetProperty(Name == "Excel" ? "TextFrame2" : "TextFrame")!.PropertyType;
                var rangeType = frameType.GetProperty("TextRange")!.PropertyType;
                var range = Proxy(rangeType, call =>
                {
                    if (call.MethodName == "get_Text") return text;
                    if (call.MethodName == "set_Text") { write((string)call.Args[0]); return null; }
                    return Unexpected(call);
                });
                var frame = Proxy(frameType, call =>
                {
                    if (call.MethodName == "get_HasText") return EnumProperty(frameType, "HasText", "msoTrue");
                    if (call.MethodName == "get_TextRange") return range;
                    return Unexpected(call);
                });
                return Proxy("Shape", call =>
                {
                    if (call.MethodName == "get_AlternativeText") return ocr ? "OfficeTranslateOCR:owner:region" : "ordinary shape";
                    if (call.MethodName == "get_Tags") return Tags(ocr);
                    if (call.MethodName == "get_Type") return EnumProperty(shapeType, "Type", "msoTextBox");
                    if (call.MethodName == "get_HasTable") return EnumProperty(shapeType, "HasTable", "msoFalse");
                    if (forbidText) return Unexpected(call);
                    if (call.MethodName == "get_HasTextFrame") return EnumProperty(shapeType, "HasTextFrame", "msoTrue");
                    if (call.MethodName == "get_TextFrame" || call.MethodName == "get_TextFrame2") return frame;
                    return Unexpected(call);
                });
            }
            private object Tags(bool ocr) => Proxy("Tags", call =>
            {
                if (call.MethodName == "get_Item")
                {
                    Assert.AreEqual("OfficeTranslateOCR", call.Args[0]);
                    return ocr ? "owner:region" : "";
                }
                return Unexpected(call);
            });
            public object Group(bool ocr, params object[] children) => Proxy("Shape", call =>
            {
                if (call.MethodName == "get_AlternativeText") return ocr ? "OfficeTranslateOCR:group" : "";
                if (call.MethodName == "get_Tags") return Tags(ocr);
                if (call.MethodName == "get_Type") return EnumProperty(Type("Shape"), "Type", "msoGroup");
                if (call.MethodName == "get_HasTable") return EnumProperty(Type("Shape"), "HasTable", "msoFalse");
                if (call.MethodName == "get_GroupItems") return Collection(Type("Shape").GetProperty("GroupItems")!.PropertyType, children);
                return Unexpected(call);
            });
            public object Slide(params object[] shapes) => Proxy("Slide", call =>
                call.MethodName == "get_Shapes" ? Collection("Shapes", shapes) : Unexpected(call));
            public object Selection(bool slides, params object[] shapes)
            {
                if (Name == "Excel") return Proxy("DrawingObjects", call =>
                    call.MethodName == "get_ShapeRange" ? Collection("ShapeRange", shapes) : Unexpected(call));
                return Proxy("Selection", call =>
                {
                    if (call.MethodName == "get_Type") return EnumProperty(Type("Selection"), "Type", slides ? "ppSelectionSlides" : "ppSelectionShapes");
                    if (slides && call.MethodName == "get_SlideRange") return Collection("SlideRange", Slide(shapes));
                    if (!slides && call.MethodName == "get_ShapeRange") return Collection("ShapeRange", shapes);
                    return Unexpected(call);
                });
            }
            public IList ReadSelection(object selection)
            {
                var window = Name == "Excel" ? null : Proxy("DocumentWindow", call =>
                    call.MethodName == "get_Selection" ? selection : Unexpected(call));
                var app = Proxy("Application", call =>
                {
                    if (Name == "Excel" && call.MethodName == "get_Selection") return selection;
                    if (Name == "PowerPoint" && call.MethodName == "get_ActiveWindow") return window;
                    return Unexpected(call);
                });
                var service = Service.GetConstructor(new[] { Type("Application") })!.Invoke(new[] { app });
                return (IList)Service.GetMethod(Name == "Excel" ? "ReadSelectionTargets" : "ReadSelection", BindingFlags.Instance | BindingFlags.NonPublic)!.Invoke(service, null)!;
            }
            public IList ReadWhole(params object[] shapes)
            {
                if (Name == "Excel")
                {
                    var range = Proxy("Range", call => call.MethodName == "get_Areas" ? Collection("Areas") : Unexpected(call));
                    var sheet = Proxy("Worksheet", call =>
                    {
                        if (call.MethodName == "get_UsedRange") return range;
                        if (call.MethodName == "get_Shapes") return Collection("Shapes", shapes);
                        return Unexpected(call);
                    });
                    return (IList)Service.GetMethod("ReadSheetTargets", BindingFlags.Static | BindingFlags.NonPublic)!.Invoke(null, new[] { sheet })!;
                }
                var presentation = Proxy("Presentation", call => call.MethodName == "get_Slides" ? Collection("Slides", Slide(shapes)) : Unexpected(call));
                var app = Proxy("Application", call => call.MethodName == "get_ActivePresentation" ? presentation : Unexpected(call));
                var service = Service.GetConstructor(new[] { Type("Application") })!.Invoke(new[] { app });
                return (IList)Service.GetMethod("ReadPresentation", BindingFlags.Instance | BindingFlags.NonPublic)!.Invoke(service, null)!;
            }
            public object Wrap(int groupDepth, bool taggedGroup, object shape)
            {
                for (var i = 0; i < groupDepth; i++) shape = Group(taggedGroup, shape);
                return shape;
            }
            public void AssertWrite(IList targets, string source, Action assertWritten)
            {
                Assert.AreEqual(1, targets.Count);
                var target = targets[0]!;
                Assert.AreEqual(source, target.GetType().GetProperty("Text")!.GetValue(target));
                target.GetType().GetMethod("Write")!.Invoke(target, Name == "Excel" ? new object[] { "译文", false } : new object[] { "译文" });
                assertWritten();
            }
        }

        [DataTestMethod]
        [DataRow("Excel", 0, false)]
        [DataRow("Excel", 2, false)]
        [DataRow("Excel", 2, true)]
        [DataRow("PowerPoint", 0, false)]
        [DataRow("PowerPoint", 2, false)]
        [DataRow("PowerPoint", 2, true)]
        public void ExplicitOcrShapeSelection_CollectsTextAndWritesTranslation(string name, int depth, bool taggedGroup)
        {
            var host = new Host(name);
            string? written = null;
            var shape = host.Wrap(depth, taggedGroup, host.TextShape("OCR source\r\n", true, text => written = text));
            host.AssertWrite(host.ReadSelection(host.Selection(false, shape)), "OCR source", () => Assert.AreEqual("译文", written));
        }

        [DataTestMethod]
        [DataRow("Excel", 0, false)]
        [DataRow("Excel", 2, false)]
        [DataRow("Excel", 2, true)]
        [DataRow("PowerPoint", 0, false)]
        [DataRow("PowerPoint", 2, false)]
        [DataRow("PowerPoint", 2, true)]
        public void WholeScan_SkipsOcrOverlaysAndStillWritesOrdinaryText(string name, int depth, bool taggedGroup)
        {
            var host = new Host(name);
            string? written = null;
            var overlay = host.Wrap(depth, taggedGroup, host.TextShape("OCR source", true, _ => Assert.Fail("Skipped overlay written"), forbidText: true));
            var ordinary = host.TextShape("ordinary source", false, text => written = text);
            host.AssertWrite(host.ReadWhole(overlay, ordinary), "ordinary source", () => Assert.AreEqual("译文", written));
        }

        [DataTestMethod]
        [DataRow(0, false)]
        [DataRow(2, false)]
        [DataRow(2, true)]
        public void SelectedSlides_SkipOcrOverlaysAndStillWriteOrdinaryText(int depth, bool taggedGroup)
        {
            var host = new Host("PowerPoint");
            string? written = null;
            var overlay = host.Wrap(depth, taggedGroup, host.TextShape("OCR source", true, _ => Assert.Fail("Skipped overlay written"), forbidText: true));
            var ordinary = host.TextShape("ordinary source", false, text => written = text);
            host.AssertWrite(host.ReadSelection(host.Selection(true, overlay, ordinary)), "ordinary source", () => Assert.AreEqual("译文", written));
        }

        [DataTestMethod]
        [DataRow("Excel")]
        [DataRow("PowerPoint")]
        public void ExplicitOrdinaryShapeSelection_StillWritesTranslation(string name)
        {
            var host = new Host(name);
            string? written = null;
            var shape = host.TextShape("ordinary source", false, text => written = text);
            host.AssertWrite(host.ReadSelection(host.Selection(false, shape)), "ordinary source", () => Assert.AreEqual("译文", written));
        }

        [DataTestMethod]
        [DataRow("Excel")]
        [DataRow("PowerPoint")]
        public void ExplicitBlankOcrShapeSelection_RemainsEmpty(string name)
        {
            var host = new Host(name);
            var shape = host.TextShape(" \r\n", true, _ => Assert.Fail("Blank shape written"));
            Assert.AreEqual(0, host.ReadSelection(host.Selection(false, shape)).Count);
        }

        [DataTestMethod]
        [DataRow("Excel")]
        [DataRow("PowerPoint")]
        public void AddShape_DefaultStillExcludesOcrOverlays(string name)
        {
            var host = new Host(name);
            var addShape = host.Service.GetMethod("AddShape", BindingFlags.Static | BindingFlags.NonPublic)!;
            Assert.IsTrue(addShape.GetParameters()[2].IsOptional);
            Assert.AreEqual(false, addShape.GetParameters()[2].DefaultValue);
            var list = Activator.CreateInstance(addShape.GetParameters()[1].ParameterType)!;
            var shape = host.TextShape("OCR source", true, _ => Assert.Fail("Skipped overlay written"), forbidText: true);
            addShape.Invoke(null, new[] { shape, list, System.Type.Missing });
            Assert.AreEqual(0, ((IList)list).Count);
        }
    }
}
