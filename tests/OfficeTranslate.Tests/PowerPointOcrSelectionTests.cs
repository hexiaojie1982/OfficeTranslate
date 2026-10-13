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
    public class PowerPointOcrSelectionTests
    {
        // Exercise the compiled OCR collector with strict Office-interface doubles.
        // Unexpected access (including unselected slides or ShapeRange in slide mode)
        // fails instead of silently returning a default value.
        private sealed class OfficeProxy : RealProxy
        {
            private readonly Func<IMethodCallMessage, object?> _call;
            public OfficeProxy(Type type, Func<IMethodCallMessage, object?> call) : base(type) { _call = call; }
            public override IMessage Invoke(IMessage message)
            {
                var call = (IMethodCallMessage)message;
                try { return new ReturnMessage(_call(call), null, 0, call.LogicalCallContext, call); }
                catch (Exception ex) { return new ReturnMessage(ex, call); }
            }
        }

        private static readonly Type Service = LoadService();
        private static readonly Assembly Interop = Service.GetField("_powerPoint", BindingFlags.Instance | BindingFlags.NonPublic)!.FieldType.Assembly;
        private static Type OfficeType(string name) => Interop.GetType("Microsoft.Office.Interop.PowerPoint." + name, true)!;
        private static object Proxy(string type, Func<IMethodCallMessage, object?> call) => new OfficeProxy(OfficeType(type), call).GetTransparentProxy();
        private static object Unexpected(IMethodCallMessage call) => throw new AssertFailedException("Unexpected Office access: " + call.MethodName);
        private static object Collection(string type, params object[] items) => Proxy(type, call =>
            call.MethodName == "GetEnumerator" ? items.GetEnumerator() : Unexpected(call));
        private static object Shape(string kind)
        {
            var shapeType = OfficeType("Shape").GetProperty("Type")!.PropertyType;
            return Proxy("Shape", call => call.MethodName == "get_Type" ? Enum.Parse(shapeType, kind) : Unexpected(call));
        }
        private static object Slide(params object[] shapes) => Proxy("Slide", call =>
            call.MethodName == "get_Shapes" ? Collection("Shapes", shapes) : Unexpected(call));
        private static object Selection(string kind, params object[] items) => Proxy("Selection", call =>
        {
            if (call.MethodName == "get_Type") return Enum.Parse(OfficeType("PpSelectionType"), kind);
            if (kind == "ppSelectionSlides" && call.MethodName == "get_SlideRange") return Collection("SlideRange", items);
            if (kind == "ppSelectionShapes" && call.MethodName == "get_ShapeRange") return Collection("ShapeRange", items);
            return Unexpected(call);
        });
        private static object CreateService(object? selection)
        {
            var window = selection == null ? null : Proxy("DocumentWindow", call =>
                call.MethodName == "get_Selection" ? selection : Unexpected(call));
            var application = Proxy("Application", call => call.MethodName == "get_ActiveWindow" ? window : Unexpected(call));
            return Service.GetConstructor(new[] { OfficeType("Application") })!.Invoke(new[] { application });
        }
        private static IList Read(object service) => (IList)Service.GetMethod("ReadSelectedOcrImages", BindingFlags.Instance | BindingFlags.NonPublic)!.Invoke(service, null)!;
        private static bool HasImages(object service) => (bool)Service.GetMethod("HasSelectedOcrImage")!.Invoke(service, null)!;
        private static Type LoadService()
        {
            var configuration = new DirectoryInfo(AppDomain.CurrentDomain.BaseDirectory).Parent!.Name;
            var path = Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory,
                "../../../../../src/OfficeTranslate.PowerPointAddIn/bin/" + configuration + "/net48/OfficeTranslate.PowerPointAddIn.dll"));
            return Assembly.LoadFrom(path).GetType("OfficeTranslate.PowerPointAddIn.PowerPointTranslationService", true)!;
        }

        [TestMethod]
        public void SelectedSingleSlide_CollectsOnlyPictures()
        {
            var picture = Shape("msoPicture"); var linked = Shape("msoLinkedPicture");
            var service = CreateService(Selection("ppSelectionSlides", Slide(Shape("msoTextBox"), picture, Shape("msoAutoShape"), linked)));
            var images = Read(service);
            Assert.AreEqual(2, images.Count);
            Assert.AreSame(picture, images[0]); Assert.AreSame(linked, images[1]);
            Assert.IsTrue(HasImages(service));
        }

        [TestMethod]
        public void SelectedMultipleSlides_CollectsAllPicturesInSlideOrderAndSkipsEmptySlides()
        {
            var first = Shape("msoPicture"); var second = Shape("msoLinkedPicture"); var third = Shape("msoPicture");
            var service = CreateService(Selection("ppSelectionSlides",
                Slide(first, Shape("msoTextBox")), Slide(), Slide(Shape("msoAutoShape")), Slide(second, third)));
            var images = Read(service);
            Assert.AreEqual(3, images.Count);
            Assert.AreSame(first, images[0]); Assert.AreSame(second, images[1]); Assert.AreSame(third, images[2]);
            Assert.IsTrue(HasImages(service));
        }

        [TestMethod]
        public void SelectedShapes_KeepsExistingPictureFiltering()
        {
            var picture = Shape("msoPicture");
            var service = CreateService(Selection("ppSelectionShapes", Shape("msoTextBox"), picture));
            var images = Read(service);
            Assert.AreEqual(1, images.Count); Assert.AreSame(picture, images[0]);
            Assert.IsTrue(HasImages(service));
        }

        [TestMethod]
        public void SelectedMultipleSlidesWithoutPictures_ReturnsEmptyAndFailsPreflight()
        {
            var service = CreateService(Selection("ppSelectionSlides", Slide(Shape("msoTextBox")), Slide(), Slide(Shape("msoAutoShape"))));
            Assert.AreEqual(0, Read(service).Count); Assert.IsFalse(HasImages(service));
        }

        [DataTestMethod]
        [DataRow("ppSelectionNone")]
        [DataRow("ppSelectionText")]
        public void UnsupportedSelection_DoesNotReadSlideOrShapeRanges(string kind)
        {
            var service = CreateService(Selection(kind));
            Assert.AreEqual(0, Read(service).Count); Assert.IsFalse(HasImages(service));
        }

        [TestMethod]
        public void NoActiveWindow_ReturnsEmptyAndFailsPreflight()
        {
            var service = CreateService(null);
            Assert.AreEqual(0, Read(service).Count); Assert.IsFalse(HasImages(service));
        }

        [TestMethod]
        public void SelectionComFailure_FailsPreflight()
        {
            var selection = Proxy("Selection", call => throw new COMException("Selection unavailable"));
            Assert.IsFalse(HasImages(CreateService(selection)));
        }
    }
}
