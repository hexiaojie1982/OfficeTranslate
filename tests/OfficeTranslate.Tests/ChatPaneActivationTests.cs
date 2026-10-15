using System;
using System.Collections;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.Remoting.Messaging;
using System.Runtime.Remoting.Proxies;
using System.Windows.Forms;
using System.Xml.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeTranslate.Core;

namespace OfficeTranslate.Tests
{
    [TestClass]
    public class ChatPaneActivationTests
    {
        private const BindingFlags PrivateInstance = BindingFlags.NonPublic | BindingFlags.Instance;
        private static Type AddIn(string host) => DocumentChatHostTests.Host(host).GetType("OfficeTranslate." + host + "AddIn.AddIn", true)!;

        [DataTestMethod]
        [DataRow("Word")] [DataRow("Excel")] [DataRow("PowerPoint")]
        public void PaneCreation_UsesDocumentWindowAndReusesHiddenPane(string host)
        {
            DocumentChatHostTests.Sta(() =>
            {
                using (var f = new Activation(host))
                {
                    f.Open();
                    Assert.AreEqual(1, f.Created);
                    Assert.AreEqual(1, f.Panes.Count);
                    Assert.IsTrue(f.Visible);
                    Assert.IsTrue(f.Content.Controls.Count > 0, "Real host control must be initialized.");
                    f.Visible = false;
                    f.Open();
                    Assert.AreEqual(1, f.Created);
                    Assert.IsTrue(f.Visible);
                }
            });
        }

        [DataTestMethod]
        [DataRow("Word", false)] [DataRow("Excel", false)] [DataRow("PowerPoint", false)]
        [DataRow("Word", true)] [DataRow("Excel", true)] [DataRow("PowerPoint", true)]
        public void FailedActivation_DiscardsPaneAndPreservesOriginalFailure(string host, bool failShow)
        {
            DocumentChatHostTests.Sta(() =>
            {
                using (var f = new Activation(host))
                {
                    f.FailShow = failShow; f.FailContent = !failShow; f.FailDelete = true;
                    var error = Assert.ThrowsException<TargetInvocationException>(() => f.Open());
                    Assert.IsInstanceOfType(error.InnerException, failShow ? typeof(COMException) : typeof(InvalidCastException));
                    Assert.AreEqual(0, f.Panes.Count);
                    Assert.AreEqual(1, f.Deleted);
                    f.FailShow = f.FailContent = f.FailDelete = false;
                    f.Open();
                    Assert.AreEqual(2, f.Created);
                    Assert.IsTrue(f.Visible);
                    Assert.AreEqual(1, f.Panes.Count);
                }
            });
        }

        [DataTestMethod]
        [DataRow("Word")] [DataRow("Excel")] [DataRow("PowerPoint")]
        public void PaneDetails_IncludeStageTypeMessageAndStackWithoutSecretsOrSourcePaths(string host)
        {
            var settings = new TranslationSettings { UiLanguage = "en", ApiKey = "fixture-secret-key", ChatQuestionSystemPrompt = "fixture private prompt" };
            Exception failure;
            try { throw new MissingMemberException("Error while invoking HWND. fixture-secret-key fixture private prompt"); }
            catch (Exception error) { failure = error; }
            var details = (string)AddIn(host).GetMethod("ChatPaneDetails", BindingFlags.Static | BindingFlags.NonPublic)!
                .Invoke(null, new object[] { failure, settings, "window-handle" })!;
            StringAssert.Contains(details, "Stage: window-handle");
            StringAssert.Contains(details, "MissingMemberException: Error while invoking HWND.");
            StringAssert.Contains(details, "HRESULT=0x");
            StringAssert.Contains(details, "at OfficeTranslate.Tests.ChatPaneActivationTests.");
            Assert.IsFalse(details.Contains("fixture-secret-key"));
            Assert.IsFalse(details.Contains("fixture private prompt"));
            Assert.IsFalse(details.Contains(".cs:"));
        }

        [DataTestMethod]
        [DataRow("Word")] [DataRow("Excel")] [DataRow("PowerPoint")]
        public void InstallerChatRegistration_MatchesCompiledHostControl(string host)
        {
            var assembly = DocumentChatHostTests.Host(host);
            var control = assembly.GetType("OfficeTranslate.WordAddIn.DocumentChatPane", true)!;
            var guid = "{" + control.GetCustomAttribute<GuidAttribute>()!.Value.ToUpperInvariant() + "}";
            var root = Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "../../../../.."));
            XNamespace ns = "http://schemas.microsoft.com/wix/2006/wi";
            var keys = XDocument.Load(Path.Combine(root, "installer/Product.wxs")).Descendants(ns + "RegistryKey");
            var key = keys.Single(k => (string?)k.Attribute("Key") == "Software\\Classes\\CLSID\\" + guid);
            var server = key.Elements(ns + "RegistryKey").Single(k => (string?)k.Attribute("Key") == "InprocServer32");
            string Value(string name) => (string)server.Elements(ns + "RegistryValue").Single(v => (string?)v.Attribute("Name") == name).Attribute("Value")!;
            Assert.AreEqual(control.FullName, Value("Class"));
            StringAssert.StartsWith(Value("Assembly"), assembly.GetName().Name + ", Version=$(var.ProductVersion),");
            Assert.AreEqual("[#" + host + "AddInDll]", Value("CodeBase"));
            Assert.IsTrue(key.Elements(ns + "RegistryKey").Any(k => (string?)k.Attribute("Key") == "Control"));
        }

        // Proxies implement the actual imported COM interfaces, exercising the real
        // compiled AddIn without installing Office controls. Real IDispatch vs
        // vtable behavior is additionally covered by RunPowerPointChatProbe.ps1.
        private sealed class Proxy : RealProxy
        {
            private readonly Func<IMethodCallMessage, object?> _call;
            internal Proxy(Type type, Func<IMethodCallMessage, object?> call) : base(type) { _call = call; }
            public override IMessage Invoke(IMessage message)
            {
                var call = (IMethodCallMessage)message;
                try { return new ReturnMessage(_call(call), null, 0, call.LogicalCallContext, call); }
                catch (Exception error) { return new ReturnMessage(error, call); }
            }
        }

        private sealed class Activation : IDisposable
        {
            private readonly Type _type;
            private readonly object _addin;
            internal readonly UserControl Content;
            internal readonly IDictionary Panes;
            internal int Created, Deleted;
            internal bool Visible, FailContent, FailShow, FailDelete;
            internal Activation(string host)
            {
                _type = AddIn(host); _addin = Activator.CreateInstance(_type)!;
                var store = _type.GetField(host == "Word" ? "_settingsStore" : "_store", PrivateInstance)!.GetValue(_addin)!;
                typeof(SettingsStore).GetField("_path", PrivateInstance)!.SetValue(store,
                    Path.Combine(Path.GetTempPath(), "OfficeTranslate-ChatActivation-" + Guid.NewGuid() + ".json"));
                Content = (UserControl)Activator.CreateInstance(_type.Assembly.GetType("OfficeTranslate.WordAddIn.DocumentChatPane", true)!)!;
                Panes = (IDictionary)_type.GetField("_chatPanes", PrivateInstance)!.GetValue(_addin)!;
                var interop = Assembly.Load(_type.Assembly.GetReferencedAssemblies().Single(a => a.Name == "Microsoft.Office.Interop." + host));
                var windowType = interop.GetType("Microsoft.Office.Interop." + host + (host == "PowerPoint" ? ".DocumentWindow" : ".Window"), true)!;
                var window = new Proxy(windowType, c => c.MethodName == "get_Hwnd" || c.MethodName == "get_HWND" ? (object)12345 : throw new NotSupportedException(c.MethodName)).GetTransparentProxy();
                var field = _type.GetField(host == "Word" ? "_word" : "_host", PrivateInstance)!;
                field.SetValue(_addin, new Proxy(field.FieldType, c => c.MethodName == "get_ActiveWindow" ? window : throw new NotSupportedException(c.MethodName)).GetTransparentProxy());
                var factoryField = _type.GetField("_chatFactory", PrivateInstance)!;
                var paneType = factoryField.FieldType.GetMethod("CreateCTP")!.ReturnType;
                var pane = new Proxy(paneType, c =>
                {
                    switch (c.MethodName)
                    {
                        case "get_Window": return window;
                        case "get_ContentControl": return FailContent ? new object() : Content;
                        case "set_Visible": if (FailShow) throw new COMException("Fixture show failed."); Visible = (bool)c.Args[0]; return null;
                        case "set_DockPosition": case "set_Width": return null;
                        case "Delete": Deleted++; if (FailDelete) throw new COMException("Fixture delete failed."); return null;
                        default: throw new NotSupportedException(c.MethodName);
                    }
                }).GetTransparentProxy();
                var factory = new Proxy(factoryField.FieldType, c =>
                {
                    Assert.AreEqual("CreateCTP", c.MethodName);
                    Assert.AreEqual("OfficeTranslate." + host + "ChatPane", c.Args[0]);
                    Assert.AreSame(window, c.Args[2]);
                    Created++; return pane;
                }).GetTransparentProxy();
                _type.GetMethod("CTPFactoryAvailable")!.Invoke(_addin, new[] { factory });
            }
            internal void Open() => _type.GetMethod("OpenChatCore", PrivateInstance)!.Invoke(_addin, new object[] { "active-window" });
            public void Dispose() { _type.GetMethod("CloseChat", PrivateInstance)!.Invoke(_addin, null); Content.Dispose(); }
        }
    }
}
