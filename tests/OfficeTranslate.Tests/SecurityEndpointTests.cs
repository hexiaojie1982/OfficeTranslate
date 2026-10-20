using System;
using System.IO;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Windows.Forms;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeTranslate.Core;

namespace OfficeTranslate.Tests
{
    [TestClass]
    public class SecurityEndpointTests
    {
        [DataTestMethod]
        [DataRow("http://localhost:11434", false)]
        [DataRow("http://LOCALHOST./v1", false)]
        [DataRow("http://test.localhost", false)]
        [DataRow("http://127.0.0.1", false)]
        [DataRow("http://127.12.34.56", false)]
        [DataRow("http://10.255.255.255", false)]
        [DataRow("http://172.16.0.1", false)]
        [DataRow("http://172.31.255.255", false)]
        [DataRow("http://192.168.1.2", false)]
        [DataRow("http://169.254.0.1", false)]
        [DataRow("http://[::1]", false)]
        [DataRow("http://[fe80::1234]", false)]
        [DataRow("http://[fd12::1]", false)]
        [DataRow("http://[fc00::1]", false)]
        [DataRow("http://[::ffff:192.168.1.2]", false)]
        [DataRow("http://172.15.255.255", true)]
        [DataRow("http://172.32.0.0", true)]
        [DataRow("http://192.169.0.1", true)]
        [DataRow("http://169.253.0.1", true)]
        [DataRow("http://8.8.8.8/v1", true)]
        [DataRow("http://[2001:4860:4860::8888]", true)]
        [DataRow("http://[::ffff:8.8.8.8]", true)]
        [DataRow("https://8.8.8.8/v1", false)]
        [DataRow("invalid", false)]
        public void HttpKeyWarning_ClassifiesLocalAndPublicEndpoints(string url, bool expected)
        {
            Assert.AreEqual(expected, EndpointSecurity.RequiresPlaintextKeyWarning(url, "fixture-key"));
            Assert.IsFalse(EndpointSecurity.RequiresPlaintextKeyWarning(url, ""));
            Assert.IsFalse(EndpointSecurity.RequiresPlaintextKeyWarning(url, "  "));
        }

        [DataTestMethod]
        [DataRow("10.0.0.1,192.168.1.1", false)]
        [DataRow("fd00::1,fe80::1", false)]
        [DataRow("10.0.0.1,8.8.8.8", true)]
        [DataRow("8.8.8.8", true)]
        [DataRow("", true)]
        public void HttpKeyWarning_LanDnsRequiresAllLocalAddresses(string addresses, bool warning)
        {
            Assert.AreEqual(warning, EndpointSecurity.RequiresPlaintextKeyWarning("http://fixture.invalid/v1", "key",
                _ => addresses.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries).Select(IPAddress.Parse).ToArray()));
        }

        [TestMethod]
        public void HttpKeyWarning_UnresolvableHostStillWarns()
        {
            Assert.IsTrue(EndpointSecurity.RequiresPlaintextKeyWarning("http://fixture.invalid", "key",
                _ => throw new System.Net.Sockets.SocketException()));
        }

        [DataTestMethod]
        [DataRow("Word", false)] [DataRow("Excel", false)] [DataRow("PowerPoint", false)]
        [DataRow("Word", true)] [DataRow("Excel", true)] [DataRow("PowerPoint", true)]
        public void BothSettingsSavePaths_RequireConfirmationAndPreserveDiskOnNo(string host, bool chat)
        {
            DocumentChatHostTests.Sta(() =>
            {
                var path = Path.Combine(Path.GetTempPath(), "security-settings-" + Guid.NewGuid() + ".json");
                var store = new SettingsStore();
                typeof(SettingsStore).GetField("_path", BindingFlags.NonPublic | BindingFlags.Instance)!.SetValue(store, path);
                try
                {
                    store.Save(new TranslationSettings { Provider = ProviderKind.OpenAiCompatible, BaseUrl = "http://8.8.8.8/v1",
                        ApiKey = "fixture-key", Model = "fixture", OpenAiBaseUrl = "http://8.8.8.8/v1", OpenAiApiKey = "fixture-key", OpenAiModel = "fixture" });
                    var before = File.ReadAllBytes(path);
                    var type = DocumentChatHostTests.Host(host).GetType("OfficeTranslate.WordAddIn." + (chat ? "ChatSettingsForm" : "SettingsForm"), true)!;
                    using (var form = (Form)Activator.CreateInstance(type, new object[] { store })!)
                    {
                        var calls = 0;
                        var accept = false;
                        type.GetProperty("ConfirmPlaintextKey", BindingFlags.NonPublic | BindingFlags.Instance)!
                            .SetValue(form, (Func<bool>)(() => { calls++; return accept; }));
                        if (chat)
                        {
                            form.Show();
                            ((Button)form.AcceptButton).PerformClick();
                        }
                        else type.GetMethod("SaveClicked", BindingFlags.NonPublic | BindingFlags.Instance)!.Invoke(form, new object[] { form, EventArgs.Empty });
                        Assert.AreEqual(1, calls); CollectionAssert.AreEqual(before, File.ReadAllBytes(path));
                        Assert.AreNotEqual(DialogResult.OK, form.DialogResult);
                        accept = true;
                        if (chat) ((Button)form.AcceptButton).PerformClick();
                        else type.GetMethod("SaveClicked", BindingFlags.NonPublic | BindingFlags.Instance)!.Invoke(form, new object[] { form, EventArgs.Empty });
                        Assert.AreEqual(2, calls);
                        Assert.AreEqual("http://8.8.8.8/v1", store.Load().BaseUrl);
                        Assert.AreEqual("fixture-key", store.Load().ApiKey);
                        if (!chat) Assert.AreEqual(DialogResult.OK, form.DialogResult);
                        else Assert.IsFalse(form.Visible);
                    }
                }
                finally { if (File.Exists(path)) File.Delete(path); }
            });
        }
    }
}
