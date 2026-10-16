using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeTranslate.Core;
using OfficeTranslate.WordAddIn;

namespace OfficeTranslate.Tests
{
    [TestClass]
    public class ChatTranslationSettingsTests
    {
        private const BindingFlags Hidden = BindingFlags.Instance | BindingFlags.NonPublic;
        private static object Field(object target, string name) => target.GetType().GetField(name, Hidden)!.GetValue(target)!;
        private static object? Call(object target, string name, params object[] args) => target.GetType().GetMethod(name, Hidden)!.Invoke(target, args);
        private static string Baseline(TranslationSettings settings, ChatMode mode) =>
            (mode == ChatMode.Question ? settings.ChatQuestionSystemPrompt : mode == ChatMode.Summarize ? settings.ChatSummarySystemPrompt : settings.ChatRewriteSystemPrompt)
            + "\n" + ChatPrompts.Safety(settings.UiLanguage);
        private static TranslationSettings Settings() => new TranslationSettings
        {
            TranslationStyle = "Technology", CustomInstructions = "Keep product codes unchanged.", Glossary = "wafer = 晶圆",
            ChatContextTokens = 8192, ChatOutputTokens = 512
        };
        private static Task<ChatReply> Send(DocumentChatSession session, TranslationSettings settings, ChatMode mode, ChatTransport transport, string text = "document") =>
            session.SendAsync(new ChatContext(text, false), "question", mode, settings, transport, _ => { }, (i, n) => { }, CancellationToken.None);

        [DataTestMethod]
        [DataRow("zh-CN", ChatMode.Question)] [DataRow("en", ChatMode.Question)]
        [DataRow("zh-CN", ChatMode.Summarize)] [DataRow("en", ChatMode.Summarize)]
        [DataRow("zh-CN", ChatMode.Rewrite)] [DataRow("en", ChatMode.Rewrite)]
        public async Task Toggle_PreservesExactDisabledPrompt_AndUsesLatestSharedValues(string language, ChatMode mode)
        {
            var settings = Settings(); settings.UiLanguage = language;
            settings.ChatQuestionSystemPrompt = "custom question prompt";
            settings.ChatSummarySystemPrompt = "custom summary prompt";
            settings.ChatRewriteSystemPrompt = "custom rewrite prompt";
            var original = Baseline(settings, mode);
            var session = new DocumentChatSession();
            var systems = new List<string>();
            ChatTransport transport = (m, s, n, o, t) => { systems.Add(m[0].Content); return Task.FromResult(new ChatReply("answer", false)); };
            Assert.IsFalse(settings.ChatUseTranslationSettings);
            await Send(session, settings, mode, transport);
            Assert.AreEqual(original, systems.Last());
            settings.ChatUseTranslationSettings = true;
            await Send(session, settings, mode, transport);
            var expected = original + "\nTranslation style:\n" + TranslationStyleCatalog.GetPrompt("Technology")
                + "\nAdditional requirements:\n" + settings.CustomInstructions + "\nGlossary:\n" + settings.Glossary;
            Assert.AreEqual(expected, systems.Last());
            settings.TranslationStyle = "AcademicPaper"; settings.CustomInstructions = "Use short sentences."; settings.Glossary = "die = 裸片";
            await Send(session, settings, mode, transport);
            StringAssert.Contains(systems.Last(), TranslationStyleCatalog.GetPrompt("AcademicPaper"));
            StringAssert.Contains(systems.Last(), settings.CustomInstructions); StringAssert.Contains(systems.Last(), settings.Glossary);
            Assert.IsFalse(systems.Last().Contains("wafer")); Assert.IsFalse(systems.Last().Contains("Keep product codes"));
            settings.ChatUseTranslationSettings = false;
            await Send(session, settings, mode, transport);
            Assert.AreEqual(original, systems.Last());
            Assert.AreEqual(original, Baseline(settings, mode), "Never write the composed prompt back into Chat settings.");
            Assert.IsFalse(session.History.Any(m => m.Content.Contains("wafer") || m.Content.Contains("Translation style:")));
        }

        [TestMethod]
        public async Task CustomStyle_EmptyValuesAddNothing_AndPresetInstructionsAreNotDuplicated()
        {
            var settings = Settings(); settings.ChatUseTranslationSettings = true; settings.TranslationStyle = "Custom";
            settings.CustomInstructions = " \r\n"; settings.Glossary = null!;
            await Send(new DocumentChatSession(), settings, ChatMode.Question, (m, s, n, o, t) =>
            { Assert.AreEqual(Baseline(settings, ChatMode.Question), m[0].Content); return Task.FromResult(new ChatReply("answer", false)); });
            settings.CustomInstructions = "Use a friendly tone.";
            await Send(new DocumentChatSession(), settings, ChatMode.Question, (m, s, n, o, t) =>
            { Assert.AreEqual(Baseline(settings, ChatMode.Question) + "\nAdditional requirements:\nUse a friendly tone.", m[0].Content); return Task.FromResult(new ChatReply("answer", false)); });
            settings.TranslationStyle = "Technology"; settings.CustomInstructions = TranslationStyleCatalog.GetPrompt("Technology");
            await Send(new DocumentChatSession(), settings, ChatMode.Question, (m, s, n, o, t) =>
            { Assert.AreEqual(Baseline(settings, ChatMode.Question) + "\nTranslation style:\n" + settings.CustomInstructions, m[0].Content); return Task.FromResult(new ChatReply("answer", false)); });
        }

        [DataTestMethod]
        [DataRow(false)] [DataRow(true)]
        public async Task LongDocument_AppliesChoiceToMapAndFinalRequests_WithinBudget(bool enabled)
        {
            var settings = Settings(); settings.ChatUseTranslationSettings = enabled;
            var maps = 0; var finals = 0;
            await Send(new DocumentChatSession(), settings, ChatMode.Summarize, (m, s, n, o, t) =>
            {
                Assert.IsTrue(ChatBudget.Cost(m) + n <= settings.ChatContextTokens);
                if (n == 256) { maps++; StringAssert.StartsWith(m[0].Content, ChatPrompts.Map(settings.UiLanguage)); }
                else { finals++; StringAssert.StartsWith(m[0].Content, Baseline(settings, ChatMode.Summarize)); }
                Assert.AreEqual(enabled, m[0].Content.Contains(settings.CustomInstructions));
                Assert.AreEqual(enabled, m[0].Content.Contains(settings.Glossary));
                Assert.AreEqual(enabled, m[0].Content.Contains(TranslationStyleCatalog.GetPrompt(settings.TranslationStyle)));
                return Task.FromResult(new ChatReply("short summary", false));
            }, new string('文', 8000));
            Assert.IsTrue(maps > 1); Assert.AreEqual(1, finals);
        }

        [TestMethod]
        public async Task OversizedTranslationSettings_FailBeforeTransport_AndDiagnosticsRedactTheirContents()
        {
            var settings = Settings(); settings.ChatUseTranslationSettings = true;
            var calls = 0; settings.Glossary = new string('文', 8000);
            var ex = await Assert.ThrowsExceptionAsync<ChatException>(() => Send(new DocumentChatSession(), settings, ChatMode.Question,
                (m, s, n, o, t) => { calls++; return Task.FromResult(new ChatReply("answer", false)); }));
            Assert.AreEqual(ChatFailure.Budget, ex.Failure); Assert.AreEqual(0, calls);
            settings.Glossary = "private glossary line one\nprivate glossary line two";
            var diagnostic = ChatErrors.Describe(new IOException(TranslationStyleCatalog.GetPrompt(settings.TranslationStyle)
                + settings.CustomInstructions + settings.Glossary + " private glossary line two"), settings);
            Assert.IsFalse(diagnostic.Contains("private glossary")); Assert.IsFalse(diagnostic.Contains(settings.CustomInstructions));
            Assert.IsFalse(diagnostic.Contains(TranslationStyleCatalog.GetPrompt(settings.TranslationStyle)));
            StringAssert.Contains(diagnostic, "[redacted]");
        }

        private sealed class SettingsSandbox : IDisposable
        {
            internal readonly string DirectoryPath = Path.Combine(Path.GetTempPath(), "OfficeTranslate-ChatSettings-" + Guid.NewGuid().ToString("N"));
            internal readonly SettingsStore Store = new SettingsStore();
            internal string PathName => Path.Combine(DirectoryPath, "settings.json");
            internal SettingsSandbox() { Directory.CreateDirectory(DirectoryPath); typeof(SettingsStore).GetField("_path", Hidden)!.SetValue(Store, PathName); }
            public void Dispose() { Directory.Delete(DirectoryPath, true); }
        }

        [DataTestMethod]
        [DataRow(false)] [DataRow(true)]
        public void SettingsStore_RoundTripsChoice_AndOldFilesDefaultToOff(bool enabled)
        {
            using (var sandbox = new SettingsSandbox())
            {
                File.WriteAllText(sandbox.PathName, "{}"); Assert.IsFalse(sandbox.Store.Load().ChatUseTranslationSettings);
                var settings = Settings(); settings.ChatUseTranslationSettings = enabled;
                sandbox.Store.Save(settings);
                var restored = sandbox.Store.Load(); Assert.AreEqual(enabled, restored.ChatUseTranslationSettings);
                Assert.AreEqual(settings.TranslationStyle, restored.TranslationStyle); Assert.AreEqual(settings.CustomInstructions, restored.CustomInstructions);
                Assert.AreEqual(settings.Glossary, restored.Glossary);
            }
        }

        private static UserControl Pane(string host, SettingsStore store, ChatHistoryStore history)
        {
            var pane = (UserControl)Activator.CreateInstance(DocumentChatHostTests.Host(host).GetType("OfficeTranslate.WordAddIn.DocumentChatPane", true)!)!;
            pane.GetType().GetProperty("ChatSettingsStore", Hidden)!.SetValue(pane, store);
            pane.GetType().GetProperty("HistoryStore", Hidden)!.SetValue(pane, history);
            Call(pane, "Initialize", (Func<ChatContext>)(() => new ChatContext("document", true)), (Func<TranslationSettings>)store.Load);
            pane.Dock = DockStyle.Fill; return pane;
        }
        private static void Pump(Func<bool> done)
        {
            var deadline = DateTime.UtcNow.AddSeconds(5);
            while (!done() && DateTime.UtcNow < deadline) { Application.DoEvents(); Thread.Sleep(5); }
            Assert.IsTrue(done(), "Chat request did not finish.");
        }

        [DataTestMethod]
        [DataRow("Word", "zh-CN")] [DataRow("Word", "en")]
        [DataRow("Excel", "zh-CN")] [DataRow("Excel", "en")]
        [DataRow("PowerPoint", "zh-CN")] [DataRow("PowerPoint", "en")]
        public void Pane_CheckBoxPersists_Restores_Localizes_AndDrivesActualRequests(string host, string language)
        {
            DocumentChatHostTests.Sta(() =>
            {
                using (var sandbox = new SettingsSandbox())
                using (var history = new HistorySandbox())
                {
                    var settings = Settings(); settings.UiLanguage = language; sandbox.Store.Save(settings);
                    using (var pane = Pane(host, sandbox.Store, history.Store))
                    using (var window = new Form { ClientSize = new Size(440, 760) })
                    {
                        window.Controls.Add(pane); window.Show(); Application.DoEvents();
                        var check = (CheckBox)Field(pane, "_useTranslationSettings");
                        Assert.IsFalse(check.Checked); Assert.AreEqual(UiText.Get(language, "ChatUseTranslationSettings"), check.Text);
                        var commands = (TableLayoutPanel)Field(pane, "_commands");
                        Assert.AreSame(commands, check.Parent); Assert.AreEqual(1, commands.GetRow(check));
                        Assert.IsTrue(check.Width <= commands.ClientSize.Width);
                        check.Checked = true; Assert.IsTrue(sandbox.Store.Load().ChatUseTranslationSettings);
                        var systems = new List<string>();
                        ChatTransport transport = (m, s, n, o, t) => { systems.Add(m[0].Content); return Task.FromResult(new ChatReply("answer", false)); };
                        pane.GetType().GetProperty("Transport", Hidden)!.SetValue(pane, transport);
                        Call(pane, "Start", ChatMode.Summarize); Pump(() => pane.GetType().GetField("_pending", Hidden)!.GetValue(pane) == null);
                        StringAssert.Contains(systems.Last(), settings.Glossary);
                        var latest = sandbox.Store.Load(); latest.UiLanguage = language == "en" ? "zh-CN" : "en";
                        latest.Glossary = "updated glossary"; sandbox.Store.Save(latest); Call(pane, "RefreshLanguage");
                        Assert.IsTrue(check.Checked); Assert.AreEqual(UiText.Get(latest.UiLanguage, "ChatUseTranslationSettings"), check.Text);
                        Call(pane, "Start", ChatMode.Summarize); Pump(() => pane.GetType().GetField("_pending", Hidden)!.GetValue(pane) == null);
                        StringAssert.Contains(systems.Last(), "updated glossary"); Assert.IsFalse(systems.Last().Contains(settings.Glossary));
                        window.Close();
                    }
                    using (var pane = Pane(host, sandbox.Store, history.Store))
                    {
                        var check = (CheckBox)Field(pane, "_useTranslationSettings"); Assert.IsTrue(check.Checked);
                        check.Checked = false; Assert.IsFalse(sandbox.Store.Load().ChatUseTranslationSettings);
                        Assert.AreEqual("updated glossary", sandbox.Store.Load().Glossary);
                        // Saving must fail safely without leaking paths or leaving an unsaved choice in the UI.
                        var currentLanguage = sandbox.Store.Load().UiLanguage;
                        File.Delete(sandbox.PathName); Directory.CreateDirectory(sandbox.PathName);
                        check.Checked = true; Assert.IsFalse(check.Checked);
                        Assert.AreEqual(UiText.Get(currentLanguage, "ChatSettingsError"), ((Label)Field(pane, "_status")).Text);
                    }
                }
            });
        }

        [DataTestMethod]
        [DataRow("Word")] [DataRow("Excel")] [DataRow("PowerPoint")]
        public void TranslationAndChatSettingsForms_PreserveThePersistedChoice(string host)
        {
            DocumentChatHostTests.Sta(() =>
            {
                using (var sandbox = new SettingsSandbox())
                {
                    var settings = Settings(); settings.ChatUseTranslationSettings = true; sandbox.Store.Save(settings);
                    var assembly = DocumentChatHostTests.Host(host);
                    using (var form = (Form)Activator.CreateInstance(assembly.GetType("OfficeTranslate.WordAddIn.SettingsForm", true)!, new object[] { sandbox.Store })!)
                    {
                        var rebuilt = (TranslationSettings)Call(form, "BuildSettings", true)!;
                        Assert.IsTrue(rebuilt.ChatUseTranslationSettings); sandbox.Store.Save(rebuilt);
                    }
                    using (var form = (Form)Activator.CreateInstance(assembly.GetType("OfficeTranslate.WordAddIn.ChatSettingsForm", true)!, new object[] { sandbox.Store })!)
                    {
                        form.Show(); Application.DoEvents(); ((Button)form.AcceptButton!).PerformClick();
                        Assert.IsTrue(sandbox.Store.Load().ChatUseTranslationSettings);
                        Assert.AreEqual(settings.Glossary, sandbox.Store.Load().Glossary);
                    }
                }
            });
        }
    }
}
