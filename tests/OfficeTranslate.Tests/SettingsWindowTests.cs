using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeTranslate.Core;

namespace OfficeTranslate.Tests
{
    [TestClass]
    public sealed class SettingsWindowTests
    {
        [DataTestMethod]
        [DataRow("Word", true, false)]
        [DataRow("Word", false, false)]
        [DataRow("Excel", true, false)]
        [DataRow("PowerPoint", true, false)]
        [DataRow("Word", true, true)]
        [DataRow("Excel", true, true)]
        [DataRow("PowerPoint", true, true)]
        public void Settings_WithLiveSummary_CanOpenRepeatedly(string host, bool activateSummary, bool closeSummary)
        {
            var configuration = new DirectoryInfo(AppDomain.CurrentDomain.BaseDirectory).Parent!.Name;
            var assembly = Assembly.LoadFrom(Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory,
                "../../../../../src/OfficeTranslate." + host + "AddIn/bin/" + configuration + "/net48/OfficeTranslate." + host + "AddIn.dll")));
            Exception? error = null;
            var thread = new Thread(() => {
                try
                {
                    using (var owner = new Form())
                    {
                        owner.Show();
                        var summary = new TranslationTaskSummary(1);
                        summary.RecordImageNoteFailure("fixture", "RECOVERABLE FIXTURE", "fixture");
                        assembly.GetType("OfficeTranslate.WordAddIn.TranslationResultForm", true)!
                            .GetMethod("ShowResult")!.Invoke(null, new object?[] { "en", summary, owner.Handle, null });
                        var result = Application.OpenForms.Cast<Form>().Single(f => f.GetType().Assembly == assembly);
                        try
                        {
                            for (var i = 0; i < (closeSummary ? 1 : 3); i++)
                            {
                                if (activateSummary) result.Activate(); else owner.Activate();
                                var store = new SettingsStore();
                                // Never load or save the user's actual settings/credentials.
                                typeof(SettingsStore).GetField("_path", BindingFlags.NonPublic | BindingFlags.Instance)!
                                    .SetValue(store, Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".json"));
                                using (var settings = (Form)Activator.CreateInstance(
                                    assembly.GetType("OfficeTranslate.WordAddIn.SettingsForm", true)!, new object[] { store })!)
                                {
                                    if (closeSummary) settings.Load += (s, e) => result.Close();
                                    IntPtr observedOwner = IntPtr.Zero;
                                    settings.Shown += (s, e) => {
                                        observedOwner = GetWindow(settings.Handle, 4);
                                        settings.BeginInvoke(new Action(() => settings.Close()));
                                    };
                                    settings.GetType().GetMethod("ShowFor")!.Invoke(settings, new object[] { owner.Handle });
                                    Assert.AreEqual(owner.Handle, observedOwner, "Settings must belong to the host, not the transient summary.");
                                }
                                if (!closeSummary) Assert.IsTrue(result.Visible && !result.IsDisposed);
                            }
                        }
                        finally { result.Close(); }
                    }
                }
                catch (Exception ex) { error = ex; }
            });
            thread.IsBackground = true;
            thread.SetApartmentState(ApartmentState.STA); thread.Start();
            Assert.IsTrue(thread.Join(15000), "Settings probe must finish without modal blocking.");
            if (error != null) throw error;
        }

        [DataTestMethod]
        [DataRow(0)]
        [DataRow(1)]
        [DataRow(2)]
        public void Settings_InvalidOwner_DoesNotAdoptActiveSummary(int mode)
        {
            Exception? error = null;
            var thread = new Thread(() => {
                try
                {
                    var configuration = new DirectoryInfo(AppDomain.CurrentDomain.BaseDirectory).Parent!.Name;
                    var assembly = Assembly.LoadFrom(Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory,
                        "../../../../../src/OfficeTranslate.WordAddIn/bin/" + configuration + "/net48/OfficeTranslate.WordAddIn.dll")));
                    IntPtr owner = mode == 1 ? new IntPtr(-1) : IntPtr.Zero;
                    if (mode == 2)
                    {
                        using (var closed = new Form()) { owner = closed.Handle; }
                    }
                    using (var summary = new Form())
                    {
                        summary.Show(); summary.Activate();
                        var store = new SettingsStore();
                        typeof(SettingsStore).GetField("_path", BindingFlags.NonPublic | BindingFlags.Instance)!
                            .SetValue(store, Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".json"));
                        using (var settings = (Form)Activator.CreateInstance(
                            assembly.GetType("OfficeTranslate.WordAddIn.SettingsForm", true)!, new object[] { store })!)
                        {
                            IntPtr observed = new IntPtr(-1);
                            settings.Shown += (s, e) => {
                                observed = GetWindow(settings.Handle, 4);
                                settings.BeginInvoke(new Action(() => settings.Close()));
                            };
                            settings.GetType().GetMethod("ShowFor")!.Invoke(settings, new object[] { owner });
                            // ShowInTaskbar=false uses a WinForms hidden owner
                            // when no host survives. It must never adopt the
                            // active summary (or any visible transient form).
                            Assert.AreNotEqual(summary.Handle, observed);
                            Assert.IsFalse(IsWindowVisible(observed));
                            Assert.IsTrue(summary.Visible);
                        }
                    }
                }
                catch (Exception ex) { error = ex; }
            });
            thread.IsBackground = true; thread.SetApartmentState(ApartmentState.STA); thread.Start();
            Assert.IsTrue(thread.Join(15000));
            if (error != null) throw error;
        }

        [DllImport("user32.dll")] private static extern IntPtr GetWindow(IntPtr hwnd, uint command);
        [DllImport("user32.dll")] private static extern bool IsWindowVisible(IntPtr hwnd);
    }
}
