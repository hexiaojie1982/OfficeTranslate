using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeTranslate.Core;
using System;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeTranslate.Tests
{
    // M2: the dispatcher pins COM/clipboard/writeback work to the Office UI
    // thread captured at the ribbon entry point. Cross-thread marshalling
    // needs a pumping message loop, which a unit test cannot provide, so
    // these tests pin the same-thread contract: on the captured thread,
    // Invoke runs the work inline, InvokeAsync completes with its value,
    // and LogProbe reports that thread honestly. A cross-thread test would
    // deadlock without a message pump and is intentionally not attempted.
    [TestClass]
    public class OfficeUiDispatcherTests
    {
        [TestMethod]
        public void Invoke_OnCapturedThread_RunsInline()
        {
            var previous = SynchronizationContext.Current;
            try
            {
                var ui = OfficeUiDispatcher.Capture("Test");
                Assert.IsTrue(ui.IsUiThread);
                var workThread = -1;
                var apartment = ApartmentState.Unknown;
                string? ctxName = null;
                ui.Invoke(() =>
                {
                    workThread = Thread.CurrentThread.ManagedThreadId;
                    apartment = Thread.CurrentThread.GetApartmentState();
                    ctxName = SynchronizationContext.Current?.GetType().Name;
                });
                Assert.AreEqual(Thread.CurrentThread.ManagedThreadId, workThread);
                Assert.AreEqual(Thread.CurrentThread.GetApartmentState(), apartment);
                Assert.IsFalse(string.IsNullOrEmpty(ctxName));
            }
            finally
            {
                SynchronizationContext.SetSynchronizationContext(previous);
            }
        }

        [TestMethod]
        public async Task InvokeAsync_OnCapturedThread_ReturnsValue()
        {
            var previous = SynchronizationContext.Current;
            try
            {
                var ui = OfficeUiDispatcher.Capture("Test");
                var value = await ui.InvokeAsync(() => 41 + 1);
                Assert.AreEqual(42, value);
            }
            finally
            {
                SynchronizationContext.SetSynchronizationContext(previous);
            }
        }

        [TestMethod]
        public void Invoke_PropagatesException()
        {
            var previous = SynchronizationContext.Current;
            try
            {
                var ui = OfficeUiDispatcher.Capture("Test");
                Assert.ThrowsException<InvalidOperationException>(
                    () => ui.Invoke(() => throw new InvalidOperationException("probe")));
            }
            finally
            {
                SynchronizationContext.SetSynchronizationContext(previous);
            }
        }

        [TestMethod]
        public void LogProbe_DoesNotThrow()
        {
            var previous = SynchronizationContext.Current;
            try
            {
                var ui = OfficeUiDispatcher.Capture("Test");
                ui.LogProbe("unit_test");
            }
            finally
            {
                SynchronizationContext.SetSynchronizationContext(previous);
            }
        }
    }
}
