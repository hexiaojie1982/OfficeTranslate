using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeTranslate.Core;
using System;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OfficeTranslate.Tests
{
    // M2/N1: the dispatcher pins COM/clipboard/writeback work to the Office
    // UI thread captured at the ribbon entry point. The ambient
    // SynchronizationContext is never trusted (N1: the real Word ribbon
    // entry carries the base System.Threading.SynchronizationContext, whose
    // Send/Post run inline on the calling thread -- adopting it made
    // cross-thread Invoke/InvokeAsync silently run on MTA pool threads).
    //
    // These tests run a real pumping STA thread (Application.Run) and assert
    // from an MTA thread that Invoke/InvokeAsync callbacks actually execute
    // on the captured STA thread with IsUiThread == true, under three
    // ambient states (null, base, real WinForms context). Same-thread
    // fast-path tests are kept and run ON the STA thread via the dispatcher
    // itself. A cross-thread test without a message pump would deadlock and
    // is not attempted; every wait carries an explicit timeout so a broken
    // dispatcher fails with a message instead of hanging the suite.
    [TestClass]
    public class OfficeUiDispatcherTests
    {
        [TestMethod]
        public void Capture_OnMtaThread_ThrowsInvalidOperation()
        {
            Exception? captured = null;
            var mta = new Thread(() =>
            {
                try { OfficeUiDispatcher.Capture("Test"); }
                catch (Exception ex) { captured = ex; }
            });
            mta.SetApartmentState(ApartmentState.MTA);
            mta.IsBackground = true;
            mta.Start();
            Assert.IsTrue(mta.Join(TimeSpan.FromSeconds(10)), "MTA probe thread did not finish");
            Assert.IsNotNull(captured, "Capture on an MTA thread must fail fast");
            Assert.IsInstanceOfType(captured, typeof(InvalidOperationException));
        }

        [TestMethod]
        public void CrossThread_NullAmbient_MarshalsToStaThread()
        {
            CrossThreadCore(null);
        }

        [TestMethod]
        public void CrossThread_BaseAmbient_MarshalsToStaThread()
        {
            // N1 regression: the base SynchronizationContext has no
            // marshaling target; the dispatcher must not adopt it.
            CrossThreadCore(() => SynchronizationContext.SetSynchronizationContext(new SynchronizationContext()));
        }

        [TestMethod]
        public void CrossThread_WinFormsAmbient_MarshalsToStaThread()
        {
            CrossThreadCore(() => SynchronizationContext.SetSynchronizationContext(new WindowsFormsSynchronizationContext()));
        }

        [TestMethod]
        public void Invoke_OnCapturedThread_RunsInline()
        {
            using (var pump = StartPumpingSta())
            {
                var dispatcher = pump.Dispatcher;
                // Runs ON the STA thread through the dispatcher, so the
                // inner Invoke exercises the inline fast path.
                dispatcher.Invoke(() =>
                {
                    Assert.IsTrue(dispatcher.IsUiThread);
                    var workThread = -1;
                    dispatcher.Invoke(() => { workThread = Thread.CurrentThread.ManagedThreadId; });
                    Assert.AreEqual(pump.StaThreadId, workThread);
                });
            }
        }

        [TestMethod]
        public void InvokeAsync_OnCapturedThread_ReturnsValue()
        {
            using (var pump = StartPumpingSta())
            {
                var dispatcher = pump.Dispatcher;
                var value = dispatcher.Invoke(() => dispatcher.InvokeAsync(() => 41 + 1).GetAwaiter().GetResult());
                Assert.AreEqual(42, value);
            }
        }

        [TestMethod]
        public void Invoke_PropagatesException()
        {
            using (var pump = StartPumpingSta())
            {
                Assert.ThrowsException<InvalidOperationException>(
                    () => pump.Dispatcher.Invoke(() => throw new InvalidOperationException("probe")));
            }
        }

        [TestMethod]
        public void LogProbe_DoesNotThrow()
        {
            using (var pump = StartPumpingSta())
            {
                pump.Dispatcher.Invoke(() => pump.Dispatcher.LogProbe("unit_test"));
            }
        }

        // Runs on the MTA test thread while the STA pump is up. Asserts that
        // both Invoke (Send) and InvokeAsync (Post) callbacks execute on the
        // captured STA thread with STA apartment and IsUiThread == true.
        private static void CrossThreadCore(Action? ambientSetup)
        {
            using (var pump = StartPumpingSta(ambientSetup))
            {
                var dispatcher = pump.Dispatcher;
                var syncFact = dispatcher.Invoke(() => new ThreadFact(dispatcher));
                Assert.AreEqual(pump.StaThreadId, syncFact.ThreadId, "Invoke did not run on the captured STA thread");
                Assert.AreEqual(ApartmentState.STA, syncFact.Apartment);
                Assert.IsTrue(syncFact.IsUi);

                var task = dispatcher.InvokeAsync(() => new ThreadFact(dispatcher));
                Assert.IsTrue(task.Wait(TimeSpan.FromSeconds(10)), "InvokeAsync callback never ran on the STA thread");
                var asyncFact = task.Result;
                Assert.AreEqual(pump.StaThreadId, asyncFact.ThreadId, "InvokeAsync did not run on the captured STA thread");
                Assert.AreEqual(ApartmentState.STA, asyncFact.Apartment);
                Assert.IsTrue(asyncFact.IsUi);
            }
        }

        private sealed class ThreadFact
        {
            public readonly int ThreadId;
            public readonly ApartmentState Apartment;
            public readonly bool IsUi;
            public ThreadFact(OfficeUiDispatcher dispatcher)
            {
                ThreadId = Thread.CurrentThread.ManagedThreadId;
                Apartment = Thread.CurrentThread.GetApartmentState();
                IsUi = dispatcher.IsUiThread;
            }
        }

        [TestMethod]
        public void RestoreGate_RealStaQueueHeldBeyondTwoSeconds_LateCallbackIsNoOp()
        {
            using (var pump = StartPumpingSta())
            using (var entered = new ManualResetEventSlim())
            using (var release = new ManualResetEventSlim())
            {
                var blocker = pump.Dispatcher.InvokeAsync(() => { entered.Set(); release.Wait(10000); });
                Assert.IsTrue(entered.Wait(5000));
                var calls = 0;
                var reason = "";
                try
                {
                    var restore = UiRestoreGate.RunAsync(pump.Dispatcher, () => true,
                        () => Interlocked.Increment(ref calls), r => reason = r);
                    Assert.IsTrue(restore.Wait(3000), "expiry must not need the blocked UI context");
                    Assert.AreEqual("expired", reason);
                }
                finally { release.Set(); }
                Assert.IsTrue(blocker.Wait(5000));
                pump.Dispatcher.Invoke(() => { }); // drains queued restore
                Assert.AreEqual(0, calls);
            }
        }

        [TestMethod]
        public void RestoreGate_SupersededGeneration_NoOpOnRealStaQueue()
        {
            using (var pump = StartPumpingSta())
            using (var entered = new ManualResetEventSlim())
            using (var release = new ManualResetEventSlim())
            {
                var blocker = pump.Dispatcher.InvokeAsync(() => { entered.Set(); release.Wait(10000); });
                Assert.IsTrue(entered.Wait(5000));
                var generation = 1;
                var calls = 0;
                var reason = "";
                var restore = UiRestoreGate.RunAsync(pump.Dispatcher, () => Volatile.Read(ref generation) == 1,
                    () => calls++, r => reason = r);
                Interlocked.Increment(ref generation);
                release.Set();
                Assert.IsTrue(restore.Wait(5000));
                Assert.AreEqual("superseded", reason);
                Assert.AreEqual(0, calls);
                Assert.IsTrue(blocker.Wait(5000));
            }
        }

        [TestMethod]
        public void RestoreGate_StartedCallbackCanFinish_AndNormalRestoreRunsOnce()
        {
            using (var pump = StartPumpingSta())
            using (var entered = new ManualResetEventSlim())
            using (var release = new ManualResetEventSlim())
            {
                var calls = 0;
                var reason = "";
                var restore = UiRestoreGate.RunAsync(pump.Dispatcher, () => true,
                    () => { entered.Set(); release.Wait(10000); calls++; }, r => reason = r, waitMs: 100);
                try
                {
                    Assert.IsTrue(entered.Wait(5000));
                    Assert.IsTrue(restore.Wait(2000));
                    Assert.AreEqual("", reason); // running callback is not expired
                }
                finally { release.Set(); }
                pump.Dispatcher.Invoke(() => { });
                Assert.AreEqual(1, calls);
                UiRestoreGate.RunAsync(pump.Dispatcher, () => true, () => calls++).Wait(5000);
                Assert.AreEqual(2, calls);
            }
        }

        [TestMethod]
        public void RestoreGate_CallbackFailureDoesNotMaskTaskOutcome()
        {
            using (var pump = StartPumpingSta())
                Assert.IsTrue(UiRestoreGate.RunAsync(pump.Dispatcher, () => true,
                    () => throw new InvalidOperationException("test")).Wait(5000));
        }

        private sealed class StaPump : IDisposable
        {
            private readonly Thread _sta;
            public StaPump(OfficeUiDispatcher dispatcher, int staThreadId, Thread sta)
            {
                Dispatcher = dispatcher;
                StaThreadId = staThreadId;
                _sta = sta;
            }
            public OfficeUiDispatcher Dispatcher { get; }
            public int StaThreadId { get; }
            public void Dispose()
            {
                // Stop the pump through the dispatcher itself; a broken
                // dispatcher fails here with a message instead of hanging.
                var stop = Dispatcher.InvokeAsync(() => { Application.ExitThread(); return 0; });
                if (!stop.Wait(TimeSpan.FromSeconds(10)))
                    throw new AssertFailedException("dispatcher could not stop the STA message pump");
                if (!_sta.Join(TimeSpan.FromSeconds(10)))
                    throw new AssertFailedException("STA pump thread did not exit");
            }
        }

        private static StaPump StartPumpingSta(Action? ambientSetup = null)
        {
            Exception? staError = null;
            var ready = new ManualResetEventSlim(false);
            OfficeUiDispatcher? dispatcher = null;
            int staThreadId = -1;
            var sta = new Thread(() =>
            {
                try
                {
                    staThreadId = Thread.CurrentThread.ManagedThreadId;
                    ambientSetup?.Invoke();
                    dispatcher = OfficeUiDispatcher.Capture("Test");
                    ready.Set();
                    Application.Run(); // pumps until ExitThread
                }
                catch (Exception ex) { staError = ex; ready.Set(); }
            });
            sta.SetApartmentState(ApartmentState.STA);
            sta.IsBackground = true;
            sta.Start();
            if (!ready.Wait(TimeSpan.FromSeconds(15)))
                throw new AssertFailedException("STA pump thread did not start");
            if (staError != null)
                throw new AssertFailedException("Capture on STA thread failed: " + staError);
            if (dispatcher == null)
                throw new AssertFailedException("dispatcher was not captured");
            return new StaPump(dispatcher, staThreadId, sta);
        }
    }
}
