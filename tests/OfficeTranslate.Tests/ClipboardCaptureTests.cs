using System;
using System.Runtime.InteropServices;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeTranslate.Core;

namespace OfficeTranslate.Tests
{
    // S1: only CLIPBRD_E_CANT_OPEN (0x800401D0, "Requested Clipboard
    // operation did not succeed") is a transient clipboard failure worth
    // retrying. The check must see through the phase wrappers that nest
    // the original error as an inner exception.
    [TestClass]
    public class ClipboardTransientFailureTests
    {
        private const int ClipbrdECantOpen = unchecked((int)0x800401D0);

        [TestMethod]
        public void Transient_ExternalExceptionWithCantOpen_ReturnsTrue()
        {
            var ex = new ExternalException(
                "Requested Clipboard operation did not succeed.", ClipbrdECantOpen);
            Assert.IsTrue(ClipboardImageCapture.IsTransientClipboardFailure(ex));
        }

        [TestMethod]
        public void Transient_ComExceptionWithCantOpen_ReturnsTrue()
        {
            var ex = new COMException("clipboard busy", ClipbrdECantOpen);
            Assert.IsTrue(ClipboardImageCapture.IsTransientClipboardFailure(ex));
        }

        [TestMethod]
        public void Transient_WrappedInPhaseException_ReturnsTrue()
        {
            var inner = new ExternalException(
                "Requested Clipboard operation did not succeed.", ClipbrdECantOpen);
            var wrapped = new InvalidOperationException("phase wrapper", inner);
            Assert.IsTrue(ClipboardImageCapture.IsTransientClipboardFailure(wrapped));
        }

        [TestMethod]
        public void NotTransient_OtherHResult_ReturnsFalse()
        {
            var ex = new ExternalException("other", unchecked((int)0x800401D1));
            Assert.IsFalse(ClipboardImageCapture.IsTransientClipboardFailure(ex));
        }

        [TestMethod]
        public void NotTransient_PlainException_ReturnsFalse()
        {
            Assert.IsFalse(ClipboardImageCapture.IsTransientClipboardFailure(
                new InvalidOperationException("nope")));
        }

        [TestMethod]
        public void NotTransient_Cancelled_ReturnsFalse()
        {
            Assert.IsFalse(ClipboardImageCapture.IsTransientClipboardFailure(
                new OperationCanceledException()));
        }

        [TestMethod]
        public void NotTransient_Null_ReturnsFalse()
        {
            Assert.IsFalse(ClipboardImageCapture.IsTransientClipboardFailure(null));
        }
    }
}
