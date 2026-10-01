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

    // S2: the no-image failure is its own category. It must never be
    // classified as transient-busy (it has no HResult and no underlying
    // exception to report), and it must stay an InvalidOperationException
    // so existing catch blocks keep catching it.
    [TestClass]
    public class ClipboardNoImageFailureTests
    {
        [TestMethod]
        public void NoImage_IsNotTransientBusy()
        {
            var ex = new ClipboardImageCapture.NoImageCaptureException("no image");
            Assert.IsFalse(ClipboardImageCapture.IsTransientClipboardFailure(ex));
        }

        [TestMethod]
        public void NoImage_WrappedAsInner_IsNotTransientBusy()
        {
            var inner = new ClipboardImageCapture.NoImageCaptureException("no image");
            var outer = new InvalidOperationException("phase wrapper", inner);
            Assert.IsFalse(ClipboardImageCapture.IsTransientClipboardFailure(outer));
        }

        [TestMethod]
        public void NoImage_IsInvalidOperationException()
        {
            var ex = new ClipboardImageCapture.NoImageCaptureException("no image");
            Assert.IsInstanceOfType(ex, typeof(InvalidOperationException));
        }
    }
}
