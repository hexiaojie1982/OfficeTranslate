using System;
using System.Threading;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeTranslate.Core;

namespace OfficeTranslate.Tests
{
    [TestClass]
    public sealed class ClipboardImageCaptureTests
    {
        [TestMethod]
        public void CapturePng_NullCopyAction_ThrowsArgumentNullException()
        {
            // Must fail before touching the clipboard.
            Assert.ThrowsException<ArgumentNullException>(() =>
                ClipboardImageCapture.CapturePng(null!, CancellationToken.None));
        }
    }
}
