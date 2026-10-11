using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeTranslate.Core;

namespace OfficeTranslate.Tests
{
    [TestClass]
    public class ImageOverlayWritebackTests
    {
        [TestMethod]
        public void CreationComFailure_ReturnsRegionFallbackWithoutDeletingOthers()
        {
            var deletes = 0;
            Assert.IsFalse(ImageOverlayWriteback.TryCreate<object>(() => throw new COMException("create failed"),
                s => Assert.Fail("no shape"), s => deletes++, out var shape, out var reason));
            Assert.IsNull(shape); Assert.AreEqual(0, deletes); StringAssert.Contains(reason, "create failed");
        }

        [TestMethod]
        public void InitializationComFailure_DeletesOnlyFailedShapeAndContinuesSiblings()
        {
            var first = new object(); var failed = new object(); var last = new object();
            var kept = new List<object>(); var deleted = new List<object>();
            foreach (var candidate in new[] { first, failed, last })
            {
                if (ImageOverlayWriteback.TryCreate(() => candidate,
                    s => { if (s == failed) throw new COMException("style failed"); },
                    s => deleted.Add(s), out var shape, out var reason)) kept.Add(shape!);
                else { Assert.IsNull(shape); StringAssert.Contains(reason, "style failed"); }
            }
            CollectionAssert.AreEqual(new[] { first, last }, kept);
            CollectionAssert.AreEqual(new[] { failed }, deleted);
        }

        [TestMethod]
        public void InitializationCancelled_CleansItsShapeButRethrowsOriginal()
        {
            var shape = new object(); var deleted = false;
            var error = new OperationCanceledException();
            Assert.AreSame(error, Assert.ThrowsException<OperationCanceledException>(() =>
                ImageOverlayWriteback.TryCreate(() => shape, s => throw error,
                    s => { Assert.AreSame(shape, s); deleted = true; }, out _, out _)));
            Assert.IsTrue(deleted);
        }

        [TestMethod]
        public void NonComFailure_IsNotSilentlyConvertedToFallback()
        {
            var error = new InvalidOperationException("bug");
            Assert.AreSame(error, Assert.ThrowsException<InvalidOperationException>(() =>
                ImageOverlayWriteback.TryCreate<object>(() => throw error, s => { }, s => { }, out _, out _)));
        }

        [TestMethod]
        public void CleanupFailure_IsExplicitInFallbackReason()
        {
            Assert.IsFalse(ImageOverlayWriteback.TryCreate(() => new object(), s => throw new COMException("style"),
                s => throw new COMException("delete"), out var shape, out var reason));
            Assert.IsNull(shape); StringAssert.Contains(reason, "未能删除失败的框");
        }
    }
}
