using System;
using System.Net;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeTranslate.Core;

namespace OfficeTranslate.Tests
{
    [TestClass]
    public sealed class ReviewFixesTests
    {
        // HttpStatusCode has no TooManyRequests member on .NET Framework 4.8;
        // the client must compare the numeric code instead.

        [TestMethod]
        public void IsTransientStatusCode_RetryableCodes_ReturnsTrue()
        {
            Assert.IsTrue(TranslationClient.IsTransientStatusCode((HttpStatusCode)429));
            Assert.IsTrue(TranslationClient.IsTransientStatusCode(HttpStatusCode.InternalServerError));
            Assert.IsTrue(TranslationClient.IsTransientStatusCode(HttpStatusCode.BadGateway));
            Assert.IsTrue(TranslationClient.IsTransientStatusCode(HttpStatusCode.ServiceUnavailable));
            Assert.IsTrue(TranslationClient.IsTransientStatusCode(HttpStatusCode.GatewayTimeout));
        }

        [TestMethod]
        public void IsTransientStatusCode_NonRetryableCodes_ReturnsFalse()
        {
            Assert.IsFalse(TranslationClient.IsTransientStatusCode(HttpStatusCode.OK));
            Assert.IsFalse(TranslationClient.IsTransientStatusCode(HttpStatusCode.BadRequest));
            Assert.IsFalse(TranslationClient.IsTransientStatusCode((HttpStatusCode)422));
            Assert.IsFalse(TranslationClient.IsTransientStatusCode(HttpStatusCode.Unauthorized));
        }

        [TestMethod]
        public void ShouldFallbackFromStructuredOutput_Handles400And422()
        {
            Assert.IsTrue(TranslationClient.ShouldFallbackFromStructuredOutput(
                400, "Unsupported parameter: 'response_format'"));
            Assert.IsTrue(TranslationClient.ShouldFallbackFromStructuredOutput(
                422, "did not understand the \"format\" field"));
        }

        [TestMethod]
        public void ShouldFallbackFromStructuredOutput_RejectsUnrelatedErrors()
        {
            Assert.IsFalse(TranslationClient.ShouldFallbackFromStructuredOutput(
                400, "Invalid API key"));
            Assert.IsFalse(TranslationClient.ShouldFallbackFromStructuredOutput(
                500, "unsupported response_format"));
            Assert.IsFalse(TranslationClient.ShouldFallbackFromStructuredOutput(422, null));
        }

        [TestMethod]
        public void RangeGrid_Normalize_ConvertsOneBasedComArray()
        {
            // Builds a real 1-based 2x2 array, exactly what Excel COM returns.
            var com = (object[,])Array.CreateInstance(typeof(object),
                new[] { 2, 2 }, new[] { 1, 1 });
            com[1, 1] = "a"; com[1, 2] = "b";
            com[2, 1] = "c"; com[2, 2] = "d";

            var grid = RangeGridHelper.Normalize(com);

            Assert.AreEqual(2, grid.GetLength(0));
            Assert.AreEqual(2, grid.GetLength(1));
            Assert.AreEqual(0, grid.GetLowerBound(0));
            Assert.AreEqual("a", grid[0, 0]);
            Assert.AreEqual("b", grid[0, 1]);
            Assert.AreEqual("c", grid[1, 0]);
            Assert.AreEqual("d", grid[1, 1]);
        }

        [TestMethod]
        public void RangeGrid_Normalize_WrapsScalar()
        {
            var grid = RangeGridHelper.Normalize("single");

            Assert.AreEqual(1, grid.GetLength(0));
            Assert.AreEqual(1, grid.GetLength(1));
            Assert.AreEqual("single", grid[0, 0]);
        }

        [TestMethod]
        public void RangeGrid_Normalize_WrapsDbNullScalar()
        {
            // Real Excel COM has returned scalar DBNull for HasFormula on a
            // multi-cell range; it must not break normalization.
            var grid = RangeGridHelper.Normalize(DBNull.Value);

            Assert.AreEqual(1, grid.GetLength(0));
            Assert.AreEqual(DBNull.Value, grid[0, 0]);
        }

        [TestMethod]
        public void IsFormulaCell_PrefersWellFormedFlagGrid()
        {
            var flags = new object[,] { { true, false }, { false, true } };
            var formulas = new object[,] { { "=A1", "x" }, { "y", "=B2" } };

            Assert.IsTrue(RangeGridHelper.IsFormulaCell(flags, formulas, 0, 0));
            Assert.IsFalse(RangeGridHelper.IsFormulaCell(flags, formulas, 0, 1));
            Assert.IsFalse(RangeGridHelper.IsFormulaCell(flags, formulas, 1, 0));
            Assert.IsTrue(RangeGridHelper.IsFormulaCell(flags, formulas, 1, 1));
        }

        [TestMethod]
        public void IsFormulaCell_ScalarDbNullFlags_FallsBackToFormulaText()
        {
            var flags = new object[,] { { DBNull.Value } };
            var formulas = new object[,] { { "=SUM(A1:A2)", "plain text" } };

            Assert.IsTrue(RangeGridHelper.IsFormulaCell(flags, formulas, 0, 0));
            Assert.IsFalse(RangeGridHelper.IsFormulaCell(flags, formulas, 0, 1));
        }

        [TestMethod]
        public void IsFormulaCell_OutOfRange_ReturnsFalse()
        {
            var flags = new object[,] { { false } };
            var formulas = new object[,] { { "=A1" } };

            Assert.IsFalse(RangeGridHelper.IsFormulaCell(flags, formulas, 5, 5));
            Assert.IsFalse(RangeGridHelper.IsFormulaCell(flags, formulas, -1, 0));
        }
    }
}
