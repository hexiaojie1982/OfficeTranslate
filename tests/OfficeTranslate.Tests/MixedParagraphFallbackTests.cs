using System;
using System.Collections.Generic;
using System.Net;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeTranslate.Core;

namespace OfficeTranslate.Tests
{
    [TestClass]
    public class MixedParagraphFallbackTests
    {
        private static readonly string[] Sources = {
            "SESS 기능은 정상적으로 동작하며 모든 검사가 완료되었습니다.",
            "PKG 장비는 정상적으로 동작하며 모든 검사가 완료되었습니다.",
            "IT 부서는 정상적으로 동작하며 모든 검사가 완료되었습니다." };
        private static readonly string[] Outputs = { "SESS 功能正常运行且所有检查已完成。", "PKG 设备正常运行且所有检查已完成。", "IT 部门正常运行且所有检查已完成。" };
        private const string BrokenTokens = "⟦OT_KEEP_9999⟧ 未通过保护验证。";
        private sealed class Handler : HttpMessageHandler
        {
            private readonly Func<string, int, string> _reply;
            private readonly ProviderKind _provider;
            public readonly List<string> Inputs = new List<string>();
            public Handler(ProviderKind provider, Func<string, int, string> reply) { _provider = provider; _reply = reply; }
            protected override async Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken token)
            {
                token.ThrowIfCancellationRequested();
                var json = new JavaScriptSerializer();
                var body = (Dictionary<string, object>)json.DeserializeObject(await request.Content.ReadAsStringAsync());
                var messages = (object[])body["messages"];
                var input = (string)((Dictionary<string, object>)messages[1])["content"];
                Inputs.Add(input);
                var response = _reply(input, Inputs.Count);
                return new HttpResponseMessage(HttpStatusCode.OK) { Content = new StringContent(json.Serialize(
                    _provider == ProviderKind.Ollama ? (object)new { message = new { content = response }, done_reason = "stop" }
                    : new { choices = new[] { new { message = new { content = response }, finish_reason = "stop" } } })) };
            }
        }
        private static TranslationSettings Settings(ProviderKind provider = ProviderKind.OpenAiCompatible, string source = "韩语") => new TranslationSettings {
            Provider = provider, BaseUrl = "http://mixed-test.invalid/v1", Model = Guid.NewGuid().ToString("N"), SourceLanguage = source, TargetLanguage = "简体中文" };
        private static string WholeFailure(string input) => input.Contains("OT_KEEP") ? BrokenTokens
            : string.Join("\r", Outputs).Replace("PKG", "pkg");
        private static string ParagraphOutput(string input, bool badMiddle = false)
        {
            if (input.Contains("OT_KEEP")) return BrokenTokens;
            for (var i = 0; i < Sources.Length; i++)
                if (input == Sources[i]) return badMiddle && i == 1 ? Outputs[i].Replace("PKG", "pkg") : Outputs[i];
            if (input == "Only English") Assert.Fail("A fully protected paragraph must not call the model.");
            throw new AssertFailedException("Unexpected test input");
        }

        [DataTestMethod]
        [DataRow(ProviderKind.OpenAiCompatible)]
        [DataRow(ProviderKind.Ollama)]
        public async Task WholeProtectionFailure_ThreeParagraphsTranslateWithoutReview(ProviderKind provider)
        {
            var source = string.Join("\r", Sources);
            var handler = new Handler(provider, (input, call) => call <= 2 ? WholeFailure(input) : ParagraphOutput(input));
            using (var client = new TranslationClient(handler))
            {
                var settings = Settings(provider);
                var result = await client.TranslateDetailedAsync(source, settings, CancellationToken.None);
                Assert.AreEqual(string.Join("\r", Outputs), result.Text);
                Assert.IsTrue(result.Changed); Assert.IsFalse(result.NeedsReview);
                Assert.AreEqual(3, result.FallbackTranslatedParagraphs); Assert.AreEqual(0, result.FallbackReviewParagraphs);
                Assert.AreEqual(5, handler.Inputs.Count);
                var cached = await client.TranslateDetailedAsync(source, settings, CancellationToken.None);
                Assert.IsTrue(cached.FromCache); Assert.AreEqual(5, handler.Inputs.Count);
            }
        }

        [DataTestMethod]
        [DataRow(ProviderKind.OpenAiCompatible)]
        [DataRow(ProviderKind.Ollama)]
        public async Task OneParagraphStillFails_PreservesOnlyItsOriginalAndMarksPartialReview(ProviderKind provider)
        {
            var source = string.Join("\r\n", Sources);
            var handler = new Handler(provider, (input, call) => input.Contains("\r\n") ? WholeFailure(input) : ParagraphOutput(input, true));
            using (var client = new TranslationClient(handler))
            {
                var settings = Settings(provider);
                var result = await client.TranslateDetailedAsync(source, settings, CancellationToken.None);
                Assert.AreEqual(Outputs[0] + "\r\n" + Sources[1] + "\r\n" + Outputs[2], result.Text);
                Assert.IsTrue(result.Changed); Assert.IsTrue(result.NeedsReview);
                Assert.AreEqual(2, result.FallbackTranslatedParagraphs); Assert.AreEqual(1, result.FallbackReviewParagraphs);
                Assert.IsFalse(result.Text.Contains("OT_KEEP")); Assert.AreEqual(6, handler.Inputs.Count);
                var summary = new TranslationTaskSummary(1); summary.Record(result);
                Assert.AreEqual(1, summary.Translated); Assert.AreEqual(1, summary.NeedsReview);
                Assert.AreEqual(1, summary.TextPartiallyTranslated); Assert.AreEqual(0, summary.TextUnchangedNeedsReview);
                Assert.AreEqual(2, summary.FallbackTranslatedParagraphs); Assert.AreEqual(1, summary.FallbackReviewParagraphs);
                var again = await client.TranslateDetailedAsync(source, settings, CancellationToken.None);
                Assert.IsFalse(again.FromCache); Assert.AreEqual(result.Text, again.Text); Assert.AreEqual(12, handler.Inputs.Count);
            }
        }

        [DataTestMethod]
        [DataRow("\r\n")]
        [DataRow("\r")]
        [DataRow("\n")]
        [DataRow("\v")]
        [DataRow("\f")]
        [DataRow("\u2028")]
        [DataRow("\u2029")]
        public async Task Fallback_PreservesSeparatorsBlankParagraphsAndEdgeWhitespace(string separator)
        {
            var source = separator + "  " + Sources[0] + "\t" + separator + " \t" + separator + Sources[1] + separator;
            var handler = new Handler(ProviderKind.OpenAiCompatible, (input, call) => call <= 2 ? WholeFailure(input) : ParagraphOutput(input));
            using (var client = new TranslationClient(handler))
            {
                var result = await client.TranslateDetailedAsync(source, Settings(), CancellationToken.None);
                Assert.AreEqual(separator + "  " + Outputs[0] + "\t" + separator + " \t" + separator + Outputs[1] + separator, result.Text);
                Assert.AreEqual(4, handler.Inputs.Count); Assert.IsFalse(result.NeedsReview);
            }
        }

        [TestMethod]
        public async Task AutomaticSource_MultilineSuccessStillUsesOneRequest()
        {
            var handler = new Handler(ProviderKind.OpenAiCompatible, (input, call) => string.Join("\r", Outputs));
            using (var client = new TranslationClient(handler))
            {
                var result = await client.TranslateDetailedAsync(string.Join("\r", Sources), Settings(source: "自动检测"), CancellationToken.None);
                Assert.IsFalse(result.NeedsReview); Assert.AreEqual(1, handler.Inputs.Count);
                Assert.AreEqual(0, result.FallbackTranslatedParagraphs);
            }
        }

        [TestMethod]
        public async Task AutomaticSource_UnchangedMultilineDoesNotEnterMixedFallback()
        {
            var source = string.Join("\r", Sources);
            var handler = new Handler(ProviderKind.OpenAiCompatible, (input, call) => source);
            using (var client = new TranslationClient(handler))
            {
                var result = await client.TranslateDetailedAsync(source, Settings(source: "自动检测"), CancellationToken.None);
                Assert.AreEqual(source, result.Text); Assert.IsTrue(result.NeedsReview);
                Assert.AreEqual(2, handler.Inputs.Count); Assert.AreEqual(0, result.FallbackReviewParagraphs);
            }
        }

        [TestMethod]
        public async Task ValidWholeProtectedRetry_DoesNotAddParagraphRequests()
        {
            var source = string.Join("\r", Sources);
            var handler = new Handler(ProviderKind.OpenAiCompatible, (input, call) => {
                if (call == 1) return WholeFailure(input);
                var translated = input;
                for (var i = 0; i < Sources.Length; i++)
                    translated = translated.Replace(Sources[i].Substring(Sources[i].IndexOf(' ')), Outputs[i].Substring(Outputs[i].IndexOf(' ')));
                return translated;
            });
            using (var client = new TranslationClient(handler))
            {
                var result = await client.TranslateDetailedAsync(source, Settings(), CancellationToken.None);
                Assert.AreEqual(string.Join("\r", Outputs), result.Text); Assert.IsFalse(result.NeedsReview);
                Assert.AreEqual(2, handler.Inputs.Count); Assert.AreEqual(0, result.FallbackTranslatedParagraphs);
            }
        }

        [TestMethod]
        public async Task ValidWholeMixedTranslation_DoesNotAddParagraphRequests()
        {
            var handler = new Handler(ProviderKind.OpenAiCompatible, (input, call) => string.Join("\r", Outputs));
            using (var client = new TranslationClient(handler))
            {
                var result = await client.TranslateDetailedAsync(string.Join("\r", Sources), Settings(), CancellationToken.None);
                Assert.IsFalse(result.NeedsReview); Assert.AreEqual(1, handler.Inputs.Count);
                Assert.AreEqual(0, result.FallbackTranslatedParagraphs);
            }
        }

        [TestMethod]
        public async Task SingleMixedParagraph_SuccessStillUsesOneRequest()
        {
            var handler = new Handler(ProviderKind.OpenAiCompatible, (input, call) => Outputs[0]);
            using (var client = new TranslationClient(handler))
            {
                var result = await client.TranslateDetailedAsync(Sources[0], Settings(), CancellationToken.None);
                Assert.AreEqual(Outputs[0], result.Text); Assert.IsFalse(result.NeedsReview); Assert.AreEqual(1, handler.Inputs.Count);
            }
        }

        [TestMethod]
        public async Task SingleMixedParagraph_FailureIsStillBoundedToTwoRequests()
        {
            var handler = new Handler(ProviderKind.OpenAiCompatible, (input, call) => input.Contains("OT_KEEP") ? BrokenTokens : Outputs[0].Replace("SESS", "sess"));
            using (var client = new TranslationClient(handler))
            {
                var result = await client.TranslateDetailedAsync(Sources[0] + "\r\n\r\n", Settings(), CancellationToken.None);
                Assert.AreEqual(Sources[0] + "\r\n\r\n", result.Text); Assert.IsTrue(result.NeedsReview);
                Assert.IsFalse(result.Changed); Assert.AreEqual(2, handler.Inputs.Count);
            }
        }

        [TestMethod]
        public async Task AllParagraphsFail_LeavesWholeOriginalAndReportsUnchanged()
        {
            var source = string.Join("\r", Sources);
            var handler = new Handler(ProviderKind.OpenAiCompatible, (input, call) => input.Contains("OT_KEEP") ? BrokenTokens : "已翻译但英文缩写全部丢失。");
            using (var client = new TranslationClient(handler))
            {
                var result = await client.TranslateDetailedAsync(source, Settings(), CancellationToken.None);
                Assert.AreEqual(source, result.Text); Assert.IsFalse(result.Changed); Assert.IsTrue(result.NeedsReview);
                Assert.AreEqual(3, result.FallbackReviewParagraphs); Assert.AreEqual(8, handler.Inputs.Count);
                var summary = new TranslationTaskSummary(1); summary.Record(result);
                Assert.AreEqual(1, summary.TextUnchangedNeedsReview); Assert.AreEqual(0, summary.TextPartiallyTranslated);
            }
        }

        [TestMethod]
        public async Task FullyProtectedParagraph_RemainsExactWithoutRequestOrReview()
        {
            var handler = new Handler(ProviderKind.OpenAiCompatible, (input, call) => call <= 2 ? WholeFailure(input) : ParagraphOutput(input));
            using (var client = new TranslationClient(handler))
            {
                var result = await client.TranslateDetailedAsync(Sources[0] + "\rOnly English\r" + Sources[1], Settings(), CancellationToken.None);
                Assert.AreEqual(Outputs[0] + "\rOnly English\r" + Outputs[1], result.Text);
                Assert.IsFalse(result.NeedsReview); Assert.AreEqual(4, handler.Inputs.Count);
            }
        }

        [TestMethod]
        public async Task UserCancelDuringParagraphFallback_StopsBeforeFollowingParagraphs()
        {
            using (var cancellation = new CancellationTokenSource())
            {
                var handler = new Handler(ProviderKind.OpenAiCompatible, (input, call) => {
                    if (call <= 2) return WholeFailure(input);
                    cancellation.Cancel(); return Outputs[0]; });
                using (var client = new TranslationClient(handler))
                    await Assert.ThrowsExceptionAsync<OperationCanceledException>(() => client.TranslateDetailedAsync(string.Join("\r", Sources), Settings(), cancellation.Token));
                Assert.AreEqual(3, handler.Inputs.Count);
            }
        }
    }
}
