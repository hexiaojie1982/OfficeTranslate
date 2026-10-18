using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeTranslate.Core;

namespace OfficeTranslate.Tests
{
    [TestClass]
    public class DocumentChatTests
    {
        private static TranslationSettings Settings() => new TranslationSettings { Provider = ProviderKind.OpenAiCompatible, BaseUrl = "http://fixture/v1", Model = "fixture", ChatContextTokens = 4096, ChatOutputTokens = 512 };
        [DataTestMethod]
        [DataRow("selected", true)] [DataRow("", false)] [DataRow(" \r\n", false)] [DataRow(null, false)]
        public void Context_PrefersSelectionAndLazilyReadsWholeDocument(string? selection, bool expected)
        {
            var reads = 0; var context = ChatContext.Capture(selection, () => { reads++; return "whole document"; });
            Assert.AreEqual(expected, context.IsSelection); Assert.AreEqual(expected ? selection : "whole document", context.Text); Assert.AreEqual(expected ? 0 : 1, reads);
        }
        [DataTestMethod]
        [DataRow(4)] [DataRow(17)] [DataRow(256)] [DataRow(8192)]
        public void Chunking_PreservesUnicodeAndEveryCharacterWithinBudget(int budget)
        {
            var source = string.Concat(Enumerable.Repeat("第一行😀\r\nsecond line.\n", 1000));
            var chunks = ChatBudget.Split(source, budget);
            Assert.AreEqual(source, string.Concat(chunks));
            foreach (var chunk in chunks) { Assert.IsTrue(ChatBudget.Estimate(chunk) <= budget); Assert.IsFalse(char.IsLowSurrogate(chunk[0])); Assert.IsFalse(char.IsHighSurrogate(chunk[chunk.Length - 1])); }
            Assert.AreEqual(0, ChatBudget.Split("", budget).Count);
        }
        [TestMethod]
        public async Task LongDocument_MapsAndRecursivelyReduces_ReservesAllRequestBudgets()
        {
            var settings = Settings(); var calls = 0; var maps = 0; var progress = new List<int>();
            var session = new DocumentChatSession();
            var reply = await session.SendAsync(new ChatContext(new string('文', 30000), false), "Summarize", ChatMode.Summarize, settings,
                (messages, s, output, display, token) => {
                    calls++; Assert.IsTrue(ChatBudget.Cost(messages) + output <= s.ChatContextTokens);
                    if (output == 256) { maps++; return Task.FromResult(new ChatReply(new string('a', 400), maps == 1)); }
                    Assert.IsFalse(messages.Last().Content.Contains(new string('文', 100))); display("answer"); return Task.FromResult(new ChatReply("answer", false));
                }, _ => { }, (i,n) => progress.Add(n), CancellationToken.None);
            Assert.IsTrue(calls > 10 && maps > 1); Assert.IsTrue(progress.Contains(0)); Assert.IsTrue(reply.UsedFallback); Assert.AreEqual(2, session.History.Count);
            Assert.IsFalse(session.History.Any(m => m.Content.Contains(new string('文', 100))));
        }
        [TestMethod]
        public async Task History_KeepsTenCompleteTurns_AndClearDiscardsThem()
        {
            var session = new DocumentChatSession();
            for (var i = 0; i < 12; i++) await session.SendAsync(new ChatContext("document", true), "Q" + i, ChatMode.Question, Settings(),
                (m,s,n,o,t) => Task.FromResult(new ChatReply("A", false)), _=>{}, (i,n)=>{}, CancellationToken.None);
            Assert.AreEqual(20, session.History.Count); Assert.AreEqual("Q2", session.History[0].Content); Assert.AreEqual("Q11", session.History[18].Content);
            session.Clear(); Assert.AreEqual(0, session.History.Count);
        }
        [TestMethod]
        public async Task FailureAndCancellation_DoNotCommitHistoryOrWrite()
        {
            var session = new DocumentChatSession(); var writes = 0; var context = new ChatContext("doc", true, (a,t)=>writes++, true);
            await Assert.ThrowsExceptionAsync<IOException>(() => session.SendAsync(context, "Q", ChatMode.Rewrite, Settings(),
                (m,s,n,o,t) => throw new IOException(), _=>{}, (i,n)=>{}, CancellationToken.None));
            using (var cts = new CancellationTokenSource()) {
                await Assert.ThrowsExceptionAsync<OperationCanceledException>(() => session.SendAsync(context, "Q", ChatMode.Rewrite, Settings(),
                    (m,s,n,o,t) => { cts.Cancel(); return Task.FromResult(new ChatReply("partial", false)); }, _=>{}, (i,n)=>{}, cts.Token));
            }
            Assert.AreEqual(0, session.History.Count); Assert.AreEqual(0, writes);
        }
        [TestMethod]
        public async Task NonReducingSummariesAndOversizeQuestion_FailWithoutHistory()
        {
            var session = new DocumentChatSession();
            await Assert.ThrowsExceptionAsync<ChatException>(() => session.SendAsync(new ChatContext(new string('a', 20000), false), "Q", ChatMode.Question, Settings(),
                (m,s,n,o,t) => Task.FromResult(new ChatReply(m.Last().Content, false)), _=>{}, (i,n)=>{}, CancellationToken.None));
            await Assert.ThrowsExceptionAsync<ChatException>(() => session.SendAsync(new ChatContext("doc", false), new string('x', 9000), ChatMode.Question, Settings(),
                (m,s,n,o,t) => throw new AssertFailedException("No HTTP expected"), _=>{}, (i,n)=>{}, CancellationToken.None));
            Assert.AreEqual(0, session.History.Count);
        }
        [DataTestMethod]
        [DataRow(ChatMode.Question, false, false, false)]
        [DataRow(ChatMode.Summarize, false, false, true)]
        [DataRow(ChatMode.Rewrite, false, true, false)]
        [DataRow(ChatMode.Rewrite, true, true, false)]
        public void WriteButtons_OnlyPermitMatchingCompletedAction(ChatMode mode, bool selected, bool below, bool start)
        {
            var writes = 0; var context = new ChatContext("source", selected, (a,t) => writes++, selected);
            Assert.AreEqual(mode == ChatMode.Rewrite && selected, context.Allows(mode, ChatWriteAction.ReplaceSelection, "result"));
            Assert.AreEqual(below, context.Allows(mode, ChatWriteAction.InsertBelow, "result")); Assert.AreEqual(start, context.Allows(mode, ChatWriteAction.InsertAtStart, "result"));
            Assert.IsFalse(context.Allows(mode, ChatWriteAction.InsertAtStart, "")); Assert.AreEqual(0, writes);
            if (start) context.Apply(mode, ChatWriteAction.InsertAtStart, "result"); else if (below) context.Apply(mode, ChatWriteAction.InsertBelow, "result");
            else Assert.ThrowsException<InvalidOperationException>(()=>context.Apply(mode, ChatWriteAction.InsertAtStart, "result"));
            Assert.AreEqual(start || below ? 1 : 0, writes);
        }
        [TestMethod]
        public async Task ClearDuringFlight_CannotResurrectConversation()
        {
            var session=new DocumentChatSession();
            await Assert.ThrowsExceptionAsync<OperationCanceledException>(()=>session.SendAsync(new ChatContext("doc",false),"Q",ChatMode.Question,Settings(),
                (m,s,n,o,t)=>{session.Clear();return Task.FromResult(new ChatReply("answer",false));},_=>{},(i,n)=>{},CancellationToken.None));
            Assert.AreEqual(0,session.History.Count);
        }
        [TestMethod]
        public async Task CustomPromptsAndChangingModelBudgets_AreAppliedPerRequest()
        {
            var session=new DocumentChatSession();var settings=Settings();settings.ChatTranslateSystemPrompt="TRANSLATE SYSTEM";settings.ChatQuestionSystemPrompt="QUESTION SYSTEM";settings.ChatSummarySystemPrompt="SUMMARY SYSTEM";settings.ChatRewriteSystemPrompt="REWRITE SYSTEM";
            foreach(ChatMode mode in Enum.GetValues(typeof(ChatMode)))
                await session.SendAsync(new ChatContext(new string('d',200),true),"Q",mode,settings,(m,s,n,o,t)=>{
                    StringAssert.StartsWith(m[0].Content,mode==ChatMode.Translate?"TRANSLATE SYSTEM":mode==ChatMode.Question?"QUESTION SYSTEM":mode==ChatMode.Summarize?"SUMMARY SYSTEM":"REWRITE SYSTEM");
                    Assert.IsTrue(ChatBudget.Cost(m)+n<=s.ChatContextTokens);return Task.FromResult(new ChatReply(new string('a',1500),false));},_=>{},(i,n)=>{},CancellationToken.None);
            settings.ChatContextTokens=2048;settings.ChatOutputTokens=256;
            await session.SendAsync(new ChatContext("doc",false),"Q",ChatMode.Question,settings,(m,s,n,o,t)=>{
                Assert.AreEqual(2,m.Count);Assert.IsTrue(ChatBudget.Cost(m)+n<=s.ChatContextTokens);return Task.FromResult(new ChatReply("trimmed",false));},_=>{},(i,n)=>{},CancellationToken.None);
        }
        private static string Event(string text) => "data: " + new JavaScriptSerializer().Serialize(new { choices = new[] { new { index = 0, delta = new { content = text }, finish_reason = (string?)null } } }) + "\n\n";
        [TestMethod]
        public async Task Sse_ParsesCommentsRoleEmptyChoicesUnicodeAndDone_Incrementally()
        {
            var wire = ": ping\n\ndata: {\"choices\":[{\"delta\":{\"role\":\"assistant\"}}]}\n\n" + Event("你好") + Event("😀\nnext") + "data: {\"choices\":[]}\n\ndata: [DONE]\n\n";
            var output = new List<string>();
            var text = await ChatSseParser.ReadAsync(new StringReader(wire), output.Add, CancellationToken.None);
            Assert.AreEqual("你好😀\nnext", text); CollectionAssert.AreEqual(new[] { "你好", "你好😀\nnext" }, output);
        }
        [DataTestMethod]
        [DataRow("data: [DONE]\n\n")]
        [DataRow("data: {\"choices\":[{\"delta\":{\"content\":\"half\"}}]}\n\n")]
        [DataRow("data: {\"error\":{}}\n\n")]
        [DataRow("data: {\"choices\":[{\"finish_reason\":\"length\"}]}\n\ndata: [DONE]\n\n")]
        public async Task Sse_RejectsEmptyInterruptedErrorAndTruncated(string wire)
        { await Assert.ThrowsExceptionAsync<ChatException>(()=>ChatSseParser.ReadAsync(new StringReader(wire), _=>{}, CancellationToken.None)); }
        private sealed class Handler : HttpMessageHandler
        {
            public readonly List<Dictionary<string,object>> Bodies = new List<Dictionary<string,object>>();
            public Func<int, CancellationToken, HttpResponseMessage> Reply = null!;
            protected override async Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken token)
            { Bodies.Add((Dictionary<string,object>)new JavaScriptSerializer().DeserializeObject(await request.Content!.ReadAsStringAsync())); return Reply(Bodies.Count, token); }
        }
        private static HttpResponseMessage Response(string body, string type = "application/json", HttpStatusCode status = HttpStatusCode.OK) => new HttpResponseMessage(status) { Content = new StringContent(body, Encoding.UTF8, type) };
        [DataTestMethod]
        [DataRow(false)] [DataRow(true)]
        public async Task Transport_SuccessStreams_OrReplacesInterruptedOutputWithFallback(bool interrupt)
        {
            var handler = new Handler { Reply = (i,t) => i == 1 ? Response(Event("half") + (interrupt ? "" : "data: [DONE]\n\n"), "text/event-stream") : Response("{\"choices\":[{\"message\":{\"content\":\"complete\"},\"finish_reason\":\"stop\"}]}") };
            using (var client = new TranslationClient(handler)) {
                var updates = new List<string>(); var result = await client.ChatAsync(new[]{new ChatMessage("user","Q")}, Settings(), 512, updates.Add, CancellationToken.None);
                Assert.AreEqual(interrupt ? "complete" : "half", result.Text); Assert.AreEqual(interrupt, result.UsedFallback);
                Assert.AreEqual(interrupt ? 2 : 1, handler.Bodies.Count); Assert.AreEqual(true, handler.Bodies[0]["stream"]);
                if (interrupt) { Assert.AreEqual(false, handler.Bodies[1]["stream"]); Assert.IsTrue(updates.Contains("")); Assert.AreEqual("complete", updates.Last()); }
            }
        }
        [DataTestMethod]
        [DataRow("length")] [DataRow("content_filter")]
        public async Task Transport_OrdinaryRecoveryRejectsPartialOutput(string reason)
        {
            var handler=new Handler { Reply=(i,t)=>i==1?Response("{}"):Response("{\"choices\":[{\"message\":{\"content\":\"partial\"},\"finish_reason\":\""+reason+"\"}]}") };
            using(var client=new TranslationClient(handler)) await Assert.ThrowsExceptionAsync<ChatException>(()=>client.ChatAsync(new[]{new ChatMessage("user","Q")},Settings(),512,_=>{},CancellationToken.None));
            Assert.AreEqual(2,handler.Bodies.Count);
        }
        [TestMethod]
        public async Task Transport_UserCancellationNeverFallsBack()
        {
            using (var cts = new CancellationTokenSource()) {
                var handler = new Handler { Reply = (i,t) => { cts.Cancel(); throw new OperationCanceledException(cts.Token); } };
                using(var client = new TranslationClient(handler)) { try { await client.ChatAsync(new[]{new ChatMessage("user","Q")},Settings(),512,_=>{},cts.Token); Assert.Fail("Expected cancellation"); } catch(OperationCanceledException) { } }
                Assert.AreEqual(1, handler.Bodies.Count);
            }
        }
        [TestMethod]
        public async Task Transport_UnsupportedSseAndOllamaUseOrdinaryCompletion()
        {
            foreach (var ollama in new[]{false,true}) {
                var handler = new Handler { Reply = (i,t) => ollama ? Response("{\"message\":{\"content\":\"ollama\"},\"done\":true,\"done_reason\":\"stop\"}") : i==1 ? Response("{}") : Response("{\"choices\":[{\"message\":{\"content\":\"recovered\"}}]}") };
                var settings=Settings(); if(ollama) settings.Provider=ProviderKind.Ollama;
                using(var client=new TranslationClient(handler)) { var result=await client.ChatAsync(new[]{new ChatMessage("user","Q")},settings,512,_=>{},CancellationToken.None); Assert.IsTrue(result.UsedFallback); Assert.AreEqual(ollama?1:2,handler.Bodies.Count); }
            }
        }
    }
}
