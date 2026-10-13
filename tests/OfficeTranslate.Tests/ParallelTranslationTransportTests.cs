using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeTranslate.Core;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;

namespace OfficeTranslate.Tests
{
    [TestClass]
    public class ParallelTranslationTransportTests
    {
        [TestMethod]
        [Timeout(25000)]
        public async Task RealHttp_FourConcurrentRequests_SharedClient_StrictRetryAndCachePreserved()
        {
            // Loopback only: no model service, API key or user documents.
            var portProbe = new TcpListener(IPAddress.Loopback, 0);
            portProbe.Start();
            var port = ((IPEndPoint)portProbe.LocalEndpoint).Port;
            portProbe.Stop();
            using (var listener = new HttpListener())
            using (var deadline = new CancellationTokenSource(15000))
            using (var client = TranslationClient.CreateForParallelCells(4))
            {
                listener.Prefixes.Add($"http://127.0.0.1:{port}/");
                listener.Start();
                var received = 0;
                var firstWave = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
                var attempts = new ConcurrentDictionary<string, int>();
                var handlers = new List<Task>();
                using (deadline.Token.Register(() => { firstWave.TrySetCanceled(); listener.Stop(); }))
                {
                    async Task ReplyAsync(HttpListenerContext context)
                    {
                        var json = new JavaScriptSerializer();
                        string body;
                        using (var reader = new StreamReader(context.Request.InputStream)) body = await reader.ReadToEndAsync();
                        var root = (Dictionary<string, object>)json.DeserializeObject(body);
                        var messages = (object[])root["messages"];
                        var source = (string)((Dictionary<string, object>)messages[1])["content"];
                        var attempt = attempts.AddOrUpdate(source, 1, (_, n) => n + 1);
                        if (Interlocked.Increment(ref received) == 4) firstWave.TrySetResult(true);
                        // No response until four real HTTP connections have reached the server.
                        await firstWave.Task;
                        var response = attempt == 1 ? source : "这是完整的中文翻译。";
                        var bytes = Encoding.UTF8.GetBytes(json.Serialize(new {
                            choices = new[] { new { message = new { content = response }, finish_reason = "stop" } }
                        }));
                        context.Response.ContentType = "application/json";
                        context.Response.ContentLength64 = bytes.Length;
                        await context.Response.OutputStream.WriteAsync(bytes, 0, bytes.Length);
                        context.Response.Close();
                    }
                    var server = Task.Run(async () =>
                    {
                        for (var i = 0; i < 8; i++) handlers.Add(ReplyAsync(await listener.GetContextAsync()));
                        await Task.WhenAll(handlers);
                    });
                    try
                    {
                        var settings = new TranslationSettings { Provider = ProviderKind.OpenAiCompatible,
                            BaseUrl = $"http://127.0.0.1:{port}/v1", Model = Guid.NewGuid().ToString(), TargetLanguage = "简体中文" };
                        var texts = new[] { "First source sentence.", "Second source sentence.", "Third source sentence.", "Fourth source sentence." };
                        var results = await Task.WhenAll(texts.Select(text => Task.Run(() =>
                            client.TranslateDetailedAsync(text, settings, deadline.Token))));
                        await server;
                        Assert.IsTrue(results.All(result => result.Changed && !result.NeedsReview && !result.FromCache));
                        Assert.AreEqual(8, received);
                        Assert.IsTrue(attempts.Values.All(attempt => attempt == 2), "Strict retry remains per cell.");
                        var cached = await client.TranslateDetailedAsync(texts[0], settings, deadline.Token);
                        Assert.IsTrue(cached.FromCache);
                        Assert.AreEqual(8, received);
                    }
                    finally
                    {
                        deadline.Cancel();
                        try { await server; } catch (Exception) when (deadline.IsCancellationRequested) { }
                        try { await Task.WhenAll(handlers); } catch (Exception) when (deadline.IsCancellationRequested) { }
                    }
                }
            }
        }
    }
}
