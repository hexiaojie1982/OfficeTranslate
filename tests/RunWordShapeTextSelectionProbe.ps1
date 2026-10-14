param([ValidateSet('Debug','Release')][string]$Configuration = 'Release')
$ErrorActionPreference = 'Stop'
# Run with Windows PowerShell -STA in the logged-in desktop. Uses only a new,
# unsaved document and a loopback model stub; does not install the candidate.
$root = Split-Path -Parent $PSScriptRoot
$output = Join-Path $root "src\OfficeTranslate.WordAddIn\bin\$Configuration\net48"
$wordInterop = Join-Path $root 'packages\microsoft.office.interop.word\15.0.4797.1004\lib\net20\Microsoft.Office.Interop.Word.dll'
$officeInterop = 'C:\Windows\assembly\GAC_MSIL\office\15.0.0.0__71e9bce111e9429c\office.dll'
[Reflection.Assembly]::LoadFrom($wordInterop) | Out-Null
[Reflection.Assembly]::LoadFrom((Join-Path $output 'OfficeTranslate.Core.dll')) | Out-Null
Add-Type -ReferencedAssemblies @($wordInterop, $officeInterop, (Join-Path $output 'OfficeTranslate.Core.dll'), 'System.Windows.Forms', 'System.Drawing') -TypeDefinition @'
using System;
using System.Collections;
using System.IO;
using System.Net;
using System.Net.Sockets;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using OfficeTranslate.Core;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;

public static class WordShapeTextProbe
{
    static void Check(bool ok, string message) { if (!ok) throw new Exception(message); }
    static string Text(Word.Shape shape) { return shape.TextFrame.TextRange.Text.TrimEnd('\r', '\a'); }
    public static string Run(string assemblyPath)
    {
        var log = new StringBuilder();
        var type = Assembly.LoadFrom(assemblyPath).GetType("OfficeTranslate.WordAddIn.WordTranslationService", true);
        var read = type.GetMethod("ReadSelection", BindingFlags.Instance | BindingFlags.NonPublic);
        var write = type.GetMethod("WriteTarget", BindingFlags.Static | BindingFlags.NonPublic);
        Word.Application word = null;
        Word.Document doc = null;
        try
        {
            word = new Word.Application(); word.Visible = false;
            word.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone;
            doc = word.Documents.Add();
            doc.Content.Text = "Unselected body paragraph.\r";
            string body = doc.Content.Text;
            object anchor = doc.Range(0, 0);
            var box = doc.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, 30, 30, 300, 70, ref anchor);
            box.TextFrame.TextRange.Text = "OCR source sentence.";
            box.AlternativeText = ImageRecognitionTags.OverlayFor("wd:probe-image");
            box.TextFrame.TextRange.Font.Name = "Arial";
            box.TextFrame.TextRange.Font.Size = 19;
            box.TextFrame.TextRange.Font.Bold = -1;
            box.TextFrame.TextRange.ParagraphFormat.SpaceAfter = 7;
            int initialBold = box.TextFrame.TextRange.Font.Bold;
            string initialAscii = box.TextFrame.TextRange.Font.NameAscii;
            string initialFarEast = box.TextFrame.TextRange.Font.NameFarEast;
            var auto = doc.Shapes.AddShape(1, 30, 130, 300, 70, ref anchor);
            auto.TextFrame.TextRange.Text = "AutoShape source sentence.";
            var empty = doc.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, 30, 230, 300, 70, ref anchor);
            var service = type.GetConstructors()[0].Invoke(new object[] { word });
            object replace = true; box.Select(ref replace);
            Check(word.Selection.Type == Word.WdSelectionType.wdSelectionShape, "Not an actual Word shape selection");
            var targets = (IList)read.Invoke(service, null);
            Check(targets.Count == 1, "OCR-marked textbox was not collected");

            // Invoke the complete selection translation pipeline, with a real
            // local HTTP response and STA dispatcher, then check the COM result.
            var probe = new TcpListener(IPAddress.Loopback, 0); probe.Start();
            int port = ((IPEndPoint)probe.LocalEndpoint).Port; probe.Stop();
            using (var listener = new HttpListener())
            using (var deadline = new CancellationTokenSource(20000))
            {
                listener.Prefixes.Add("http://127.0.0.1:" + port + "/"); listener.Start();
                var reply = Task.Run(() => {
                    var context = listener.GetContext();
                    using (var reader = new StreamReader(context.Request.InputStream)) {
                        Check(reader.ReadToEnd().Contains("OCR source sentence."), "Wrong text sent to translator");
                    }
                    byte[] bytes = Encoding.UTF8.GetBytes("{\"choices\":[{\"message\":{\"content\":\"这是识别文本的中文译文。\"},\"finish_reason\":\"stop\"}]}");
                    context.Response.ContentType = "application/json";
                    context.Response.ContentLength64 = bytes.Length;
                    context.Response.OutputStream.Write(bytes, 0, bytes.Length); context.Response.Close();
                });
                var settings = new TranslationSettings { Provider = ProviderKind.OpenAiCompatible,
                    BaseUrl = "http://127.0.0.1:" + port + "/v1", Model = Guid.NewGuid().ToString(),
                    TargetLanguage = "简体中文", ImageOcrEnabled = false };
                var ui = OfficeUiDispatcher.Capture("WordShapeTextProbe");
                var task = (Task<TranslationTaskSummary>)type.GetMethod("TranslateAsync").Invoke(service,
                    new object[] { false, false, settings, deadline.Token, new Action<string>(_ => {}), ui });
                while (!task.IsCompleted && !deadline.IsCancellationRequested) {
                    System.Windows.Forms.Application.DoEvents(); Thread.Sleep(10);
                }
                Check(task.IsCompleted, "TranslateAsync timed out");
                var summary = task.GetAwaiter().GetResult();
                Check(summary.Total == 1 && summary.Translated == 1, "Unexpected task summary");
                Check(reply.Wait(2000), "Loopback server did not finish");
            }
            Check(Text(box) == "这是识别文本的中文译文。", "Translated text did not reach selected textbox");
            Check(doc.Content.Text == body, "Unselected body text changed");
            Check(box.TextFrame.TextRange.Font.Size == 19 && box.TextFrame.TextRange.Font.Bold == initialBold
                && box.TextFrame.TextRange.Font.NameAscii == initialAscii
                && box.TextFrame.TextRange.Font.NameFarEast == initialFarEast,
                "Textbox font was not retained: size=" + box.TextFrame.TextRange.Font.Size
                + " bold=" + box.TextFrame.TextRange.Font.Bold + "/" + initialBold
                + " ascii=" + box.TextFrame.TextRange.Font.NameAscii + "/" + initialAscii
                + " farEast=" + box.TextFrame.TextRange.Font.NameFarEast + "/" + initialFarEast);
            Check(box.TextFrame.TextRange.ParagraphFormat.SpaceAfter == 7, "Paragraph formatting changed");
            Check(box.AlternativeText == ImageRecognitionTags.OverlayFor("wd:probe-image"), "OCR marker changed");
            log.AppendLine("PASS actual Word selected OCR textbox -> TranslateAsync -> HTTP -> writeback; 19pt, bold, Western/East Asian font settings and paragraph format retained");

            box.Select(ref replace);
            object extend = false;
            auto.Select(ref extend); empty.Select(ref extend);
            targets = (IList)read.Invoke(service, null);
            Check(targets.Count == 2, "Multi-shape selection did not skip the empty box");
            write.Invoke(null, new object[] { targets[0], "第一个文本框。", false });
            write.Invoke(null, new object[] { targets[1], "第二个图形。", false });
            Check(Text(box) == "第一个文本框。" && Text(auto) == "第二个图形。", "Multi-shape writeback went to wrong frame");
            Check(doc.Content.Text == body, "Multi-shape selection changed anchor body");
            log.AppendLine("PASS actual Word multi-shape selection: textbox + AutoShape + empty box; correct per-shape writeback");

            box.Select(ref replace); targets = (IList)read.Invoke(service, null);
            write.Invoke(null, new object[] { targets[0], "Bilingual translation.", true });
            Check(Text(box) == "第一个文本框。\rBilingual translation.", "Bilingual text escaped the selected box");
            Check(box.TextFrame.TextRange.Font.Size == 19, "Bilingual append lost font size");
            log.AppendLine("PASS actual Word bilingual append stays inside textbox and preserves 19pt");
            return log.ToString();
        }
        finally
        {
            if (doc != null) ((Word._Document)doc).Close(Word.WdSaveOptions.wdDoNotSaveChanges);
            if (word != null) ((Word._Application)word).Quit(Word.WdSaveOptions.wdDoNotSaveChanges);
        }
    }
}
'@
[WordShapeTextProbe]::Run((Join-Path $output 'OfficeTranslate.WordAddIn.dll'))
