using System;
using System.Collections.Generic;
using System.Dynamic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeTranslate.Core;

namespace OfficeTranslate.Tests
{
    [TestClass]
    public class DocumentChatHostTests
    {
        internal static Assembly Host(string host)
        {
            var configuration = new DirectoryInfo(AppDomain.CurrentDomain.BaseDirectory).Parent!.Name;
            return Assembly.LoadFrom(Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory,
                "../../../../../src/OfficeTranslate." + host + "AddIn/bin/" + configuration + "/net48/OfficeTranslate." + host + "AddIn.dll")));
        }
        public static dynamic Obj(params (string, object?)[] values)
        { var result = new ExpandoObject(); foreach(var pair in values) ((IDictionary<string,object?>)result).Add(pair.Item1,pair.Item2); return result; }
        public static dynamic Range(string text)
        {
            dynamic range=Obj(("Text",text),("Start",0),("End",text.Length)); range.Duplicate=range; range.Font=Obj(("Size",18));
            range.Collapse=(Action<int>)(i=>{}); range.InsertAfter=(Action<string>)(s=>range.Text+=s); range.InsertBefore=(Action<string>)(s=>range.Text=s+range.Text); return range;
        }
        private static ChatContext Capture(string host, object window) => (ChatContext)Host(host).GetType("OfficeTranslate.WordAddIn.DocumentChatHost",true)!
            .GetMethod("Capture",BindingFlags.Static|BindingFlags.NonPublic)!.Invoke(null,new[]{window})!;
        [DataTestMethod]
        [DataRow(true, ChatMode.Rewrite)] [DataRow(false, ChatMode.Rewrite)]
        [DataRow(true, ChatMode.Translate)] [DataRow(false, ChatMode.Translate)]
        public void Word_RealAdapterSelectsSelectionOrFullTextAndWritesOnlyOnApply(bool selected, ChatMode mode)
        {
            dynamic target=Range(selected?"selection":""); dynamic content=Range("whole document\r");
            dynamic document=Obj(("Name","sample.docx"),("Content",content),("Shapes",new ChatItems())); dynamic window=Obj(("Document",document),("Selection",Obj(("Range",target))));
            ChatContext context=Capture("Word",(object)window); Assert.AreEqual("sample.docx",context.DocumentName); Assert.AreEqual(selected,context.IsSelection); Assert.AreEqual(selected?"selection":"whole document\r",context.Text);
            Assert.AreEqual("Word",context.Host); Assert.AreEqual(selected,context.CanReplace); Assert.AreEqual(selected?"selection":"",(string)target.Text);
            if(selected) { context.Apply(mode,ChatWriteAction.ReplaceSelection,"rewritten"); Assert.AreEqual("rewritten",(string)target.Text); }
        }
        [TestMethod]
        public void Word_ChangedSelectionRejectsWriteback_AndShapeIsContext()
        {
            dynamic range=Range("selected"); dynamic frame=Obj(("HasText",1),("TextRange",range)); dynamic shape=Obj(("TextFrame",frame),("Type",1));
            dynamic window=Obj(("Document",Obj(("Content",Range("body")),("Shapes",new ChatItems(shape)))),("Selection",Obj(("Range",Range("")),("ShapeRange",new ChatItems(shape)))));
            ChatContext context=Capture("Word",(object)window); Assert.AreEqual("selected",context.Text); Assert.IsTrue(context.CanReplace);
            range.Text="edited"; Assert.ThrowsException<InvalidOperationException>(()=>context.Apply(ChatMode.Rewrite,ChatWriteAction.ReplaceSelection,"result")); Assert.AreEqual("edited",(string)range.Text);
        }
        [DataTestMethod]
        [DataRow(true, ChatMode.Rewrite)] [DataRow(false, ChatMode.Rewrite)]
        [DataRow(true, ChatMode.Translate)] [DataRow(false, ChatMode.Translate)]
        public void Excel_RealAdapterExtractsSelectionOrEntireWorkbook(bool selected, ChatMode mode)
        {
            dynamic range=Obj(("Value2",selected?"selection":""),("CountLarge",1d),("MergeCells",false),("Formula",selected?"selection":""),("Row",2),("Rows",Obj(("Count",1)))); range.Areas=new ChatItems(range);
            dynamic used=Obj(("Value2","whole sheet")); used.Areas=new ChatItems(used);
            dynamic sheet=Obj(("UsedRange",used),("Shapes",new ChatItems()),("Name","Sheet1")); dynamic workbook=Obj(("Name","sample.xlsx"),("Worksheets",new ChatItems(sheet))); sheet.Parent=workbook;
            dynamic window=Obj(("ActiveSheet",sheet),("RangeSelection",range)); dynamic app=Obj(); app.ActiveWindow=window; app.Intersect=(Func<object,object,object>)((a,b)=>range); window.Application=app;
            ChatContext context=Capture("Excel",(object)window); Assert.AreEqual("sample.xlsx",context.DocumentName); Assert.AreEqual(selected,context.IsSelection); StringAssert.Contains(context.Text,selected?"selection":"whole sheet");
            Assert.AreEqual("Excel",context.Host);
            if(selected) { context.Apply(mode,ChatWriteAction.ReplaceSelection,"=2+2"); Assert.AreEqual("'=2+2",(string)range.Value2); }
        }
        [TestMethod]
        public void Excel_MultipleCellsCannotReplace_AndSummaryInsertsTopRow()
        {
            var insertions=new List<int>(); dynamic top=Obj(("Insert",(Action)(()=>insertions.Add(1)))); dynamic cell=Obj(("Value2","old"));
            dynamic range=Obj(("Value2","selection"),("CountLarge",2d),("Row",2),("Rows",Obj(("Count",1)))); range.Areas=new ChatItems(range);
            dynamic sheet=Obj(("UsedRange",range),("Rows",new ChatItems(top)),("Cells",new ChatCells(cell)),("Name","sheet")); dynamic book=Obj(("Worksheets",new ChatItems(sheet))); sheet.Parent=book;
            dynamic app=Obj(("Intersect",(Func<object,object,object>)((a,b)=>range))); dynamic window=Obj(("ActiveSheet",sheet),("Application",app),("RangeSelection",range)); app.ActiveWindow=window;
            ChatContext context=Capture("Excel",(object)window); Assert.IsFalse(context.CanReplace); Assert.AreEqual(0,insertions.Count);
            context.Apply(ChatMode.Summarize,ChatWriteAction.InsertAtStart,"summary"); CollectionAssert.AreEqual(new[]{1},insertions); Assert.AreEqual("'summary",(string)cell.Value2);
        }
        [DataTestMethod]
        [DataRow(true, ChatMode.Rewrite)] [DataRow(false, ChatMode.Rewrite)]
        [DataRow(true, ChatMode.Translate)] [DataRow(false, ChatMode.Translate)]
        public void PowerPoint_RealAdapterSelectsTextOrEntirePresentation(bool selected, ChatMode mode)
        {
            dynamic target=Range("selected text"); dynamic shape=Obj(("Type",1),("HasTable",0),("HasTextFrame",-1),("TextFrame",Obj(("TextRange",Range("whole slide")))));
            dynamic slide=Obj(("Shapes",new ChatItems(shape))); dynamic presentation=Obj(("Name","sample.pptx"),("Slides",new ChatItems(slide)));
            dynamic selection=Obj(("Type",selected?3:0),("TextRange",target)); dynamic window=Obj(("Presentation",presentation),("Selection",selection));
            ChatContext context=Capture("PowerPoint",(object)window); Assert.AreEqual("sample.pptx",context.DocumentName); Assert.AreEqual(selected,context.IsSelection); StringAssert.Contains(context.Text,selected?"selected text":"whole slide");
            Assert.AreEqual("PowerPoint",context.Host);
            if(selected) { context.Apply(mode,ChatWriteAction.InsertBelow,"rewrite"); Assert.AreEqual("selected text\rrewrite",(string)target.Text); }
        }
        [TestMethod]
        public void PowerPoint_SummaryCreatesNewFirstSlide_AndDoesNotModifyExistingSlide()
        {
            dynamic text=Range(""); dynamic added=Obj(("Shapes",new ChatShapes(text))); var slides=new ChatSlides(added);
            dynamic presentation=Obj(("Slides",slides),("PageSetup",Obj(("SlideWidth",960f),("SlideHeight",540f))));
            dynamic target=Range("selection"); dynamic window=Obj(("Presentation",presentation),("Selection",Obj(("Type",3),("TextRange",target))));
            ChatContext context=Capture("PowerPoint",(object)window); Assert.AreEqual(0,slides.Added.Count);
            context.Apply(ChatMode.Summarize,ChatWriteAction.InsertAtStart,"summary"); CollectionAssert.AreEqual(new[]{1},slides.Added); Assert.AreEqual("summary",(string)text.Text); Assert.AreEqual("selection",(string)target.Text);
        }
        [DataTestMethod]
        [DataRow("Word")] [DataRow("Excel")] [DataRow("PowerPoint")]
        public void AllHosts_RegisterNativePaneAndRibbonEntry(string host)
        {
            var assembly=Host(host); var addin=assembly.GetType("OfficeTranslate."+host+"AddIn.AddIn",true)!;
            Assert.IsTrue(addin.GetInterfaces().Any(i=>i.Name=="ICustomTaskPaneConsumer")); Assert.IsNotNull(addin.GetMethod("OpenChat"));
            var pane=assembly.GetType("OfficeTranslate.WordAddIn.DocumentChatPane",true)!; Assert.IsTrue(typeof(UserControl).IsAssignableFrom(pane));
            Assert.AreEqual("OfficeTranslate."+host+"ChatPane",pane.GetCustomAttribute<System.Runtime.InteropServices.ProgIdAttribute>()!.Value);
            using(var stream=assembly.GetManifestResourceStream("OfficeTranslate."+host+"AddIn.Ribbon.xml")) using(var reader=new StreamReader(stream!)) StringAssert.Contains(reader.ReadToEnd(),"onAction=\"OpenChat\"");
            Assert.AreEqual(new Version(2,3,8,0),assembly.GetName().Version);
        }
        internal static void Sta(Action action)
        {
            Exception? failure=null; var t=new Thread(()=>{try{action();}catch(Exception e){failure=e;}}) { IsBackground=true }; t.SetApartmentState(ApartmentState.STA); t.Start(); Assert.IsTrue(t.Join(15000)); if(failure!=null) System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(failure).Throw();
        }
        [DataTestMethod]
        [DataRow("Word","zh-CN")] [DataRow("Excel","en")] [DataRow("PowerPoint","en")]
        public void Pane_EmptyLoadingSuccessAndFailure_MatchExistingVisualsAndGateWriteButtons(string host,string language)
        { Sta(()=>{
            var type=Host(host).GetType("OfficeTranslate.WordAddIn.DocumentChatPane",true)!;
            using(var history = new HistorySandbox())
            using(var pane=(UserControl)Activator.CreateInstance(type)!) using(var form=new Form()) {
                type.GetProperty("HistoryStore", BindingFlags.Instance | BindingFlags.NonPublic)!.SetValue(pane, history.Store);
                var settings=new TranslationSettings {UiLanguage=language}; var writeCount=0; var target=new ChatContext("document",true,(a,s)=>writeCount++,true);
                form.ClientSize=new System.Drawing.Size(440,760);
                type.GetMethod("Initialize",BindingFlags.Instance|BindingFlags.NonPublic)!.Invoke(pane,new object[]{(Func<ChatContext>)(()=>target),(Func<TranslationSettings>)(()=>settings)});
                form.Controls.Add(pane); pane.Dock=DockStyle.Fill; form.Show(); Application.DoEvents();
                Assert.AreEqual("Microsoft YaHei UI",pane.Font.Name); Assert.AreEqual(System.Drawing.Color.FromArgb(240,244,250),pane.BackColor);
                var transcript=(Control)Field(type,pane,"_transcript"); Assert.IsTrue(transcript.Text.Length>20); var writes=(FlowLayoutPanel)Field(type,pane,"_writes"); Assert.AreEqual(0,writes.Controls.Count);
                var completion=new TaskCompletionSource<ChatReply>();
                SynchronizationContext.SetSynchronizationContext(new SynchronizationContext());
                ChatTransport transport=(m,s,n,o,t)=>Task.Run(async()=>{o("streamed");return await completion.Task;}); type.GetProperty("Transport",BindingFlags.NonPublic|BindingFlags.Instance)!.SetValue(pane,transport);
                type.GetMethod("Start",BindingFlags.NonPublic|BindingFlags.Instance)!.Invoke(pane,new object[]{ChatMode.Rewrite}); Pump(()=>transcript.Text.Contains("streamed"));
                Assert.AreEqual(0,writes.Controls.Count); Assert.IsTrue(((Button)Field(type,pane,"_cancel")).Enabled);
                Task.Run(()=>completion.SetResult(new ChatReply("finished",false))); Pump(()=>writes.Controls.Count==2 && Field(type,pane,"_pending")==null); Assert.IsTrue(transcript.Text.Contains("finished")); Assert.AreEqual(0,writeCount);
                foreach(Button button in writes.Controls) { Assert.AreEqual(FlatStyle.Flat,button.FlatStyle); Assert.AreEqual(0,button.FlatAppearance.BorderSize); Assert.AreEqual(System.Drawing.Color.White,button.ForeColor); }
                var evidence=Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory,"../../../../../TestResults/doc-chat-2.3.0")); Directory.CreateDirectory(evidence);
                using(var bitmap=new System.Drawing.Bitmap(pane.Width,pane.Height)) { pane.DrawToBitmap(bitmap,pane.ClientRectangle); bitmap.Save(Path.Combine(evidence,"chat-"+host+"-"+language+".png")); }
                type.GetProperty("ConfirmWrite",BindingFlags.Instance|BindingFlags.NonPublic)!.SetValue(pane,(Func<ChatWriteAction,bool>)(_=>false));
                ((Button)writes.Controls[0]).PerformClick(); Assert.AreEqual(0,writeCount);
                type.GetProperty("ConfirmWrite",BindingFlags.Instance|BindingFlags.NonPublic)!.SetValue(pane,(Func<ChatWriteAction,bool>)(_=>true));
                ((Button)writes.Controls[0]).PerformClick(); Assert.AreEqual(1,writeCount); Assert.IsFalse(writes.Enabled);
                ((Button)writes.Controls[0]).PerformClick(); Assert.AreEqual(1,writeCount);
                type.GetMethod("Clear",BindingFlags.NonPublic|BindingFlags.Instance)!.Invoke(pane,null); Assert.AreEqual(0,writes.Controls.Count);
                settings.ApiKey="SECRET KEY MUST NOT DISPLAY";
                ChatTransport failed=(m,s,n,o,t)=>throw new IOException("Read failed: "+target.Text+" "+settings.ApiKey);type.GetProperty("Transport",BindingFlags.NonPublic|BindingFlags.Instance)!.SetValue(pane,failed);
                type.GetMethod("Start",BindingFlags.NonPublic|BindingFlags.Instance)!.Invoke(pane,new object[]{ChatMode.Summarize}); Pump(()=>Field(type,pane,"_pending")==null);
                Assert.AreEqual(0,writes.Controls.Count); Assert.IsFalse(((Label)Field(type,pane,"_status")).Text.Contains("SECRET"));
                Assert.IsFalse(transcript.Text.Contains("SECRET")); StringAssert.Contains(transcript.Text,"IOException: Read failed: [redacted] [redacted]");
                form.Close();
            }
        }); }
        [DataTestMethod]
        [DataRow("Word", "zh-CN", ChatFailure.Authentication, "ChatAuthError")]
        [DataRow("Excel", "en", ChatFailure.Network, "ChatNetworkError")]
        [DataRow("PowerPoint", "en", ChatFailure.Model, "ChatModelError")]
        [DataRow("Word", "en", ChatFailure.Budget, "ChatBudgetError")]
        public void Pane_ClassifiesErrorsAndShowsCopyableSafeDetails(string host,string language,ChatFailure failure,string key)
        { Sta(()=>{
            var type=Host(host).GetType("OfficeTranslate.WordAddIn.DocumentChatPane",true)!;
            using(var history = new HistorySandbox())
            using(var pane=(UserControl)Activator.CreateInstance(type)!) using(var form=new Form()) {
                type.GetProperty("HistoryStore", BindingFlags.Instance | BindingFlags.NonPublic)!.SetValue(pane, history.Store);
                var settings=new TranslationSettings {UiLanguage=language};
                type.GetMethod("Initialize",BindingFlags.Instance|BindingFlags.NonPublic)!.Invoke(pane,new object[]{(Func<ChatContext>)(()=>new ChatContext("document",false)),(Func<TranslationSettings>)(()=>settings)});
                form.Controls.Add(pane); form.Show(); Application.DoEvents();
                ChatTransport failed=(m,s,n,o,t)=>throw new ChatException(failure,"HTTP 401: fixture diagnosis.");
                type.GetProperty("Transport",BindingFlags.NonPublic|BindingFlags.Instance)!.SetValue(pane,failed);
                type.GetMethod("Start",BindingFlags.NonPublic|BindingFlags.Instance)!.Invoke(pane,new object[]{ChatMode.Summarize});
                Pump(()=>Field(type,pane,"_pending")==null);
                Assert.AreEqual(OfficeTranslate.WordAddIn.UiText.Get(language,key),((Label)Field(type,pane,"_status")).Text);
                StringAssert.Contains(((Control)Field(type,pane,"_transcript")).Text,"ChatException: HTTP 401: fixture diagnosis.");
                Assert.AreEqual(0,((FlowLayoutPanel)Field(type,pane,"_writes")).Controls.Count);
                Assert.AreEqual(0,((DocumentChatSession)Field(type,pane,"_session")).History.Count);
                form.Close();
            }
        }); }
        private static object Field(Type type,object target,string name)=>type.GetField(name,BindingFlags.NonPublic|BindingFlags.Instance)!.GetValue(target)!;
        private static void Pump(Func<bool> done) { var end=DateTime.UtcNow.AddSeconds(5); while(!done()&&DateTime.UtcNow<end){Application.DoEvents();Thread.Sleep(10);}Assert.IsTrue(done()); }
    }
    public class ChatItems
    {
        protected readonly List<object> Items;
        public ChatItems(params object[] items) { Items=items.ToList(); }
        public int Count=>Items.Count;
        public object this[int index]=>Items[index-1];
    }
    public class ChatCells { private readonly object _cell;public ChatCells(object cell){_cell=cell;}public object this[int r,int c]=>_cell; }
    public class ChatSlides : ChatItems { private readonly object _added;public readonly List<int> Added=new List<int>();public ChatSlides(object added){_added=added;}public object Add(int index,int layout){Added.Add(index);return _added;} }
    public class ChatShapes : ChatItems
    {
        private readonly object _text;public ChatShapes(object text){_text=text;}
        public object AddTextbox(int direction,float left,float top,float width,float height)=>DocumentChatHostTests.Obj(("TextFrame",DocumentChatHostTests.Obj(("TextRange",_text),("WordWrap",0),("AutoSize",0))),("TextFrame2",DocumentChatHostTests.Obj(("AutoSize",0))));
    }
}
