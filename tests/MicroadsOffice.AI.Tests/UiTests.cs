using System.Reflection;
using System.Runtime.InteropServices;
using MicroadsOffice.AI.Providers;
using MicroadsOffice.AI.UI;
using MicroadsOffice.Core;
using MicroadsOffice.Core.Office;
using MicroadsOffice.Excel.Model;
using MicroadsOffice.Excel.Ui;
using MicroadsOffice.UI;
using MicroadsOffice.Word.Model;
using MicroadsOffice.Word.UI;

namespace MicroadsOffice.AI.Tests;

/// <summary>An AI that answers from a script (streams in pieces, honours Stop).</summary>
public sealed class ScriptedProvider : AiProvider
{
    public readonly Queue<string> Replies = new();
    public readonly List<AiRequest> Requests = [];
    public Exception? Fail;
    public int DelayMs = 15;
    public bool Hang;
    public override AiProviderKind Kind => AiProviderKind.Compatible;

    public override async Task<string> CompleteAsync(AiRequest request, Action<string> onDelta, CancellationToken ct)
    {
        lock (Requests) Requests.Add(request);
        if (Fail != null) { var f = Fail; Fail = null; throw f; }
        if (Hang) await Task.Delay(Timeout.Infinite, ct);
        var text = Replies.Count > 0 ? Replies.Dequeue() : "OK";
        for (int i = 0; i < text.Length; i += 12) { ct.ThrowIfCancellationRequested(); onDelta(text.Substring(i, Math.Min(12, text.Length - i))); await Task.Delay(DelayMs, ct); }
        return text;
    }

    public override Task<IReadOnlyList<AiModel>> ListModelsAsync(CancellationToken ct) => Task.FromResult<IReadOnlyList<AiModel>>([new AiModel("fake-model", "Fake")]);
}

public static class Ui
{
    delegate bool EnumWindowsProc(IntPtr h, IntPtr l);
    [DllImport("user32.dll")] static extern bool EnumThreadWindows(uint threadId, EnumWindowsProc cb, IntPtr l);
    [DllImport("kernel32.dll")] static extern uint GetCurrentThreadId();
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern int GetClassName(IntPtr h, System.Text.StringBuilder s, int n);
    [DllImport("user32.dll")] static extern bool IsWindowVisible(IntPtr h);
    [DllImport("user32.dll")] static extern bool PostMessage(IntPtr h, int msg, IntPtr w, IntPtr l);

    public static readonly Dictionary<string, Action<Form>> DialogHandlers = new();
    public static readonly List<string> SeenDialogs = new();
    public static int NativeDialogs;

    /// <summary>Closes modal dialogs (Cancel) unless a handler is registered; counts native message boxes.</summary>
    public static System.Windows.Forms.Timer StartCloser()
    {
        var t = new System.Windows.Forms.Timer { Interval = 150 };
        t.Tick += (_, _) =>
        {
            foreach (var f in Application.OpenForms.Cast<Form>().ToList())
            {
                if (!f.Modal || !f.Visible) continue;
                SeenDialogs.Add(f.Name);
                if (DialogHandlers.Remove(f.Name, out var h)) h(f);
                else { f.DialogResult = DialogResult.Cancel; f.Close(); }
            }
            EnumThreadWindows(GetCurrentThreadId(), (h, _) =>
            {
                var sb = new System.Text.StringBuilder(64);
                GetClassName(h, sb, 64);
                if (sb.ToString() == "#32770" && IsWindowVisible(h))
                {
                    PostMessage(h, 0x0111, (IntPtr)2, IntPtr.Zero);
                    PostMessage(h, 0x0010, IntPtr.Zero, IntPtr.Zero);
                    NativeDialogs++;
                }
                return true;
            }, IntPtr.Zero);
        };
        t.Start();
        return t;
    }

    static readonly MethodInfo OnClick = typeof(Control).GetMethod("OnClick", BindingFlags.NonPublic | BindingFlags.Instance)!;

    public static void Click(Control c)
    {
        T.True(c.Enabled, "button enabled: " + c.Name);
        OnClick.Invoke(c, [EventArgs.Empty]);
        Pump();
    }

    public static void Pump(int ms = 30)
    {
        var until = DateTime.Now.AddMilliseconds(ms);
        do { Application.DoEvents(); Thread.Sleep(3); } while (DateTime.Now < until);
    }

    public static void Wait(Func<bool> cond, string what, int timeoutMs = 15000) => Wait(cond, () => what, timeoutMs);

    /// <summary>Waits; on timeout the message is computed then (e.g. the status label's text at that moment).</summary>
    public static void Wait(Func<bool> cond, Func<string> what, int timeoutMs = 15000)
    {
        var until = DateTime.Now.AddMilliseconds(timeoutMs);
        while (!cond()) { if (DateTime.Now > until) throw new AssertException("Timed out waiting for " + what()); Pump(20); }
    }

    public static void WaitRequest(AiPane pane)
    {
        Pump(50);
        Wait(() => pane.CurrentRequest == null || pane.CurrentRequest.IsCompleted, "assistant answer");
        Wait(() => !pane.Busy, "pane idle");
        Pump(60);
    }

    public static T2 Find<T2>(Control root, string name) where T2 : Control =>
        root.Controls.Find(name, true).OfType<T2>().FirstOrDefault() ?? throw new AssertException("control not found: " + name);

    [DllImport("user32.dll")] static extern bool PrintWindow(IntPtr hwnd, IntPtr hdc, uint flags);

    /// <summary>Captures our own window exactly as drawn (PrintWindow, full content) — never the screen, so no other app is captured.</summary>
    public static void Shot(Form f, string name)
    {
        Pump(250);
        using var bmp = new Bitmap(f.Width, f.Height);
        using (var g = Graphics.FromImage(bmp))
        {
            var hdc = g.GetHdc();
            PrintWindow(f.Handle, hdc, 2);
            g.ReleaseHdc(hdc);
        }
        bmp.Save(Path.Combine(T.Out, name + ".png"));
    }
}

public static class WordUiTests
{
    public static void Run(Runner r)
    {
        var closer = Ui.StartCloser();
        OfficeShell.AiSettingsHandler = owner => { using var d = new AiSettingsDialog(); d.ShowDialog(owner); };
        WordMainForm? w = null;
        var ai = new ScriptedProvider();
        try
        {
            AiService.Override = null;
            AiSettings.Update(x => x.Provider = AiProviderKind.None);

            r.Test("W01 Home › Trợ lý AI opens the pane; not connected → setup card with the three ways to connect", () =>
            {
                w = (WordMainForm)Word.WordApp.CreateMainForm([]);
                w.Show(); Ui.Pump(300); w.HideBackstage(); Ui.Pump(100);
                var btn = (RibbonButton)w.Commands["btnAiAssistant"];
                T.True(btn.ToolTipText.Contains("ChatGPT"), "tooltip mentions ChatGPT");
                Ui.Click(btn);
                T.True(w.AiPane is { Visible: true }, "pane visible");
                T.True(btn.Checked, "ribbon button checked");
                foreach (var n in new[] { "btnAiSetupChatGpt", "btnAiSetupClaude", "btnAiSetupOther", "btnAiSend", "btnAiStop", "btnAiNewChat", "btnAiSettings", "btnAiClose", "txtAiInput", "chkAiSendDoc", "chkAiAutoApply" })
                    T.True(w.AiPane!.Controls.Find(n, true).Length == 1, n);
                T.Eq(8, Ui.Find<FlowLayoutPanel>(w.AiPane!, "flowAiQuick").Controls.Count);
                T.True(Ui.Find<LinkLabel>(w.AiPane!, "lnkAiStatus").Text.Contains("Chưa kết nối"), "status");
                Ui.Shot(w, "word-ai-setup");
            });

            r.Test("W02 the setup buttons open AI settings at the right place; asking while not connected explains what to do", () =>
            {
                Ui.SeenDialogs.Clear();
                foreach (var n in new[] { "btnAiSetupChatGpt", "btnAiSetupClaude", "btnAiSetupOther", "btnAiSettings" })
                {
                    Ui.Click(Ui.Find<Button>(w!.AiPane!, n));
                    Ui.Pump(200);
                }
                T.Eq(4, Ui.SeenDialogs.Count(d => d == "aiSettingsDialog"));
                var box = Ui.Find<TextBox>(w!.AiPane!, "txtAiInput");
                box.Text = "Xin chào";
                Ui.Click(Ui.Find<Button>(w.AiPane!, "btnAiSend"));
                T.True(w.AiPane!.Controls.Find("flowAiChat", true)[0].Controls.OfType<AiMessageView>().Any(v => v.Text2.Contains("chưa được kết nối")), "explains");
            });

            AiService.Override = _ => ai;
            AiSettings.Update(x => { x.Provider = AiProviderKind.Compatible; x.CompatibleModel = "fake-model"; });

            r.Test("W03 asking for content writes it into the document with Heading styles as ONE undo step; steps are listed", () =>
            {
                w!.AiPane!.RefreshStatus();
                w.Editor.Document = Document.CreateBlank();
                w.Editor.InsertText("Đoạn đầu tiên.");
                var before = w.Editor.Document.PlainText;
                ai.Replies.Enqueue("Tôi đã lập dàn ý.\n```microads-document mode=insert\n# Mở đầu\nGiới thiệu chung.\n## Nội dung chính\n- Ý thứ nhất\n- Ý thứ hai\n# Kết luận\nTóm lại.\n```");
                w.AiPane.Ask("Lập dàn ý cho báo cáo");
                Ui.WaitRequest(w.AiPane);
                var d = w.Editor.Document;
                var heads = d.Paragraphs.Where(p => p.Props.OutlineLevel == 1).Select(p => p.Text).ToList();
                T.True(heads.Contains("Mở đầu") && heads.Contains("Kết luận"), "Heading 1 paragraphs: " + string.Join("|", heads));
                T.True(d.Paragraphs.Any(p => p.Text == "Nội dung chính" && p.Props.Style == StyleIds.Heading2), "Heading 2 style");
                T.Eq(2, d.Paragraphs.Count(p => p.Props.List == ListType.Bullet));
                T.Eq("Đoạn đầu tiên.", d.Paragraphs[0].Text);
                var back = Word.AI.WordMarkdown.Export(d, 10000);
                T.True(back.Contains("# Mở đầu") && !back.Contains("**Mở đầu**") && back.Contains("- Ý thứ nhất"), "round trip without style noise: " + back);
                var ans = w.AiPane.LastAnswer!;
                T.True(ans.Footnote.StartsWith("✓"), "footnote " + ans.Footnote);
                T.Eq("Tôi đã lập dàn ý.", ans.Text2.Trim());
                T.True(ans.Controls.Find("flowAiSteps" + ans.Seq, true)[0].Controls.Count == 4, "3 steps + title");
                T.True(ans.FindAction("btnAiUndo") != null, "undo offered");
                T.True(ai.Requests.Last().Messages.Last().Text.Contains("<document"), "document sent as context");
                Ui.Shot(w, "word-ai-written");
                Ui.Click(ans.FindAction("btnAiUndo")!);
                T.Eq(before, w.Editor.Document.PlainText);
                T.True(ans.FindAction("btnAiApply") != null, "can write again");
                Ui.Click(ans.FindAction("btnAiApply")!);
                T.True(w.Editor.Document.Paragraphs.Any(p => p.Text == "Kết luận"), "written again");
                w.Editor.Undo();
                T.Eq(before, w.Editor.Document.PlainText);
            });

            r.Test("W04 'Viết lại' replaces exactly the selection, with formatting, keeping the rest of the paragraph", () =>
            {
                w!.Editor.Document = Document.CreateBlank();
                w.Editor.InsertText("Câu một. Câu cũ cần sửa. Câu ba.");
                var p0 = w.Editor.Document.Paragraphs[0].Text;
                int i = p0.IndexOf("Câu cũ cần sửa.");
                w.Editor.Select(new DocPos(0, i), new DocPos(0, i + "Câu cũ cần sửa.".Length));
                ai.Replies.Enqueue("Đã viết lại.\n```microads-document mode=replace_selection\nCâu **mới** đã sửa.\n```");
                Ui.Click(Ui.Find<Button>(w.AiPane!, "btnAiQ_Rewrite"));
                Ui.WaitRequest(w.AiPane!);
                var p = w.Editor.Document.Paragraphs[0];
                T.Eq("Câu một. Câu mới đã sửa. Câu ba.", p.Text);
                T.True(p.Runs.Any(x => x.Text == "mới" && x.Props.Bold), "bold run");
                T.True(ai.Requests.Last().Messages.Last().Text.Contains("<selection>\nCâu cũ cần sửa."), "selection sent");
                T.True(ai.Requests.Last().Instructions.Contains("replace_selection"), "action guide sent");
            });

            r.Test("W05 quick actions that need a selection explain it instead of sending", () =>
            {
                w!.Editor.Document = Document.CreateBlank();
                int n = ai.Requests.Count;
                Ui.Click(Ui.Find<Button>(w.AiPane!, "btnAiQ_ToEnglish"));
                T.Eq(n, ai.Requests.Count);
                T.True(w.AiPane!.Controls.Find("flowAiChat", true)[0].Controls.OfType<AiMessageView>().Last().Text2.Contains("Dịch sang tiếng Anh"), "explains");
            });

            r.Test("W06 a chat answer offers Insert / Replace selection / Copy, and each works", () =>
            {
                w!.Editor.Document = Document.CreateBlank();
                w.Editor.InsertText("Tiêu đề cũ");
                ai.Replies.Enqueue("Gợi ý:\n\n**Báo cáo quý III**");
                w.AiPane!.Ask("Gợi ý tiêu đề");
                Ui.WaitRequest(w.AiPane);
                var ans = w.AiPane.LastAnswer!;
                T.True(ans.FindAction("btnAiInsert") != null && ans.FindAction("btnAiCopy") != null, "insert + copy");
                T.True(ans.FindAction("btnAiReplace") == null, "no replace without a selection");
                IDataObject? saved = null;
                try { saved = Clipboard.GetDataObject(); } catch { }
                string copied;
                try
                {
                    Ui.Click(ans.FindAction("btnAiCopy")!);
                    copied = Clipboard.GetText();
                }
                finally
                {
                    // give the user's clipboard back
                    try { if (saved != null) Clipboard.SetDataObject(saved, true); else Clipboard.Clear(); } catch { }
                }
                T.True(copied.Contains("Báo cáo quý III"), "copied as plain text");
                Ui.Click(ans.FindAction("btnAiInsert")!);
                T.True(w.Editor.Document.Paragraphs.Any(p => p.Text == "Báo cáo quý III" && p.Runs.All(x => x.Props.Bold)), "inserted bold");
                Ui.Click(ans.FindAction("btnAiUndo")!);
                T.False(w.Editor.Document.PlainText.Contains("Báo cáo quý III"), "undone");
                w.Editor.SelectAll();
                ai.Replies.Enqueue("Báo cáo mới");
                w.AiPane.Ask("Đổi tiêu đề");
                Ui.WaitRequest(w.AiPane);
                var a2 = w.AiPane.LastAnswer!;
                T.True(a2.FindAction("btnAiReplace") != null, "replace offered with a selection");
                Ui.Click(a2.FindAction("btnAiReplace")!);
                T.Eq("Báo cáo mới", w.Editor.Document.PlainText.Trim());
            });

            r.Test("W07 administrative document: centred motto block and a table with a header row", () =>
            {
                w!.Editor.Document = Document.CreateBlank();
                ai.Replies.Enqueue("```microads-document mode=replace_document\n:::center\n**CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM**\n**Độc lập - Tự do - Hạnh phúc**\n:::\n\n# BIÊN BẢN\n\n| STT | Nội dung | Ghi chú |\n|---|---|---|\n| 1 | Họp | |\n```");
                w.AiPane!.Ask("Soạn biên bản");
                Ui.WaitRequest(w.AiPane);
                var d = w.Editor.Document;
                var motto = d.Paragraphs.First(p => p.Text.Contains("CỘNG HÒA"));
                T.Eq(ParaAlign.Center, motto.Props.Align);
                T.True(motto.Text.Contains('\v'), "two lines in one centred paragraph");
                var t = d.Blocks.OfType<Table>().Single();
                T.Eq(3, t.ColumnCount); T.Eq(2, t.Rows.Count);
                T.True(t.Rows[0].HeaderRow && t.Rows[0].Cells[0].Paragraphs[0].Runs.All(x => x.Props.Bold), "bold header row");
                T.True(t.ColumnWidths[1] > t.ColumnWidths[0], "content column wider than STT");
                Ui.Shot(w, "word-ai-admin");
            });

            r.Test("W08 errors show a message with Retry and account settings; Retry sends again", () =>
            {
                ai.Fail = new AiException("Phiên đăng nhập ChatGPT đã hết hạn.", "x", 401) { NeedsSignIn = true };
                w!.AiPane!.Ask("Chào");
                Ui.WaitRequest(w.AiPane);
                var ans = w.AiPane.LastAnswer!;
                T.True(ans.IsError && ans.Text2.Contains("hết hạn"), "error shown");
                T.True(ans.FindAction("btnAiFix") != null, "settings offered");
                Ui.SeenDialogs.Clear();
                Ui.Click(ans.FindAction("btnAiFix")!);
                Ui.Pump(250);
                T.True(Ui.SeenDialogs.Contains("aiSettingsDialog"), "settings opened");
                int n = ai.Requests.Count;
                ai.Replies.Enqueue("Đã thử lại.");
                Ui.Click(ans.FindAction("btnAiRetry")!);
                Ui.WaitRequest(w.AiPane);
                T.Eq(n + 1, ai.Requests.Count);
                T.Eq("Đã thử lại.", w.AiPane.LastAnswer!.Text2.Trim());
            });

            r.Test("W09 Stop cancels a running answer; Send is disabled while busy", () =>
            {
                ai.Hang = true;
                w!.AiPane!.Ask("Viết dài");
                Ui.Pump(150);
                T.False(Ui.Find<Button>(w.AiPane, "btnAiSend").Enabled, "send disabled while busy");
                T.False(Ui.Find<Button>(w.AiPane, "btnAiQ_Summarize").Enabled, "quick actions disabled while busy");
                Ui.Click(Ui.Find<Button>(w.AiPane, "btnAiStop"));
                Ui.WaitRequest(w.AiPane);
                ai.Hang = false;
                T.True(w.AiPane.LastAnswer!.Text2.Contains("Đã dừng"), "stopped");
                T.True(Ui.Find<Button>(w.AiPane, "btnAiSend").Enabled, "send enabled again");
            });

            r.Test("W10 Enter sends, every quick action sends a request, New conversation clears history, Close hides the pane", () =>
            {
                var box = Ui.Find<TextBox>(w!.AiPane!, "txtAiInput");
                box.Text = "Câu hỏi bằng Enter";
                var m = typeof(Control).GetMethod("OnKeyDown", BindingFlags.NonPublic | BindingFlags.Instance)!;
                m.Invoke(box, [new KeyEventArgs(Keys.Enter)]);
                Ui.WaitRequest(w.AiPane!);
                T.True(ai.Requests.Last().Messages.Last().Text.EndsWith("Câu hỏi bằng Enter"), "sent by Enter");
                w.Editor.Document = Document.CreateBlank(); w.Editor.InsertText("Một đoạn văn để thử."); w.Editor.SelectAll();
                foreach (var b in Ui.Find<FlowLayoutPanel>(w.AiPane!, "flowAiQuick").Controls.OfType<Button>().ToList())
                {
                    int n = ai.Requests.Count;
                    ai.Replies.Enqueue("Trả lời cho " + b.Name);
                    Ui.Click(b);
                    Ui.WaitRequest(w.AiPane!);
                    T.Eq(n + 1, ai.Requests.Count);
                    w.Editor.SelectAll();
                }
                T.True(w.AiPane!.History.Count > 4, "history kept");
                Ui.Click(Ui.Find<Button>(w.AiPane, "btnAiNewChat"));
                T.Eq(0, w.AiPane.History.Count);
                Ui.Click(Ui.Find<Button>(w.AiPane, "btnAiClose"));
                T.False(w.AiPane.Visible, "hidden");
                T.False(((RibbonButton)w.Commands["btnAiAssistant"]).Checked, "ribbon unchecked");
                Ui.Click(w.Commands["btnAiAssistant"]);
                T.True(w.AiPane.Visible, "shown again");
            });

            r.Test("W11 untick 'send document' → only the question is sent; untick auto-write → 'Ghi vào văn bản' button", () =>
            {
                var send = Ui.Find<CheckBox>(w!.AiPane!, "chkAiSendDoc");
                send.Checked = false; Ui.Pump();
                T.False(AiSettings.Load().SendDocument, "setting saved");
                ai.Replies.Enqueue("```microads-document\nNội dung\n```");
                var auto = Ui.Find<CheckBox>(w.AiPane!, "chkAiAutoApply");
                auto.Checked = false;
                w.Editor.Document = Document.CreateBlank();
                w.AiPane!.Ask("Viết");
                Ui.WaitRequest(w.AiPane);
                T.Eq("Viết", ai.Requests.Last().Messages.Last().Text);
                T.Eq("", w.Editor.Document.PlainText.Trim());
                Ui.Click(w.AiPane.LastAnswer!.FindAction("btnAiApply")!);
                T.Eq("Nội dung", w.Editor.Document.PlainText.Trim());
                send.Checked = true; auto.Checked = true; Ui.Pump();
            });

            r.Test("W12 AI settings dialog: own server → Test connection lists models → Save; keys are stored encrypted", () =>
            {
                using var srv = new FakeServer();
                using var d = new AiSettingsDialog(AiSettingsDialog.Section.Other);
                d.Show(); Ui.Pump(100);
                Ui.Find<RadioButton>(d, "rbAiCompat").Checked = true;
                Ui.Find<TextBox>(d, "txtAiCompatUrl").Text = srv.Base + "/compat/v1";
                Ui.Click(Ui.Find<Button>(d, "btnAiTest"));
                Ui.Wait(() => d.Pending?.IsCompleted == true, "test connection");
                Ui.Pump(50);
                var cbo = Ui.Find<ComboBox>(d, "cboAiModel");
                T.Eq(2, cbo.Items.Count);
                T.True(Ui.Find<Label>(d, "lblAiTestResult").Text.Contains("2"), "result text");
                cbo.Text = "llama-local";
                Ui.Find<TextBox>(d, "txtAiOpenAiKey").Text = "sk-saved-openai";
                Ui.Click(Ui.Find<Button>(d, "btnOk"));
                var s = AiSettings.Load();
                T.Eq(AiProviderKind.Compatible, s.Provider);
                T.Eq("llama-local", s.CompatibleModel);
                T.Eq(srv.Base + "/compat/v1", s.CompatibleBaseUrl);
                T.Eq("sk-saved-openai", SecretStore.Read(SecretNames.OpenAiKey));
                using var d2 = new AiSettingsDialog();
                d2.Show(); Ui.Pump(50);
                T.True(Ui.Find<TextBox>(d2, "txtAiOpenAiKey").PlaceholderText.Contains("Đã lưu"), "saved key indicated");
                var clear = Ui.Find<LinkLabel>(d2, "lnkAiOpenAiKeyClear");
                typeof(LinkLabel).GetMethod("OnLinkClicked", BindingFlags.NonPublic | BindingFlags.Instance)!.Invoke(clear, [new LinkLabelLinkClickedEventArgs(clear.Links[0])]);
                T.False(SecretStore.Has(SecretNames.OpenAiKey), "key removed");
                foreach (var n in new[] { "rbAiChatGpt", "btnAiChatGptSignIn", "btnAiChatGptSignOut", "lnkAiChatGptOther", "lnkAiChatGptUsage", "btnAiClaudeDesktop", "lnkAiClaudeDownload", "rbAiClaudeKey", "txtAiClaudeKey", "lnkAiClaudeKeyGet", "rbAiOpenAiKey", "rbAiCompat", "txtAiInstructions", "lnkAiPrivacy", "lblAiPrivacy" })
                    T.True(d2.Controls.Find(n, true).Length == 1, n);
                T.True(Ui.Find<Label>(d2, "lblAiPrivacy").Text.Contains("không qua máy chủ Microads"), "privacy explained");
                d2.Close();
            });

            r.Test("W13 Continue with ChatGPT in the dialog signs in through the browser, then Sign out signs out", () =>
            {
                AiService.Override = null;
                using var srv = new FakeServer();
                Environment.SetEnvironmentVariable("MICROADS_AI_AUTH_BASE", srv.Base);
                Environment.SetEnvironmentVariable("MICROADS_AI_OPENAI_BASE", srv.Base + "/v1");
                ChatGptAuth.Forget();
                ChatGptAuth.OpenBrowser = url =>
                {
                    var q = ChatGptAuth.ParseQuery(new Uri(url).Query.TrimStart('?'));
                    srv.Challenge = q["code_challenge"]; srv.Nonce = q["nonce"]; srv.LastRedirect = q["redirect_uri"];
                    _ = Task.Run(async () => { await Task.Delay(100); using var h = new HttpClient(); await h.GetAsync(q["redirect_uri"] + "?code=CODE-1&state=" + Uri.EscapeDataString(q["state"]) + "&client_id=" + srv.IssuedClientId + "&scope=" + Uri.EscapeDataString(srv.Scope)); });
                    return true;
                };
                using var d = new AiSettingsDialog(AiSettingsDialog.Section.ChatGpt);
                d.Show(); Ui.Pump(100);
                T.False(Ui.Find<Button>(d, "btnAiChatGptSignOut").Enabled, "sign out disabled before sign-in");
                Ui.Click(Ui.Find<Button>(d, "btnAiChatGptSignIn"));
                Ui.Wait(() => d.Pending?.IsCompleted == true, () => "sign-in; status: " + Ui.Find<Label>(d, "lblAiChatGptStatus").Text
                    + " | fake server saw: " + string.Join(", ", srv.Requests.Select(q => q.Method + " " + q.Path)));
                Ui.Wait(() => Ui.Find<Label>(d, "lblAiTestResult").Text.Contains("1"), "models loaded after sign-in", 10000);
                T.True(Ui.Find<Label>(d, "lblAiChatGptStatus").Text.Contains("nhanvien@example.com"), Ui.Find<Label>(d, "lblAiChatGptStatus").Text);
                T.Eq(AiProviderKind.ChatGpt, AiSettings.Load().Provider);
                T.Eq("gpt-test", Ui.Find<ComboBox>(d, "cboAiModel").Text);
                Ui.Click(Ui.Find<Button>(d, "btnOk"));
                T.Eq("gpt-test", AiSettings.Load().ChatGptModel);
                // the pane now says "Using ChatGPT plan · Manage usage" and answers through the Responses API
                w!.AiPane!.RefreshStatus();
                T.True(Ui.Find<LinkLabel>(w.AiPane!, "lnkAiPlanUsage").Visible, "plan usage line");
                srv.Reply = "```microads-document\n# Từ ChatGPT\n```";
                w.Editor.Document = Document.CreateBlank();
                w.AiPane!.Ask("Viết tiêu đề");
                Ui.WaitRequest(w.AiPane);
                T.True(w.Editor.Document.Paragraphs.Any(p => p.Text == "Từ ChatGPT" && p.Props.OutlineLevel == 1), "answer from ChatGPT written: " + w.AiPane.LastAnswer?.Text2);
                using var d2 = new AiSettingsDialog();
                d2.Show(); Ui.Pump(50);
                T.True(Ui.Find<Button>(d2, "btnAiChatGptSignOut").Enabled, "sign out enabled");
                Ui.Click(Ui.Find<Button>(d2, "btnAiChatGptSignOut"));
                Ui.Wait(() => d2.Pending?.IsCompleted == true, "sign-out");
                T.False(ChatGptAuth.IsSignedIn, "signed out");
                T.Eq(1, srv.Revoked.Count);
                T.True(Ui.Find<Label>(d2, "lblAiChatGptStatus").Text.Contains("Chưa đăng nhập"), "status updated");
                d2.Close();
                ChatGptAuth.OpenBrowser = _ => false;
            });

            r.Test("W14 Claude Desktop button asks first (Cancel keeps it disconnected); Disconnect works without asking", () =>
            {
                using var d = new AiSettingsDialog(AiSettingsDialog.Section.Claude);
                d.Show(); Ui.Pump(50);
                var b = Ui.Find<Button>(d, "btnAiClaudeDesktop");
                T.True(b.Text.Contains("Kết nối"), b.Text);
                int native = Ui.NativeDialogs;
                Ui.Click(b); Ui.Pump(400);
                T.True(Ui.NativeDialogs > native, "confirmation shown");
                T.False(Mcp.ClaudeDesktop.IsConnected(), "cancel = not connected");
                Mcp.ClaudeDesktop.Connect();
                AiSettings.Update(x => x.AllowClaudeDesktop = true);
                d.Close();
                using var d2 = new AiSettingsDialog(AiSettingsDialog.Section.Claude);
                d2.Show(); Ui.Pump(50);
                var b2 = Ui.Find<Button>(d2, "btnAiClaudeDesktop");
                T.True(b2.Text.Contains("Ngắt"), b2.Text);
                Ui.Click(b2);
                T.False(Mcp.ClaudeDesktop.IsConnected(), "disconnected");
                T.False(AiSettings.Load().AllowClaudeDesktop, "access withdrawn");
                T.True(b2.Text.Contains("Kết nối"), b2.Text);
                d2.Close();
            });

            r.Test("W16 'Báo cáo' sends the question, the answer and the reason to Microads (Store policy 11.16)", () =>
            {
                AiService.Override = _ => ai;
                AiSettings.Update(x => { x.Provider = AiProviderKind.Compatible; x.CompatibleModel = "fake-model"; });
                using var srv = new FakeServer();
                Environment.SetEnvironmentVariable("MICROADS_AI_REPORT_URL", srv.Base + "/report.php");
                ai.Replies.Enqueue("Một câu trả lời cần báo cáo.");
                w!.AiPane!.Ask("Câu hỏi thử");
                Ui.WaitRequest(w.AiPane);
                var ans = w.AiPane.LastAnswer!;
                T.True(ans.FindAction("btnAiReport") != null, "report button on every answer");
                bool sentOk = false;
                Ui.DialogHandlers["aiReportDialog"] = f =>
                {
                    var d = (AiReportDialog)f;
                    Ui.Find<RadioButton>(d, "rbAiReport_wrong").Checked = true;
                    Ui.Find<TextBox>(d, "txtAiReportComment").Text = "Sai số liệu";
                    Ui.Click(Ui.Find<Button>(d, "btnAiReportSend"));
                    Ui.Wait(() => d.Pending?.IsCompleted == true, "report sent");
                    sentOk = d.Sent;
                    d.Close();
                };
                Ui.Click(ans.FindAction("btnAiReport")!);
                Ui.Wait(() => !Ui.DialogHandlers.ContainsKey("aiReportDialog"), "report dialog handled");
                Ui.Pump(200);
                T.True(sentOk, "server accepted");
                var j = srv.Last("/report.php")!.Json!;
                T.Eq("wrong", j["reason"]!.GetValue<string>());
                T.Eq("Sai số liệu", j["comment"]!.GetValue<string>());
                T.Eq("word", j["app"]!.GetValue<string>());
                T.True(j["question"]!.GetValue<string>().Contains("Câu hỏi thử") && j["answer"]!.GetValue<string>().Contains("cần báo cáo"), "question + answer");
                T.False(j.ToJsonString().Contains("sk-"), "no key sent");
                T.True(ans.Footnote.Contains("báo cáo"), ans.Footnote);
                Ui.DialogHandlers["aiReportDialog"] = f =>
                {
                    var d = (AiReportDialog)f;
                    Ui.Find<CheckBox>(d, "chkAiReportInclude").Checked = false;
                    Ui.Click(Ui.Find<Button>(d, "btnAiReportSend"));
                    T.True(d.Pending == null && Ui.Find<Label>(d, "lblAiReportResult").Text.Length > 0, "needs a comment");
                    Ui.Find<TextBox>(d, "txtAiReportComment").Text = "Chỉ ghi chú";
                    Ui.Click(Ui.Find<Button>(d, "btnAiReportSend"));
                    Ui.Wait(() => d.Pending?.IsCompleted == true, "sent");
                    d.Close();
                };
                Ui.Click(ans.FindAction("btnAiReport")!);
                Ui.Wait(() => !Ui.DialogHandlers.ContainsKey("aiReportDialog"), "second report");
                Ui.Pump(200);
                var j2 = srv.Last("/report.php")!.Json!;
                T.Eq("", j2["question"]!.GetValue<string>()); T.Eq("", j2["answer"]!.GetValue<string>());
                T.Eq("Chỉ ghi chú", j2["comment"]!.GetValue<string>());
                Environment.SetEnvironmentVariable("MICROADS_AI_REPORT_URL", null);
            });

            r.Test("W17 a whole-document rewrite is refused when the document has a picture (it would be lost)", () =>
            {
                AiService.Override = _ => ai;
                AiSettings.Update(x => { x.Provider = AiProviderKind.Compatible; x.CompatibleModel = "fake-model"; });
                w!.Editor.Document = Document.CreateBlank();
                w.Editor.InsertText("Đoạn có hình.");
                using (var bmp = new Bitmap(8, 8)) using (var ms = new MemoryStream()) { bmp.Save(ms, System.Drawing.Imaging.ImageFormat.Png); w.Editor.InsertImage(ms.ToArray(), ".png"); }
                var before = w.Editor.Document.PlainText;
                ai.Replies.Enqueue("```microads-document mode=replace_document\n# Viết lại\n```");
                w.AiPane!.Ask("Viết lại toàn bộ");
                Ui.WaitRequest(w.AiPane);
                T.Eq(before, w.Editor.Document.PlainText);
                T.True(w.AiPane.LastAnswer!.Footnote.Contains("hình ảnh"), w.AiPane.LastAnswer.Footnote);
            });

            r.Test("W18 the pane's Undo never undoes the user's own later typing", () =>
            {
                w!.Editor.Document = Document.CreateBlank();
                ai.Replies.Enqueue("```microads-document\n# Mục mới\n```");
                w.AiPane!.Ask("Thêm mục");
                Ui.WaitRequest(w.AiPane);
                var ans = w.AiPane.LastAnswer!;
                w.Editor.InsertText("Người dùng gõ thêm");
                Ui.Click(ans.FindAction("btnAiUndo")!);
                T.True(w.Editor.Document.PlainText.Contains("Người dùng gõ thêm") && w.Editor.Document.PlainText.Contains("Mục mới"), "nothing undone");
                T.True(ans.Footnote.Contains("Ctrl+Z"), ans.Footnote);
            });

            r.Test("W19 an answer cut off inside its content block writes nothing and offers Try again", () =>
            {
                w!.Editor.Document = Document.CreateBlank();
                w.Editor.InsertText("Nội dung gốc");
                ai.Replies.Enqueue("Đang soạn\n```microads-document mode=replace_document\n# Bị cắt");
                w.AiPane!.Ask("Viết lại");
                Ui.WaitRequest(w.AiPane);
                T.Eq("Nội dung gốc", w.Editor.Document.PlainText.Trim());
                var ans = w.AiPane.LastAnswer!;
                T.True(ans.Footnote.Contains("cắt ngang") && ans.FindAction("btnAiRetry") != null, ans.Footnote);
            });

            r.Test("W23 an answer with a table in the MIDDLE: headings and text after the table stay after it (Store screenshot bug)", () =>
            {
                AiService.Override = _ => ai;
                AiSettings.Update(x => { x.Provider = AiProviderKind.Compatible; x.CompatibleModel = "fake-model"; });
                w!.Editor.Document = Document.CreateBlank();
                ai.Replies.Enqueue("```microads-document\n## Điều 1. Công việc\n| Nội dung | Thỏa thuận |\n|---|---|\n| Chức danh | Kế toán |\n## Điều 2. Chế độ làm việc\nThời giờ làm việc 8 giờ/ngày.\n```");
                w.AiPane!.Ask("Soạn điều khoản");
                Ui.WaitRequest(w.AiPane);
                var blocks = w.Editor.Document.Blocks;
                int ti = blocks.FindIndex(b => b is Table);
                T.True(ti > 0, "table written");
                var table = (Table)blocks[ti];
                T.Eq("Nội dung", table.Rows[0].Cells[0].Paragraphs.Single().Text);
                T.True(blocks[ti + 1] is Paragraph { Text: "Điều 2. Chế độ làm việc" } h && h.Props.OutlineLevel == 2, "heading after the table");
                T.True(blocks[ti + 2] is Paragraph { Text: "Thời giờ làm việc 8 giờ/ngày." }, "text after the heading");
                w.Editor.Undo();
                T.False(w.Editor.Document.Blocks.Any(b => b is Table), "one undo step removes it all");
            });

            r.Test("W20 'Hủy đăng nhập' stops a ChatGPT sign-in that waits in the browser; the dialog stays usable", () =>
            {
                AiService.Override = null;
                ChatGptAuth.Forget();
                string? opened = null;
                ChatGptAuth.OpenBrowser = url => { opened = url; return true; }; // the user never finishes in the browser
                using var d = new AiSettingsDialog(AiSettingsDialog.Section.ChatGpt);
                d.Show(); Ui.Pump(100);
                var cancel = Ui.Find<Button>(d, "btnAiChatGptCancel");
                T.False(cancel.Visible, "cancel hidden before sign-in");
                Ui.Click(Ui.Find<Button>(d, "btnAiChatGptSignIn"));
                Ui.Wait(() => opened != null, "browser opened");
                T.True(cancel.Visible && cancel.Enabled, "cancel offered while waiting");
                T.False(Ui.Find<Button>(d, "btnAiChatGptSignIn").Enabled, "sign-in disabled while waiting");
                var pendingBefore = d.Pending;
                string trace = ""; int uiThread = Environment.CurrentManagedThreadId;
                Ui.Find<Label>(d, "lblAiChatGptStatus").TextChanged += (_, _) =>
                {
                    if (Ui.Find<Label>(d, "lblAiChatGptStatus").Text.Contains("Đã hủy"))
                        trace = $"thread {Environment.CurrentManagedThreadId} (ui {uiThread}) " + string.Join(" < ", Environment.StackTrace.Split('\n').Select(l => l.Trim()).Where(l => l.StartsWith("at ")).Skip(1).Take(14).Select(l => l.Split(" in ")[0][3..]));
                };
                Ui.Click(cancel);
                Ui.Wait(() => d.Pending?.IsCompleted == true, () => $"sign-in cancelled; pending {d.Pending?.Status} (same task: {ReferenceEquals(pendingBefore, d.Pending)}), status '{Ui.Find<Label>(d, "lblAiChatGptStatus").Text}', cancel visible {cancel.Visible}, signIn enabled {Ui.Find<Button>(d, "btnAiChatGptSignIn").Enabled}, other enabled {Ui.Find<LinkLabel>(d, "lnkAiChatGptOther").Enabled} | TRACE {trace}", 5000);
                T.True(Ui.Find<Label>(d, "lblAiChatGptStatus").Text.Contains("Đã hủy đăng nhập"), Ui.Find<Label>(d, "lblAiChatGptStatus").Text);
                T.False(cancel.Visible, "cancel hidden again");
                T.True(Ui.Find<Button>(d, "btnAiChatGptSignIn").Enabled, "can sign in again");
                T.False(ChatGptAuth.IsSignedIn, "nothing saved");
                // the loopback listener was closed (nothing keeps waiting for a callback in the background)
                int port = new Uri(ChatGptAuth.ParseQuery(new Uri(opened!).Query.TrimStart('?'))["redirect_uri"]).Port;
                using (var probe = new System.Net.Sockets.TcpClient())
                    T.Throws<System.Net.Sockets.SocketException>(() => probe.Connect(System.Net.IPAddress.Loopback, port), "callback port closed after cancel");
                opened = null;
                Ui.Click(Ui.Find<Button>(d, "btnAiChatGptSignIn"));
                Ui.Wait(() => opened != null, "second sign-in started");
                int port2 = new Uri(ChatGptAuth.ParseQuery(new Uri(opened!).Query.TrimStart('?'))["redirect_uri"]).Port;
                d.Close(); Ui.Pump(300);
                using (var probe2 = new System.Net.Sockets.TcpClient())
                    T.Throws<System.Net.Sockets.SocketException>(() => probe2.Connect(System.Net.IPAddress.Loopback, port2), "closing the dialog ends the sign-in (port closed)");
                ChatGptAuth.OpenBrowser = _ => false;
            });

            r.Test("W21 Save refuses a Claude / OpenAI API-key choice without a key and says what to do", () =>
            {
                SecretStore.Write(SecretNames.ClaudeKey, null); SecretStore.Write(SecretNames.OpenAiKey, null);
                AiSettings.Update(x => x.Provider = AiProviderKind.None);
                using var d = new AiSettingsDialog(AiSettingsDialog.Section.Claude);
                d.Show(); Ui.Pump(50);
                Ui.Find<RadioButton>(d, "rbAiClaudeKey").Checked = true;
                int native = Ui.NativeDialogs;
                Ui.Click(Ui.Find<Button>(d, "btnOk")); Ui.Pump(300);
                T.True(Ui.NativeDialogs > native, "warning shown");
                T.False(d.IsDisposed || !d.Visible, "dialog stays open");
                T.Eq(AiProviderKind.None, AiSettings.Load().Provider);
                Ui.Find<RadioButton>(d, "rbAiOpenAiKey").Checked = true;
                native = Ui.NativeDialogs;
                Ui.Click(Ui.Find<Button>(d, "btnOk")); Ui.Pump(300);
                T.True(Ui.NativeDialogs > native, "warning shown for OpenAI too");
                // with a key it saves, and the key is stored encrypted, not in office.json
                Ui.Find<RadioButton>(d, "rbAiClaudeKey").Checked = true;
                Ui.Find<TextBox>(d, "txtAiClaudeKey").Text = "sk-ant-test";
                Ui.Click(Ui.Find<Button>(d, "btnOk"));
                T.Eq(AiProviderKind.ClaudeKey, AiSettings.Load().Provider);
                T.Eq("sk-ant-test", SecretStore.Read(SecretNames.ClaudeKey));
                T.False(File.ReadAllText(Path.Combine(MicroadsOffice.Core.AppSettings.DataDirectory, "office.json")).Contains("sk-ant-test"), "key not in office.json");
                SecretStore.Write(SecretNames.ClaudeKey, null);
                AiSettings.Update(x => x.Provider = AiProviderKind.None);
            });

            r.Test("W22 Claude Desktop connected: the pane says to ask in Claude Desktop; Store edition status asks to paste the configuration", () =>
            {
                AiService.Override = null;
                AiSettings.Update(x => { x.Provider = AiProviderKind.None; x.AllowClaudeDesktop = false; });
                w!.AiPane!.Ask("x"); Ui.Pump(50); // makes sure the setup card is there
                var note = Ui.Find<Label>(w.AiPane!, "lblAiDesktopConnected");
                T.False(note.Visible, "no note while Claude Desktop is not connected");
                AiSettings.Update(x => x.AllowClaudeDesktop = true); Ui.Pump(100);
                T.True(note.Visible, "note shown once Claude Desktop is allowed");
                T.True(Ui.Find<LinkLabel>(w.AiPane!, "lnkAiStatus").Text.Contains("Claude Desktop"), Ui.Find<LinkLabel>(w.AiPane!, "lnkAiStatus").Text);
                // Store edition (cannot write Claude's file): connecting copies the configuration and the status says to paste it
                Environment.SetEnvironmentVariable("MICROADS_CLAUDE_MANUAL", "1");
                try
                {
                    AiSettings.Update(x => x.AllowClaudeDesktop = false);
                    using var d = new AiSettingsDialog(AiSettingsDialog.Section.Claude);
                    d.Show(); Ui.Pump(50);
                    Ui.DialogHandlers.Clear();
                    var b = Ui.Find<Button>(d, "btnAiClaudeDesktop");
                    T.True(b.Text.Contains("Kết nối"), b.Text);
                    // the confirmation (native) is cancelled by the closer: still not allowed
                    Ui.Click(b); Ui.Pump(400);
                    T.False(AiSettings.Load().AllowClaudeDesktop, "cancel = not allowed");
                    AiSettings.Update(x => x.AllowClaudeDesktop = true);
                    d.Close();
                    using var d2 = new AiSettingsDialog(AiSettingsDialog.Section.Claude);
                    d2.Show(); Ui.Pump(50);
                    var st = Ui.Find<Label>(d2, "lblAiClaudeDesktopStatus").Text;
                    T.True(st.Contains("Sao chép cấu hình") && !st.StartsWith("Đã ghi cấu hình"), st);
                    T.True(Ui.Find<Button>(d2, "btnAiClaudeDesktop").Text.Contains("Ngắt"), "shows Disconnect");
                    d2.Close();
                }
                finally { Environment.SetEnvironmentVariable("MICROADS_CLAUDE_MANUAL", null); AiSettings.Update(x => x.AllowClaudeDesktop = false); }
                Ui.Pump(100);
                T.False(note.Visible, "note hidden after disconnecting");
            });

            r.Test("W15 Office settings: 'Trợ lý AI…' opens AI settings; 'install updates by itself' is saved", () =>
            {
                using var d = new OfficeSettingsDialog();
                d.Show(); Ui.Pump(50);
                var btn = Ui.Find<Button>(d, "btnSettingsAi");
                T.True(btn.Visible, "visible");
                Ui.SeenDialogs.Clear();
                Ui.Click(btn); Ui.Pump(200);
                T.True(Ui.SeenDialogs.Contains("aiSettingsDialog"), "AI settings opened");
                var auto = Ui.Find<CheckBox>(d, "chkAutoUpdate"); var inst = Ui.Find<CheckBox>(d, "chkAutoInstall");
                auto.Checked = true; inst.Checked = false;
                Ui.Click(Ui.Find<Button>(d, "btnOk"));
                OfficeConfig.Reload();
                T.False(OfficeConfig.Current.AutoInstallForcedUpdates, "saved as off");
                using var d2 = new OfficeSettingsDialog();
                Ui.Find<CheckBox>(d2, "chkAutoInstall").Checked = true;
                d2.Show(); Ui.Click(Ui.Find<Button>(d2, "btnOk"));
                OfficeConfig.Reload();
                T.True(OfficeConfig.Current.AutoInstallForcedUpdates, "saved as on");
            });
        }
        finally
        {
            AiService.Override = null;
            // close while the dialog closer still runs, so a "save changes?" prompt can never wait on screen
            try { foreach (var x in Application.OpenForms.OfType<WordMainForm>().ToList()) { x.Editor.MarkSaved(); x.Close(); } Ui.Pump(300); } catch { }
            closer.Stop();
        }
    }
}

public static class ExcelUiTests
{
    public static void Run(Runner r)
    {
        var closer = Ui.StartCloser();
        ExcelForm? f = null;
        var ai = new ScriptedProvider();
        AiService.Override = _ => ai;
        AiSettings.Update(x => { x.Provider = AiProviderKind.Compatible; x.CompatibleModel = "fake-model"; });
        try
        {
            r.Test("X01 Home › Trợ lý AI opens the pane with the Excel quick actions", () =>
            {
                f = (ExcelForm)Excel.ExcelApp.CreateMainForm([]);
                f.Show(); Ui.Pump(300); f.HideBackstage(); Ui.Pump(100);
                var btn = (RibbonButton)f.CommandControls["btnAiAssistant"];
                Ui.Click(btn);
                T.True(f.AiPane is { Visible: true }, "pane visible");
                T.True(btn.Checked, "checked");
                var names = Ui.Find<FlowLayoutPanel>(f.AiPane!, "flowAiQuick").Controls.Cast<Control>().Select(c => c.Name).ToList();
                foreach (var n in new[] { "btnAiQ_Explain", "btnAiQ_WriteFormula", "btnAiQ_Analyze", "btnAiQ_Summary", "btnAiQ_Chart", "btnAiQ_Clean", "btnAiQ_Template" }) T.True(names.Contains(n), n);
                T.True(Ui.Find<CheckBox>(f.AiPane!, "chkAiSendDoc").Text.Contains("trang tính"), "sheet wording");
            });

            r.Test("X02 the assistant builds a table: values, formulas, formats, autofit, chart — all undone by ONE Undo", () =>
            {
                ai.Replies.Enqueue("""
                    Đã tạo bảng doanh thu.
                    ```microads-sheet
                    {"actions":[
                      {"range":"A1","values":[["Chi nhánh","Quý 1","Quý 2","Tổng"],["Hà Nội",120,150,"=SUM(B2:C2)"],["HCM",340,310,"=SUM(B3:C3)"]]},
                      {"cell":"A4","value":"Cộng"},{"cell":"D4","formula":"=SUM(D2:D3)"},
                      {"range":"A1:D1","bold":true,"fill":"#D9E1F2","align":"center","border":true},
                      {"range":"B2:D4","numberFormat":"#,##0"},
                      {"autofit":"A:D"},
                      {"chart":{"type":"column","range":"A1:C3","title":"Doanh thu","at":"F2"}}
                    ]}
                    ```
                    """);
                f!.AiPane!.Ask("Tạo bảng doanh thu");
                Ui.WaitRequest(f.AiPane);
                var s = f.Doc.Sheet;
                T.Eq(270.0, s.GetValue(1, 3).Num);
                T.Eq(920.0, s.GetValue(3, 3).Num);
                T.Eq("SUM(D2:D3)", s.GetCell(3, 3)!.Formula);
                T.True(s.StyleAt(0, 0).Bold && s.StyleAt(0, 0).Fill != 0, "header style");
                T.Eq("#,##0", s.StyleAt(1, 1).NumberFormat);
                T.True(s.RawColWidth(0) != Sheet.DefaultColWidth, "autofit changed the width");
                T.Eq(1, s.Charts.Count); T.Eq("Doanh thu", s.Charts[0].Title);
                T.True(f.AiPane.LastAnswer!.Footnote.Contains("biểu đồ"), f.AiPane.LastAnswer.Footnote);
                Ui.Shot(f, "excel-ai-table");
                Ui.Click(f.AiPane.LastAnswer.FindAction("btnAiUndo")!);
                T.True(f.Doc.Sheet.GetValue(0, 0).IsBlank && f.Doc.Sheet.Charts.Count == 0, "one undo removes everything");
            });

            r.Test("X03 an unknown function or a bad address writes nothing and explains the problem", () =>
            {
                ai.Replies.Enqueue("```microads-sheet\n{\"actions\":[{\"cell\":\"A1\",\"value\":\"x\"},{\"cell\":\"B1\",\"formula\":\"=FOOBAR(A1)\"},{\"cell\":\"ZZZZ0\",\"value\":1}]}\n```");
                f!.AiPane!.Ask("Thử");
                Ui.WaitRequest(f.AiPane);
                T.True(f.Doc.Sheet.GetValue(0, 0).IsBlank, "nothing written");
                var foot = f.AiPane.LastAnswer!.Footnote;
                T.True(foot.Contains("FOOBAR") && foot.Contains("ZZZZ0"), foot);
            });

            r.Test("X04 sheet context: values with formulas and the selection are sent; 'Giải thích công thức' works", () =>
            {
                var s = f!.Doc.Sheet;
                f.Grid.Select(0, 0); f.Grid.BeginEdit("5", true); f.Grid.CommitEdit(0, 0);
                f.Grid.Select(0, 1); f.Grid.BeginEdit("=A1*2", true); f.Grid.CommitEdit(0, 0);
                f.Grid.Select(0, 1);
                ai.Replies.Enqueue("Công thức nhân A1 với 2.");
                Ui.Click(Ui.Find<Button>(f.AiPane!, "btnAiQ_Explain"));
                Ui.WaitRequest(f.AiPane!);
                var sent = ai.Requests.Last().Messages.Last().Text;
                T.True(sent.Contains("A1: 5") && sent.Contains("B1: 10  [=A1*2]"), sent);
                T.True(sent.Contains("<selection>") && sent.Contains("[=A1*2]"), "selection with formula");
                T.True(ai.Requests.Last().Instructions.Contains("VLOOKUP") && ai.Requests.Last().Instructions.Contains("SUM"), "supported functions listed");
            });

            r.Test("X05 a chat answer with a table or a formula can be written at the active cell", () =>
            {
                f!.Grid.Select(5, 0);
                ai.Replies.Enqueue("| Tên | Điểm |\n|---|---|\n| An | 8 |");
                f.AiPane!.Ask("Bảng điểm");
                Ui.WaitRequest(f.AiPane);
                Ui.Click(f.AiPane.LastAnswer!.FindAction("btnAiInsert")!);
                T.Eq("Tên", f.Doc.Sheet.GetValue(5, 0).Str);
                T.Eq(8.0, f.Doc.Sheet.GetValue(6, 1).Num);
                f.Grid.Select(9, 0);
                ai.Replies.Enqueue("Dùng công thức:\n\n`=SUM(A1:B1)`");
                f.AiPane.Ask("Tổng");
                Ui.WaitRequest(f.AiPane);
                Ui.Click(f.AiPane.LastAnswer!.FindAction("btnAiInsert")!);
                T.Eq("SUM(A1:B1)", f.Doc.Sheet.GetCell(9, 0)!.Formula);
                T.Eq(15.0, f.Doc.Sheet.GetValue(9, 0).Num);
            });

            r.Test("X06 new sheets: addSheet creates and activates a summary sheet with formulas to the data", () =>
            {
                ai.Replies.Enqueue("```microads-sheet\n{\"actions\":[{\"addSheet\":\"Tổng hợp\"},{\"cell\":\"A1\",\"value\":\"Tổng A1\"},{\"cell\":\"B1\",\"formula\":\"=Sheet1!A1+0\"}]}\n```");
                var first = f!.Doc.Book.Sheets[0].Name;
                ai.Replies.Clear();
                ai.Replies.Enqueue("```microads-sheet\n{\"actions\":[{\"addSheet\":\"Tổng hợp\"},{\"cell\":\"A1\",\"value\":\"Tổng A1\"},{\"cell\":\"B1\",\"formula\":\"='" + first + "'!A1+0\"}]}\n```");
                f.AiPane!.Ask("Tổng hợp");
                Ui.WaitRequest(f.AiPane);
                T.Eq("Tổng hợp", f.Doc.Sheet.Name);
                T.Eq(5.0, f.Doc.Sheet.GetValue(0, 1).Num);
                T.True(f.Tabs.Book!.Sheets.Count == 2, "tabs show the new sheet");
            });

            r.Test("X08 'B:B' formats only the used part of the column; huge explicit ranges are refused; Rollback restores Redo and the modified flag", () =>
            {
                var s = f!.Doc.Sheet;
                int before = s.UsedRange().R2;
                ai.Replies.Enqueue("```microads-sheet\n{\"actions\":[{\"range\":\"B:B\",\"bold\":true}]}\n```");
                f.AiPane!.Ask("In đậm cột B");
                Ui.WaitRequest(f.AiPane);
                T.True(f.Doc.Sheet.UsedRange().R2 <= Math.Max(before, 20), $"bounded: used range now ends at row {f.Doc.Sheet.UsedRange().R2 + 1}");
                T.True(f.Doc.Sheet.StyleAt(0, 1).Bold, "B1 bold");
                ai.Replies.Enqueue("```microads-sheet\n{\"actions\":[{\"range\":\"A1:Z20000\",\"fill\":\"#FFFF00\"}]}\n```");
                f.AiPane.Ask("Tô vàng");
                Ui.WaitRequest(f.AiPane);
                T.True(f.AiPane.LastAnswer!.Footnote.Contains("too large"), f.AiPane.LastAnswer.Footnote);
                // Rollback (used when applying fails half-way) puts back Redo and Modified exactly
                f.CmdUndo();
                T.True(f.Doc.CanRedo, "a redo step exists");
                var redo = f.Doc.SaveRedo(); bool mod = f.Doc.Modified; long seq = f.Doc.EditSeq;
                f.Doc.Checkpoint("test");
                f.Doc.Sheet.SetValue(30, 0, Excel.Model.CellValue.Text("rác"));
                f.Doc.Rollback(redo, mod);
                T.True(f.Doc.Sheet.GetCell(30, 0) == null, "half-applied change gone");
                T.True(f.Doc.CanRedo && f.Doc.RedoLabel == redo[0].Label, "redo history kept");
                T.Eq(mod, f.Doc.Modified);
                T.True(f.Doc.EditSeq > seq, "edit sequence moved");
            });

            r.Test("X09 autofit sizes a formula column by its CALCULATED values (no '####' in a Tổng column)", () =>
            {
                ai.Replies.Enqueue("```microads-sheet\n{\"actions\":[{\"addSheet\":\"Doanh thu\"},{\"range\":\"A1\",\"values\":[[\"Chi nhánh\",\"Q1\",\"Q2\",\"Tổng\"],[\"Hà Nội\",125000000,138000000,\"=SUM(B2:C2)\"]]},{\"range\":\"B2:D2\",\"numberFormat\":\"#,##0\"},{\"autofit\":\"A:D\"}]}\n```");
                f!.AiPane!.Ask("Bảng doanh thu");
                Ui.WaitRequest(f.AiPane);
                var s = f.Doc.Sheet;
                T.Eq("Doanh thu", s.Name);
                T.True(s.ColWidths.ContainsKey(3) && s.ColWidths.ContainsKey(1), "widths set");
                double wTotal = s.ColWidths[3], wQ1 = s.ColWidths[1];
                T.True(wTotal >= wQ1, $"Tổng (263,000,000) at least as wide as Q1 (125,000,000): {wTotal} vs {wQ1}");
            });

            r.Test("X07 inserting a chat answer with an invalid formula changes nothing and says why", () =>
            {
                f!.Grid.Select(12, 0);
                bool couldUndo = f.Doc.CanUndo; var label = f.Doc.UndoLabel;
                ai.Replies.Enqueue("`=KHONGCO(A1)`");
                f.AiPane!.Ask("Công thức");
                Ui.WaitRequest(f.AiPane);
                Ui.Click(f.AiPane.LastAnswer!.FindAction("btnAiInsert")!);
                T.True(f.Doc.Sheet.GetCell(12, 0) == null, "cell untouched");
                T.Eq(label, f.Doc.UndoLabel);
                T.True(f.AiPane.LastAnswer!.Footnote.Contains("KHONGCO"), f.AiPane.LastAnswer.Footnote);
            });
        }
        finally
        {
            AiService.Override = null;
            try
            {
                if (f != null) { foreach (var d in f.Docs) d.Modified = false; f.Close(); }
                Ui.Pump(300);
            }
            catch { }
            closer.Stop();
        }
    }
}
