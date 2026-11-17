using System.Diagnostics;
using System.Text;
using MicroadsOffice.AI.Providers;
using MicroadsOffice.Core.Office;
using MicroadsOffice.UI;

namespace MicroadsOffice.AI.UI;

/// <summary>
/// The "Trợ lý AI" side pane (Word, Excel, PDF): status line, quick actions, conversation with live "steps", the question box.
/// Answers that contain an action block are written straight into the document as one undo step (like Claude for Word).
/// </summary>
public sealed class AiPane : Panel
{
    readonly IAiHost _host;
    readonly Label _title;
    readonly LinkLabel _status;
    readonly LinkLabel _planLine;
    readonly FlowLayoutPanel _quick;
    readonly FlowLayoutPanel _chat;
    readonly TextBox _input;
    readonly Button _send, _stop;
    readonly CheckBox _sendDoc, _autoApply;
    readonly List<AiMessage> _history = [];
    CancellationTokenSource? _cts;
    int _seq;

    /// <summary>Raised when the user closes the pane (the host hides it and unticks its ribbon button).</summary>
    public event EventHandler? CloseRequested;
    /// <summary>For tests: completes when the current request (including applying) is finished.</summary>
    public Task? CurrentRequest { get; private set; }
    public bool Busy => _cts != null;
    public IReadOnlyList<AiMessage> History => _history;
    /// <summary>The last assistant message view (tests and the bridge read its state).</summary>
    public AiMessageView? LastAnswer { get; private set; }

    public AiPane(IAiHost host)
    {
        _host = host;
        Name = "aiPane";
        Dock = DockStyle.Right;
        Width = Dpi.Scale(380);
        BackColor = Color.White;
        Padding = new Padding(1, 0, 0, 0);
        Font = Theme.UiFont;

        // ---- header ----
        var header = new Panel { Dock = DockStyle.Top, Height = Dpi.Scale(40), BackColor = Theme.PanelBack, Padding = new Padding(Dpi.Scale(10), 0, Dpi.Scale(4), 0) };
        _title = new Label { Text = Loc.T("Trợ lý AI", "AI assistant"), Font = new Font("Segoe UI Semibold", 11f), AutoSize = false, Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft, Name = "lblAiTitle" };
        var tools = new FlowLayoutPanel { Dock = DockStyle.Right, AutoSize = true, FlowDirection = FlowDirection.LeftToRight, WrapContents = false, Padding = new Padding(0, Dpi.Scale(6), 0, 0) };
        tools.Controls.Add(IconButton("btnAiNewChat", "", Loc.T("Cuộc trò chuyện mới", "New conversation"), (_, _) => NewConversation()));
        tools.Controls.Add(IconButton("btnAiSettings", "", Loc.T("Cài đặt Trợ lý AI (tài khoản, mô hình)", "AI assistant settings (account, model)"), (_, _) => OpenSettings()));
        tools.Controls.Add(IconButton("btnAiClose", "", Loc.T("Đóng Trợ lý AI", "Close the AI assistant"), (_, _) => CloseRequested?.Invoke(this, EventArgs.Empty)));
        header.Controls.Add(_title);
        header.Controls.Add(tools);

        _status = new LinkLabel { Dock = DockStyle.Top, Height = Dpi.Scale(22), Padding = new Padding(Dpi.Scale(10), Dpi.Scale(3), 0, 0), Name = "lnkAiStatus", LinkBehavior = LinkBehavior.HoverUnderline, BackColor = Theme.PanelBack };
        _status.LinkClicked += (_, _) => OpenSettings();
        _planLine = new LinkLabel { Dock = DockStyle.Top, Height = Dpi.Scale(22), Padding = new Padding(Dpi.Scale(10), Dpi.Scale(3), 0, 0), Name = "lnkAiPlanUsage", Visible = false, BackColor = Color.FromArgb(232, 240, 254) };
        _planLine.LinkClicked += (_, _) => OpenUrl(AiDefaults.ChatGptUsage);

        // ---- quick actions ----
        _quick = new FlowLayoutPanel { Dock = DockStyle.Top, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, Padding = new Padding(Dpi.Scale(8), Dpi.Scale(6), Dpi.Scale(6), Dpi.Scale(2)), Name = "flowAiQuick", MaximumSize = new Size(Width, 0) };
        foreach (var qa in host.QuickActions)
        {
            var b = new Button
            {
                Text = qa.Label, Name = "btnAiQ_" + qa.Name, AutoSize = true, FlatStyle = FlatStyle.Flat, BackColor = Color.White,
                Margin = new Padding(0, 0, Dpi.Scale(4), Dpi.Scale(4)), Cursor = Cursors.Hand, Font = Theme.SmallFont, Tag = qa,
            };
            b.FlatAppearance.BorderColor = Theme.Border;
            b.AccessibleName = qa.Label;
            new ToolTip().SetToolTip(b, qa.Tooltip ?? qa.Prompt);
            b.Click += (_, _) => RunQuick(qa);
            _quick.Controls.Add(b);
        }

        // ---- conversation ----
        _chat = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown, WrapContents = false, AutoScroll = true, Name = "flowAiChat", Padding = new Padding(Dpi.Scale(8)), BackColor = Color.White };
        _chat.Resize += (_, _) => FitChildren();

        // ---- composer ----
        var composer = new Panel { Dock = DockStyle.Bottom, Height = Dpi.Scale(150), Padding = new Padding(Dpi.Scale(8), Dpi.Scale(4), Dpi.Scale(8), Dpi.Scale(8)), BackColor = Theme.PanelBack, Name = "pnlAiComposer" };
        var s = AiSettings.Load();
        _sendDoc = new CheckBox { Text = host.Kind == AppKind.Excel ? Loc.T("Gửi kèm dữ liệu trang tính đang mở", "Include the open sheet's data") : Loc.T("Gửi kèm nội dung văn bản đang mở", "Include the open document"), Checked = s.SendDocument, Dock = DockStyle.Top, Height = Dpi.Scale(22), Name = "chkAiSendDoc" };
        _sendDoc.CheckedChanged += (_, _) => AiSettings.Update(x => x.SendDocument = _sendDoc.Checked);
        _autoApply = new CheckBox
        {
            Text = host.Kind == AppKind.Excel ? Loc.T("Tự ghi kết quả vào trang tính", "Write results into the sheet") : Loc.T("Tự ghi kết quả vào văn bản", "Write results into the document"),
            Checked = host.CanEdit, Enabled = host.CanEdit, Dock = DockStyle.Top, Height = Dpi.Scale(22), Name = "chkAiAutoApply",
        };
        if (!host.CanEdit) _autoApply.Text = Loc.T("Tệp PDF chỉ đọc: trợ lý trả lời trong khung này", "PDF is read-only: answers stay in this pane");
        _input = new TextBox { Multiline = true, Dock = DockStyle.Fill, ScrollBars = ScrollBars.Vertical, Name = "txtAiInput", Font = new Font("Segoe UI", 10f), AcceptsReturn = true };
        _input.PlaceholderText = Loc.T("Hỏi hoặc yêu cầu Trợ lý AI… (Enter để gửi, Shift+Enter xuống dòng)", "Ask or tell the AI assistant… (Enter to send, Shift+Enter for a new line)");
        _input.KeyDown += (_, e) =>
        {
            if (e.KeyCode == Keys.Enter && !e.Shift) { e.SuppressKeyPress = true; Send(); }
        };
        var buttons = new FlowLayoutPanel { Dock = DockStyle.Bottom, Height = Dpi.Scale(36), FlowDirection = FlowDirection.RightToLeft, Padding = new Padding(0, Dpi.Scale(4), 0, 0) };
        _send = new Button { Text = Loc.T("Gửi", "Send"), Name = "btnAiSend", AutoSize = true, MinimumSize = new Size(Dpi.Scale(80), Dpi.Scale(28)), FlatStyle = FlatStyle.Flat, BackColor = Theme.Accent, ForeColor = Color.White, Font = Theme.UiFontBold, Cursor = Cursors.Hand };
        _send.FlatAppearance.BorderSize = 0;
        _send.Click += (_, _) => Send();
        _stop = new Button { Text = Loc.T("Dừng", "Stop"), Name = "btnAiStop", AutoSize = true, MinimumSize = new Size(Dpi.Scale(80), Dpi.Scale(28)), FlatStyle = FlatStyle.Flat, Enabled = false, Cursor = Cursors.Hand };
        _stop.Click += (_, _) => _cts?.Cancel();
        buttons.Controls.Add(_send);
        buttons.Controls.Add(_stop);
        composer.Controls.Add(_input);
        composer.Controls.Add(buttons);
        composer.Controls.Add(_autoApply);
        composer.Controls.Add(_sendDoc);

        var border = new Panel { Dock = DockStyle.Left, Width = 1, BackColor = Theme.Border };
        Controls.Add(_chat);
        Controls.Add(composer);
        Controls.Add(_quick);
        Controls.Add(_planLine);
        Controls.Add(_status);
        Controls.Add(header);
        Controls.Add(border);

        AiSettings.Changed += OnSettingsChanged;
        Disposed += (_, _) => { AiSettings.Changed -= OnSettingsChanged; _cts?.Cancel(); };
        RefreshStatus();
        ShowWelcome();
    }

    void OnSettingsChanged() { if (IsHandleCreated) BeginInvoke(RefreshStatus); else RefreshStatus(); }

    Button IconButton(string name, string glyph, string tip, EventHandler click)
    {
        var b = new Button
        {
            Name = name, Size = new Size(Dpi.Scale(28), Dpi.Scale(28)), FlatStyle = FlatStyle.Flat, Image = Icons.Get(glyph, Dpi.Scale(16), Theme.Ink),
            Margin = new Padding(Dpi.Scale(2), 0, 0, 0), Cursor = Cursors.Hand, AccessibleName = tip, BackColor = Theme.PanelBack,
        };
        b.FlatAppearance.BorderSize = 0;
        new ToolTip().SetToolTip(b, tip);
        b.Click += click;
        return b;
    }

    public void RefreshStatus()
    {
        var s = AiSettings.Load();
        bool ready = AiService.IsReady(s);
        _status.Text = AiService.StatusText(s) + "  ·  " + (ready ? Loc.T("Đổi", "Change") : Loc.T("Kết nối…", "Connect…"));
        _status.LinkArea = new LinkArea(_status.Text.LastIndexOf("·", StringComparison.Ordinal) + 2, _status.Text.Length - _status.Text.LastIndexOf("·", StringComparison.Ordinal) - 2);
        _status.ForeColor = ready ? Theme.OkGreen : Theme.SubtleInk;
        bool plan = ready && s.Provider == AiProviderKind.ChatGpt;
        _planLine.Visible = plan;
        if (plan)
        {
            _planLine.Text = Loc.T("Đang dùng gói ChatGPT  ·  Quản lý hạn mức", "Using ChatGPT plan  ·  Manage usage");
            _planLine.LinkArea = new LinkArea(_planLine.Text.IndexOf('·') + 3, _planLine.Text.Length - _planLine.Text.IndexOf('·') - 3);
        }
        if (_setupCard != null && ready) { _chat.Controls.Remove(_setupCard); _setupCard.Dispose(); _setupCard = null; _desktopNote = null; }
        if (!ready && _setupCard == null && _chat.Controls.Count == 0) ShowWelcome();
        if (_desktopNote != null && _desktopNote.Visible != s.AllowClaudeDesktop) { _desktopNote.Visible = s.AllowClaudeDesktop; FitChildren(); }
    }

    Control? _setupCard;
    Label? _desktopNote;

    void ShowWelcome()
    {
        if (_setupCard != null) return; // never two setup cards
        var s = AiSettings.Load();
        if (AiService.IsReady(s))
        {
            if (_chat.Controls.Count == 0)
                AddNote(_host.CanEdit
                    ? Loc.T("Hãy hỏi, hoặc yêu cầu soạn/sửa văn bản. Ví dụ: \"Soạn hợp đồng lao động mẫu\", \"Viết lại đoạn đang chọn cho trang trọng hơn\". Kết quả được ghi thẳng vào văn bản; bấm Hoàn tác (Ctrl+Z) để bỏ.",
                            "Ask a question or ask for writing/editing. For example: \"Draft a sample employment contract\", \"Rewrite the selected paragraph more formally\". Results go straight into the document; press Undo (Ctrl+Z) to revert.")
                    : Loc.T("Hãy hỏi về tệp PDF đang mở, ví dụ: \"Tóm tắt tài liệu này\".", "Ask about the open PDF, e.g. \"Summarise this document\"."));
            return;
        }
        _setupCard = SetupCard();
        _chat.Controls.Add(_setupCard);
        FitChildren();
    }

    Control SetupCard()
    {
        var card = new FlowLayoutPanel { FlowDirection = FlowDirection.TopDown, WrapContents = false, AutoSize = true, Name = "pnlAiSetup", BackColor = Color.FromArgb(250, 250, 252), Padding = new Padding(Dpi.Scale(10)), Margin = new Padding(0, 0, 0, Dpi.Scale(8)) };
        card.Controls.Add(new Label { Text = Loc.T("Kết nối Trợ lý AI", "Connect the AI assistant"), Font = Theme.UiFontBold, AutoSize = true });
        card.Controls.Add(new Label
        {
            Text = Loc.T("Chọn một cách. Microads Office không thu phí và không lưu tài khoản của bạn trên máy chủ Microads; nội dung chỉ được gửi tới dịch vụ AI bạn chọn khi bạn bấm Gửi.",
                         "Choose one. Microads Office charges nothing and keeps no account on Microads servers; content is sent only to the AI service you choose, when you click Send."),
            AutoSize = true, MaximumSize = new Size(Dpi.Scale(320), 0), ForeColor = Theme.SubtleInk, Margin = new Padding(0, Dpi.Scale(4), 0, Dpi.Scale(8)),
        });
        // Claude Desktop works from Claude's own window: say so, otherwise "not connected" looks like the connection failed
        _desktopNote = new Label
        {
            Name = "lblAiDesktopConnected", AutoSize = true, MaximumSize = new Size(Dpi.Scale(320), 0), ForeColor = Theme.OkGreen,
            Margin = new Padding(0, 0, 0, Dpi.Scale(8)), Visible = AiSettings.Load().AllowClaudeDesktop,
            Text = Loc.T("Claude Desktop đã được kết nối: hãy yêu cầu ngay trong cửa sổ Claude Desktop, ví dụ \"Đọc văn bản đang mở trong Microads Word và tóm tắt\". Muốn hỏi ngay trong khung này, chọn thêm một cách bên dưới.",
                         "Claude Desktop is connected: ask in the Claude Desktop window, e.g. \"Read the document open in Microads Word and summarise it\". To ask right here in this pane, also choose a way below."),
        };
        card.Controls.Add(_desktopNote);
        card.Controls.Add(BigButton("btnAiSetupChatGpt", "Continue with ChatGPT", Loc.T("Dùng gói ChatGPT Plus/Pro của bạn", "Use your ChatGPT Plus/Pro plan"), Color.Black, Color.White, (_, _) => OpenSettings(AiSettingsDialog.Section.ChatGpt, autoStart: true)));
        card.Controls.Add(BigButton("btnAiSetupClaude", Loc.T("Dùng Claude", "Use Claude"), Loc.T("Gói Claude Pro/Max (qua Claude Desktop) hoặc khóa API", "Claude Pro/Max plan (via Claude Desktop) or an API key"), Color.FromArgb(217, 119, 87), Color.White, (_, _) => OpenSettings(AiSettingsDialog.Section.Claude)));
        card.Controls.Add(BigButton("btnAiSetupOther", Loc.T("Khóa API OpenAI / máy chủ AI riêng", "OpenAI API key / own AI server"), Loc.T("Kể cả Ollama chạy ngay trên máy (miễn phí)", "Including Ollama on this computer (free)"), Color.White, Theme.Ink, (_, _) => OpenSettings(AiSettingsDialog.Section.Other)));
        return card;
    }

    Button BigButton(string name, string text, string sub, Color back, Color fore, EventHandler click)
    {
        var b = new Button
        {
            Name = name, Text = text + "\n" + sub, TextAlign = ContentAlignment.MiddleLeft, Size = new Size(Dpi.Scale(320), Dpi.Scale(48)),
            FlatStyle = FlatStyle.Flat, BackColor = back, ForeColor = fore, Cursor = Cursors.Hand, Margin = new Padding(0, 0, 0, Dpi.Scale(6)),
            Padding = new Padding(Dpi.Scale(8), 0, 0, 0), AccessibleName = text,
        };
        b.FlatAppearance.BorderColor = back == Color.White ? Theme.Border : back;
        b.Click += click;
        return b;
    }

    public void OpenSettings() => OpenSettings(AiSettingsDialog.Section.Auto);

    public void OpenSettings(AiSettingsDialog.Section section, bool autoStart = false)
    {
        using var d = new AiSettingsDialog(section, autoStart);
        d.ShowDialog(FindForm());
        RefreshStatus();
    }

    static void OpenUrl(string url) { try { Process.Start(new ProcessStartInfo(url) { UseShellExecute = true }); } catch { } }

    void FitChildren()
    {
        int w = Math.Max(Dpi.Scale(200), _chat.ClientSize.Width - _chat.Padding.Horizontal - SystemInformation.VerticalScrollBarWidth);
        _chat.SuspendLayout();
        foreach (Control c in _chat.Controls) { c.Width = w; if (c is AiMessageView v) v.Relayout(); }
        _chat.ResumeLayout();
        _quick.MaximumSize = new Size(Width - Dpi.Scale(4), 0);
    }

    void AddNote(string text)
    {
        var v = new AiMessageView(AiMessageView.Kind.Note, ++_seq) { };
        v.SetText(text, final: true);
        _chat.Controls.Add(v);
        FitChildren();
    }

    void NewConversation()
    {
        if (Busy) _cts?.Cancel();
        _history.Clear();
        foreach (Control c in _chat.Controls.Cast<Control>().ToList()) { _chat.Controls.Remove(c); c.Dispose(); }
        _setupCard = null; _desktopNote = null;
        LastAnswer = null;
        ShowWelcome();
        _input.Focus();
    }

    void RunQuick(AiQuickAction qa)
    {
        if (qa.NeedsSelection && string.IsNullOrWhiteSpace(_host.SelectionText))
        {
            AddNote(Loc.T("Hãy chọn (bôi đen) phần cần xử lý trước, rồi bấm lại \"", "Select the part to work on first, then click \"") + qa.Label + "\".");
            return;
        }
        Ask(qa.Prompt);
    }

    /// <summary>Sends what is in the question box.</summary>
    public void Send()
    {
        var q = _input.Text.Trim();
        if (q.Length == 0) { _input.Focus(); return; }
        _input.Clear();
        Ask(q);
    }

    /// <summary>Asks a question (public for quick actions, the ribbon and tests).</summary>
    public void Ask(string question)
    {
        if (Busy) return;
        var s = AiSettings.Load();
        if (!AiService.IsReady(s))
        {
            if (_setupCard == null) { _setupCard = SetupCard(); _chat.Controls.Add(_setupCard); FitChildren(); }
            AddNote(Loc.T("Trợ lý AI chưa được kết nối. Chọn một cách kết nối ở trên.", "The AI assistant is not connected yet. Choose a way to connect above."));
            return;
        }
        var provider = AiService.Create(s);
        if (provider == null) { OpenSettings(); return; }

        var user = new AiMessageView(AiMessageView.Kind.User, ++_seq);
        user.SetText(question, final: true);
        _chat.Controls.Add(user);
        var answer = new AiMessageView(AiMessageView.Kind.Assistant, ++_seq);
        answer.SetText(Loc.T("Đang suy nghĩ…", "Thinking…"), final: false);
        _chat.Controls.Add(answer);
        LastAnswer = answer;
        FitChildren();
        _chat.ScrollControlIntoView(answer);

        bool sendDoc = _sendDoc.Checked;
        var messages = new List<AiMessage>(_history) { new(AiRole.User, Prompts.WithContext(_host, question, sendDoc)) };
        var instructions = Prompts.Instructions(_host, s);
        _cts = new CancellationTokenSource();
        var ct = _cts.Token;
        SetBusy(true);
        var buffer = new StringBuilder();
        var uiTimer = new System.Windows.Forms.Timer { Interval = 90 };
        string lastShown = "";
        uiTimer.Tick += (_, _) =>
        {
            string now; lock (buffer) now = buffer.ToString();
            if (now == lastShown) return;
            lastShown = now;
            var (blocks, chatText) = ActionBlocks.Extract(now);
            answer.SetText(chatText.Length > 0 ? chatText : (blocks.Count > 0 ? Loc.T("Đang soạn…", "Writing…") : now), final: false);
            answer.SetSteps(ActionBlocks.Steps(now), done: false);
            _chat.ScrollControlIntoView(answer);
        };
        uiTimer.Start();

        CurrentRequest = RunAsync();

        async Task RunAsync()
        {
            string? full = null; Exception? error = null;
            try
            {
                var model = s.ModelFor(provider.Kind);
                if (string.IsNullOrWhiteSpace(model))
                {
                    var models = await UiThread.On(this, provider.ListModelsAsync(ct));
                    model = provider.DefaultModel(models);
                    if (string.IsNullOrEmpty(model)) throw new AiException(Loc.T("Dịch vụ AI không có mô hình nào dùng được.", "The AI service offers no usable model."));
                    var kind = provider.Kind; var m2 = model;
                    AiSettings.Update(x => x.SetModel(kind, m2));
                }
                full = await UiThread.On(this, Task.Run(() => provider.CompleteAsync(new AiRequest(instructions, messages, model), d => { lock (buffer) buffer.Append(d); }, ct), ct));
            }
            catch (Exception e) { error = e; }
            uiTimer.Stop(); uiTimer.Dispose();
            if (IsDisposed) return;
            try { Finish(question, answer, full, error, ct.IsCancellationRequested, buffer); }
            finally { _cts?.Dispose(); _cts = null; SetBusy(false); }
        }
    }

    void Finish(string question, AiMessageView answer, string? full, Exception? error, bool cancelled, StringBuilder buffer)
    {
        if (error != null || full == null)
        {
            string partial; lock (buffer) partial = buffer.ToString();
            if (cancelled)
            {
                answer.SetText((partial.Length > 0 ? ActionBlocks.Extract(partial).Chat + "\n\n" : "") + Loc.T("(Đã dừng.)", "(Stopped.)"), final: true);
                answer.SetSteps([], done: true);
                return;
            }
            var ae = error as AiException;
            answer.SetError(ae?.Message ?? Loc.T("Lỗi không mong muốn: ", "Unexpected error: ") + error?.Message);
            answer.AddAction("btnAiRetry", Loc.T("Thử lại", "Try again"), () => Ask(question));
            if (ae?.NeedsSignIn == true) answer.AddAction("btnAiFix", Loc.T("Mở cài đặt tài khoản", "Open account settings"), () => OpenSettings());
            if (ae?.HelpUrl is { } url) answer.AddAction("btnAiHelp", Loc.T("Xem hạn mức", "View usage"), () => OpenUrl(url));
            return;
        }

        _history.Add(new AiMessage(AiRole.User, question));
        _history.Add(new AiMessage(AiRole.Assistant, full));
        var (blocks, chat) = ActionBlocks.Extract(full);
        answer.SetText(chat.Length > 0 ? chat : (blocks.Count > 0 ? Loc.T("Đã soạn xong.", "Done.") : full), final: true);
        answer.SetSteps(ActionBlocks.Steps(full), done: true);
        answer.Markdown = chat;

        if (blocks.Any(b => !b.Closed))
        {
            // the answer was cut off inside the content: never write half a document / half an action list
            answer.SetFootnote(Loc.T("Câu trả lời bị cắt ngang nên không ghi gì vào tài liệu. Bấm Thử lại.", "The answer was cut off, so nothing was written. Click Try again."), error: true);
            answer.AddAction("btnAiRetry", Loc.T("Thử lại", "Try again"), () => Ask(question));
        }
        else if (blocks.Count > 0 && _host.CanEdit)
        {
            if (_autoApply.Checked) ApplyBlocks(answer, blocks);
            else answer.AddAction("btnAiApply", _host.Kind == AppKind.Excel ? Loc.T("Ghi vào trang tính", "Write to the sheet") : Loc.T("Ghi vào văn bản", "Write to the document"), () => ApplyBlocks(answer, blocks));
        }
        else if (_host.CanEdit && chat.Length > 0)
        {
            answer.AddAction("btnAiInsert", _host.Kind == AppKind.Excel ? Loc.T("Ghi vào ô đang chọn", "Write to the active cell") : Loc.T("Chèn vào văn bản", "Insert into the document"), () => Insert(answer, chat, false));
            if (_host.Kind == AppKind.Word && !string.IsNullOrEmpty(_host.SelectionText))
                answer.AddAction("btnAiReplace", Loc.T("Thay vùng chọn", "Replace selection"), () => Insert(answer, chat, true));
        }
        answer.AddAction("btnAiCopy", Loc.T("Sao chép", "Copy"), () =>
        {
            try { Clipboard.SetText(Markdown.ToPlainText(chat.Length > 0 ? chat : full)); answer.SetFootnote(Loc.T("Đã sao chép.", "Copied.")); } catch { }
        });
        // Microsoft Store policy 11.16: users can report inappropriate AI content to the developer
        answer.AddAction("btnAiReport", Loc.T("Báo cáo", "Report"), () =>
        {
            using var d = new AiReportDialog(_host.Kind, question, full);
            d.ShowDialog(FindForm());
            if (d.Sent) answer.SetFootnote(Loc.T("Đã gửi báo cáo tới Microads. Cảm ơn bạn.", "Report sent to Microads. Thank you."));
        });
    }

    void ApplyBlocks(AiMessageView answer, List<AiActionBlock> blocks)
    {
        var done = new List<string>();
        foreach (var b in blocks)
        {
            try { done.Add(_host.Apply(b).Summary); }
            catch (Exception e)
            {
                answer.SetFootnote((done.Count > 0 ? string.Join(" ", done) + " " : "") + Loc.T("Không ghi được: ", "Could not write: ") + e.Message, error: true);
                if (done.Count > 0) OfferUndo(answer, done.Count, null);
                return;
            }
        }
        answer.SetFootnote("✓ " + string.Join(" ", done));
        answer.RemoveAction("btnAiApply");
        OfferUndo(answer, blocks.Count, () => ApplyBlocks(answer, blocks));
    }

    /// <summary>
    /// "Hoàn tác" under an answer: undoes exactly the assistant's steps, and only while nothing else changed the document
    /// since (otherwise it would undo the user's own typing or a Claude Desktop edit).
    /// </summary>
    void OfferUndo(AiMessageView answer, int steps, Action? writeAgain)
    {
        long seq = _host.EditSeq;
        answer.RemoveAction("btnAiUndo");
        answer.AddAction("btnAiUndo", Loc.T("Hoàn tác", "Undo"), () =>
        {
            if (_host.EditSeq != seq)
            {
                answer.SetFootnote(Loc.T("Tài liệu đã có thay đổi khác sau lần ghi này, nên không hoàn tác tự động để tránh mất thay đổi đó. Hãy dùng Ctrl+Z để hoàn tác từng bước.",
                                         "The document changed after this write, so it is not undone automatically (that would lose the newer change). Use Ctrl+Z step by step."), error: true);
                answer.RemoveAction("btnAiUndo");
                return;
            }
            for (int i = 0; i < steps; i++) _host.Undo();
            answer.SetFootnote(Loc.T("Đã hoàn tác.", "Undone."));
            answer.RemoveAction("btnAiUndo");
            if (writeAgain != null) answer.AddAction("btnAiApply", Loc.T("Ghi lại", "Write again"), writeAgain);
        });
    }

    void Insert(AiMessageView answer, string md, bool replace)
    {
        try
        {
            var r = _host.InsertAnswer(md, replace);
            answer.SetFootnote("✓ " + r.Summary);
            OfferUndo(answer, 1, null);
        }
        catch (Exception e) { answer.SetFootnote(Loc.T("Không chèn được: ", "Could not insert: ") + e.Message, error: true); }
    }

    void SetBusy(bool busy)
    {
        _send.Enabled = !busy; _stop.Enabled = busy;
        foreach (Control c in _quick.Controls) c.Enabled = !busy;
        if (busy) return;
        if (AiSettings.Load() is { Provider: AiProviderKind.ChatGpt, ChatGptWelcomeShown: false } s && AiService.IsReady(s) && _history.Count > 0)
        {
            AiSettings.Update(x => x.ChatGptWelcomeShown = true);
            ShowPlanWelcome(FindForm());
        }
    }

    /// <summary>OpenAI UI guideline: shown once after the first use of the ChatGPT plan.</summary>
    public static void ShowPlanWelcome(IWin32Window? owner)
    {
        if (Environment.GetEnvironmentVariable("MICROADS_AI_NO_MODALS") == "1") return;
        Msg.Info(owner, Loc.T("Các yêu cầu AI trong Microads Office sẽ dùng hạn mức gói ChatGPT của bạn. Bạn có thể xem và giới hạn mức dùng trong ChatGPT › Settings › Usage.",
                              "Eligible AI requests in Microads Office use your ChatGPT plan. You can review and limit usage in ChatGPT › Settings › Usage."),
                 Loc.T("Bạn đang dùng gói ChatGPT", "You're using your ChatGPT plan"));
    }

    protected override void OnVisibleChanged(EventArgs e)
    {
        base.OnVisibleChanged(e);
        if (Visible) { RefreshStatus(); BeginInvokeSafe(() => _input.Focus()); }
    }

    void BeginInvokeSafe(Action a) { if (IsHandleCreated) BeginInvoke(a); }

    /// <summary>Puts text in the question box (used by the ribbon "Hỏi AI về vùng chọn").</summary>
    public void SetQuestion(string text) { _input.Text = text; _input.SelectionStart = text.Length; _input.Focus(); }
}

/// <summary>One bubble in the conversation: the text (Markdown rendered), live steps, a footnote and action buttons.</summary>
public sealed class AiMessageView : Panel
{
    public enum Kind { User, Assistant, Note }
    readonly Kind _kind;
    readonly Label _who;
    readonly RichTextBox _text;
    readonly FlowLayoutPanel _steps;
    readonly Label _foot;
    readonly FlowLayoutPanel _actions;
    public int Seq { get; }
    public string Markdown { get; set; } = "";
    public string Text2 => _text.Text;
    public string Footnote => _foot.Text;
    public bool IsError { get; private set; }
    public IEnumerable<Button> Actions => _actions.Controls.OfType<Button>();

    public AiMessageView(Kind kind, int seq)
    {
        _kind = kind; Seq = seq;
        Name = "aiMsg" + seq;
        Margin = new Padding(0, 0, 0, Dpi.Scale(8));
        BackColor = kind switch { Kind.User => Color.FromArgb(240, 237, 232), Kind.Note => Color.FromArgb(248, 249, 250), _ => Color.White };
        Padding = new Padding(Dpi.Scale(8), Dpi.Scale(6), Dpi.Scale(8), Dpi.Scale(6));
        _who = new Label { AutoSize = false, Height = Dpi.Scale(18), Font = Theme.UiFontBold, ForeColor = Theme.SubtleInk, Text = kind switch { Kind.User => Loc.T("Bạn", "You"), Kind.Assistant => Loc.T("Trợ lý AI", "AI assistant"), _ => "" }, Visible = kind != Kind.Note };
        _text = new RichTextBox { ReadOnly = true, BorderStyle = BorderStyle.None, BackColor = BackColor, ScrollBars = RichTextBoxScrollBars.None, Font = new Font("Segoe UI", 10f), DetectUrls = true, Name = "rtbAiMsg" + seq, TabStop = false };
        _text.ContentsResized += (_, e) => { _contentH = e.NewRectangle.Height; Relayout(); };
        _text.LinkClicked += (_, e) =>
        {
            // AI text can be steered by a malicious document: open only web and e-mail links, never file:/UNC/other schemes
            if (Uri.TryCreate(e.LinkText, UriKind.Absolute, out var u) && (u.Scheme == Uri.UriSchemeHttps || u.Scheme == Uri.UriSchemeHttp || u.Scheme == Uri.UriSchemeMailto))
                try { Process.Start(new ProcessStartInfo(u.AbsoluteUri) { UseShellExecute = true }); } catch { }
        };
        _steps = new FlowLayoutPanel { FlowDirection = FlowDirection.TopDown, WrapContents = false, AutoSize = true, BackColor = Color.FromArgb(250, 247, 242), Visible = false, Name = "flowAiSteps" + seq, Padding = new Padding(Dpi.Scale(6)) };
        _foot = new Label { AutoSize = false, ForeColor = Theme.OkGreen, Visible = false, Name = "lblAiFoot" + seq };
        _actions = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.LeftToRight, WrapContents = true, Name = "flowAiActions" + seq };
        Controls.AddRange([_who, _text, _steps, _foot, _actions]);
    }

    int _contentH;

    public void Relayout()
    {
        int w = Math.Max(50, Width - Padding.Horizontal);
        int y = Padding.Top;
        if (_who.Visible) { _who.SetBounds(Padding.Left, y, w, _who.Height); y += _who.Height; }
        _text.SetBounds(Padding.Left, y, w, Math.Max(Dpi.Scale(20), _contentH + Dpi.Scale(4))); y += _text.Height;
        if (_steps.Visible)
        {
            _steps.MaximumSize = new Size(w, 0);
            foreach (Control c in _steps.Controls) c.MaximumSize = new Size(w - Dpi.Scale(12), 0);
            _steps.Location = new Point(Padding.Left, y + Dpi.Scale(4)); _steps.Width = w; y = _steps.Bottom;
        }
        if (_foot.Visible)
        {
            var sz = TextRenderer.MeasureText(_foot.Text, _foot.Font, new Size(w, 0), TextFormatFlags.WordBreak);
            _foot.SetBounds(Padding.Left, y + Dpi.Scale(4), w, sz.Height + 2); y = _foot.Bottom;
        }
        if (_actions.Controls.Count > 0) { _actions.MaximumSize = new Size(w, 0); _actions.Location = new Point(Padding.Left, y + Dpi.Scale(4)); y = _actions.Bottom; }
        Height = y + Padding.Bottom;
    }

    public void SetText(string text, bool final)
    {
        _text.SuspendLayout();
        if (final && _kind == Kind.Assistant) RichMarkdown.Render(_text, text);
        else _text.Text = text;
        _text.ResumeLayout();
        Relayout();
    }

    public void SetError(string message)
    {
        IsError = true;
        _text.Text = message;
        _text.SelectAll(); _text.SelectionColor = Theme.ErrRed; _text.Select(0, 0);
        Relayout();
    }

    public void SetSteps(List<string> steps, bool done)
    {
        _steps.SuspendLayout();
        _steps.Controls.Clear();
        if (steps.Count > 0)
        {
            _steps.Controls.Add(new Label { Text = Loc.T("Các bước", "Steps"), Font = Theme.UiFontBold, AutoSize = true });
            for (int i = 0; i < steps.Count; i++)
            {
                bool complete = done || i < steps.Count - 1;
                _steps.Controls.Add(new Label
                {
                    Text = (complete ? "✓  " : "⋯  ") + steps[i], AutoSize = true, Font = complete ? new Font(Theme.UiFont, FontStyle.Strikeout) : Theme.UiFont,
                    ForeColor = complete ? Theme.SubtleInk : Theme.Ink, Name = $"lblAiStep{Seq}_{i}",
                });
            }
        }
        _steps.Visible = steps.Count > 0;
        _steps.ResumeLayout();
        Relayout();
    }

    public void SetFootnote(string text, bool error = false)
    {
        _foot.Text = text; _foot.ForeColor = error ? Theme.ErrRed : Theme.OkGreen; _foot.Visible = text.Length > 0;
        Relayout();
    }

    public void AddAction(string name, string text, Action click)
    {
        RemoveAction(name);
        var b = new Button { Name = name, Text = text, AutoSize = true, FlatStyle = FlatStyle.Flat, BackColor = Color.White, Cursor = Cursors.Hand, Margin = new Padding(0, 0, Dpi.Scale(6), Dpi.Scale(4)), Font = Theme.SmallFont, AccessibleName = text };
        b.FlatAppearance.BorderColor = Theme.Border;
        b.Click += (_, _) => click();
        _actions.Controls.Add(b);
        Relayout();
    }

    public void RemoveAction(string name)
    {
        foreach (var b in _actions.Controls.OfType<Button>().Where(x => x.Name == name).ToList()) { _actions.Controls.Remove(b); b.Dispose(); }
        Relayout();
    }

    public Button? FindAction(string name) => _actions.Controls.OfType<Button>().FirstOrDefault(b => b.Name == name);
}

/// <summary>Renders assistant Markdown into a RichTextBox (headings, bold/italic, bullets, tables as tab-separated rows, code).</summary>
public static class RichMarkdown
{
    public static void Render(RichTextBox box, string md)
    {
        box.Clear();
        var baseFont = box.Font;
        void Append(string text, FontStyle style = FontStyle.Regular, float size = 0, Color? color = null, string family = "")
        {
            box.SelectionStart = box.TextLength; box.SelectionLength = 0;
            box.SelectionFont = new Font(family.Length > 0 ? family : baseFont.FontFamily.Name, size > 0 ? size : baseFont.Size, style);
            box.SelectionColor = color ?? Theme.Ink;
            box.AppendText(text);
        }
        void Spans(IEnumerable<MdSpan> spans, FontStyle extra = FontStyle.Regular, float size = 0)
        {
            foreach (var s in spans)
            {
                var st = extra;
                if (s.Bold) st |= FontStyle.Bold;
                if (s.Italic) st |= FontStyle.Italic;
                if (s.Underline) st |= FontStyle.Underline;
                if (s.Strike) st |= FontStyle.Strikeout;
                Append(s.Text.Replace('\v', '\n'), st, size, s.Code ? Color.FromArgb(120, 40, 40) : null, s.Code ? "Consolas" : "");
            }
        }
        bool first = true;
        foreach (var b in AI.Markdown.Parse(md))
        {
            if (!first) Append("\n");
            first = false;
            switch (b)
            {
                case MdHeading h: Spans(h.Spans, FontStyle.Bold, h.Level <= 1 ? 12.5f : h.Level == 2 ? 11.5f : 10.5f); break;
                case MdParagraph p: Spans(p.Spans); break;
                case MdListItem l: Append(new string(' ', l.Level * 4) + (l.Ordered ? "▸ " : "• ")); Spans(l.Spans); break;
                case MdQuote q: Append("│ ", color: Theme.SubtleInk); Spans(q.Spans, FontStyle.Italic); break;
                case MdCode c: Append(c.Text, family: "Consolas", color: Color.FromArgb(60, 60, 60)); break;
                case MdRule: Append("────────────", color: Theme.Border); break;
                case MdPageBreak: Append("— — —", color: Theme.Border); break;
                case MdTable t:
                    for (int r = 0; r < t.Rows.Count; r++)
                    {
                        if (r > 0) Append("\n");
                        for (int c = 0; c < t.Rows[r].Count; c++) { if (c > 0) Append("  │  ", color: Theme.Border); Spans(t.Rows[r][c], r == 0 && t.HasHeader ? FontStyle.Bold : FontStyle.Regular); }
                    }
                    break;
            }
        }
        box.SelectionStart = 0; box.SelectionLength = 0;
    }
}
