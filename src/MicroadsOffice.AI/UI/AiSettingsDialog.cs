using System.Diagnostics;
using MicroadsOffice.AI.Mcp;
using MicroadsOffice.AI.Providers;
using MicroadsOffice.Core.Office;
using MicroadsOffice.Dialogs;
using MicroadsOffice.UI;

namespace MicroadsOffice.AI.UI;

/// <summary>Settings › AI assistant: how to connect (ChatGPT account, Claude, API keys, own server), model, standing instructions.</summary>
public sealed class AiSettingsDialog : DialogBase
{
    public enum Section { Auto, ChatGpt, Claude, Other }

    readonly RadioButton _rbChatGpt, _rbClaudeKey, _rbOpenAi, _rbCompat;
    readonly Label _chatGptStatus, _desktopStatus, _testResult;
    readonly Button _signIn, _signOut, _cancelSignIn, _desktop, _test;
    readonly LinkLabel _otherAccount;
    readonly TextBox _claudeKey, _openAiKey, _compatUrl, _compatKey, _instructions;
    readonly ComboBox _model;
    readonly FlowLayoutPanel _flow;
    readonly bool _autoStart;
    readonly Section _section;
    AiProviderKind _modelFor;
    CancellationTokenSource? _cts;

    /// <summary>For tests: the running sign-in / test task.</summary>
    public Task? Pending { get; private set; }

    public AiSettingsDialog(Section section = Section.Auto, bool autoStart = false) : base(Loc.T("Cài đặt Trợ lý AI", "AI assistant settings"), 640, 660)
    {
        Name = "aiSettingsDialog";
        _section = section; _autoStart = autoStart;
        var s = AiSettings.Load();

        _flow = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown, WrapContents = false, AutoScroll = true, Name = "flowAiSettings" };
        Body.Controls.Add(_flow);
        int w = Dpi.Scale(580);

        // ---------- ChatGPT ----------
        var gChat = Group(Loc.T("Tài khoản ChatGPT (gói Plus hoặc Pro)", "ChatGPT account (Plus or Pro plan)"), "grpAiChatGpt", w);
        _rbChatGpt = Radio(Loc.T("Dùng tài khoản ChatGPT của tôi", "Use my ChatGPT account"), "rbAiChatGpt");
        gChat.Controls.Add(_rbChatGpt);
        gChat.Controls.Add(Note(Loc.T("Đăng nhập bằng tài khoản ChatGPT; các yêu cầu dùng hạn mức gói ChatGPT của bạn, không cần khóa API. (Chương trình \"Sign in with ChatGPT\" chính thức của OpenAI cho phần mềm mã nguồn mở.)",
                                     "Sign in with your ChatGPT account; requests use your ChatGPT plan, no API key needed. (OpenAI's official \"Sign in with ChatGPT\" program for open-source apps.)"), w));
        var chatRow = Row();
        _signIn = new Button { Text = "Continue with ChatGPT", Name = "btnAiChatGptSignIn", AutoSize = true, MinimumSize = new Size(Dpi.Scale(200), Dpi.Scale(32)), FlatStyle = FlatStyle.Flat, BackColor = Color.Black, ForeColor = Color.White, Font = Theme.UiFontBold, Cursor = Cursors.Hand };
        _signIn.FlatAppearance.BorderSize = 0;
        _signIn.Click += (_, _) => StartSignIn(newAccount: false);
        _signOut = Btn(Loc.T("Đăng xuất", "Sign out"), "btnAiChatGptSignOut", (_, _) => SignOut());
        // while the browser sign-in is waiting (up to 10 minutes) the user can give up without closing the dialog
        _cancelSignIn = Btn(Loc.T("Hủy đăng nhập", "Cancel sign-in"), "btnAiChatGptCancel", (_, _) => _cts?.Cancel());
        _cancelSignIn.Visible = false;
        chatRow.Controls.Add(_signIn); chatRow.Controls.Add(_signOut); chatRow.Controls.Add(_cancelSignIn);
        gChat.Controls.Add(chatRow);
        _chatGptStatus = new Label { AutoSize = true, MaximumSize = new Size(w - Dpi.Scale(20), 0), Name = "lblAiChatGptStatus", Margin = new Padding(Dpi.Scale(3), Dpi.Scale(4), 0, Dpi.Scale(2)) };
        gChat.Controls.Add(_chatGptStatus);
        var chatLinks = Row();
        _otherAccount = Link(Loc.T("Dùng tài khoản ChatGPT khác", "Use another ChatGPT account"), "lnkAiChatGptOther", (_, _) => StartSignIn(newAccount: true));
        chatLinks.Controls.Add(_otherAccount);
        chatLinks.Controls.Add(Link(Loc.T("Quản lý hạn mức (ChatGPT › Usage)", "Manage usage (ChatGPT › Usage)"), "lnkAiChatGptUsage", (_, _) => Open(AiDefaults.ChatGptUsage)));
        gChat.Controls.Add(chatLinks);

        // ---------- Claude ----------
        var gClaude = Group("Claude", "grpAiClaude", w);
        gClaude.Controls.Add(new Label { Text = Loc.T("Gói Claude Pro / Max — qua ứng dụng Claude Desktop", "Claude Pro / Max plan — through the Claude Desktop app"), Font = Theme.UiFontBold, AutoSize = true, Margin = new Padding(Dpi.Scale(3), Dpi.Scale(4), 0, 0) });
        gClaude.Controls.Add(Note(Loc.T("Anthropic không cho phần mềm khác đăng nhập tài khoản Claude. Thay vào đó, kết nối Microads Office với ứng dụng Claude Desktop (bạn đăng nhập Claude ở đó): Claude sẽ đọc và chỉnh văn bản, bảng tính đang mở trong Microads Office khi bạn yêu cầu trong Claude Desktop.",
                                        "Anthropic does not let other apps sign in to Claude accounts. Instead, connect Microads Office to the Claude Desktop app (you sign in to Claude there): Claude can then read and edit the documents and workbooks open in Microads Office when you ask it in Claude Desktop."), w));
        var desktopRow = Row();
        _desktop = Btn("", "btnAiClaudeDesktop", (_, _) => ToggleDesktop());
        _desktop.BackColor = Color.FromArgb(217, 119, 87); _desktop.ForeColor = Color.White; _desktop.FlatAppearance.BorderSize = 0; _desktop.Font = Theme.UiFontBold;
        desktopRow.Controls.Add(_desktop);
        desktopRow.Controls.Add(Link(Loc.T("Tải Claude Desktop", "Download Claude Desktop"), "lnkAiClaudeDownload", (_, _) => Open(AiDefaults.ClaudeDesktopDownload)));
        gClaude.Controls.Add(desktopRow);
        var desktopTools = Row();
        desktopTools.Controls.Add(Link(Loc.T("Sao chép cấu hình", "Copy configuration"), "lnkAiClaudeCopyConfig", (_, _) =>
        {
            try { Clipboard.SetText(ClaudeDesktop.ConfigSnippet()); _desktopStatus.Text = Loc.T("Đã sao chép. Trong Claude Desktop: Settings › Developer › Edit Config → dán → lưu → thoát hẳn Claude (khay hệ thống › Quit) rồi mở lại.", "Copied. In Claude Desktop: Settings › Developer › Edit Config → paste → save → quit Claude completely (system tray › Quit) and reopen it."); }
            catch { }
        }));
        desktopTools.Controls.Add(Link(Loc.T("Mở thư mục cấu hình Claude", "Open Claude's config folder"), "lnkAiClaudeConfigFolder", (_, _) =>
        {
            try { var d = ClaudeDesktop.PrimaryConfigDir; Directory.CreateDirectory(d); Process.Start(new ProcessStartInfo(d) { UseShellExecute = true }); } catch { }
        }));
        gClaude.Controls.Add(desktopTools);
        _desktopStatus = new Label { AutoSize = true, MaximumSize = new Size(w - Dpi.Scale(20), 0), Name = "lblAiClaudeDesktopStatus", Margin = new Padding(Dpi.Scale(3), Dpi.Scale(2), 0, Dpi.Scale(8)) };
        gClaude.Controls.Add(_desktopStatus);
        _rbClaudeKey = Radio(Loc.T("Dùng khóa API Claude (Claude Console, trả theo mức dùng)", "Use a Claude API key (Claude Console, pay as you go)"), "rbAiClaudeKey");
        gClaude.Controls.Add(_rbClaudeKey);
        _claudeKey = KeyBox("txtAiClaudeKey", SecretNames.ClaudeKey, "sk-ant-…", w);
        gClaude.Controls.Add(KeyRow(_claudeKey, SecretNames.ClaudeKey, "lnkAiClaudeKeyGet", AiDefaults.AnthropicConsole, "lnkAiClaudeKeyClear"));

        // ---------- other ----------
        var gOther = Group(Loc.T("Khác", "Other"), "grpAiOther", w);
        _rbOpenAi = Radio(Loc.T("Dùng khóa API OpenAI", "Use an OpenAI API key"), "rbAiOpenAiKey");
        gOther.Controls.Add(_rbOpenAi);
        _openAiKey = KeyBox("txtAiOpenAiKey", SecretNames.OpenAiKey, "sk-…", w);
        gOther.Controls.Add(KeyRow(_openAiKey, SecretNames.OpenAiKey, "lnkAiOpenAiKeyGet", AiDefaults.OpenAiKeysPage, "lnkAiOpenAiKeyClear"));
        _rbCompat = Radio(Loc.T("Máy chủ tương thích OpenAI (Ollama, LM Studio, máy chủ nội bộ)", "OpenAI-compatible server (Ollama, LM Studio, internal server)"), "rbAiCompat");
        gOther.Controls.Add(_rbCompat);
        _compatUrl = new TextBox { Text = s.CompatibleBaseUrl, Width = w - Dpi.Scale(40), Name = "txtAiCompatUrl", Margin = new Padding(Dpi.Scale(22), 0, 0, Dpi.Scale(4)), PlaceholderText = "http://127.0.0.1:11434/v1" };
        gOther.Controls.Add(_compatUrl);
        _compatKey = KeyBox("txtAiCompatKey", SecretNames.CompatibleKey, Loc.T("Khóa (nếu máy chủ yêu cầu)", "Key (if the server needs one)"), w);
        gOther.Controls.Add(KeyRow(_compatKey, SecretNames.CompatibleKey, null, null, "lnkAiCompatKeyClear"));

        // ---------- model ----------
        var gModel = Group(Loc.T("Mô hình", "Model"), "grpAiModel", w);
        var modelRow = Row();
        _model = new ComboBox { Width = Dpi.Scale(300), Name = "cboAiModel", DropDownStyle = ComboBoxStyle.DropDown };
        _test = Btn(Loc.T("Kiểm tra kết nối", "Test connection"), "btnAiTest", (_, _) => TestConnection());
        modelRow.Controls.Add(_model); modelRow.Controls.Add(_test);
        gModel.Controls.Add(modelRow);
        _testResult = new Label { AutoSize = true, MaximumSize = new Size(w - Dpi.Scale(20), 0), Name = "lblAiTestResult", Margin = new Padding(Dpi.Scale(3), Dpi.Scale(2), 0, Dpi.Scale(2)) };
        gModel.Controls.Add(_testResult);

        // ---------- instructions ----------
        var gIns = Group(Loc.T("Hướng dẫn luôn áp dụng (không bắt buộc)", "Standing instructions (optional)"), "grpAiInstructions", w);
        _instructions = new TextBox { Multiline = true, Height = Dpi.Scale(54), Width = w - Dpi.Scale(20), Text = s.Instructions, Name = "txtAiInstructions", ScrollBars = ScrollBars.Vertical, PlaceholderText = Loc.T("Ví dụ: Văn phong hành chính, trang trọng; công ty tôi là …", "e.g. Formal tone; my company is …") };
        gIns.Controls.Add(_instructions);

        // ---------- privacy ----------
        var priv = Note(Loc.T("Quyền riêng tư: Trợ lý AI chỉ kết nối Internet khi bạn bấm Gửi (hoặc khi Claude Desktop yêu cầu, nếu bạn đã kết nối). Câu hỏi và nội dung văn bản bạn cho phép được gửi thẳng tới dịch vụ AI bạn chọn ở trên, không qua máy chủ Microads. Khóa API và phiên đăng nhập được mã hóa (Windows DPAPI) và chỉ lưu trên máy này.",
                                "Privacy: the AI assistant connects to the Internet only when you click Send (or when Claude Desktop asks, if you connected it). Your question and the document content you allow go directly to the AI service you chose above, never through Microads servers. API keys and sign-ins are encrypted (Windows DPAPI) and stored only on this computer."), w);
        priv.Name = "lblAiPrivacy";
        _flow.Controls.Add(priv);
        _flow.Controls.Add(Link(Loc.T("Chính sách quyền riêng tư", "Privacy policy"), "lnkAiPrivacy", (_, _) => Open(Product.PrivacyUrl)));

        foreach (var rb in new[] { _rbChatGpt, _rbClaudeKey, _rbOpenAi, _rbCompat }) rb.CheckedChanged += (_, _) => { if (((RadioButton)rb).Checked) LoadModelFor(SelectedKind); };
        switch (s.Provider)
        {
            case AiProviderKind.ChatGpt: _rbChatGpt.Checked = true; break;
            case AiProviderKind.ClaudeKey: _rbClaudeKey.Checked = true; break;
            case AiProviderKind.OpenAiKey: _rbOpenAi.Checked = true; break;
            case AiProviderKind.Compatible: _rbCompat.Checked = true; break;
            default:
                (section switch { Section.Claude => _rbClaudeKey, Section.Other => _rbOpenAi, _ => _rbChatGpt }).Checked = true; break;
        }
        // "Continue with ChatGPT" on the pane's setup card means: use ChatGPT (even if another way was chosen before)
        if (autoStart && section == Section.ChatGpt) _rbChatGpt.Checked = true;
        LoadModelFor(SelectedKind);
        RefreshAccountState();

        AddButton(Loc.T("Hủy", "Cancel"), "btnCancel", DialogResult.Cancel);
        OkButton = AddButton(Loc.T("Lưu", "Save"), "btnOk", DialogResult.OK, primary: true, (_, _) => { if (SaveAll()) { DialogResult = DialogResult.OK; Close(); } });
        AcceptButton = null; // Enter inside text boxes must not close the dialog
        CancelButton = null;
        FormClosing += (_, _) => _cts?.Cancel();
        Shown += (_, _) =>
        {
            var target = section switch { Section.Claude => gClaude, Section.Other => gOther, Section.ChatGpt => gChat, _ => null };
            if (target != null) _flow.ScrollControlIntoView(target);
            // Shown is posted: the dialog may already be closed by then — never start a sign-in nobody can see or cancel
            if (_autoStart && section == Section.ChatGpt && !ChatGptAuth.IsSignedIn && Visible && !IsDisposed) StartSignIn(newAccount: false);
        };
    }

    public AiProviderKind SelectedKind =>
        _rbChatGpt.Checked ? AiProviderKind.ChatGpt : _rbClaudeKey.Checked ? AiProviderKind.ClaudeKey : _rbOpenAi.Checked ? AiProviderKind.OpenAiKey : _rbCompat.Checked ? AiProviderKind.Compatible : AiProviderKind.None;

    // ---------- small builders ----------

    FlowLayoutPanel Group(string title, string name, int w)
    {
        var g = new FlowLayoutPanel { FlowDirection = FlowDirection.TopDown, WrapContents = false, AutoSize = true, Width = w, Name = name, Margin = new Padding(0, 0, 0, Dpi.Scale(10)), Padding = new Padding(0, 0, 0, Dpi.Scale(6)) };
        g.Controls.Add(new Label { Text = title, Font = new Font("Segoe UI Semibold", 10.5f), AutoSize = true, ForeColor = Theme.Ink, Margin = new Padding(0, Dpi.Scale(4), 0, Dpi.Scale(2)) });
        g.Controls.Add(new Panel { Height = 1, Width = w - Dpi.Scale(10), BackColor = Theme.Border, Margin = new Padding(0, 0, 0, Dpi.Scale(4)) });
        _flow.Controls.Add(g);
        return g;
    }

    static FlowLayoutPanel Row() => new() { FlowDirection = FlowDirection.LeftToRight, AutoSize = true, WrapContents = false, Margin = new Padding(0, Dpi.Scale(2), 0, Dpi.Scale(2)) };
    static RadioButton Radio(string text, string name) => new() { Text = text, Name = name, AutoSize = true, Font = Theme.UiFontBold, Margin = new Padding(Dpi.Scale(3), Dpi.Scale(6), 0, Dpi.Scale(2)) };
    static Label Note(string text, int w) => new() { Text = text, AutoSize = true, MaximumSize = new Size(w - Dpi.Scale(20), 0), ForeColor = Theme.SubtleInk, Margin = new Padding(Dpi.Scale(3), Dpi.Scale(2), 0, Dpi.Scale(4)) };

    static Button Btn(string text, string name, EventHandler click)
    {
        var b = new Button { Text = text, Name = name, AutoSize = true, MinimumSize = new Size(Dpi.Scale(100), Dpi.Scale(32)), FlatStyle = FlatStyle.Flat, BackColor = Color.White, Cursor = Cursors.Hand, Margin = new Padding(0, 0, Dpi.Scale(8), 0) };
        b.FlatAppearance.BorderColor = Theme.Border;
        b.Click += click;
        return b;
    }

    static LinkLabel Link(string text, string name, LinkLabelLinkClickedEventHandler click)
    {
        var l = new LinkLabel { Text = text, Name = name, AutoSize = true, Margin = new Padding(Dpi.Scale(3), Dpi.Scale(9), Dpi.Scale(12), 0) };
        l.LinkClicked += click;
        return l;
    }

    TextBox KeyBox(string name, string secret, string placeholder, int w)
    {
        bool has = SecretStore.Has(secret);
        return new TextBox
        {
            Name = name, Width = Dpi.Scale(330), UseSystemPasswordChar = true, Margin = new Padding(Dpi.Scale(22), 0, Dpi.Scale(8), 0),
            PlaceholderText = has ? Loc.T("Đã lưu khóa — nhập để thay", "Key saved — type to replace") : placeholder,
        };
    }

    FlowLayoutPanel KeyRow(TextBox box, string secret, string? getName, string? getUrl, string clearName)
    {
        var r = Row();
        r.Controls.Add(box);
        if (getName != null && getUrl != null) r.Controls.Add(Link(Loc.T("Lấy khóa", "Get a key"), getName, (_, _) => Open(getUrl)));
        var clear = Link(Loc.T("Xóa khóa đã lưu", "Remove saved key"), clearName, (_, _) =>
        {
            SecretStore.Write(secret, null);
            box.Text = ""; box.PlaceholderText = Loc.T("Đã xóa", "Removed");
            AiSettings.RaiseChanged();
        });
        r.Controls.Add(clear);
        return r;
    }

    static void Open(string url) { try { Process.Start(new ProcessStartInfo(url) { UseShellExecute = true }); } catch { } }

    // ---------- state ----------

    void RefreshAccountState()
    {
        var c = ChatGptAuth.Current;
        if (c is { SignedIn: true })
        {
            _chatGptStatus.Text = Loc.T("Đã đăng nhập: ", "Signed in: ") + (c.Email.Length > 0 ? c.Email : c.Subject) + (c.PlanUsage ? "" : Loc.T(" — chưa cho phép dùng gói ChatGPT (đăng nhập lại để cho phép)", " — ChatGPT plan use not allowed (sign in again to allow it)"));
            _chatGptStatus.ForeColor = c.PlanUsage ? Theme.OkGreen : Theme.WarnAmber;
            _signIn.Text = "Continue with ChatGPT";
            _signOut.Enabled = true;
        }
        else
        {
            _chatGptStatus.Text = c is { Email.Length: > 0 } ? Loc.T("Chưa đăng nhập (tài khoản đã dùng: ", "Signed out (last account: ") + c.Email + ")" : Loc.T("Chưa đăng nhập.", "Not signed in.");
            _chatGptStatus.ForeColor = Theme.SubtleInk;
            _signOut.Enabled = false;
        }
        _otherAccount.Enabled = true;
        bool connected = DesktopConnected();
        _desktop.Text = connected ? Loc.T("Ngắt kết nối Claude Desktop", "Disconnect Claude Desktop") : Loc.T("Kết nối với Claude Desktop", "Connect to Claude Desktop");
        var hs = ClaudeDesktop.LastHandshake();
        bool working = connected && hs != null;
        _desktopStatus.Text = working
            ? Loc.F("Đang hoạt động — Claude Desktop đã dùng Microads Office lúc {0:HH:mm dd/MM/yyyy}. Trong Claude Desktop hãy yêu cầu, ví dụ: \"Đọc văn bản đang mở trong Microads Word và tóm tắt\".",
                    "Working — Claude Desktop last used Microads Office at {0:HH:mm dd/MM/yyyy}. In Claude Desktop ask e.g. \"Read the document open in Microads Word and summarise it\".", hs!.Time.LocalDateTime)
            : connected && !ClaudeDesktop.CanEditConfig
                // Store edition: nothing was written for the user, they paste the configuration themselves
                ? Loc.T("Đã cho phép, nhưng Claude Desktop chưa khởi động Microads lần nào. Bấm \"Sao chép cấu hình\", dán vào Claude Desktop › Settings › Developer › Edit Config, lưu, rồi THOÁT HẲN Claude Desktop (biểu tượng Claude ở khay hệ thống góc phải dưới › Quit — bấm X chỉ thu nhỏ) và mở lại.",
                        "Allowed, but Claude Desktop has not started Microads yet. Click \"Copy configuration\", paste it in Claude Desktop › Settings › Developer › Edit Config, save, then QUIT Claude Desktop completely (Claude icon in the system tray › Quit — closing the window only minimises it) and open it again.")
            : connected
                ? Loc.T("Đã ghi cấu hình, Claude Desktop chưa khởi động Microads lần nào. Hãy THOÁT HẲN Claude Desktop (biểu tượng Claude ở khay hệ thống góc phải dưới › Quit — bấm X chỉ thu nhỏ) rồi mở lại. Vẫn chưa được: bấm \"Sao chép cấu hình\" và dán trong Claude Desktop › Settings › Developer › Edit Config.",
                        "Configuration written; Claude Desktop has not started Microads yet. QUIT Claude Desktop completely (Claude icon in the system tray › Quit — closing the window only minimises it) and open it again. Still nothing: click \"Copy configuration\" and paste it in Claude Desktop › Settings › Developer › Edit Config.")
                : ClaudeDesktop.IsInstalled() ? Loc.T("Claude Desktop đã cài trên máy, chưa kết nối.", "Claude Desktop is installed, not connected.")
                                               : Loc.T("Chưa thấy Claude Desktop trên máy (vẫn có thể kết nối trước rồi cài sau).", "Claude Desktop was not found (you can still connect first and install it later).");
        _desktopStatus.ForeColor = working ? Theme.OkGreen : connected ? Theme.WarnAmber : Theme.SubtleInk;
    }

    void LoadModelFor(AiProviderKind k)
    {
        if (_modelFor != AiProviderKind.None && _model.Text.Trim().Length > 0) _pendingModels[_modelFor] = _model.Text.Trim();
        _modelFor = k;
        _model.Items.Clear();
        var s = AiSettings.Load();
        _model.Text = _pendingModels.TryGetValue(k, out var m) ? m : s.ModelFor(k);
        _testResult.Text = "";
    }
    readonly Dictionary<AiProviderKind, string> _pendingModels = [];

    AiProvider? ProviderFromInputs()
    {
        var s = AiSettings.Load();
        switch (SelectedKind)
        {
            case AiProviderKind.ClaudeKey:
                var ck = _claudeKey.Text.Trim(); if (ck.Length == 0) ck = SecretStore.Read(SecretNames.ClaudeKey) ?? "";
                return ck.Length == 0 ? null : new ClaudeProvider(ck);
            case AiProviderKind.OpenAiKey:
                var ok = _openAiKey.Text.Trim(); if (ok.Length == 0) ok = SecretStore.Read(SecretNames.OpenAiKey) ?? "";
                return ok.Length == 0 ? null : new OpenAiResponsesProvider(AiDefaults.OpenAiApiBase, _ => Task.FromResult(ok), false);
            case AiProviderKind.Compatible:
                var url = _compatUrl.Text.Trim(); var key = _compatKey.Text.Trim(); if (key.Length == 0) key = SecretStore.Read(SecretNames.CompatibleKey) ?? "";
                return url.Length == 0 ? null : new CompatibleProvider(url, key);
            case AiProviderKind.ChatGpt:
                return ChatGptAuth.IsSignedIn ? AiService.Create(s, AiProviderKind.ChatGpt) : null;
        }
        return null;
    }

    void TestConnection()
    {
        var p = ProviderFromInputs();
        if (p == null)
        {
            _testResult.ForeColor = Theme.WarnAmber;
            _testResult.Text = SelectedKind == AiProviderKind.ChatGpt ? Loc.T("Hãy bấm Continue with ChatGPT để đăng nhập trước.", "Click Continue with ChatGPT to sign in first.")
                                                                    : Loc.T("Hãy nhập khóa API / địa chỉ máy chủ trước.", "Enter the API key / server address first.");
            return;
        }
        _cts?.Cancel(); _cts = new CancellationTokenSource();
        var ct = _cts.Token;
        _test.Enabled = false;
        _testResult.ForeColor = Theme.SubtleInk; _testResult.Text = Loc.T("Đang kiểm tra…", "Testing…");
        Pending = Run();
        async Task Run()
        {
            try
            {
                var models = await UiThread.On(this, Task.Run(() => p.ListModelsAsync(ct), ct));
                if (IsDisposed) return;
                var current = _model.Text.Trim();
                _model.Items.Clear();
                foreach (var m in models) _model.Items.Add(m.Id);
                if (current.Length == 0 || !models.Any(m => m.Id == current)) _model.Text = p.DefaultModel(models);
                _testResult.ForeColor = Theme.OkGreen;
                _testResult.Text = Loc.F("Kết nối được — {0} mô hình. Chọn mô hình ở trên rồi bấm Lưu.", "Connected — {0} models. Pick a model above, then click Save.", models.Count);
            }
            catch (Exception e) when (!IsDisposed)
            {
                _testResult.ForeColor = Theme.ErrRed;
                _testResult.Text = e is AiException ? e.Message : Loc.T("Lỗi: ", "Error: ") + e.Message;
            }
            finally { if (!IsDisposed) _test.Enabled = true; }
        }
    }

    void StartSignIn(bool newAccount)
    {
        if (IsDisposed) return;
        if (newAccount && ChatGptAuth.IsSignedIn &&
            !Msg.Confirm(this, Loc.T("Đăng xuất tài khoản ChatGPT hiện tại và đăng nhập tài khoản khác?", "Sign out of the current ChatGPT account and sign in with another one?"), Loc.T("Tiếp tục", "Continue"), Loc.T("Hủy", "Cancel")))
            return;
        _cts?.Cancel(); _cts = new CancellationTokenSource();
        var ct = _cts.Token;
        _signIn.Enabled = false; _signOut.Enabled = false; _otherAccount.Enabled = false;
        _cancelSignIn.Visible = true; _cancelSignIn.Enabled = true;
        _chatGptStatus.ForeColor = Theme.SubtleInk;
        Pending = Run();
        async Task Run()
        {
            try
            {
                if (newAccount) { if (ChatGptAuth.IsSignedIn) await UiThread.On(this, ChatGptAuth.SignOutAsync(ct)); ChatGptAuth.Forget(); }
                void Status(string msg)
                {
                    try { if (!IsDisposed && IsHandleCreated) BeginInvoke(() => { if (!IsDisposed) _chatGptStatus.Text = msg; }); } catch (InvalidOperationException) { }
                }
                var cred = await UiThread.On(this, Task.Run(() => ChatGptAuth.SignInAsync(newAccount, Status, ct), ct));
                if (IsDisposed) return;
                _rbChatGpt.Checked = true;
                AiSettings.Update(x => x.Provider = AiProviderKind.ChatGpt);
                RefreshAccountState();
                Activate();
                if (!cred.PlanUsage)
                    Msg.Warn(this, Loc.T("Bạn đã đăng nhập nhưng chưa đồng ý cho Microads Office dùng gói ChatGPT, nên Trợ lý AI chưa dùng được. Bấm Continue with ChatGPT lần nữa và đồng ý quyền này, hoặc dùng khóa API.",
                                         "You signed in but did not allow Microads Office to use your ChatGPT plan, so the assistant cannot run. Click Continue with ChatGPT again and allow it, or use an API key."));
                else TestConnection();
            }
            catch (OperationCanceledException) when (!IsDisposed)
            {
                // "Hủy đăng nhập": the previous sign-in (if any) is untouched
                RefreshAccountState();
                _chatGptStatus.Text = Loc.T("Đã hủy đăng nhập. ", "Sign-in cancelled. ") + _chatGptStatus.Text;
            }
            catch (OperationCanceledException) { }
            catch (Exception e) when (!IsDisposed)
            {
                _chatGptStatus.ForeColor = Theme.ErrRed;
                _chatGptStatus.Text = e is AiException ? e.Message : Loc.T("Lỗi đăng nhập: ", "Sign-in error: ") + e.Message;
            }
            finally
            {
                if (!IsDisposed) { _signIn.Enabled = true; _otherAccount.Enabled = true; _signOut.Enabled = ChatGptAuth.IsSignedIn; _cancelSignIn.Visible = false; }
            }
        }
    }

    void SignOut()
    {
        _cts?.Cancel(); _cts = new CancellationTokenSource();
        var ct = _cts.Token;
        _signOut.Enabled = false;
        Pending = Run();
        async Task Run()
        {
            bool confirmed = await UiThread.On(this, ChatGptAuth.SignOutAsync(ct));
            if (IsDisposed) return;
            RefreshAccountState();
            AiSettings.RaiseChanged();
            if (!confirmed) Msg.Warn(this, Loc.T("Đã đăng xuất trên máy này, nhưng chưa xác nhận được việc thu hồi phiên với ChatGPT (mất mạng?). Bạn có thể ngắt kết nối Microads Office trong ChatGPT › Settings.",
                                                 "Signed out on this computer, but ChatGPT did not confirm the session was revoked (offline?). You can disconnect Microads Office in ChatGPT › Settings."));
        }
    }

    /// <summary>Connected = Claude Desktop knows the server AND the user allows access (the Store edition can only check the latter reliably).</summary>
    static bool DesktopConnected() => AiSettings.Load().AllowClaudeDesktop && (ClaudeDesktop.IsConnected() || !ClaudeDesktop.CanEditConfig);

    void ToggleDesktop()
    {
        ClaudeDesktop.ClearHandshake();
        try
        {
            if (DesktopConnected())
            {
                if (ClaudeDesktop.CanEditConfig) ClaudeDesktop.Disconnect();
                AiSettings.Update(x => x.AllowClaudeDesktop = false);
                if (!ClaudeDesktop.CanEditConfig)
                    Msg.Info(this, Loc.T("Đã ngắt: Microads Office từ chối mọi yêu cầu từ Claude Desktop. Muốn gỡ hẳn, mở Claude Desktop › Settings › Developer › Edit Config và xóa mục \"microads-office\".",
                                         "Disconnected: Microads Office now refuses every request from Claude Desktop. To remove it completely, open Claude Desktop › Settings › Developer › Edit Config and delete the \"microads-office\" entry."));
            }
            else
            {
                if (!Msg.Confirm(this, Loc.T("Cho phép Claude Desktop đọc và chỉnh sửa văn bản, bảng tính đang mở trong Microads Office khi bạn yêu cầu trong Claude Desktop?\n\nClaude Desktop sẽ hỏi bạn trước mỗi thao tác. Nội dung được gửi tới Anthropic theo tài khoản Claude của bạn.",
                                               "Allow Claude Desktop to read and edit the documents and workbooks open in Microads Office when you ask it in Claude Desktop?\n\nClaude Desktop asks you before each action. Content is sent to Anthropic under your Claude account."),
                                 Loc.T("Kết nối", "Connect"), Loc.T("Hủy", "Cancel"))) return;
                if (ClaudeDesktop.CanEditConfig) ClaudeDesktop.Connect();
                else
                {
                    try { Clipboard.SetText(ClaudeDesktop.ConfigSnippet()); } catch { }
                    Msg.Info(this, Loc.T("Đã sao chép cấu hình. Làm tiếp trong Claude Desktop:\n1. Settings › Developer › Edit Config (mở tệp claude_desktop_config.json).\n2. Dán (Ctrl+V) đè lên nội dung — hoặc, nếu tệp đã có \"mcpServers\", chỉ thêm mục \"microads-office\".\n3. Lưu tệp, thoát hẳn rồi mở lại Claude Desktop.",
                                         "Configuration copied. Continue in Claude Desktop:\n1. Settings › Developer › Edit Config (opens claude_desktop_config.json).\n2. Paste (Ctrl+V) over the content — or, if the file already has \"mcpServers\", add only the \"microads-office\" entry.\n3. Save, quit Claude Desktop completely and open it again."),
                             Loc.T("Kết nối Claude Desktop", "Connect Claude Desktop"));
                }
                AiSettings.Update(x => x.AllowClaudeDesktop = true);
            }
        }
        catch (Exception e) { Msg.Error(this, Loc.T("Không cập nhật được cấu hình Claude Desktop: ", "Could not update the Claude Desktop configuration: ") + e.Message); }
        RefreshAccountState();
    }

    bool SaveAll()
    {
        var kind = SelectedKind;
        if (_claudeKey.Text.Trim().Length > 0) SecretStore.Write(SecretNames.ClaudeKey, _claudeKey.Text.Trim());
        if (_openAiKey.Text.Trim().Length > 0) SecretStore.Write(SecretNames.OpenAiKey, _openAiKey.Text.Trim());
        if (_compatKey.Text.Trim().Length > 0) SecretStore.Write(SecretNames.CompatibleKey, _compatKey.Text.Trim());
        var url = _compatUrl.Text.Trim();
        // an API-key way without a key would leave the assistant silently "not connected"
        if (kind == AiProviderKind.ClaudeKey && !SecretStore.Has(SecretNames.ClaudeKey))
        { Msg.Warn(this, Loc.T("Hãy nhập khóa API Claude (bấm \"Lấy khóa\" để tạo khóa trong Claude Console), hoặc chọn cách kết nối khác.", "Enter a Claude API key (click \"Get a key\" to create one in the Claude Console), or choose another way to connect.")); _claudeKey.Focus(); return false; }
        if (kind == AiProviderKind.OpenAiKey && !SecretStore.Has(SecretNames.OpenAiKey))
        { Msg.Warn(this, Loc.T("Hãy nhập khóa API OpenAI (bấm \"Lấy khóa\"), hoặc chọn cách kết nối khác.", "Enter an OpenAI API key (click \"Get a key\"), or choose another way to connect.")); _openAiKey.Focus(); return false; }
        if (kind == AiProviderKind.Compatible)
        {
            if (!Uri.TryCreate(url, UriKind.Absolute, out var u) || (u.Scheme != "http" && u.Scheme != "https"))
            { Msg.Warn(this, Loc.T("Địa chỉ máy chủ không hợp lệ (ví dụ http://127.0.0.1:11434/v1).", "Invalid server address (e.g. http://127.0.0.1:11434/v1).")); return false; }
            if (_model.Text.Trim().Length == 0)
            { Msg.Warn(this, Loc.T("Hãy nhập hoặc chọn tên mô hình (bấm Kiểm tra kết nối để tải danh sách).", "Enter or choose a model name (click Test connection to load the list).")); return false; }
        }
        _pendingModels[kind] = _model.Text.Trim();
        AiSettings.Update(x =>
        {
            x.Provider = kind;
            x.CompatibleBaseUrl = url.Length > 0 ? url : x.CompatibleBaseUrl;
            foreach (var (k, m) in _pendingModels) x.SetModel(k, m);
            x.Instructions = _instructions.Text.Trim();
        });
        return true;
    }
}
