using MicroadsOffice.Core.Office;
using MicroadsOffice.Core;
using MicroadsOffice.Core.Pdf;
using MicroadsOffice.Core.Pdfium;
using MicroadsOffice.Core.Signing;
using MicroadsOffice.UI;
using MicroadsOffice.Viewer;

namespace MicroadsOffice.Views;

public enum NavKind { Thumbnails, Bookmarks, Signatures, Search }

/// <summary>One open document: navigation pane, signature banner and the page viewer.</summary>
public sealed class DocumentView : UserControl
{
    public string? FilePath { get; private set; }
    public string DisplayName { get; private set; }
    public byte[] Bytes { get; private set; }
    public PdfiumDocument Pdf { get; private set; }
    public string? Password { get; }
    public bool Modified { get; private set; }
    public PdfViewer Viewer { get; }
    public List<SignatureValidationResult> Signatures { get; private set; } = new();
    public bool IsValidating { get; private set; }

    public event EventHandler? StateChanged;
    public event Action<int>? SignaturePropertiesRequested;
    public event Action<SignatureValidationResult>? ViewSignedVersionRequested;
    public event EventHandler? ValidationCompleted;

    private readonly SplitContainer _split;
    private readonly ToolStrip _navStrip;
    private readonly Panel _navContent;
    private readonly ThumbnailList _thumbs;
    private readonly TreeView _bookmarks;
    private readonly TreeView _sigTree;
    private readonly Panel _sigPanel;
    private readonly Panel _searchPanel;
    private readonly Panel _banner;
    private readonly Label _bannerText;
    private readonly PictureBox _bannerIcon;
    private readonly Label _navTitle;
    public TextBox SearchBox { get; }
    private readonly CheckBox _chkCase, _chkWord;
    private readonly ListBox _results;
    private readonly Label _searchStatus;
    private List<PdfSearchHit> _hits = new();
    private int _hitIndex = -1;
    private CancellationTokenSource? _searchCts;
    private NavKind _nav = NavKind.Thumbnails;
    private readonly ImageList _sigImages;
    private int _validationGen;

    public NavKind CurrentNav => _nav;
    public bool NavVisible => !_split.Panel1Collapsed;
    public int SearchHitCount => _hits.Count;
    public int SearchHitIndex => _hitIndex;

    public DocumentView(string? path, string displayName, byte[] bytes, PdfiumDocument pdf, string? password)
    {
        FilePath = path;
        DisplayName = displayName;
        Bytes = bytes;
        Pdf = pdf;
        Password = password;
        Dock = DockStyle.Fill;
        BackColor = Theme.ViewerBack;

        Viewer = new PdfViewer { Dock = DockStyle.Fill, Name = "pdfViewer" };
        _banner = new Panel { Dock = DockStyle.Top, Height = Dpi.Scale(34), Visible = false, Padding = new Padding(Dpi.Scale(8), 0, Dpi.Scale(8), 0), Name = "sigBanner" };
        _bannerIcon = new PictureBox { Dock = DockStyle.Left, Width = Dpi.Scale(26), SizeMode = PictureBoxSizeMode.CenterImage };
        _bannerText = new Label { Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft, Font = Theme.UiFontBold, Name = "lblSigBanner" };
        var bannerBtn = new Button
        {
            Text = Loc.T("Bảng chữ ký", "Signature panel"), Name = "btnBannerSignatures", Dock = DockStyle.Right, AutoSize = true, FlatStyle = FlatStyle.Flat,
            BackColor = Color.White, Font = Theme.UiFont, Cursor = Cursors.Hand
        };
        bannerBtn.FlatAppearance.BorderColor = Theme.Border;
        bannerBtn.Click += (_, _) => ShowNav(NavKind.Signatures);
        _banner.Controls.Add(_bannerText);
        _banner.Controls.Add(_bannerIcon);
        _banner.Controls.Add(bannerBtn);

        var viewerArea = new Panel { Dock = DockStyle.Fill };
        viewerArea.Controls.Add(Viewer);
        viewerArea.Controls.Add(_banner);

        // ---- navigation pane
        _navStrip = new ToolStrip
        {
            Dock = DockStyle.Left, LayoutStyle = ToolStripLayoutStyle.VerticalStackWithOverflow, GripStyle = ToolStripGripStyle.Hidden,
            BackColor = Theme.TabStripBack, Renderer = new RibbonRenderer(false), ImageScalingSize = new Size(Dpi.Scale(20), Dpi.Scale(20)),
            Padding = new Padding(Dpi.Scale(2)), Name = "navStrip", AccessibleName = Loc.T("Ngăn điều hướng", "Navigation pane")
        };
        AddNavButton("navThumbs", Loc.T("Hình thu nhỏ trang", "Page thumbnails"), Icons.Thumbnails, NavKind.Thumbnails);
        AddNavButton("navBookmarks", Loc.T("Dấu trang (mục lục)", "Bookmarks (outline)"), Icons.Bookmarks, NavKind.Bookmarks);
        AddNavButton("navSignatures", Loc.T("Bảng chữ ký số", "Signature panel"), Icons.Sign, NavKind.Signatures);
        AddNavButton("navSearch", Loc.T("Tìm kiếm", "Search"), Icons.Search, NavKind.Search);

        _navTitle = new Label
        {
            Dock = DockStyle.Top, Height = Dpi.Scale(28), Font = Theme.UiFontBold, ForeColor = Theme.Ink, TextAlign = ContentAlignment.MiddleLeft,
            Padding = new Padding(Dpi.Scale(8), 0, 0, 0), BackColor = Theme.PanelBack
        };
        _navContent = new Panel { Dock = DockStyle.Fill, BackColor = Theme.PanelBack };

        _thumbs = new ThumbnailList { Dock = DockStyle.Fill, Name = "thumbnailList" };
        _thumbs.PageClicked += p => Viewer.GoToPage(p);

        _bookmarks = new TreeView
        {
            Dock = DockStyle.Fill, BorderStyle = BorderStyle.None, BackColor = Theme.PanelBack, Name = "bookmarkTree",
            HideSelection = false, ShowLines = false, FullRowSelect = true, Font = Theme.UiFont, ItemHeight = Dpi.Scale(24)
        };
        _bookmarks.AfterSelect += (_, e) =>
        {
            if (e.Node?.Tag is PdfBookmark b && b.PageIndex >= 0 && e.Action != TreeViewAction.Unknown) Viewer.GoToPage(b.PageIndex, b.Y);
        };
        _bookmarks.NodeMouseClick += (_, e) =>
        {
            if (e.Node?.Tag is PdfBookmark b && b.PageIndex >= 0) Viewer.GoToPage(b.PageIndex, b.Y);
        };

        _sigImages = new ImageList { ImageSize = new Size(Dpi.Scale(16), Dpi.Scale(16)), ColorDepth = ColorDepth.Depth32Bit };
        _sigImages.Images.Add("valid", Icons.Small(Icons.Check, Theme.OkGreen));
        _sigImages.Images.Add("unknown", Icons.Small(Icons.Unknown, Theme.WarnAmber));
        _sigImages.Images.Add("invalid", Icons.Small(Icons.Error, Theme.ErrRed));
        _sigImages.Images.Add("empty", Icons.Small(Icons.Sign, Theme.SubtleInk));
        _sigImages.Images.Add("info", Icons.Small(Icons.Info, Theme.SubtleInk));
        _sigTree = new TreeView
        {
            Dock = DockStyle.Fill, BorderStyle = BorderStyle.None, BackColor = Theme.PanelBack, ImageList = _sigImages, Name = "signatureTree",
            HideSelection = false, Font = Theme.UiFont, ItemHeight = Dpi.Scale(22), ShowNodeToolTips = true
        };
        _sigTree.NodeMouseDoubleClick += (_, e) => { if (TagIndex(e.Node) is int i) SignaturePropertiesRequested?.Invoke(i); };
        _sigTree.AfterSelect += (_, e) =>
        {
            if (TagIndex(e.Node) is int i && i < Signatures.Count && Signatures[i].PageIndex >= 0 && Signatures[i].Rect.Width > 0)
                Viewer.EnsureVisible(Signatures[i].PageIndex, Signatures[i].Rect);
        };
        var sigMenu = new ContextMenuStrip { Renderer = new RibbonRenderer(false) };
        sigMenu.Items.Add(Loc.T("Xem thuộc tính chữ ký...", "Signature properties..."), Icons.Small(Icons.Properties), (_, _) => { if (TagIndex(_sigTree.SelectedNode) is int i) SignaturePropertiesRequested?.Invoke(i); });
        sigMenu.Items.Add(Loc.T("Đi tới chữ ký", "Go to signature"), Icons.Small(Icons.GoTo), (_, _) =>
        {
            if (TagIndex(_sigTree.SelectedNode) is int i && Signatures[i].PageIndex >= 0) Viewer.EnsureVisible(Signatures[i].PageIndex, Signatures[i].Rect);
        });
        sigMenu.Items.Add(Loc.T("Xem phiên bản đã ký", "View signed version"), Icons.Small(Icons.Document), (_, _) =>
        {
            if (TagIndex(_sigTree.SelectedNode) is int i && Signatures[i].IsSigned) ViewSignedVersionRequested?.Invoke(Signatures[i]);
        });
        sigMenu.Items.Add(Loc.T("Xác thực lại tất cả chữ ký", "Re-validate all signatures"), Icons.Small(Icons.Shield), (_, _) => Revalidate());
        _sigTree.ContextMenuStrip = sigMenu;
        _sigTree.NodeMouseClick += (_, e) => { if (e.Button == MouseButtons.Right) _sigTree.SelectedNode = e.Node; };
        var sigTop = new FlowLayoutPanel { Dock = DockStyle.Top, Height = Dpi.Scale(34), Padding = new Padding(Dpi.Scale(4)), BackColor = Theme.PanelBack };
        var btnValidate = new Button { Text = Loc.T("Xác thực tất cả", "Validate all"), Name = "btnValidateAllPanel", AutoSize = true, FlatStyle = FlatStyle.Flat, BackColor = Color.White };
        btnValidate.FlatAppearance.BorderColor = Theme.Border;
        btnValidate.Click += (_, _) => Revalidate();
        var btnProps = new Button { Text = Loc.T("Thuộc tính", "Properties"), Name = "btnSigPropsPanel", AutoSize = true, FlatStyle = FlatStyle.Flat, BackColor = Color.White };
        btnProps.FlatAppearance.BorderColor = Theme.Border;
        btnProps.Click += (_, _) =>
        {
            int? i = TagIndex(_sigTree.SelectedNode) ?? (Signatures.Count > 0 ? 0 : null);
            if (i is int k) SignaturePropertiesRequested?.Invoke(k);
        };
        sigTop.Controls.Add(btnValidate);
        sigTop.Controls.Add(btnProps);
        _sigPanel = new Panel { Dock = DockStyle.Fill };
        _sigPanel.Controls.Add(_sigTree);
        _sigPanel.Controls.Add(sigTop);

        // search panel
        _searchPanel = new Panel { Dock = DockStyle.Fill, Padding = new Padding(Dpi.Scale(6)) };
        SearchBox = new TextBox { Dock = DockStyle.Top, Name = "txtSearch", PlaceholderText = Loc.T("Nhập từ cần tìm rồi nhấn Enter", "Type the text to find and press Enter"), Font = new Font("Segoe UI", 10f) };
        SearchBox.KeyDown += (_, e) =>
        {
            if (e.KeyCode == Keys.Enter) { e.SuppressKeyPress = true; if (e.Shift) PrevHit(); else StartSearch(); }
        };
        // Row 1: buttons (always visible), rows 2-3: options, then status and results.
        var btnRow = new FlowLayoutPanel { Dock = DockStyle.Top, Height = Dpi.Scale(36), FlowDirection = FlowDirection.LeftToRight, WrapContents = false, Padding = new Padding(0, Dpi.Scale(4), 0, 0) };
        var btnFind = new Button { Text = Loc.T("Tìm", "Find"), Name = "btnSearchGo", Width = Dpi.Scale(64), Height = Dpi.Scale(28), Margin = new Padding(0, 0, Dpi.Scale(4), 0) };
        btnFind.Click += (_, _) => StartSearch();
        var btnPrev = new Button { Text = "◀", Name = "btnSearchPrev", Width = Dpi.Scale(34), Height = Dpi.Scale(28), AccessibleName = Loc.T("Kết quả trước", "Previous result"), Margin = new Padding(0, 0, Dpi.Scale(4), 0) };
        btnPrev.Click += (_, _) => PrevHit();
        var btnNext = new Button { Text = "▶", Name = "btnSearchNext", Width = Dpi.Scale(34), Height = Dpi.Scale(28), AccessibleName = Loc.T("Kết quả sau", "Next result"), Margin = Padding.Empty };
        btnNext.Click += (_, _) => NextHit();
        btnRow.Controls.AddRange([btnFind, btnPrev, btnNext]);
        _chkCase = new CheckBox { Text = Loc.T("Phân biệt hoa/thường", "Match case"), Name = "chkMatchCase", Dock = DockStyle.Top, Height = Dpi.Scale(24) };
        _chkWord = new CheckBox { Text = Loc.T("Chỉ khớp nguyên từ", "Whole words only"), Name = "chkWholeWord", Dock = DockStyle.Top, Height = Dpi.Scale(24) };
        var opts = new Panel { Dock = DockStyle.Top, Height = Dpi.Scale(48) };
        opts.Controls.Add(_chkWord);
        opts.Controls.Add(_chkCase);
        _searchStatus = new Label { Dock = DockStyle.Top, Height = Dpi.Scale(24), ForeColor = Theme.SubtleInk, Name = "lblSearchStatus", TextAlign = ContentAlignment.MiddleLeft };
        _results = new ListBox { Dock = DockStyle.Fill, IntegralHeight = false, BorderStyle = BorderStyle.None, Name = "lstSearchResults", Font = Theme.UiFont };
        _results.SelectedIndexChanged += (_, _) =>
        {
            if (_results.SelectedIndex >= 0 && _results.SelectedIndex != _hitIndex) { _hitIndex = _results.SelectedIndex; Viewer.SetSearchHits(_hits, _hitIndex); }
        };
        _searchPanel.Controls.Add(_results);
        _searchPanel.Controls.Add(_searchStatus);
        _searchPanel.Controls.Add(opts);
        _searchPanel.Controls.Add(btnRow);
        _searchPanel.Controls.Add(SearchBox);

        var navHost = new Panel { Dock = DockStyle.Fill };
        navHost.Controls.Add(_navContent);
        navHost.Controls.Add(_navTitle);

        _split = new SplitContainer
        {
            Dock = DockStyle.Fill, FixedPanel = FixedPanel.Panel1, SplitterWidth = Dpi.Scale(4), BackColor = Theme.Border, Name = "docSplit"
        };
        _split.Panel1.BackColor = Theme.PanelBack;
        _split.Panel2.BackColor = Theme.ViewerBack;
        _split.Panel1.Controls.Add(navHost);
        _split.Panel2.Controls.Add(viewerArea);
        Controls.Add(_split);
        Controls.Add(_navStrip);

        Viewer.PageChanged += (_, _) => { _thumbs.SetCurrent(Viewer.CurrentPage); StateChanged?.Invoke(this, EventArgs.Empty); };
        Viewer.ZoomChanged += (_, _) => StateChanged?.Invoke(this, EventArgs.Empty);
        Viewer.SignatureClicked += i => SignaturePropertiesRequested?.Invoke(i);
        Load += (_, _) =>
        {
            _split.Panel1MinSize = Dpi.Scale(140);
            BeginInvoke(() => { if (_split.Width > Dpi.Scale(500)) _split.SplitterDistance = Dpi.Scale(250); });
            LoadIntoViews(keepView: false);
            bool hasOutline = _bookmarks.Nodes.Count > 0;
            ShowNav(hasOutline ? NavKind.Bookmarks : NavKind.Thumbnails);
            Revalidate();
        };
    }

    private void AddNavButton(string name, string tip, string glyph, NavKind kind)
    {
        var b = new ToolStripButton(Icons.Get(glyph, Dpi.Scale(20), Theme.Ink))
        {
            Name = name, ToolTipText = tip, AccessibleName = tip, DisplayStyle = ToolStripItemDisplayStyle.Image,
            ImageScaling = ToolStripItemImageScaling.None, Margin = new Padding(0, Dpi.Scale(2), 0, Dpi.Scale(2)), Tag = kind
        };
        b.Click += (_, _) =>
        {
            if (_nav == kind && NavVisible) SetNavVisible(false);
            else ShowNav(kind);
        };
        _navStrip.Items.Add(b);
    }

    private static int? TagIndex(TreeNode? n)
    {
        while (n != null) { if (n.Tag is int i) return i; n = n.Parent; }
        return null;
    }

    public void ShowNav(NavKind kind)
    {
        _nav = kind;
        SetNavVisible(true);
        _navContent.Controls.Clear();
        (Control c, string title) = kind switch
        {
            NavKind.Thumbnails => ((Control)_thumbs, Loc.T("Hình thu nhỏ", "Thumbnails")),
            NavKind.Bookmarks => (_bookmarks, Loc.T("Dấu trang", "Bookmarks")),
            NavKind.Signatures => (_sigPanel, Loc.T("Chữ ký số", "Digital signatures")),
            _ => (_searchPanel, Loc.T("Tìm kiếm", "Search"))
        };
        _navTitle.Text = title;
        _navContent.Controls.Add(c);
        foreach (ToolStripItem it in _navStrip.Items)
            if (it is ToolStripButton b) b.Checked = (NavKind)b.Tag! == kind;
        if (kind == NavKind.Search) { SearchBox.Focus(); SearchBox.SelectAll(); }
        if (kind == NavKind.Bookmarks && _bookmarks.Nodes.Count == 0)
        {
            _bookmarks.Nodes.Add(new TreeNode(Loc.T("(Tài liệu không có dấu trang)", "(The document has no bookmarks)")) { ForeColor = Theme.SubtleInk });
        }
        StateChanged?.Invoke(this, EventArgs.Empty);
    }

    public void SetNavVisible(bool visible)
    {
        _split.Panel1Collapsed = !visible;
        if (!visible) foreach (ToolStripItem it in _navStrip.Items) if (it is ToolStripButton b) b.Checked = false;
        StateChanged?.Invoke(this, EventArgs.Empty);
    }

    private void LoadIntoViews(bool keepView)
    {
        Viewer.SetDocument(Pdf, keepView);
        _thumbs.SetDocument(Pdf);
        _bookmarks.BeginUpdate();
        _bookmarks.Nodes.Clear();
        try
        {
            void Add(TreeNodeCollection nodes, IReadOnlyList<PdfBookmark> list)
            {
                foreach (var b in list)
                {
                    var n = new TreeNode(string.IsNullOrWhiteSpace(b.Title) ? Loc.T("(không tên)", "(untitled)") : b.Title) { Tag = b };
                    nodes.Add(n);
                    Add(n.Nodes, b.Children);
                }
            }
            Add(_bookmarks.Nodes, Pdf.GetBookmarks());
            if (_bookmarks.Nodes.Count < 30) foreach (TreeNode n in _bookmarks.Nodes) n.Expand();
        }
        catch { }
        _bookmarks.EndUpdate();
    }

    /// <summary>Replaces the document content (after signing / stamping) keeping the current view.</summary>
    public void ReplaceContent(byte[] bytes, string? newPath, bool modified)
    {
        var newPdf = PdfiumDocument.Load(bytes, Password);
        var old = Pdf;
        Bytes = bytes;
        Pdf = newPdf;
        if (newPath != null)
        {
            FilePath = newPath;
            DisplayName = Path.GetFileName(newPath);
        }
        Modified = modified;
        LoadIntoViews(keepView: true);
        old.Dispose();
        ClearSearch();
        Revalidate();
        StateChanged?.Invoke(this, EventArgs.Empty);
    }

    public void MarkSaved(string path)
    {
        FilePath = path;
        DisplayName = Path.GetFileName(path);
        Modified = false;
        StateChanged?.Invoke(this, EventArgs.Empty);
    }

    // ------------------------------------------------------------------ signatures

    public void Revalidate()
    {
        int gen = ++_validationGen;
        var bytes = Bytes;
        IsValidating = true;
        _sigTree.Nodes.Clear();
        _sigTree.Nodes.Add(new TreeNode(Loc.T("Đang xác thực chữ ký...", "Validating signatures..."), 4, 4));
        var settings = AppSettings.Load();
        var opts = new VerifyOptions { OnlineRevocation = settings.OnlineRevocationCheck };
        Task.Run(() =>
        {
            try { return SignatureVerifier.Verify(bytes, opts); }
            catch { return new List<SignatureValidationResult>(); }
        }).ContinueWith(t =>
        {
            if (IsDisposed || gen != _validationGen) return;
            try { BeginInvoke(() => { if (gen == _validationGen) ApplyValidation(t.Result); }); } catch { }
        });
    }

    private void ApplyValidation(List<SignatureValidationResult> results)
    {
        IsValidating = false;
        Signatures = results;
        _sigTree.BeginUpdate();
        _sigTree.Nodes.Clear();
        var signed = results.Where(r => r.IsSigned).ToList();
        if (results.Count == 0)
        {
            _sigTree.Nodes.Add(new TreeNode(Loc.T("Tài liệu chưa có chữ ký số.", "The document has no digital signatures yet."), 4, 4));
        }
        for (int i = 0; i < results.Count; i++)
        {
            var r = results[i];
            string img = r.Status switch
            {
                SignatureStatus.Valid => "valid",
                SignatureStatus.Unknown => "unknown",
                SignatureStatus.Invalid => "invalid",
                _ => "empty"
            };
            string title = r.IsSigned ? Loc.T($"Rev.{r.Revision}: Ký bởi {r.SignerName}", $"Rev.{r.Revision}: Signed by {r.SignerName}") : Loc.T($"{r.FieldName} (chưa ký)", $"{r.FieldName} (unsigned)");
            var node = new TreeNode(title) { Tag = i, ImageKey = img, SelectedImageKey = img, ToolTipText = r.StatusText };
            if (r.IsSigned)
            {
                void Add(string text) => node.Nodes.Add(new TreeNode(text) { ImageKey = "info", SelectedImageKey = "info" });
                Add(r.StatusText);
                if (r.ClaimedSigningTime is { } t) Add(Loc.T("Thời gian ký: ", "Signing time: ") + t.ToString("dd/MM/yyyy HH:mm:ss zzz"));
                if (r.HasTimestamp && r.TimestampTime is { } ts) Add(Loc.T("Dấu thời gian: ", "Timestamp: ") + ts.ToString("dd/MM/yyyy HH:mm:ss zzz") + (r.TimestampValid ? Loc.T(" (hợp lệ)", " (valid)") : Loc.T(" (lỗi)", " (error)")));
                if (!string.IsNullOrEmpty(r.Reason)) Add(Loc.T("Lý do: ", "Reason: ") + r.Reason);
                if (!string.IsNullOrEmpty(r.Location)) Add(Loc.T("Vị trí: ", "Location: ") + r.Location);
                if (!string.IsNullOrEmpty(r.Application)) Add(Loc.T("Nền tảng ký: ", "Signing platform: ") + r.Application);
                Add(r.CoversWholeDocument ? Loc.T("Chữ ký bao phủ toàn bộ tài liệu", "The signature covers the whole document")
                    : r.ChangesAfterAreSignaturesOnly ? Loc.T("Sau chữ ký này chỉ có các chữ ký bổ sung (được phép)", "Only additional signatures follow this one (permitted)") : Loc.T("Tài liệu đã thay đổi sau khi ký", "The document changed after signing"));
                Add(r.IntegrityValid ? Loc.T("Tài liệu không bị sửa đổi kể từ khi ký", "The document has not been modified since signing") : Loc.T("Tài liệu đã bị sửa đổi/hỏng", "The document has been modified/corrupted"));
                Add(r.CertificateTrusted ? Loc.T("Danh tính người ký hợp lệ", "Signer identity is valid") : Loc.T("Danh tính người ký chưa xác thực (CA chưa được tin cậy)", "Signer identity not verified (CA not trusted)"));
                Add(Loc.T($"Trang {r.PageIndex + 1}, trường: {r.FieldName}", $"Page {r.PageIndex + 1}, field: {r.FieldName}"));
            }
            _sigTree.Nodes.Add(node);
        }
        _sigTree.EndUpdate();
        if (_sigTree.Nodes.Count is > 0 and < 4) foreach (TreeNode n in _sigTree.Nodes) if (n.Tag is int) n.Expand();

        Viewer.SetSignatureMarks(
            results.Select((r, i) => (r, i)).Where(x => x.r.IsSigned && x.r.PageIndex >= 0)
                .Select(x => new SignatureMark(x.i, x.r.PageIndex, x.r.Rect, x.r.Status, x.r.FieldName)).ToList(),
            ReadEmptyFields());

        // Banner like Foxit / Adobe
        if (signed.Count == 0) _banner.Visible = false;
        else
        {
            _banner.Visible = true;
            if (signed.Any(s => s.Status == SignatureStatus.Invalid))
                SetBanner(Theme.ErrRed, Color.FromArgb(253, 236, 236), Icons.Error, Loc.T("Có ít nhất một chữ ký KHÔNG hợp lệ — tài liệu có thể đã bị sửa đổi.", "At least one signature is NOT valid — the document may have been modified."));
            else if (signed.All(s => s.Status == SignatureStatus.Valid))
                SetBanner(Theme.OkGreen, Color.FromArgb(232, 245, 236), Icons.Check, Loc.T($"Đã ký và tất cả chữ ký đều hợp lệ ({signed.Count} chữ ký).", $"Signed, and all signatures are valid ({signed.Count} signature(s))."));
            else
                SetBanner(Theme.WarnAmber, Color.FromArgb(255, 246, 225), Icons.Unknown,
                    Loc.T($"Tài liệu có {signed.Count} chữ ký toàn vẹn; có chữ ký chưa xác thực được danh tính hoặc tài liệu có thay đổi sau khi ký.", $"The document has {signed.Count} intact signature(s); some signer identities could not be verified or the document changed after signing."));
        }
        StateChanged?.Invoke(this, EventArgs.Empty);
        ValidationCompleted?.Invoke(this, EventArgs.Empty);
    }

    private List<SignatureFieldInfo> ReadEmptyFields()
    {
        try { return SignatureFields.Find(PdfFile.Load(Bytes)).Where(f => !f.IsSigned).ToList(); }
        catch { return new(); }
    }

    private void SetBanner(Color fg, Color bg, string glyph, string text)
    {
        _banner.BackColor = bg;
        _bannerText.ForeColor = fg;
        _bannerText.Text = text;
        _bannerIcon.Image = Icons.Get(glyph, Dpi.Scale(18), fg);
    }

    // ------------------------------------------------------------------ search

    public void FocusSearch(string? text = null)
    {
        ShowNav(NavKind.Search);
        if (!string.IsNullOrEmpty(text)) SearchBox.Text = text;
        SearchBox.Focus();
        SearchBox.SelectAll();
    }

    public void StartSearch()
    {
        var q = SearchBox.Text.Trim();
        _searchCts?.Cancel();
        ClearSearch();
        if (q.Length == 0) return;
        var cts = _searchCts = new CancellationTokenSource();
        bool mc = _chkCase.Checked, ww = _chkWord.Checked;
        var pdf = Pdf;
        _searchStatus.Text = Loc.T("Đang tìm...", "Searching...");
        Task.Run(() =>
        {
            var all = new List<PdfSearchHit>();
            for (int p = 0; p < pdf.PageCount && !cts.IsCancellationRequested; p++)
            {
                try { all.AddRange(pdf.Search(p, q, mc, ww)); } catch { }
                if (all.Count >= 5000) break;
            }
            return all;
        }, cts.Token).ContinueWith(t =>
        {
            if (cts.IsCancellationRequested || t.IsFaulted || IsDisposed) return;
            try
            {
                BeginInvoke(() =>
                {
                    _hits = t.Result;
                    _results.BeginUpdate();
                    _results.Items.Clear();
                    foreach (var h in _hits) _results.Items.Add($"Tr.{h.PageIndex + 1}: …{h.Context}…");
                    _results.EndUpdate();
                    _searchStatus.Text = _hits.Count == 0 ? Loc.T($"Không tìm thấy \"{q}\".", $"\"{q}\" was not found.") : Loc.T($"Tìm thấy {_hits.Count} kết quả.", $"Found {_hits.Count} result(s).");
                    _hitIndex = _hits.Count > 0 ? 0 : -1;
                    if (_hitIndex >= 0) _results.SelectedIndex = 0;
                    Viewer.SetSearchHits(_hits, _hitIndex);
                    StateChanged?.Invoke(this, EventArgs.Empty);
                });
            }
            catch { }
        });
    }

    public void NextHit()
    {
        if (_hits.Count == 0) { StartSearch(); return; }
        _hitIndex = (_hitIndex + 1) % _hits.Count;
        _results.SelectedIndex = _hitIndex;
        Viewer.SetSearchHits(_hits, _hitIndex);
    }

    public void PrevHit()
    {
        if (_hits.Count == 0) return;
        _hitIndex = (_hitIndex - 1 + _hits.Count) % _hits.Count;
        _results.SelectedIndex = _hitIndex;
        Viewer.SetSearchHits(_hits, _hitIndex);
    }

    public void ClearSearch()
    {
        _hits = new();
        _hitIndex = -1;
        _results.Items.Clear();
        _searchStatus.Text = "";
        Viewer.SetSearchHits(_hits, -1);
    }

    protected override void Dispose(bool disposing)
    {
        if (disposing)
        {
            _searchCts?.Cancel();
            _validationGen++;
            Viewer.SetDocument(null);
            _thumbs.SetDocument(null);
            Pdf.Dispose();
        }
        base.Dispose(disposing);
    }
}
