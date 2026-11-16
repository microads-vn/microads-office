using System.Diagnostics;
using MicroadsOffice.Core;
using MicroadsOffice.Core.Office;
using MicroadsOffice.UI;
using MicroadsOffice.Word.Editor;
using MicroadsOffice.Word.IO;
using MicroadsOffice.Word.Model;

namespace MicroadsOffice.Word.UI;

/// <summary>Microads Office Word main window (one document per window, like Word).</summary>
public sealed partial class WordMainForm : Form, IOfficeAppWindow
{
    private static readonly List<WordMainForm> Windows = new();
    private static WordMainForm? _primary;

    private readonly WordEditor _ed;
    private readonly RibbonBar _ribbon;
    private readonly RulerControl _ruler;
    private readonly Panel _editorHost;
    private readonly Panel _navPane;
    private readonly ListBox _navList;
    private readonly Backstage _backstage;
    private readonly SpellService _spell;
    private readonly System.Windows.Forms.Timer _autoSave = new() { Interval = 120_000 };
    private readonly System.Windows.Forms.Timer _navTimer = new() { Interval = 600 };
    private readonly string _recoverId = Guid.NewGuid().ToString("N");
    private string? _path;
    private string _untitledName;
    private FindReplaceDialog? _find;
    private bool _fullScreen;
    private FormWindowState _prevState;
    private FormBorderStyle _prevBorder;
    private bool _forceClose;
    private static int _untitledCounter;

    public WordEditor Editor => _ed;
    public string? FilePath => _path;
    public Backstage BackstageView => _backstage;
    public RibbonBar Ribbon => _ribbon;
    public static IReadOnlyList<WordMainForm> OpenWindows => Windows;
    public bool IsPristine => _path == null && !_ed.Modified && _ed.Document.PlainText.Length == 0 && _ed.Document.Blocks.Count <= 1;

    public WordMainForm(string[] files) : this(files, primary: _primary == null) { }

    private WordMainForm(string[] files, bool primary)
    {
        if (primary) _primary = this;
        Windows.Add(this);
        _untitledName = Loc.T("Tài liệu", "Document") + (++_untitledCounter);
        Name = "WordMainForm";
        Font = Theme.UiFont;
        AutoScaleMode = AutoScaleMode.None;
        KeyPreview = true;
        AllowDrop = true;
        BackColor = Theme.RibbonBack;
        Icon = OfficeShell.AppIcon;
        var wa = Screen.PrimaryScreen?.WorkingArea ?? new Rectangle(0, 0, 1366, 728);
        Size = new Size(Math.Min(Dpi.Scale(1300), (int)(wa.Width * 0.92)), Math.Min(Dpi.Scale(860), (int)(wa.Height * 0.92)));
        MinimumSize = new Size(Dpi.Scale(700), Dpi.Scale(480));
        StartPosition = Windows.Count > 1 ? FormStartPosition.WindowsDefaultLocation : FormStartPosition.CenterScreen;

        _spell = new SpellService();
        _ed = new WordEditor { Dock = DockStyle.Fill };
        _ed.Misspelled = p => _spell.Errors(p);
        _ruler = new RulerControl(_ed);
        _editorHost = new Panel { Dock = DockStyle.Fill, Name = "editorHost" };
        _editorHost.Controls.Add(_ed);
        _editorHost.Controls.Add(_ruler);

        _navList = new ListBox { Dock = DockStyle.Fill, Name = "lstNavHeadings", BorderStyle = BorderStyle.None, IntegralHeight = false, DrawMode = DrawMode.OwnerDrawFixed, ItemHeight = Dpi.Scale(24), BackColor = Theme.PanelBack };
        _navList.DrawItem += DrawNavItem;
        _navList.Click += (_, _) => { if (_navList.SelectedItem is NavItem n) GoToParagraph(n.Para); };
        var navTitle = new Label { Text = Loc.T("Điều hướng", "Navigation"), Dock = DockStyle.Top, Height = Dpi.Scale(32), Font = new Font("Segoe UI Semibold", 11f), Padding = new Padding(Dpi.Scale(8), Dpi.Scale(6), 0, 0), BackColor = Theme.PanelBack };
        var navHint = new Label { Text = Loc.T("Dùng kiểu Tiêu đề 1–3 để tạo dàn ý.", "Use Heading 1–3 styles to build the outline."), Dock = DockStyle.Bottom, Height = Dpi.Scale(40), ForeColor = Theme.SubtleInk, Padding = new Padding(Dpi.Scale(8), 0, Dpi.Scale(4), 0), BackColor = Theme.PanelBack, Name = "lblNavHint" };
        _navPane = new Panel { Dock = DockStyle.Left, Width = Dpi.Scale(240), Visible = false, Name = "navPane", BackColor = Theme.PanelBack };
        _navPane.Controls.Add(_navList);
        _navPane.Controls.Add(navHint);
        _navPane.Controls.Add(navTitle);
        var split = new Splitter { Dock = DockStyle.Left, Width = 3, BackColor = Theme.Border, Visible = false, Name = "navSplitter" };
        _navPane.VisibleChanged += (_, _) => split.Visible = _navPane.Visible;

        _ribbon = new RibbonBar();
        _backstage = new Backstage(AppKind.Word);
        BuildRibbon();
        var status = BuildStatusBar();
        BuildBackstage();

        var content = new Panel { Dock = DockStyle.Fill, Name = "contentHost" };
        content.Controls.Add(_editorHost);
        content.Controls.Add(split);
        content.Controls.Add(_navPane);
        InitAi(content);
        Controls.Add(content);
        Controls.Add(status);
        Controls.Add(_ribbon);
        Controls.Add(_backstage);
        _backstage.BringToFront();

        _ed.SelectionChanged += (_, _) => UpdateUi();
        _ed.DocumentChanged += (_, _) => { UpdateUi(); _navTimer.Stop(); _navTimer.Start(); };
        _ed.ZoomChanged += (_, _) => UpdateZoomUi();
        _ed.MouseSelectionDone += (_, _) => ApplyFormatPainter();
        _navTimer.Tick += (_, _) => { _navTimer.Stop(); RefreshNav(); UpdateSpellMode(); };
        _autoSave.Tick += (_, _) => AutoRecoverSave();
        _autoSave.Start();

        DragEnter += (_, e) => { if (e.Data?.GetDataPresent(DataFormats.FileDrop) == true) e.Effect = DragDropEffects.Copy; };
        DragDrop += (_, e) =>
        {
            if (e.Data?.GetData(DataFormats.FileDrop) is not string[] dropped) return;
            var images = dropped.Where(WordEditor.IsImageFile).ToList();
            foreach (var img in images) { HideBackstage(); _ed.InsertImageFile(img); }
            OpenFiles(dropped.Except(images).ToList(), false);
        };
        _ed.AllowDrop = true;
        _ed.DragEnter += (_, e) => { if (e.Data?.GetDataPresent(DataFormats.FileDrop) == true) e.Effect = DragDropEffects.Copy; };
        _ed.DragDrop += (s, e) => HandleDrop(e);

        Load += (_, _) =>
        {
            RestoreWindow();
            if (files.Length > 0) OpenFiles(files, false);
            else if (primary) ShowBackstage("home");
            UpdateUi();
            if (primary) BeginInvoke(OfferRecovery);
        };
        Shown += (_, _) => { if (!_backstage.Visible) _ed.Focus(); UpdateUi(); };
        FormClosing += OnClosingWindow;
        FormClosed += OnClosedWindow;
        UpdateTitle();
    }

    // ------------------------------------------------------------------ IOfficeAppWindow

    public IReadOnlyList<string> OpenDocumentPaths => Windows.Where(w => w._path != null && !w.IsDisposed).Select(w => w._path!).Distinct().ToList();

    public void OpenFiles(IReadOnlyList<string> paths, bool fromShell)
    {
        foreach (var raw in paths)
        {
            string path;
            try { path = Path.GetFullPath(raw); } catch { continue; }
            var existing = Windows.FirstOrDefault(w => !w.IsDisposed && string.Equals(w._path, path, StringComparison.OrdinalIgnoreCase));
            if (existing != null) { existing.ActivateWindow(); continue; }
            var target = Windows.FirstOrDefault(w => !w.IsDisposed && w.IsPristine && (w == this || !w.Visible)) ?? (IsPristine ? this : null);
            if (target == null)
            {
                target = new WordMainForm([], primary: false);
                target.Show();
            }
            target.LoadFile(path);
            target.ActivateWindow();
        }
    }

    private void ActivateWindow()
    {
        if (!Visible) Show();
        if (WindowState == FormWindowState.Minimized) WindowState = FormWindowState.Normal;
        Activate();
    }

    // ------------------------------------------------------------------ files

    public bool LoadFile(string path)
    {
        try
        {
            Cursor = Cursors.WaitCursor;
            var doc = WordFiles.Open(path);
            _ed.Document = doc;
            _path = path;
            UpdateSpellMode();
            OfficeConfig.AddRecent(AppKind.Word, path);
            HideBackstage();
            UpdateTitle();
            RefreshNav();
            UpdateUi();
            return true;
        }
        catch (NotSupportedException ex) { Msg.Warn(this, ex.Message); }
        catch (Exception ex)
        {
            Msg.Error(this, Loc.F("Không mở được tệp \"{0}\".\n{1}", "Could not open \"{0}\".\n{1}", Path.GetFileName(path), ex.Message));
            OfficeConfig.RemoveRecent(AppKind.Word, path);
        }
        finally { Cursor = Cursors.Default; }
        return false;
    }

    public void NewDocument(Document? doc = null, bool newWindow = false)
    {
        if (newWindow && !IsPristine)
        {
            var w = new WordMainForm([], primary: false);
            w.Show();
            w.NewDocument(doc);
            return;
        }
        _ed.Document = doc ?? Document.CreateBlank();
        _path = null;
        UpdateSpellMode();
        _untitledName = Loc.T("Tài liệu", "Document") + (++_untitledCounter);
        HideBackstage();
        UpdateTitle();
        RefreshNav();
        UpdateUi();
        _ed.Focus();
    }

    public void CmdOpen()
    {
        using var d = new OpenFileDialog { Title = Loc.T("Mở tài liệu", "Open document"), Filter = WordFiles.OpenFilter, Multiselect = true };
        if (d.ShowDialog(this) == DialogResult.OK) OpenFiles(d.FileNames, false);
    }

    public bool CmdSave()
    {
        if (_path == null || Path.GetExtension(_path).Equals(".doc", StringComparison.OrdinalIgnoreCase)) return CmdSaveAs();
        return SaveTo(_path);
    }

    public bool CmdSaveAs()
    {
        using var d = new SaveFileDialog
        {
            Title = Loc.T("Lưu thành", "Save As"), Filter = WordFiles.SaveFilter, AddExtension = true, OverwritePrompt = true,
            FileName = _path != null ? Path.GetFileNameWithoutExtension(_path) : SuggestName(),
            InitialDirectory = _path != null ? Path.GetDirectoryName(_path) : Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
        };
        var ext = Path.GetExtension(_path ?? ".docx").ToLowerInvariant();
        d.FilterIndex = ext switch { ".pdf" => 2, ".rtf" => 3, ".html" or ".htm" => 4, ".txt" => 5, _ => 1 };
        if (d.ShowDialog(this) != DialogResult.OK) return false;
        if (Path.GetExtension(d.FileName).Equals(".pdf", StringComparison.OrdinalIgnoreCase)) return ExportPdfTo(d.FileName);
        return SaveTo(d.FileName);
    }

    private string SuggestName()
    {
        var first = _ed.Document.Paragraphs.Select(p => p.Text.Replace(Run.ObjectChar.ToString(), "").Trim()).FirstOrDefault(t => t.Length > 0);
        if (first == null) return _untitledName;
        foreach (var c in Path.GetInvalidFileNameChars()) first = first.Replace(c, ' ');
        return first.Length > 40 ? first[..40].Trim() : first;
    }

    public bool SaveTo(string path)
    {
        try
        {
            Cursor = Cursors.WaitCursor;
            WordFiles.Save(_ed.Document, path);
            _path = path;
            _ed.MarkSaved();
            OfficeConfig.AddRecent(AppKind.Word, path);
            DeleteRecovery();
            UpdateTitle();
            return true;
        }
        catch (Exception ex)
        {
            Msg.Error(this, Loc.F("Không lưu được tệp.\n{0}", "Could not save the file.\n{0}", ex.Message));
            return false;
        }
        finally { Cursor = Cursors.Default; }
    }

    public bool ExportPdfTo(string path)
    {
        try
        {
            Cursor = Cursors.WaitCursor;
            WordFiles.ExportPdf(_ed.Document, path);
            Cursor = Cursors.Default;
            if (Msg.Confirm(this, Loc.F("Đã xuất PDF: {0}\nMở tệp PDF ngay?", "PDF exported: {0}\nOpen it now?", path), Loc.T("Mở", "Open"), Loc.T("Đóng", "Close")))
                OpenPdf(path);
            return true;
        }
        catch (Exception ex)
        {
            Msg.Error(this, Loc.F("Không xuất được PDF.\n{0}", "Could not export PDF.\n{0}", ex.Message));
            return false;
        }
        finally { Cursor = Cursors.Default; }
    }

    public void CmdExportPdf()
    {
        using var d = new SaveFileDialog
        {
            Title = Loc.T("Xuất PDF", "Export to PDF"), Filter = "PDF (*.pdf)|*.pdf", AddExtension = true,
            FileName = Path.GetFileNameWithoutExtension(_path ?? SuggestName()) + ".pdf",
            InitialDirectory = _path != null ? Path.GetDirectoryName(_path) : Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
        };
        if (d.ShowDialog(this) == DialogResult.OK) ExportPdfTo(d.FileName);
    }

    private static void OpenPdf(string path)
    {
        try
        {
            var exe = Environment.ProcessPath;
            if (exe != null && Path.GetFileName(exe).StartsWith("MicroadsOffice", StringComparison.OrdinalIgnoreCase) && !exe.Contains("Tests", StringComparison.OrdinalIgnoreCase))
                Process.Start(new ProcessStartInfo(exe, $"--app=pdf \"{path}\"") { UseShellExecute = false });
            else Process.Start(new ProcessStartInfo(path) { UseShellExecute = true });
        }
        catch { }
    }

    /// <summary>Asks to save; returns false when the user cancelled.</summary>
    public bool ConfirmDiscard()
    {
        if (!_ed.Modified) return true;
        ActivateWindow();
        switch (Msg.AskSave(this, DisplayName))
        {
            case SaveChoice.Save: return CmdSave();
            case SaveChoice.DontSave: return true;
            default: return false;
        }
    }

    public string DisplayName => _path != null ? Path.GetFileName(_path) : _untitledName;

    private void UpdateTitle() => Text = (_ed.Modified ? "• " : "") + DisplayName + " - " + Product.AppName(AppKind.Word);

    // ------------------------------------------------------------------ closing / windows

    private void OnClosingWindow(object? sender, FormClosingEventArgs e)
    {
        if (_forceClose) return;
        if (this == _primary && OfficeShell.PendingRestartArgs != null)
        {
            foreach (var w in Windows.Where(w => w != this && !w.IsDisposed).ToList())
            {
                w.Close();
                if (!w.IsDisposed && w.Visible) { e.Cancel = true; return; }
            }
        }
        if (!ConfirmDiscard()) { e.Cancel = true; return; }
        if (_fullScreen) ToggleFullScreen();
        SaveWindowState();
        if (this == _primary && e.CloseReason == CloseReason.UserClosing && OfficeShell.PendingRestartArgs == null && Windows.Any(w => w != this && !w.IsDisposed && w.Visible))
        {
            // keep the primary window alive (it owns the message loop) while other documents are open
            e.Cancel = true;
            Hide();
            _ed.Document = Document.CreateBlank();
            _path = null;
            DeleteRecovery();
            return;
        }
        DeleteRecovery();
    }

    private void OnClosedWindow(object? sender, FormClosedEventArgs e)
    {
        Windows.Remove(this);
        _autoSave.Stop();
        _find?.Close();
        if (this != _primary && _primary != null && !_primary.IsDisposed)
        {
            bool anyVisible = Windows.Any(w => !w.IsDisposed && w.Visible);
            if (OfficeShell.PendingRestartArgs != null || (!_primary.Visible && !anyVisible))
            {
                var p = _primary;
                p.BeginInvoke(() => { p._forceClose = !p._ed.Modified; p.Close(); });
            }
        }
        if (this == _primary) _primary = null;
    }

    private void SaveWindowState()
    {
        try
        {
            var b = WindowState == FormWindowState.Normal ? Bounds : RestoreBounds;
            OfficeConfig.SetAppData("word.window", $"{b.X},{b.Y},{b.Width},{b.Height},{(WindowState == FormWindowState.Maximized ? 1 : 0)}");
        }
        catch { }
    }

    private void RestoreWindow()
    {
        if (Windows.Count > 1 || this != _primary) return;
        var s = OfficeConfig.GetAppData("word.window")?.Split(',');
        if (s is not { Length: 5 }) return;
        if (!int.TryParse(s[0], out var x) || !int.TryParse(s[1], out var y) || !int.TryParse(s[2], out var w) || !int.TryParse(s[3], out var h)) return;
        var r = new Rectangle(x, y, w, h);
        if (Screen.AllScreens.Any(sc => sc.WorkingArea.IntersectsWith(r)) && w > 400 && h > 300) { StartPosition = FormStartPosition.Manual; Bounds = r; }
        if (s[4] == "1") WindowState = FormWindowState.Maximized;
    }

    private void HandleDrop(DragEventArgs e)
    {
        if (e.Data?.GetData(DataFormats.FileDrop) is not string[] dropped) return;
        var images = dropped.Where(WordEditor.IsImageFile).ToList();
        foreach (var img in images) _ed.InsertImageFile(img);
        OpenFiles(dropped.Except(images).ToList(), false);
    }

    // ------------------------------------------------------------------ AutoRecover

    private static string RecoverDir => Path.Combine(AppSettings.DataDirectory, "AutoRecover", "word");

    private void AutoRecoverSave()
    {
        if (!_ed.Modified) return;
        try
        {
            Directory.CreateDirectory(RecoverDir);
            DocxWriter.Write(_ed.Document, Path.Combine(RecoverDir, _recoverId + ".docx"));
            File.WriteAllText(Path.Combine(RecoverDir, _recoverId + ".txt"), (_path ?? "") + "\n" + DisplayName + "\n" + Environment.ProcessId);
        }
        catch { }
    }

    public void ForceAutoRecoverSave() => AutoRecoverSave();

    private void DeleteRecovery()
    {
        try
        {
            File.Delete(Path.Combine(RecoverDir, _recoverId + ".docx"));
            File.Delete(Path.Combine(RecoverDir, _recoverId + ".txt"));
        }
        catch { }
    }

    private void OfferRecovery()
    {
        try
        {
            if (!Directory.Exists(RecoverDir)) return;
            var items = new List<(string Docx, string Info)>();
            foreach (var info in Directory.GetFiles(RecoverDir, "*.txt"))
            {
                var docx = Path.ChangeExtension(info, ".docx");
                if (!File.Exists(docx)) { File.Delete(info); continue; }
                var lines = File.ReadAllLines(info);
                if (lines.Length >= 3 && int.TryParse(lines[2], out var pid) && pid != Environment.ProcessId && IsRunning(pid)) continue;
                if (Path.GetFileNameWithoutExtension(info) == _recoverId) continue;
                items.Add((docx, info));
            }
            if (items.Count == 0) return;
            var names = string.Join("\n", items.Select(i => "• " + (File.ReadAllLines(i.Info).ElementAtOrDefault(1) ?? "?")));
            bool ok = Msg.Confirm(this, Loc.F("Microads Office Word đã đóng đột ngột. Tìm thấy {0} tài liệu chưa lưu:\n{1}\n\nKhôi phục ngay?",
                "Microads Office Word closed unexpectedly. {0} unsaved document(s) were found:\n{1}\n\nRecover them now?", items.Count, names), Loc.T("Khôi phục", "Recover"), Loc.T("Bỏ qua", "Discard"));
            foreach (var (docx, info) in items)
            {
                if (ok)
                {
                    var lines = File.ReadAllLines(info);
                    var doc = DocxReader.Read(docx);
                    var w = IsPristine ? this : new WordMainForm([], primary: false);
                    if (w != this) w.Show();
                    w.NewDocument(doc);
                    w._path = string.IsNullOrWhiteSpace(lines.ElementAtOrDefault(0)) ? null : lines[0];
                    w._untitledName = (lines.ElementAtOrDefault(1) ?? w._untitledName) + Loc.T(" (đã khôi phục)", " (recovered)");
                    w._ed.Edit("recover", () => { });
                    w.UpdateTitle();
                }
                File.Delete(docx);
                File.Delete(info);
            }
        }
        catch { }
    }

    private static bool IsRunning(int pid)
    {
        try { using var p = Process.GetProcessById(pid); return !p.HasExited && p.ProcessName.StartsWith("MicroadsOffice", StringComparison.OrdinalIgnoreCase); }
        catch { return false; }
    }

    /// <summary>With only an English dictionary, Vietnamese documents are not spell-checked (avoids false red lines).</summary>
    private void UpdateSpellMode()
    {
        if (_spell.Language.StartsWith("vi", StringComparison.OrdinalIgnoreCase)) { _spell.Suppressed = false; return; }
        var text = _ed.Document.PlainText;
        bool vi = SpellService.LooksVietnamese(text.Length > 8000 ? text[..8000] : text);
        if (vi != _spell.Suppressed) { _spell.Suppressed = vi; _ed.Invalidate(); }
    }

    // ------------------------------------------------------------------ navigation pane

    private sealed record NavItem(Paragraph Para, int Level, string Text);

    private void RefreshNav()
    {
        if (!_navPane.Visible) return;
        var items = _ed.Document.Paragraphs.Where(p => p.Props.OutlineLevel is >= 1 and <= 3 && p.Cell == null && p.Text.Trim().Length > 0)
            .Select(p => new NavItem(p, p.Props.OutlineLevel, p.Text.Replace(Run.ObjectChar.ToString(), "").Trim())).ToList();
        _navList.BeginUpdate();
        _navList.Items.Clear();
        foreach (var i in items) _navList.Items.Add(i);
        _navList.EndUpdate();
    }

    private void DrawNavItem(object? sender, DrawItemEventArgs e)
    {
        if (e.Index < 0) return;
        var it = (NavItem)_navList.Items[e.Index];
        bool sel = (e.State & DrawItemState.Selected) != 0;
        using (var b = new SolidBrush(sel ? Theme.Hover : Theme.PanelBack)) e.Graphics.FillRectangle(b, e.Bounds);
        var r = new Rectangle(e.Bounds.X + Dpi.Scale(8 + 14 * (it.Level - 1)), e.Bounds.Y, e.Bounds.Width - Dpi.Scale(12), e.Bounds.Height);
        TextRenderer.DrawText(e.Graphics, it.Text, it.Level == 1 ? Theme.UiFontBold : Theme.UiFont, r, Theme.Ink, TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis);
    }

    private void GoToParagraph(Paragraph p)
    {
        int i = _ed.Document.IndexOf(p);
        if (i < 0) return;
        _ed.SetCaret(new DocPos(i, 0), false);
        _ed.Focus();
    }

    // ------------------------------------------------------------------ backstage

    public void ShowBackstage(string page)
    {
        _backstage.Visible = true;
        _backstage.BringToFront();
        _backstage.ShowPage(page);
    }

    public void HideBackstage()
    {
        if (!_backstage.Visible) return;
        _backstage.Visible = false;
        _ed.Focus();
    }

    // ------------------------------------------------------------------ keyboard

    protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
    {
        if (_backstage.Visible)
        {
            if (keyData == Keys.Escape && !(IsPristine && this == _primary && Windows.Count == 1 && _path == null && false)) { HideBackstage(); return true; }
            return base.ProcessCmdKey(ref msg, keyData);
        }
        switch (keyData)
        {
            case Keys.Control | Keys.N: NewDocument(null, newWindow: true); return true;
            case Keys.Control | Keys.O: CmdOpen(); return true;
            case Keys.Control | Keys.S: CmdSave(); return true;
            case Keys.F12: CmdSaveAs(); return true;
            case Keys.Control | Keys.P: ShowBackstage("print"); return true;
            case Keys.Control | Keys.F: ShowFind(0); return true;
            case Keys.Control | Keys.H: ShowFind(1); return true;
            case Keys.Control | Keys.G: case Keys.F5: ShowFind(2); return true;
            case Keys.Control | Keys.K: CmdHyperlink(); return true;
            case Keys.F7: CmdSpelling(); return true;
            case Keys.Control | Keys.F1: _ribbon.SetMinimized(!_ribbon.Minimized); return true;
            case Keys.Control | Keys.D: CmdFontDialog(); return true;
            case Keys.Control | Keys.Shift | Keys.D8: ToggleMarks(); return true;
            case Keys.Control | Keys.Shift | Keys.C: CopyFormat(); return true;
            case Keys.Control | Keys.Shift | Keys.V: PasteFormat(); return true;
            case Keys.Control | Keys.W: Close(); return true;
            case Keys.F1: ShowHelp(); return true;
            case Keys.F11: ToggleFullScreen(); return true;
            case Keys.Escape when _fullScreen: ToggleFullScreen(); return true;
            case Keys.Escape when _painter != null: _painter = null; UpdateUi(); return true;
        }
        return base.ProcessCmdKey(ref msg, keyData);
    }

    public void ToggleFullScreen()
    {
        _fullScreen = !_fullScreen;
        if (_fullScreen)
        {
            _prevState = WindowState; _prevBorder = FormBorderStyle;
            _ribbon.Visible = false; _statusBar.Visible = false;
            FormBorderStyle = FormBorderStyle.None;
            WindowState = FormWindowState.Normal;
            WindowState = FormWindowState.Maximized;
        }
        else
        {
            _ribbon.Visible = true; _statusBar.Visible = true;
            FormBorderStyle = _prevBorder;
            WindowState = _prevState;
        }
        if (_cmd.TryGetValue("btnFullScreen", out var b) && b is RibbonButton rb) rb.Checked = _fullScreen;
    }

    protected override void Dispose(bool disposing)
    {
        if (disposing) { _autoSave.Dispose(); _navTimer.Dispose(); }
        base.Dispose(disposing);
    }
}
