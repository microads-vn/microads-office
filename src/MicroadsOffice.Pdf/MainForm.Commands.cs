using MicroadsOffice.Core.Office;
using System.Diagnostics;
using System.Security.Cryptography.X509Certificates;
using MicroadsOffice.Core;
using MicroadsOffice.Core.Pdfium;
using MicroadsOffice.Core.Signing;
using MicroadsOffice.Dialogs;
using MicroadsOffice.Printing;
using MicroadsOffice.UI;
using MicroadsOffice.Viewer;
using MicroadsOffice.Views;

namespace MicroadsOffice;

public sealed partial class MainForm
{
    private sealed class PendingSign
    {
        public required X509Certificate2 Certificate { get; init; }
        public required SignatureStyle Style { get; init; }
        public Image? Graphic { get; init; }
        public string Reason { get; init; } = "";
        public string Location { get; init; } = "";
        public string Contact { get; init; } = "";
        public SignatureStandard Standard { get; init; }
        public string? TsaUrl { get; init; }
        public required SignatureAppearanceInfo Info { get; init; }
    }

    private readonly Dictionary<DocumentView, PendingSign> _pendingSigns = new();

    /// <summary>Writes a file atomically: temp file in the same folder, then replace (the original is never truncated on failure).</summary>
    private static void SafeWrite(string path, byte[] bytes)
    {
        var dir = Path.GetDirectoryName(Path.GetFullPath(path))!;
        var tmp = Path.Combine(dir, "." + Path.GetFileName(path) + "." + Guid.NewGuid().ToString("N")[..8] + ".tmp");
        try
        {
            File.WriteAllBytes(tmp, bytes);
            File.Move(tmp, path, overwrite: true);
        }
        finally { try { if (File.Exists(tmp)) File.Delete(tmp); } catch { } }
    }
    private Image? _pendingStamp;
    private float _pendingStampOpacity = 1f;

    // ================================================================== open / save / close

    private void CmdOpen()
    {
        using var ofd = new OpenFileDialog
        {
            Filter = Loc.T("Tài liệu PDF (*.pdf)|*.pdf|Tất cả tệp (*.*)|*.*", "PDF documents (*.pdf)|*.pdf|All files (*.*)|*.*"), Title = Loc.T("Mở tài liệu PDF", "Open PDF document"), Multiselect = true,
            InitialDirectory = _settings.RecentFiles.Select(Path.GetDirectoryName).FirstOrDefault(Directory.Exists) ?? ""
        };
        if (ofd.ShowDialog(this) != DialogResult.OK) return;
        foreach (var f in ofd.FileNames) OpenFile(f);
    }

    /// <summary>Opens files coming from the command line, another instance, the hub or drag &amp; drop.</summary>
    public void OpenFiles(IReadOnlyList<string> paths, bool fromShell)
    {
        if (WindowState == FormWindowState.Minimized) WindowState = FormWindowState.Normal;
        if (fromShell) { Activate(); BringToFront(); }
        foreach (var p in paths) OpenFile(p);
    }

    /// <summary>Saved documents currently open (reopened after a language switch or an update restart).</summary>
    public IReadOnlyList<string> OpenDocumentPaths =>
        _tabs.TabPages.Cast<TabPage>().Select(p => p.Controls.OfType<DocumentView>().FirstOrDefault()?.FilePath)
            .Where(p => p != null && File.Exists(p)).Select(p => p!).ToList();

    /// <summary>Remembers a document in the PDF recent list and in the suite-wide list (hub / start screens).</summary>
    private void RememberRecent(string path)
    {
        _settings.AddRecent(path);
        try { OfficeConfig.AddRecent(AppKind.Pdf, path); } catch { }
    }

    public bool OpenFile(string path)
    {
        try { path = Path.GetFullPath(path); } catch { }
        foreach (TabPage p in _tabs.TabPages)
        {
            if (p.Controls.OfType<DocumentView>().FirstOrDefault() is { } v && string.Equals(v.FilePath, path, StringComparison.OrdinalIgnoreCase))
            {
                _tabs.SelectedTab = p;
                return true;
            }
        }
        byte[] bytes;
        try { bytes = File.ReadAllBytes(path); }
        catch (Exception ex)
        {
            Msg.Error(this, Loc.T($"Không mở được tệp:\n{path}\n\n{ex.Message}", $"Could not open the file:\n{path}\n\n{ex.Message}"));
            _settings.RecentFiles.RemoveAll(x => string.Equals(x, path, StringComparison.OrdinalIgnoreCase));
            _settings.Save();
            BuildRecentList();
            return false;
        }
        var view = OpenBytes(bytes, path, Path.GetFileName(path));
        if (view == null) return false;
        RememberRecent(path);
        _settings.Save();
        BuildRecentList();
        return true;
    }

    private DocumentView? OpenBytes(byte[] bytes, string? path, string displayName)
    {
        string? password = null;
        PdfiumDocument? pdf = null;
        for (int attempt = 0; attempt < 4 && pdf == null; attempt++)
        {
            try { pdf = PdfiumDocument.Load(bytes, password); }
            catch (PdfPasswordRequiredException ex)
            {
                password = InputDialog.Ask(this, Loc.T("Tài liệu được bảo vệ", "Protected document"),
                    (ex.WrongPassword ? Loc.T("Mật khẩu không đúng. ", "Incorrect password. ") : "") + Loc.T($"Nhập mật khẩu để mở \"{displayName}\":", $"Enter the password to open \"{displayName}\":"), "", password: true);
                if (password == null) return null;
            }
            catch (Exception ex)
            {
                Msg.Error(this, Loc.T($"Không mở được \"{displayName}\":\n{ex.Message}", $"Could not open \"{displayName}\":\n{ex.Message}"));
                return null;
            }
        }
        if (pdf == null) return null;
        var view = new DocumentView(path, displayName, bytes, pdf, password) { Name = "docView" };
        HookView(view);
        var page = new TabPage(displayName) { ToolTipText = path ?? displayName, BackColor = Theme.ViewerBack };
        page.Controls.Add(view);
        _tabs.TabPages.Add(page);
        _tabs.SelectedTab = page;
        OnCurrentChanged();
        view.Viewer.Focus();
        return view;
    }

    private void HookView(DocumentView view)
    {
        var v = view.Viewer;
        view.StateChanged += (_, _) => { if (view == Current) UpdateStatus(); UpdateTabTitle(view); };
        view.SignaturePropertiesRequested += i => ShowSignatureProperties(view, i);
        view.ViewSignedVersionRequested += r => OpenSignedVersion(view, r);
        v.ToolChanged += (_, _) => UpdateStatus();
        v.SelectionChanged += (_, _) => UpdateStatus();
        v.PlacementDrawn += (page, rect) => OnPlacementDrawn(view, page, rect);
        v.PlacementConfirmed += (page, rect) => OnPlacementConfirmed(view, page, rect);
        v.PlacementCancelled += (_, _) => { SetStatus(Loc.T("Đã hủy đặt chữ ký / con dấu.", "Signature / stamp placement cancelled.")); };
        v.EmptyFieldClicked += f => OnEmptyFieldClicked(view, f);
        v.UriClicked += uri =>
        {
            if (Msg.Confirm(this, Loc.T($"Tài liệu muốn mở liên kết:\n{uri}\n\nBạn có tin tưởng liên kết này?", $"The document wants to open a link:\n{uri}\n\nDo you trust this link?"),
                    Loc.T("Mở liên kết", "Open link"), Loc.T("Không", "No")))
                CertificateUi.OpenUrl(uri);
        };
    }

    private void UpdateTabTitle(DocumentView v)
    {
        if (v.Parent is TabPage p)
        {
            var t = (v.Modified ? "● " : "") + v.DisplayName;
            if (p.Text != t) p.Text = t;
            p.ToolTipText = v.FilePath ?? v.DisplayName;
        }
        if (v == Current) Text = $"{v.DisplayName}{(v.Modified ? " *" : "")} - Microads Office PDF";
    }

    private void CmdSave()
    {
        var v = Current;
        if (v == null) return;
        SaveView(v, forceDialog: v.FilePath == null);
    }

    private void CmdSaveAs()
    {
        var v = Current;
        if (v == null) return;
        SaveView(v, forceDialog: true);
    }

    private bool SaveView(DocumentView v, bool forceDialog)
    {
        string? path = v.FilePath;
        if (forceDialog || path == null)
        {
            using var sfd = new SaveFileDialog
            {
                Filter = Loc.T("Tài liệu PDF (*.pdf)|*.pdf", "PDF documents (*.pdf)|*.pdf"), Title = Loc.T("Lưu tài liệu", "Save document"), FileName = Path.GetFileNameWithoutExtension(v.DisplayName) + ".pdf",
                InitialDirectory = v.FilePath != null ? Path.GetDirectoryName(v.FilePath) : "", OverwritePrompt = true
            };
            if (sfd.ShowDialog(this) != DialogResult.OK) return false;
            path = sfd.FileName;
        }
        try
        {
            SafeWrite(path!, v.Bytes);
            v.MarkSaved(path!);
            RememberRecent(path!);
            _settings.Save();
            BuildRecentList();
            SetStatus(Loc.T("Đã lưu: ", "Saved: ") + path);
            return true;
        }
        catch (Exception ex)
        {
            Msg.Error(this, Loc.T("Không lưu được tệp: ", "Could not save the file: ") + ex.Message);
            return false;
        }
    }

    private void CloseTab(TabPage page)
    {
        var v = page.Controls.OfType<DocumentView>().FirstOrDefault();
        if (v is { Modified: true })
        {
            _tabs.SelectedTab = page;
            var r = Msg.AskSave(this, v.DisplayName);
            if (r == SaveChoice.Cancel) return;
            if (r == SaveChoice.Save && !SaveView(v, false)) return;
        }
        if (v != null) _pendingSigns.Remove(v);
        _tabs.TabPages.Remove(page);
        page.Dispose();
        OnCurrentChanged();
    }

    // ================================================================== state

    private void OnCurrentChanged()
    {
        var v = Current;
        bool has = v != null;
        _startPage.Visible = !has;
        _tabs.Visible = has;
        if (!has) BuildRecentList();
        foreach (var (name, item) in _cmd)
        {
            if (name is "btnOpen" or "miOpen" or "btnStyles" or "btnCreateId" or "btnCertificates" or "btnSignSettings" or "miSignSettings"
                or "btnHelp" or "btnAbout" or "btnFileAssoc" or "miExit" or "btnCheckUpdate" or "btnSettings" or "btnVisitWebsite"
                or "miCheckUpdate" or "miSettings" or "miAbout") continue;
            item.Enabled = has;
        }
        Text = has ? $"{v!.DisplayName}{(v.Modified ? " *" : "")} - Microads Office PDF" : "Microads Office PDF";
        UpdateStatus();
    }

    private void UpdateStatus()
    {
        var v = Current;
        if (v == null)
        {
            _pageBox.Text = "";
            _pageCount.Text = "/ 0";
            return;
        }
        var viewer = v.Viewer;
        if (!_pageBox.Focused) _pageBox.Text = (viewer.CurrentPage + 1).ToString();
        _pageCount.Text = "/ " + viewer.PageCount;
        _updatingZoom = true;
        var z = $"{Math.Round(viewer.Zoom * 100)}%";
        if (!_zoomBox.Focused) _zoomBox.Text = z;
        if (!_ribbonZoom.Focused) _ribbonZoom.Text = z;
        _updatingZoom = false;
        ((ToolStripButton)_cmd["btnHand"]).Checked = viewer.Tool == ViewerTool.Hand;
        ((ToolStripButton)_cmd["btnSelect"]).Checked = viewer.Tool == ViewerTool.SelectText;
        ((ToolStripButton)_cmd["btnSign"]).Checked = viewer.Tool == ViewerTool.PlaceSignature;
        ((ToolStripButton)_cmd["btnSignHome"]).Checked = viewer.Tool == ViewerTool.PlaceSignature;
        ((ToolStripButton)_cmd["btnStamp"]).Checked = viewer.Tool == ViewerTool.PlaceStamp;
        ((ToolStripButton)_cmd["btnStampHome"]).Checked = viewer.Tool == ViewerTool.PlaceStamp;
        ((ToolStripButton)_cmd["btnSinglePage"]).Checked = viewer.LayoutMode == PageLayoutMode.SinglePage;
        ((ToolStripButton)_cmd["btnContinuous"]).Checked = viewer.LayoutMode == PageLayoutMode.Continuous;
        ((ToolStripButton)_cmd["btnNavPane"]).Checked = v.NavVisible;
        ((ToolStripButton)_cmd["btnFitPage"]).Checked = viewer.ZoomMode == ZoomMode.FitPage;
        ((ToolStripButton)_cmd["btnFitWidth"]).Checked = viewer.ZoomMode == ZoomMode.FitWidth;
        _cmd["btnCopy"].Enabled = viewer.HasSelection;
        _cmd["btnSave"].Enabled = _cmd["miSave"].Enabled = true;
        if (viewer.Tool == ViewerTool.PlaceSignature && !viewer.HasPending)
            SetStatus(Loc.T("KÝ SỐ: kéo chuột vẽ khung chữ ký trên trang (hoặc bấm vào trường chữ ký có sẵn). Esc để hủy.", "SIGN: drag on the page to draw the signature box (or click an existing signature field). Esc to cancel."));
        else if (viewer.Tool == ViewerTool.PlaceStamp && !viewer.HasPending)
            SetStatus(Loc.T("CHÈN DẤU: kéo chuột vẽ vị trí con dấu trên trang. Esc để hủy.", "STAMP: drag on the page to place the stamp. Esc to cancel."));
        else if ((_statusText.Text ?? "").StartsWith(Loc.T("KÝ SỐ", "SIGN")) || (_statusText.Text ?? "").StartsWith(Loc.T("CHÈN DẤU", "STAMP")))
            SetStatus(Loc.T("Sẵn sàng", "Ready"));
    }

    private void SetStatus(string text) => _statusText.Text = text;

    // ================================================================== view commands

    private void CmdTool(ViewerTool tool)
    {
        var v = Current?.Viewer;
        if (v == null) return;
        v.CancelPending();
        v.SetTool(tool);
        v.Focus();
    }

    private void CmdCopy()
    {
        var v = Current?.Viewer;
        if (v == null) return;
        if (v.CopySelection()) SetStatus(Loc.T("Đã sao chép văn bản vào bộ nhớ tạm.", "Text copied to the clipboard."));
        else SetStatus(Loc.T("Chưa chọn văn bản. Dùng công cụ \"Chọn văn bản\" để bôi đen.", "No text selected. Use the \"Select text\" tool to highlight text."));
    }

    private void CmdSearch()
    {
        var v = Current;
        if (v == null) return;
        var sel = v.Viewer.HasSelection ? v.Viewer.GetSelectedText() : null;
        v.FocusSearch(sel is { Length: > 0 and < 80 } ? sel.Trim() : null);
    }

    private void CmdGoTo()
    {
        var v = Current;
        if (v == null) return;
        var s = InputDialog.Ask(this, Loc.T("Đi tới trang", "Go to page"), Loc.T($"Nhập số trang (1 – {v.Viewer.PageCount}):", $"Enter a page number (1 – {v.Viewer.PageCount}):"), (v.Viewer.CurrentPage + 1).ToString());
        if (s == null) return;
        if (int.TryParse(s.Trim(), out int p) && p >= 1 && p <= v.Viewer.PageCount) v.Viewer.GoToPage(p - 1);
        else Msg.Error(this, Loc.T("Số trang không hợp lệ.", "Invalid page number."));
    }

    private void CmdToggleNav()
    {
        var v = Current;
        if (v == null) return;
        if (v.NavVisible) v.SetNavVisible(false);
        else v.ShowNav(v.CurrentNav);
    }

    private void CmdFullScreen()
    {
        if (!_fullScreen)
        {
            _prevState = WindowState;
            _prevBorder = FormBorderStyle;
            _ribbon.Visible = false;
            _status.Visible = false;
            FormBorderStyle = FormBorderStyle.None;
            WindowState = FormWindowState.Normal;
            WindowState = FormWindowState.Maximized;
            _fullScreen = true;
            Current?.SetNavVisible(false);
            SetStatus(Loc.T("Toàn màn hình — nhấn Esc hoặc F11 để thoát.", "Full screen — press Esc or F11 to exit."));
        }
        else
        {
            FormBorderStyle = _prevBorder;
            WindowState = _prevState;
            _ribbon.Visible = true;
            _status.Visible = true;
            _fullScreen = false;
        }
        Current?.Viewer.Focus();
    }

    private void CmdProperties()
    {
        var v = Current;
        if (v == null) return;
        using var d = new DocumentPropertiesDialog(v);
        d.ShowDialog(this);
    }

    private void CmdPrint()
    {
        var v = Current;
        if (v == null) return;
        if (System.Drawing.Printing.PrinterSettings.InstalledPrinters.Count == 0)
        {
            Msg.Error(this, Loc.T("Máy tính chưa cài đặt máy in nào.", "No printer is installed on this computer."));
            return;
        }
        using var d = new PrintDialogForm(v.Pdf, v.Viewer.CurrentPage);
        if (d.ShowDialog(this) != DialogResult.OK || d.Options == null) return;
        try
        {
            Cursor = Cursors.WaitCursor;
            using var doc = PdfPrinter.CreateDocument(v.Pdf, d.Options, v.DisplayName);
            doc.PrintController = new System.Drawing.Printing.StandardPrintController();
            doc.Print();
            SetStatus(Loc.T($"Đã gửi {d.Options.Pages.Count} trang × {d.Options.Copies} bản tới máy in \"{d.Options.PrinterName}\".", $"Sent {d.Options.Pages.Count} page(s) × {d.Options.Copies} copy(ies) to printer \"{d.Options.PrinterName}\"."));
        }
        catch (Exception ex) { Msg.Error(this, Loc.T("Lỗi khi in: ", "Printing failed: ") + ex.Message); }
        finally { Cursor = Cursors.Default; }
    }

    // ================================================================== signatures

    private void CmdSign()
    {
        var v = Current;
        if (v == null) return;
        if (v.Pdf.IsEncrypted)
        {
            Msg.Error(this, Loc.T("Tài liệu đang được bảo vệ bằng mật khẩu/mã hóa. Microads Office không ký trực tiếp lên tài liệu mã hóa để tránh làm hỏng cấu trúc bảo mật.", "The document is password-protected/encrypted. Microads Office PDF does not sign encrypted documents directly, to avoid breaking their security structure."));
            return;
        }
        v.Viewer.CancelPending();
        if (v.Viewer.Tool == ViewerTool.PlaceSignature) { v.Viewer.SetTool(ViewerTool.Hand); return; }
        v.Viewer.SetTool(ViewerTool.PlaceSignature);
        v.Viewer.Focus();
        UpdateStatus();
    }

    private void OnPlacementDrawn(DocumentView view, int page, Rectangle docRect)
    {
        var viewer = view.Viewer;
        float k = viewer.Zoom * viewer.DeviceDpi / 72f;
        var box = viewer.ViewRotation % 2 == 1 ? new SizeF(docRect.Height / k, docRect.Width / k) : new SizeF(docRect.Width / k, docRect.Height / k);
        if (viewer.Tool == ViewerTool.PlaceSignature)
        {
            using var d = new SignDialog(_settings, box);
            if (d.ShowDialog(this) != DialogResult.OK || d.SelectedCertificate == null) { UpdateStatus(); return; }
            var info = d.BuildInfo();
            var img = d.TakeStyleImage();
            info.Graphic = img;
            var newPending = new PendingSign
            {
                Certificate = d.SelectedCertificate, Style = d.SelectedStyle.Clone(), Graphic = img, Reason = d.Reason, Location = d.SignLocation,
                Contact = d.Contact, Standard = d.Standard, TsaUrl = d.TsaUrl, Info = info
            };
            _pendingSigns[view] = newPending;
            var ps = newPending;
            viewer.BeginPending(page, docRect,
                size => AppearanceRenderer.RenderBitmap(ps.Style, ps.Info, Math.Max(10, size.Width), Math.Max(10, size.Height), k),
                Loc.T("✔  Ký và lưu", "✔  Sign and save"));
            SetStatus(Loc.T("Kéo để di chuyển, kéo các ô vuông để đổi kích thước khung chữ ký. Bấm \"Ký và lưu\" (Enter) để ký, Esc để hủy.", "Drag to move, drag the handles to resize the signature box. Click \"Sign and save\" (Enter) to sign, Esc to cancel."));
        }
        else if (viewer.Tool == ViewerTool.PlaceStamp && _pendingStamp != null)
        {
            // keep the image aspect ratio
            var img = _pendingStamp;
            float ar = (float)img.Width / Math.Max(1, img.Height);
            var r = docRect;
            int h = (int)(r.Width / ar);
            r = new Rectangle(r.X, r.Y, r.Width, Math.Max(10, h));
            var opacity = _pendingStampOpacity;
            viewer.BeginPending(page, r, size =>
            {
                var bmp = new Bitmap(Math.Max(1, (int)(size.Width * k)), Math.Max(1, (int)(size.Height * k)));
                using var g = Graphics.FromImage(bmp);
                using var ia = new System.Drawing.Imaging.ImageAttributes();
                ia.SetColorMatrix(new System.Drawing.Imaging.ColorMatrix { Matrix33 = opacity });
                g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                g.DrawImage(img, new Rectangle(0, 0, bmp.Width, bmp.Height), 0, 0, img.Width, img.Height, GraphicsUnit.Pixel, ia);
                return bmp;
            }, Loc.T("✔  Chèn dấu", "✔  Insert stamp"));
            SetStatus(Loc.T("Điều chỉnh vị trí / kích thước con dấu rồi bấm \"Chèn dấu\" (Enter). Esc để hủy.", "Adjust the stamp position / size, then click \"Insert stamp\" (Enter). Esc to cancel."));
        }
    }

    private void OnPlacementConfirmed(DocumentView view, int page, RectangleF pageRect)
    {
        if (view.Viewer.Tool == ViewerTool.PlaceStamp && _pendingStamp != null)
        {
            try
            {
                int rot = view.Pdf.GetPageRotation(page) * 90;
                var bytes = ImageStamper.AddImageStamp(view.Bytes, page, pageRect, rot, _pendingStamp, _pendingStampOpacity, Environment.UserName);
                try { view.ReplaceContent(bytes, null, modified: true); }
                catch (Exception ex) { Msg.Error(this, Loc.T("Không hiển thị được tài liệu sau khi chèn dấu: ", "Could not display the document after stamping: ") + ex.Message); return; }
                view.Viewer.SetTool(ViewerTool.Hand);
                SetStatus(Loc.T("Đã chèn dấu hình ảnh. Nhấn Lưu (Ctrl+S) để ghi vào tệp.", "Image stamp inserted. Press Save (Ctrl+S) to write it to the file."));
            }
            catch (Exception ex) { Msg.Error(this, Loc.T("Không chèn được dấu: ", "Could not insert the stamp: ") + ex.Message); }
            return;
        }
        if (!_pendingSigns.TryGetValue(view, out var ps)) return;
        int rotation = view.Pdf.GetPageRotation(page) * 90;
        DoSign(view, new SignRequest
        {
            Certificate = ps.Certificate, PageIndex = page, PageRect = pageRect, PageRotation = rotation, Style = ps.Style, Graphic = ps.Graphic,
            Reason = ps.Reason, Location = ps.Location, ContactInfo = ps.Contact, Standard = ps.Standard, TsaUrl = ps.TsaUrl, SigningTime = DateTimeOffset.Now
        });
    }

    private void OnEmptyFieldClicked(DocumentView view, SignatureFieldInfo field)
    {
        var size = field.Rect.Size;
        int rot = field.PageIndex >= 0 ? view.Pdf.GetPageRotation(field.PageIndex) * 90 : 0;
        if (rot % 180 != 0) size = new SizeF(size.Height, size.Width);
        using var d = new SignDialog(_settings, size, field.FullName);
        if (d.ShowDialog(this) != DialogResult.OK || d.SelectedCertificate == null) return;
        var img = d.TakeStyleImage();
        DoSign(view, new SignRequest
        {
            Certificate = d.SelectedCertificate, ExistingFieldName = field.FullName, PageIndex = Math.Max(0, field.PageIndex), PageRotation = rot,
            Style = d.SelectedStyle, Graphic = img, Reason = d.Reason, Location = d.SignLocation, ContactInfo = d.Contact,
            Standard = d.Standard, TsaUrl = d.TsaUrl, SigningTime = DateTimeOffset.Now
        });
    }

    /// <summary>Signs, asks where to save (like Foxit), writes the file and reloads it.</summary>
    private void DoSign(DocumentView view, SignRequest req)
    {
        string baseName = Path.GetFileNameWithoutExtension(view.DisplayName);
        string dir = view.FilePath != null ? Path.GetDirectoryName(view.FilePath)! : Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
        using var sfd = new SaveFileDialog
        {
            Title = Loc.T("Lưu tài liệu đã ký", "Save signed document"), Filter = Loc.T("Tài liệu PDF (*.pdf)|*.pdf", "PDF documents (*.pdf)|*.pdf"), InitialDirectory = dir,
            FileName = baseName.EndsWith(_settings.SignedFileSuffix) ? baseName + ".pdf" : baseName + _settings.SignedFileSuffix + ".pdf",
            OverwritePrompt = true
        };
        if (sfd.ShowDialog(this) != DialogResult.OK)
        {
            SetStatus(Loc.T("Đã hủy ký (chưa chọn nơi lưu).", "Signing cancelled (no save location chosen)."));
            return;
        }
        byte[] signed;
        try
        {
            Cursor = Cursors.WaitCursor;
            SetStatus(Loc.T("Đang ký số... (nhập mã PIN nếu USB Token yêu cầu)", "Signing... (enter the PIN if your USB token asks for it)"));
            Application.DoEvents();
            signed = PdfSigner.Sign(view.Bytes, req);
        }
        catch (Exception ex)
        {
            Cursor = Cursors.Default;
            Msg.Error(this, Loc.T("Ký số không thành công:\n", "Signing failed:\n") + ex.Message);
            SetStatus(Loc.T("Ký số không thành công.", "Signing failed."));
            return;
        }
        finally { Cursor = Cursors.Default; }
        try { SafeWrite(sfd.FileName, signed); }
        catch (Exception ex)
        {
            Msg.Error(this, Loc.T("Đã ký nhưng không ghi được tệp: ", "Signed, but the file could not be written: ") + ex.Message);
            return;
        }
        _pendingSigns.Remove(view);
        try { view.ReplaceContent(signed, sfd.FileName, modified: false); }
        catch (Exception ex) { Msg.Error(this, Loc.T($"Đã ký và lưu tệp:\n{sfd.FileName}\nnhưng không tải lại được để hiển thị: {ex.Message}", $"Signed and saved:\n{sfd.FileName}\nbut it could not be reloaded for display: {ex.Message}")); return; }
        view.Viewer.SetTool(ViewerTool.Hand);
        RememberRecent(sfd.FileName);
        _settings.Save();
        BuildRecentList();
        view.ShowNav(NavKind.Signatures);
        SetStatus(Loc.T($"Đã ký số thành công và lưu: {sfd.FileName}", $"Signed successfully and saved: {sfd.FileName}"));
    }

    private void CmdValidate()
    {
        var v = Current;
        if (v == null) return;
        v.ShowNav(NavKind.Signatures);
        v.Revalidate();
        SetStatus(Loc.T("Đang xác thực lại các chữ ký...", "Re-validating signatures..."));
        void Done(object? s, EventArgs e)
        {
            v.ValidationCompleted -= Done;
            int n = v.Signatures.Count(x => x.IsSigned);
            SetStatus(n == 0 ? Loc.T("Tài liệu không có chữ ký số.", "The document has no digital signatures.") : Loc.T($"Đã xác thực {n} chữ ký.", $"Validated {n} signature(s)."));
        }
        v.ValidationCompleted += Done;
    }

    private void ShowSignatureProperties(DocumentView v, int index)
    {
        if (index < 0 || index >= v.Signatures.Count) return;
        var r = v.Signatures[index];
        if (!r.IsSigned)
        {
            if (Msg.Confirm(this, Loc.T($"Trường chữ ký \"{r.FieldName}\" chưa được ký. Ký vào trường này ngay?", $"Signature field \"{r.FieldName}\" is not signed yet. Sign it now?"),
                    Loc.T("Ký ngay", "Sign now"), Loc.T("Không", "No"), Loc.T("Trường chữ ký", "Signature field")))
            {
                try
                {
                    var f = SignatureFields.Find(Core.Pdf.PdfFile.Load(v.Bytes)).FirstOrDefault(x => x.FullName == r.FieldName);
                    if (f != null) OnEmptyFieldClicked(v, f);
                }
                catch (Exception ex) { Msg.Error(this, ex.Message); }
            }
            return;
        }
        using var d = new SignaturePropertiesDialog(r);
        d.ViewSignedVersion += res => OpenSignedVersion(v, res);
        d.ShowDialog(this);
    }

    private void OpenSignedVersion(DocumentView v, SignatureValidationResult r)
    {
        if (r.RevisionLength <= 0 || r.RevisionLength > v.Bytes.Length) return;
        var bytes = v.Bytes.AsSpan(0, (int)r.RevisionLength).ToArray();
        OpenBytes(bytes, null, Loc.T($"{Path.GetFileNameWithoutExtension(v.DisplayName)} (bản đã ký Rev.{r.Revision}).pdf", $"{Path.GetFileNameWithoutExtension(v.DisplayName)} (signed version Rev.{r.Revision}).pdf"));
    }

    private void CmdStyles()
    {
        using var d = new StyleManagerDialog(_settings);
        d.ShowDialog(this);
    }

    private void CmdCreateId()
    {
        using var d = new CreateIdDialog();
        if (d.ShowDialog(this) == DialogResult.OK && d.PfxPath != null)
            Msg.Info(this, Loc.T($"Đã tạo ID số tự ký:\n{d.PfxPath}\n\nID này xuất hiện trong danh sách chứng thư khi ký (cần nhập mật khẩu).", $"Self-signed digital ID created:\n{d.PfxPath}\n\nIt appears in the certificate list when signing (password required)."));
    }

    private void CmdCertificates()
    {
        using var d = new CertificatesDialog(_settings);
        d.ShowDialog(this);
    }

    /// <summary>
    /// Ctrl+V or a dropped picture (screenshot, copied image, image file, image dragged from a browser or chat app):
    /// it becomes an image stamp that the next click places on the page, like Insert › Image stamp.
    /// Returns false when the data holds no picture.
    /// </summary>
    public bool PlaceImageFrom(IDataObject? data)
    {
        var v = Current;
        if (v == null || data == null) return false;
        var files = DataTransfer.Files(data);
        var img = files.Length > 0 ? DataTransfer.ImageFiles(files).FirstOrDefault() : DataTransfer.Picture(data);
        if (img == null) return false;
        if (v.Pdf.IsEncrypted) { Msg.Error(this, Loc.T("Không thể chèn ảnh vào tài liệu được mã hóa.", "Pictures cannot be inserted into an encrypted document.")); return true; }
        v.Viewer.CancelPending();
        _pendingStamp?.Dispose();
        using (var loaded = img.Load()) _pendingStamp = new Bitmap(loaded);
        _pendingStampOpacity = 1f;
        v.Viewer.SetTool(ViewerTool.PlaceStamp);
        v.Viewer.Focus();
        UpdateStatus();
        SetStatus(Loc.T("Bấm vào trang để đặt ảnh (Esc để hủy). Nhấn Lưu (Ctrl+S) để ghi vào tệp.", "Click on the page to place the picture (Esc to cancel). Press Save (Ctrl+S) to write it to the file."));
        return true;
    }

    private void CmdStamp()
    {
        var v = Current;
        if (v == null) return;
        if (v.Pdf.IsEncrypted) { Msg.Error(this, Loc.T("Không thể chèn dấu vào tài liệu được mã hóa.", "Stamps cannot be inserted into an encrypted document.")); return; }
        v.Viewer.CancelPending();
        if (v.Viewer.Tool == ViewerTool.PlaceStamp) { v.Viewer.SetTool(ViewerTool.Hand); return; }
        using var d = new StampDialog();
        if (d.ShowDialog(this) != DialogResult.OK || d.Image == null) return;
        _pendingStamp?.Dispose();
        _pendingStamp = new Bitmap(d.Image);
        _pendingStampOpacity = d.StampOpacity;
        v.Viewer.SetTool(ViewerTool.PlaceStamp);
        v.Viewer.Focus();
        UpdateStatus();
    }

    private void CmdSignSettings()
    {
        using var d = new SettingsDialog(_settings);
        if (d.ShowDialog(this) == DialogResult.OK)
        {
            SetStatus(Loc.T("Đã lưu thiết lập ký số.", "Signing settings saved."));
            Current?.Revalidate();
        }
    }

    private void CmdFileAssociation()
    {
        if (StorePackage.IsPackaged)
        {
            // Store edition: file types are declared by the package; Windows lets the user pick the default.
            Msg.Info(this, Loc.T("Microads Office PDF đã có trong danh sách \"Mở bằng\" của tệp PDF. Hãy chọn nó làm ứng dụng mặc định trong Cài đặt Windows ▸ Ứng dụng mặc định.",
                "Microads Office PDF is already in the \"Open with\" list for PDF files. Choose it as the default app in Windows Settings ▸ Default apps."));
            StorePackage.OpenDefaultAppsSettings();
            return;
        }
        if (!Msg.Confirm(this,
                Loc.T("Thêm Microads Office PDF vào danh sách \"Mở bằng\" của tệp PDF cho tài khoản Windows hiện tại?\n\n", "Add Microads Office PDF to the \"Open with\" list for PDF files of the current Windows account?\n\n") +
                Loc.T("Sau đó bạn có thể chọn Microads Office PDF làm trình đọc PDF mặc định trong Cài đặt Windows ▸ Ứng dụng mặc định.", "You can then choose Microads Office PDF as the default PDF reader in Windows Settings ▸ Default apps."),
                Loc.T("Đồng ý", "Yes"), Loc.T("Không", "No"), Loc.T("Mở PDF bằng Microads Office", "Open PDF with Microads Office"))) return;
        try
        {
            var exe = Environment.ProcessPath ?? Application.ExecutablePath;
            using (var k = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(@"Software\Classes\MicroadsOffice.PDF"))
            {
                k.SetValue("", Loc.T("Tài liệu PDF (Microads Office)", "PDF document (Microads Office)"));
                using (var ic = k.CreateSubKey("DefaultIcon")) ic.SetValue("", $"\"{exe}\",0");
                using (var cmd = k.CreateSubKey(@"shell\open\command")) cmd.SetValue("", $"\"{exe}\" --app=pdf \"%1\"");
            }
            using (var k = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(@"Software\Classes\.pdf\OpenWithProgids")) k.SetValue("MicroadsOffice.PDF", Array.Empty<byte>(), Microsoft.Win32.RegistryValueKind.None);
            using (var k = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(@"Software\Classes\Applications\MicroadsOffice.exe\shell\open\command")) k.SetValue("", $"\"{exe}\" \"%1\"");
            SHChangeNotify(0x08000000, 0, IntPtr.Zero, IntPtr.Zero);
            if (Msg.Confirm(this, Loc.T("Đã đăng ký. Mở Cài đặt Windows để chọn làm ứng dụng mặc định?", "Registered. Open Windows Settings to make it the default app?"),
                    Loc.T("Mở Cài đặt", "Open Settings"), Loc.T("Không", "No")))
                Process.Start(new ProcessStartInfo("ms-settings:defaultapps") { UseShellExecute = true });
        }
        catch (Exception ex) { Msg.Error(this, Loc.T("Không đăng ký được: ", "Registration failed: ") + ex.Message); }
    }

    [System.Runtime.InteropServices.DllImport("shell32.dll")]
    private static extern void SHChangeNotify(int wEventId, int uFlags, IntPtr dwItem1, IntPtr dwItem2);

    // ================================================================== keyboard

    protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
    {
        bool inText = ActiveControlDeep() is TextBoxBase or ComboBox;
        var v = Current;
        switch (keyData)
        {
            case Keys.Control | Keys.O: CmdOpen(); return true;
            case Keys.Control | Keys.S: if (v != null) CmdSave(); return true;
            case Keys.Control | Keys.Shift | Keys.S: if (v != null) CmdSign(); return true;
            case Keys.Control | Keys.P: if (v != null) CmdPrint(); return true;
            case Keys.Control | Keys.W: if (_tabs.SelectedTab != null) CloseTab(_tabs.SelectedTab); return true;
            case Keys.Control | Keys.F: if (v != null) CmdSearch(); return true;
            case Keys.F3: v?.NextHit(); return v != null;
            case Keys.Shift | Keys.F3: v?.PrevHit(); return v != null;
            case Keys.Control | Keys.G: if (v != null) CmdGoTo(); return true;
            case Keys.Control | Keys.D: if (v != null) CmdProperties(); return true;
            case Keys.F11: if (v != null || _fullScreen) CmdFullScreen(); return true;
            case Keys.Escape when _fullScreen: CmdFullScreen(); return true;
            case Keys.F4: CmdToggleNav(); return true;
            case Keys.F1: { using var d = new HelpDialog(); d.ShowDialog(this); return true; }
            case Keys.Control | Keys.Q: Close(); return true;
            case Keys.Control | Keys.Tab when _tabs.TabCount > 1: _tabs.SelectedIndex = (_tabs.SelectedIndex + 1) % _tabs.TabCount; return true;
            case Keys.Control | Keys.Shift | Keys.Tab when _tabs.TabCount > 1: _tabs.SelectedIndex = (_tabs.SelectedIndex - 1 + _tabs.TabCount) % _tabs.TabCount; return true;
        }
        if (v != null && !inText)
        {
            switch (keyData)
            {
                case Keys.Control | Keys.Oemplus: case Keys.Control | Keys.Add: v.Viewer.ZoomStep(true); return true;
                case Keys.Control | Keys.OemMinus: case Keys.Control | Keys.Subtract: v.Viewer.ZoomStep(false); return true;
                case Keys.Control | Keys.D0: case Keys.Control | Keys.NumPad0: v.Viewer.SetZoomMode(ZoomMode.FitPage); return true;
                case Keys.Control | Keys.D1: case Keys.Control | Keys.NumPad1: v.Viewer.SetZoom(1f); return true;
                case Keys.Control | Keys.D2: case Keys.Control | Keys.NumPad2: v.Viewer.SetZoomMode(ZoomMode.FitWidth); return true;
                case Keys.Control | Keys.Shift | Keys.Oemplus: case Keys.Control | Keys.Shift | Keys.Add: v.Viewer.Rotate(true); return true;
                case Keys.Control | Keys.Shift | Keys.OemMinus: case Keys.Control | Keys.Shift | Keys.Subtract: v.Viewer.Rotate(false); return true;
                case Keys.Control | Keys.C: CmdCopy(); return true;
                case Keys.Control | Keys.V: case Keys.Shift | Keys.Insert:
                    if (PlaceImageFrom(DataTransfer.ReadClipboard())) return true;
                    SetStatus(Loc.T("Không có ảnh trong bộ nhớ tạm để chèn vào PDF.", "There is no picture on the clipboard to insert into the PDF."));
                    return true;
                case Keys.Control | Keys.A: v.Viewer.SetTool(ViewerTool.SelectText); v.Viewer.SelectAll(); return true;
            }
        }
        return base.ProcessCmdKey(ref msg, keyData);
    }

    private Control? ActiveControlDeep()
    {
        Control? c = ActiveControl;
        while (c is ContainerControl cc && cc.ActiveControl != null) c = cc.ActiveControl;
        return c;
    }
}

/// <summary>Lists available signing certificates (Windows store / USB token / PFX).</summary>
public sealed class CertificatesDialog : DialogBase
{
    public CertificatesDialog(AppSettings settings) : base(Loc.T("Chứng thư số của tôi", "My certificates"), 760, 420)
    {
        var lv = new ListView { Dock = DockStyle.Fill, View = View.Details, FullRowSelect = true, Name = "lstCertificates", BorderStyle = BorderStyle.FixedSingle };
        lv.Columns.Add(Loc.T("Người sở hữu", "Owner"), Dpi.Scale(200));
        lv.Columns.Add(Loc.T("Cấp bởi", "Issued by"), Dpi.Scale(180));
        lv.Columns.Add(Loc.T("Hết hạn", "Expires"), Dpi.Scale(90));
        lv.Columns.Add(Loc.T("Nguồn", "Source"), Dpi.Scale(240));
        void Fill()
        {
            lv.Items.Clear();
            foreach (var e in CertificateService.GetStoreCertificates())
                lv.Items.Add(new ListViewItem([e.DisplayName, e.Issuer, e.Certificate.NotAfter.ToString("dd/MM/yyyy"), e.Source]) { Tag = e.Certificate, ForeColor = e.IsExpired ? Color.Gray : Theme.Ink });
            var pfx = new List<string>(settings.PfxFiles);
            if (Directory.Exists(AppSettings.IdsDirectory)) pfx.AddRange(Directory.GetFiles(AppSettings.IdsDirectory, "*.pfx"));
            foreach (var p in pfx.Where(File.Exists).Distinct(StringComparer.OrdinalIgnoreCase))
                lv.Items.Add(new ListViewItem([Path.GetFileNameWithoutExtension(p), Loc.T("(tệp được bảo vệ bằng mật khẩu)", "(password-protected file)"), "", Loc.T("Tệp PFX: ", "PFX file: ") + p]));
            if (lv.Items.Count == 0) lv.Items.Add(new ListViewItem([Loc.T("Không có chứng thư nào. Cắm USB Token rồi bấm Làm mới, hoặc tạo ID số.", "No certificates. Plug in your USB token and click Refresh, or create a digital ID."), "", "", ""]));
        }
        Fill();
        var note = new Label
        {
            Dock = DockStyle.Bottom, Height = Dpi.Scale(44), ForeColor = Theme.SubtleInk,
            Text = Loc.T("Chứng thư trên USB Token xuất hiện khi Token được cắm và driver của nhà cung cấp đã được cài. Chọn một dòng rồi bấm \"Xem chi tiết\".", "USB token certificates appear when the token is plugged in and the vendor driver is installed. Select a row and click \"View details\".")
        };
        Body.Controls.Add(lv);
        Body.Controls.Add(note);
        AddButton(Loc.T("Đóng", "Close"), "btnClose", DialogResult.OK, true);
        AddButton(Loc.T("Làm mới", "Refresh"), "btnRefresh", DialogResult.None, false, (_, _) => Fill());
        AddButton(Loc.T("Xem chi tiết...", "View details..."), "btnDetails", DialogResult.None, false, (_, _) =>
        {
            if (lv.SelectedItems.Count > 0 && lv.SelectedItems[0].Tag is X509Certificate2 c) CertificateUi.Show(this, c);
        });
        lv.DoubleClick += (_, _) => { if (lv.SelectedItems.Count > 0 && lv.SelectedItems[0].Tag is X509Certificate2 c) CertificateUi.Show(this, c); };
    }
}

/// <summary>Manage signature appearance styles.</summary>
public sealed class StyleManagerDialog : DialogBase
{
    private readonly AppSettings _settings;
    private readonly ListBox _list;
    private readonly PictureBox _preview;

    public StyleManagerDialog(AppSettings settings) : base(Loc.T("Mẫu chữ ký", "Signature styles"), 720, 420)
    {
        _settings = settings;
        _list = new ListBox { Dock = DockStyle.Left, Width = Dpi.Scale(260), Name = "lstStyles", DisplayMember = nameof(SignatureStyle.Title), IntegralHeight = false };
        _preview = new PictureBox { Dock = DockStyle.Fill, SizeMode = PictureBoxSizeMode.Zoom, BackColor = Color.White, BorderStyle = BorderStyle.FixedSingle, Name = "picStyleManagerPreview" };
        var spacer = new Panel { Dock = DockStyle.Left, Width = Dpi.Scale(12) };
        Body.Controls.Add(_preview);
        Body.Controls.Add(spacer);
        Body.Controls.Add(_list);
        _list.SelectedIndexChanged += (_, _) => Preview();
        _list.DoubleClick += (_, _) => Edit(_list.SelectedItem as SignatureStyle);
        AddButton(Loc.T("Đóng", "Close"), "btnClose", DialogResult.OK, true);
        AddButton(Loc.T("Xóa", "Delete"), "btnDeleteStyleMgr", DialogResult.None, false, (_, _) =>
        {
            if (_list.SelectedItem is not SignatureStyle s) return;
            if (_settings.Styles.Count <= 1) { Msg.Info(this, Loc.T("Cần giữ lại ít nhất một mẫu.", "At least one style must be kept.")); return; }
            if (!Msg.Confirm(this, Loc.T($"Xóa mẫu \"{s.Title}\"?", $"Delete style \"{s.Title}\"?"), Loc.T("Xóa", "Delete"), Loc.T("Hủy", "Cancel"))) return;
            _settings.Styles.RemoveAll(x => x.Id == s.Id); _settings.Save(); Fill(null);
        });
        AddButton(Loc.T("Sửa...", "Edit..."), "btnEditStyleMgr", DialogResult.None, false, (_, _) => Edit(_list.SelectedItem as SignatureStyle));
        AddButton(Loc.T("Tạo mới...", "New..."), "btnNewStyleMgr", DialogResult.None, false, (_, _) => Edit(null));
        Fill(null);
    }

    private void Fill(string? selectId)
    {
        _list.Items.Clear();
        foreach (var s in _settings.Styles) _list.Items.Add(s);
        int idx = _settings.Styles.FindIndex(s => s.Id == selectId);
        _list.SelectedIndex = idx >= 0 ? idx : (_list.Items.Count > 0 ? 0 : -1);
    }

    private void Edit(SignatureStyle? s)
    {
        using var d = new SignatureStyleDialog(s);
        if (d.ShowDialog(this) != DialogResult.OK) return;
        int idx = s == null ? -1 : _settings.Styles.FindIndex(x => x.Id == s.Id);
        if (idx >= 0) _settings.Styles[idx] = d.Style; else _settings.Styles.Add(d.Style);
        _settings.Save();
        Fill(d.Style.Id);
    }

    private void Preview()
    {
        if (_list.SelectedItem is not SignatureStyle s) return;
        using var img = StyleImages.Load(s);
        var info = new SignatureAppearanceInfo { SignerName = Loc.T("Nguyễn Văn An", "Nguyen Van An"), Reason = Loc.T("Phê duyệt", "Approved"), Location = Loc.T("Hà Nội", "Hanoi"), SigningTime = DateTimeOffset.Now, Graphic = img, DistinguishedName = Loc.T("CN=Nguyễn Văn An, C=VN", "CN=Nguyen Van An, C=VN") };
        var old = _preview.Image;
        _preview.Image = AppearanceRenderer.RenderBitmap(s, info, 240, 90, 2, false);
        old?.Dispose();
    }
}
