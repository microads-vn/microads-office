using MicroadsOffice.Core.Office;
using System.Security.Cryptography.X509Certificates;
using MicroadsOffice.Core;
using MicroadsOffice.Core.Signing;
using MicroadsOffice.UI;

namespace MicroadsOffice.Dialogs;

/// <summary>Certificate choice: a loaded certificate or a .pfx file that still has to be unlocked.</summary>
public sealed class CertChoice
{
    public X509Certificate2? Certificate { get; set; }
    public string? PfxPath { get; init; }
    public string Source { get; init; } = "";

    public override string ToString()
    {
        if (Certificate != null) return new CertificateEntry(Certificate, Source).ToString();
        return Loc.T($"🔒 {Path.GetFileName(PfxPath)}  [Tệp PFX — chọn để nhập mật khẩu]", $"🔒 {Path.GetFileName(PfxPath)}  [PFX file — select to enter the password]");
    }
}

/// <summary>"Ký tài liệu" dialog (Foxit: Sign Document).</summary>
public sealed class SignDialog : DialogBase
{
    private readonly AppSettings _settings;
    private readonly SizeF _boxPts;
    private readonly ComboBox _cert, _style, _reason, _standard;
    private readonly TextBox _location, _contact, _tsa;
    private readonly CheckBox _useTs;
    private readonly PictureBox _preview;
    private readonly Label _certInfo;
    private int _lastCertIndex = -1;
    private bool _suppressPrompt;
    private Image? _styleImage;
    private readonly Label _styleHint;

    public X509Certificate2? SelectedCertificate => (_cert.SelectedItem as CertChoice)?.Certificate;
    public SignatureStyle SelectedStyle => (_style.SelectedItem as SignatureStyle) ?? SignatureStyle.CreateDefault();
    public Image? StyleImage => _styleImage;
    public string Reason => _reason.Text.Trim();
    public string SignLocation => _location.Text.Trim();
    public string Contact => _contact.Text.Trim();
    public SignatureStandard Standard => _standard.SelectedIndex == 1 ? SignatureStandard.Pkcs7 : SignatureStandard.PAdES;
    public string? TsaUrl => _useTs.Checked && !string.IsNullOrWhiteSpace(_tsa.Text) ? _tsa.Text.Trim() : null;
    public bool Invisible { get; }

    public SignDialog(AppSettings settings, SizeF boxPts, string? fieldName = null) : base(Loc.T("Ký tài liệu", "Sign document"), 860, 600)
    {
        _settings = settings;
        _boxPts = boxPts.Width < 10 ? new SizeF(210, 70) : boxPts;
        Text = fieldName != null ? Loc.T($"Ký tài liệu — trường \"{fieldName}\"", $"Sign document — field \"{fieldName}\"") : Loc.T("Ký tài liệu", "Sign document");

        var left = new TableLayoutPanel { Dock = DockStyle.Left, Width = Dpi.Scale(500), ColumnCount = 2 };
        left.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        left.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));

        _cert = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = Dpi.Scale(390), Name = "cboCertificate", DropDownWidth = Dpi.Scale(640) };
        left.Controls.Add(L(Loc.T("Ký với chứng thư số:", "Sign with certificate:"), true)); left.Controls.Add(_cert);
        var certBtns = new FlowLayoutPanel { AutoSize = true, Margin = Padding.Empty };
        var btnPfx = SmallBtn(Loc.T("Nạp tệp .pfx/.p12...", "Load .pfx/.p12 file..."), "btnLoadPfx", (_, _) => LoadPfx());
        var btnNewId = SmallBtn(Loc.T("Tạo ID số mới...", "Create new digital ID..."), "btnCreateId", (_, _) => CreateId());
        var btnView = SmallBtn(Loc.T("Xem chứng thư", "View certificate"), "btnViewCertificate", (_, _) => { if (SelectedCertificate != null) CertificateUi.Show(this, SelectedCertificate); });
        var btnRefresh = SmallBtn(Loc.T("Làm mới (cắm Token)", "Refresh (plug in token)"), "btnRefreshCerts", (_, _) => LoadCertificates());
        certBtns.Controls.AddRange([btnPfx, btnNewId, btnView, btnRefresh]);
        left.Controls.Add(new Label()); left.Controls.Add(certBtns);
        _certInfo = new Label { AutoSize = true, MaximumSize = new Size(Dpi.Scale(390), 0), ForeColor = Theme.SubtleInk, Name = "lblCertInfo" };
        left.Controls.Add(new Label()); left.Controls.Add(_certInfo);

        _style = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = Dpi.Scale(390), Name = "cboStyle" };
        left.Controls.Add(L(Loc.T("Mẫu hiển thị:", "Appearance:"), true)); left.Controls.Add(_style);
        var styleBtns = new FlowLayoutPanel { AutoSize = true, Margin = Padding.Empty };
        styleBtns.Controls.AddRange([
            SmallBtn(Loc.T("Tạo mẫu mới...", "New style..."), "btnNewStyle", (_, _) => EditStyle(null)),
            SmallBtn(Loc.T("Sửa mẫu...", "Edit style..."), "btnEditStyle", (_, _) => EditStyle(_style.SelectedItem as SignatureStyle)),
            SmallBtn(Loc.T("Xóa mẫu", "Delete style"), "btnDeleteStyle", (_, _) => DeleteStyle())]);
        left.Controls.Add(new Label()); left.Controls.Add(styleBtns);
        _styleHint = new Label { AutoSize = true, MaximumSize = new Size(Dpi.Scale(390), 0), ForeColor = Theme.WarnAmber, Name = "lblStyleHint" };
        left.Controls.Add(new Label()); left.Controls.Add(_styleHint);

        _reason = new ComboBox { Width = Dpi.Scale(390), Name = "cboReason" };
        _reason.Items.AddRange([
            Loc.T("Tôi đồng ý với nội dung văn bản", "I agree with the content of this document"), Loc.T("Tôi là tác giả của tài liệu này", "I am the author of this document"), Loc.T("Tôi phê duyệt tài liệu này", "I approve this document"),
            Loc.T("Tôi đã xem xét tài liệu này", "I have reviewed this document"), Loc.T("Xác nhận đã nhận văn bản", "Receipt confirmed"), Loc.T("Ký nháy", "Initialled"), Loc.T("Ký thay", "Signed on behalf"), Loc.T("Sao y bản chính", "Certified true copy")]);
        _reason.Text = settings.LastReason;
        left.Controls.Add(L(Loc.T("Lý do:", "Reason:"))); left.Controls.Add(_reason);
        _location = new TextBox { Width = Dpi.Scale(390), Text = settings.LastLocation, Name = "txtLocation", PlaceholderText = Loc.T("VD: Hà Nội", "e.g. Hanoi") };
        left.Controls.Add(L(Loc.T("Vị trí:", "Location:"))); left.Controls.Add(_location);
        _contact = new TextBox { Width = Dpi.Scale(390), Text = settings.LastContact, Name = "txtContact", PlaceholderText = Loc.T("Email / số điện thoại", "Email / phone number") };
        left.Controls.Add(L(Loc.T("Thông tin liên hệ:", "Contact info:"))); left.Controls.Add(_contact);
        _standard = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = Dpi.Scale(390), Name = "cboSignStandard" };
        _standard.Items.AddRange([Loc.T("PAdES — ETSI.CAdES.detached (khuyến nghị)", "PAdES — ETSI.CAdES.detached (recommended)"), "PKCS#7 — adbe.pkcs7.detached"]);
        _standard.SelectedIndex = settings.Standard == SignatureStandard.Pkcs7 ? 1 : 0;
        left.Controls.Add(L(Loc.T("Chuẩn chữ ký:", "Signature standard:"))); left.Controls.Add(_standard);
        var tsRow = new FlowLayoutPanel { AutoSize = true, Margin = Padding.Empty };
        _useTs = new CheckBox { Text = Loc.T("Dấu thời gian", "Timestamp"), Checked = settings.UseTimestamp, AutoSize = true, Name = "chkUseTimestamp" };
        _tsa = new TextBox { Width = Dpi.Scale(270), Text = settings.TsaUrl, Name = "txtTsa" };
        tsRow.Controls.AddRange([_useTs, _tsa]);
        left.Controls.Add(L("TSA:")); left.Controls.Add(tsRow);

        var right = new Panel { Dock = DockStyle.Fill, Padding = new Padding(Dpi.Scale(14), 0, 0, 0) };
        var pvl = new Label { Text = Loc.T("Xem trước hình chữ ký", "Signature appearance preview"), Dock = DockStyle.Top, Font = Theme.UiFontBold, Height = Dpi.Scale(24) };
        _preview = new PictureBox { Dock = DockStyle.Top, Height = Dpi.Scale(200), SizeMode = PictureBoxSizeMode.Zoom, BackColor = Color.White, BorderStyle = BorderStyle.FixedSingle, Name = "picSignPreview" };
        var hint = new Label
        {
            Dock = DockStyle.Top, Height = Dpi.Scale(170), ForeColor = Theme.SubtleInk, Padding = new Padding(0, Dpi.Scale(8), 0, 0),
            Text = Loc.T("Sau khi bấm \"Đặt chữ ký\", khung chữ ký xuất hiện trên trang: kéo để di chuyển, kéo các ô vuông để đổi kích thước, rồi bấm \"Ký và lưu\".\n\n", "After clicking \"Place signature\", the signature box appears on the page: drag to move it, drag the handles to resize it, then click \"Sign and save\".\n\n") +
                   Loc.T("Nền tảng ký được ghi nhận: ", "Recorded signing platform: ") + AppearanceRenderer.PlatformFullName + Loc.T(".\n\nUSB Token: cắm Token và cài driver của nhà cung cấp (VNPT-CA, Viettel-CA, FPT-CA, BKAV-CA...) — chứng thư sẽ hiện trong danh sách. Mã PIN được hỏi khi ký.", ".\n\nUSB token: plug in the token and install the vendor driver (VNPT-CA, Viettel-CA, FPT-CA, BKAV-CA...) — its certificate will appear in the list. The PIN is asked when signing.")
        };
        right.Controls.Add(hint);
        right.Controls.Add(_preview);
        right.Controls.Add(pvl);
        Body.Controls.Add(right);
        Body.Controls.Add(left);

        LoadCertificates();
        LoadStyles(settings.LastStyleId);
        _cert.SelectedIndexChanged += (_, _) => OnCertChanged();
        _style.SelectedIndexChanged += (_, _) => { LoadStyleImage(); UpdatePreview(); };
        _reason.TextChanged += (_, _) => UpdatePreview();
        _location.TextChanged += (_, _) => UpdatePreview();
        _contact.TextChanged += (_, _) => UpdatePreview();
        LoadStyleImage();
        UpdateCertInfo();
        UpdatePreview();

        AddOkCancel(Loc.T("Đặt chữ ký", "Place signature"), Loc.T("Hủy", "Cancel"), (_, _) =>
        {
            if (SelectedCertificate == null && _cert.SelectedItem is CertChoice { PfxPath: not null } && !UnlockSelected()) return;
            if (SelectedCertificate == null) { Msg.Error(this, Loc.T("Vui lòng chọn chứng thư số để ký (USB Token, chứng thư Windows hoặc tệp .pfx).", "Please choose a certificate to sign with (USB token, Windows certificate or .pfx file).")); return; }
            if (DateTime.Now > SelectedCertificate.NotAfter &&
                !Msg.Confirm(this, Loc.T("Chứng thư số đã HẾT HẠN. Chữ ký sẽ bị đánh giá không hợp lệ. Vẫn tiếp tục?", "The certificate has EXPIRED. The signature will be considered invalid. Continue anyway?"),
                    Loc.T("Vẫn tiếp tục", "Continue"), Loc.T("Hủy", "Cancel"), Loc.T("Cảnh báo", "Warning"), warning: true)) return;
            _settings.LastReason = Reason;
            _settings.LastLocation = SignLocation;
            _settings.LastContact = Contact;
            _settings.LastCertThumbprint = SelectedCertificate.Thumbprint;
            _settings.LastStyleId = SelectedStyle.Id;
            _settings.Standard = Standard;
            _settings.UseTimestamp = _useTs.Checked;
            if (!string.IsNullOrWhiteSpace(_tsa.Text)) _settings.TsaUrl = _tsa.Text.Trim();
            _settings.Save();
            DialogResult = DialogResult.OK;
        });
    }

    private Button SmallBtn(string text, string name, EventHandler click)
    {
        var b = new Button { Text = text, Name = name, AutoSize = true, FlatStyle = FlatStyle.Flat, BackColor = Color.White, Margin = new Padding(0, 2, 4, 2) };
        b.FlatAppearance.BorderColor = Theme.Border;
        b.Click += click;
        return b;
    }

    // ------------------------------------------------------------------ certificates

    private void LoadCertificates()
    {
        var keepThumb = SelectedCertificate?.Thumbprint ?? _settings.LastCertThumbprint;
        var loaded = _cert.Items.OfType<CertChoice>().Where(c => c.Certificate != null && c.PfxPath != null).ToList();
        _cert.Items.Clear();
        foreach (var e in CertificateService.GetStoreCertificates())
            _cert.Items.Add(new CertChoice { Certificate = e.Certificate, Source = e.Source });
        var pfxPaths = new List<string>(_settings.PfxFiles);
        if (Directory.Exists(AppSettings.IdsDirectory)) pfxPaths.AddRange(Directory.GetFiles(AppSettings.IdsDirectory, "*.pfx"));
        foreach (var p in pfxPaths.Where(File.Exists).Distinct(StringComparer.OrdinalIgnoreCase))
        {
            var already = loaded.FirstOrDefault(l => string.Equals(l.PfxPath, p, StringComparison.OrdinalIgnoreCase));
            _cert.Items.Add(already ?? new CertChoice { PfxPath = p, Source = Loc.T("Tệp PFX", "PFX file") });
        }
        int sel = -1;
        for (int i = 0; i < _cert.Items.Count; i++)
            if (_cert.Items[i] is CertChoice { Certificate: { } c } && c.Thumbprint == keepThumb) { sel = i; break; }
        if (sel < 0) for (int i = 0; i < _cert.Items.Count; i++) if (_cert.Items[i] is CertChoice { Certificate: not null }) { sel = i; break; }
        if (sel < 0 && _cert.Items.Count > 0) sel = 0; // a locked .pfx: password is asked when signing
        _suppressPrompt = true;
        _cert.SelectedIndex = sel;
        _suppressPrompt = false;
        UpdateCertInfo();
        _lastCertIndex = sel;
        if (_cert.Items.Count == 0) _certInfo.Text = Loc.T("Không tìm thấy chứng thư số. Hãy cắm USB Token rồi bấm \"Làm mới\", nạp tệp .pfx hoặc tạo ID số mới.", "No certificate found. Plug in your USB token and click \"Refresh\", load a .pfx file or create a new digital ID.");
    }

    private void OnCertChanged()
    {
        if (!_suppressPrompt && _cert.SelectedItem is CertChoice { Certificate: null, PfxPath: not null })
        {
            if (!UnlockSelected()) { _cert.SelectedIndex = _lastCertIndex; return; }
        }
        _lastCertIndex = _cert.SelectedIndex;
        UpdateCertInfo();
        UpdatePreview();
    }

    /// <summary>Asks the password of the selected .pfx and loads it. Returns false if cancelled or wrong.</summary>
    private bool UnlockSelected()
    {
        if (_cert.SelectedItem is CertChoice { Certificate: null, PfxPath: { } path } choice)
        {
            var pw = InputDialog.Ask(this, Loc.T("Mở khóa tệp chứng thư", "Unlock certificate file"), Loc.T($"Nhập mật khẩu của tệp:\n{Path.GetFileName(path)}", $"Enter the password of the file:\n{Path.GetFileName(path)}"), "", password: true);
            if (pw == null) return false;
            try { choice.Certificate = CertificateService.LoadPfx(path, pw); }
            catch (Exception ex) { Msg.Error(this, ex.Message); return false; }
            int idx = _cert.SelectedIndex;
            _suppressPrompt = true;
            _cert.Items[idx] = choice; // refresh text
            _cert.SelectedIndex = idx;
            _suppressPrompt = false;
            _lastCertIndex = idx;
            UpdateCertInfo();
            UpdatePreview();
        }
        return SelectedCertificate != null;
    }

    private void UpdateCertInfo()
    {
        var c = SelectedCertificate;
        if (c == null)
        {
            if (_cert.SelectedItem is CertChoice { PfxPath: { } p }) _certInfo.Text = Loc.T($"Tệp chứng thư {Path.GetFileName(p)} đang khóa — mật khẩu sẽ được hỏi khi bấm \"Đặt chữ ký\".", $"Certificate file {Path.GetFileName(p)} is locked — the password will be asked when you click \"Place signature\".");
            return;
        }
        _certInfo.Text = Loc.T($"Chủ thể: {c.GetNameInfo(X509NameType.SimpleName, false)}\nCấp bởi: {c.GetNameInfo(X509NameType.SimpleName, true)}\n", $"Subject: {c.GetNameInfo(X509NameType.SimpleName, false)}\nIssued by: {c.GetNameInfo(X509NameType.SimpleName, true)}\n") +
                         Loc.T($"Hiệu lực: {c.NotBefore:dd/MM/yyyy} – {c.NotAfter:dd/MM/yyyy}", $"Valid: {c.NotBefore:dd/MM/yyyy} – {c.NotAfter:dd/MM/yyyy}") + (DateTime.Now > c.NotAfter ? Loc.T("  (ĐÃ HẾT HẠN)", "  (EXPIRED)") : "");
    }

    /// <summary>Adds a PFX file (used by the UI and by automated tests).</summary>
    public bool AddPfx(string path, string password)
    {
        try
        {
            var cert = CertificateService.LoadPfx(path, password);
            if (!_settings.PfxFiles.Contains(path, StringComparer.OrdinalIgnoreCase)) { _settings.PfxFiles.Add(path); _settings.Save(); }
            var existing = _cert.Items.OfType<CertChoice>().FirstOrDefault(c => string.Equals(c.PfxPath, path, StringComparison.OrdinalIgnoreCase));
            var choice = new CertChoice { Certificate = cert, PfxPath = path, Source = Loc.T("Tệp PFX", "PFX file") };
            if (existing != null) _cert.Items[_cert.Items.IndexOf(existing)] = choice;
            else _cert.Items.Add(choice);
            _cert.SelectedItem = choice;
            return true;
        }
        catch (Exception ex) { Msg.Error(this, ex.Message); return false; }
    }

    private void LoadPfx()
    {
        using var ofd = new OpenFileDialog { Filter = Loc.T("Chứng thư số (*.pfx;*.p12)|*.pfx;*.p12", "Certificates (*.pfx;*.p12)|*.pfx;*.p12"), Title = Loc.T("Chọn tệp chứng thư số", "Choose a certificate file") };
        if (ofd.ShowDialog(this) != DialogResult.OK) return;
        var pw = InputDialog.Ask(this, Loc.T("Mật khẩu chứng thư", "Certificate password"), Loc.T("Nhập mật khẩu của tệp .pfx/.p12:", "Enter the password of the .pfx/.p12 file:"), "", password: true);
        if (pw == null) return;
        AddPfx(ofd.FileName, pw);
    }

    private void CreateId()
    {
        using var d = new CreateIdDialog();
        if (d.ShowDialog(this) != DialogResult.OK || d.Certificate == null) return;
        var choice = new CertChoice { Certificate = d.Certificate, PfxPath = d.PfxPath, Source = Loc.T("ID tự ký", "Self-signed ID") };
        _cert.Items.Add(choice);
        _cert.SelectedItem = choice;
        Msg.Info(this, Loc.T($"Đã tạo ID số và lưu tại:\n{d.PfxPath}", $"Digital ID created and saved to:\n{d.PfxPath}"));
    }

    // ------------------------------------------------------------------ styles

    private void LoadStyles(string? selectId)
    {
        _style.Items.Clear();
        foreach (var s in _settings.Styles) _style.Items.Add(s);
        _style.DisplayMember = nameof(SignatureStyle.Title);
        int idx = _settings.Styles.FindIndex(s => s.Id == selectId);
        _style.SelectedIndex = idx >= 0 ? idx : (_style.Items.Count > 0 ? 0 : -1);
    }

    private void EditStyle(SignatureStyle? s)
    {
        using var d = new SignatureStyleDialog(s);
        if (d.ShowDialog(this) != DialogResult.OK) return;
        int idx = s == null ? -1 : _settings.Styles.FindIndex(x => x.Id == s.Id);
        if (idx >= 0) _settings.Styles[idx] = d.Style;
        else _settings.Styles.Add(d.Style);
        _settings.Save();
        LoadStyles(d.Style.Id);
        LoadStyleImage();
        UpdatePreview();
    }

    /// <summary>For automated tests: adds a style programmatically.</summary>
    public void AddStyle(SignatureStyle s)
    {
        _settings.Styles.Add(s);
        _settings.Save();
        LoadStyles(s.Id);
        LoadStyleImage();
        UpdatePreview();
    }

    private void DeleteStyle()
    {
        if (_style.SelectedItem is not SignatureStyle s) return;
        if (_settings.Styles.Count <= 1) { Msg.Info(this, Loc.T("Cần giữ lại ít nhất một mẫu chữ ký.", "At least one signature style must be kept.")); return; }
        if (!Msg.Confirm(this, Loc.T($"Xóa mẫu \"{s.Title}\"?", $"Delete style \"{s.Title}\"?"), Loc.T("Xóa mẫu", "Delete style"), Loc.T("Hủy", "Cancel"))) return;
        _settings.Styles.RemoveAll(x => x.Id == s.Id);
        _settings.Save();
        LoadStyles(null);
    }

    private void LoadStyleImage()
    {
        _styleImage?.Dispose();
        _styleImage = StyleImages.Load(SelectedStyle);
        bool needsImage = SelectedStyle.Graphic is SignatureGraphic.Image or SignatureGraphic.Drawn;
        _styleHint.Text = needsImage && _styleImage == null ? Loc.T("Mẫu này chưa có ảnh con dấu/chữ ký — bấm \"Sửa mẫu...\" để chọn ảnh.", "This style has no seal/signature image yet — click \"Edit style...\" to choose one.") : "";
    }

    private void UpdatePreview()
    {
        var c = SelectedCertificate;
        var info = new SignatureAppearanceInfo
        {
            SignerName = c != null ? PdfSigner.SignerDisplayName(c) : Loc.T("(chưa chọn chứng thư)", "(no certificate selected)"),
            DistinguishedName = c?.Subject ?? "",
            SigningTime = DateTimeOffset.Now,
            Reason = Reason, Location = SignLocation, Contact = Contact, Graphic = _styleImage
        };
        float k = Math.Min(3f, 600f / Math.Max(_boxPts.Width, _boxPts.Height));
        var old = _preview.Image;
        _preview.Image = AppearanceRenderer.RenderBitmap(SelectedStyle, info, _boxPts.Width, _boxPts.Height, k, transparent: false);
        old?.Dispose();
    }

    public SignatureAppearanceInfo BuildInfo() => new()
    {
        SignerName = SelectedCertificate != null ? PdfSigner.SignerDisplayName(SelectedCertificate) : "",
        DistinguishedName = SelectedCertificate?.Subject ?? "",
        SigningTime = DateTimeOffset.Now,
        Reason = Reason, Location = SignLocation, Contact = Contact, Graphic = _styleImage
    };

    /// <summary>Detach the style image so it outlives the dialog.</summary>
    public Image? TakeStyleImage()
    {
        var img = _styleImage;
        _styleImage = null;
        return img;
    }

    protected override void Dispose(bool disposing)
    {
        if (disposing) { _styleImage?.Dispose(); _preview.Image?.Dispose(); }
        base.Dispose(disposing);
    }
}
