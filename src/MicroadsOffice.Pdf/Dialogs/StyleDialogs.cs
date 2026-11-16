using MicroadsOffice.Core.Office;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using MicroadsOffice.Core;
using MicroadsOffice.Core.Signing;
using MicroadsOffice.UI;

namespace MicroadsOffice.Dialogs;

public static class StyleImages
{
    /// <summary>Loads an image fully into memory (no file lock).</summary>
    public static Image? LoadFile(string path)
    {
        try
        {
            var bytes = File.ReadAllBytes(path);
            using var ms = new MemoryStream(bytes);
            using var img = Image.FromStream(ms);
            return new Bitmap(img);
        }
        catch { return null; }
    }

    public static Image? Load(SignatureStyle s) =>
        s.Graphic is SignatureGraphic.Image or SignatureGraphic.Drawn && !string.IsNullOrEmpty(s.ImageFile)
            ? LoadFile(Path.Combine(AppSettings.StylesDirectory, s.ImageFile!)) : null;

    public static string Import(string sourcePath)
    {
        Directory.CreateDirectory(AppSettings.StylesDirectory);
        var name = Guid.NewGuid().ToString("N") + ".png";
        using var img = LoadFile(sourcePath) ?? throw new InvalidOperationException(Loc.T("Không đọc được ảnh.", "Could not read the image."));
        img.Save(Path.Combine(AppSettings.StylesDirectory, name), ImageFormat.Png);
        return name;
    }

    public static string Save(Image img)
    {
        Directory.CreateDirectory(AppSettings.StylesDirectory);
        var name = Guid.NewGuid().ToString("N") + ".png";
        img.Save(Path.Combine(AppSettings.StylesDirectory, name), ImageFormat.Png);
        return name;
    }

    public static string ImageFilter => Loc.T("Ảnh (*.png;*.jpg;*.jpeg;*.bmp;*.gif)|*.png;*.jpg;*.jpeg;*.bmp;*.gif|Tất cả tệp|*.*", "Images (*.png;*.jpg;*.jpeg;*.bmp;*.gif)|*.png;*.jpg;*.jpeg;*.bmp;*.gif|All files|*.*");
}

/// <summary>Create / edit a signature appearance style (Foxit "Create Signature Style").</summary>
public sealed class SignatureStyleDialog : DialogBase
{
    public SignatureStyle Style { get; }
    private Image? _image;
    private readonly PictureBox _preview;
    private readonly TextBox _title, _custom, _dateFmt;
    private readonly RadioButton _gNone, _gImage, _gName, _gDrawn;
    private readonly ComboBox _layout, _font;
    private readonly NumericUpDown _size;
    private readonly TrackBar _opacity;
    private readonly CheckBox _name, _date, _reason, _location, _dn, _contact, _platform, _labels, _bold, _border;
    private readonly Button _color;
    private readonly Label _imgLabel;
    private bool _loading = true;

    public SignatureStyleDialog(SignatureStyle? existing) : base(existing == null ? Loc.T("Tạo mẫu chữ ký", "Create signature style") : Loc.T("Sửa mẫu chữ ký", "Edit signature style"), 820, 600)
    {
        Style = existing?.Clone() ?? new SignatureStyle { Title = Loc.T("Mẫu chữ ký mới", "New signature style") };
        _image = StyleImages.Load(Style);

        var left = new TableLayoutPanel { Dock = DockStyle.Left, Width = Dpi.Scale(430), ColumnCount = 2, AutoScroll = true };
        left.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        left.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        _title = new TextBox { Text = Style.Title, Width = Dpi.Scale(280), Name = "txtStyleTitle" };
        left.Controls.Add(L(Loc.T("Tên mẫu:", "Style name:"), true)); left.Controls.Add(_title);

        var gBox = new GroupBox { Text = Loc.T("Hình ảnh (đồ họa)", "Graphic"), Dock = DockStyle.Top, Height = Dpi.Scale(150), Width = Dpi.Scale(410) };
        var gFlow = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown, WrapContents = false };
        _gNone = new RadioButton { Text = Loc.T("Không có hình", "No graphic"), AutoSize = true, Name = "rdoGraphicNone" };
        var row1 = new FlowLayoutPanel { AutoSize = true, Margin = Padding.Empty };
        _gImage = new RadioButton { Text = Loc.T("Ảnh nhập (con dấu / chữ ký scan)", "Imported image (seal / scanned signature)"), AutoSize = true, Name = "rdoGraphicImage" };
        var btnImg = new Button { Text = Loc.T("Chọn ảnh...", "Choose image..."), AutoSize = true, Name = "btnChooseImage" };
        btnImg.Click += (_, _) => ChooseImage();
        row1.Controls.AddRange([_gImage, btnImg]);
        var row2 = new FlowLayoutPanel { AutoSize = true, Margin = Padding.Empty };
        _gDrawn = new RadioButton { Text = Loc.T("Chữ ký tay (vẽ)", "Handwritten (drawn)"), AutoSize = true, Name = "rdoGraphicDrawn" };
        var btnDraw = new Button { Text = Loc.T("Vẽ chữ ký...", "Draw signature..."), AutoSize = true, Name = "btnDrawSignature" };
        btnDraw.Click += (_, _) => DrawSignature();
        row2.Controls.AddRange([_gDrawn, btnDraw]);
        _gName = new RadioButton { Text = Loc.T("Tên người ký (chữ nghệ thuật)", "Signer name (script font)"), AutoSize = true, Name = "rdoGraphicName" };
        _imgLabel = new Label { AutoSize = true, ForeColor = Theme.SubtleInk };
        gFlow.Controls.AddRange([_gNone, row1, row2, _gName, _imgLabel]);
        gBox.Controls.Add(gFlow);
        left.Controls.Add(gBox); left.SetColumnSpan(gBox, 2);

        _layout = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = Dpi.Scale(280), Name = "cboLayout" };
        _layout.Items.AddRange([Loc.T("Hình bên trái, chữ bên phải", "Graphic left, text right"), Loc.T("Hình phía trên, chữ phía dưới", "Graphic above, text below"), Loc.T("Hình làm nền (chữ đè lên con dấu)", "Graphic as background (text over the seal)"), Loc.T("Chữ bên trái, hình bên phải", "Text left, graphic right")]);
        left.Controls.Add(L(Loc.T("Bố cục:", "Layout:"))); left.Controls.Add(_layout);
        _opacity = new TrackBar { Minimum = 10, Maximum = 100, TickFrequency = 10, Width = Dpi.Scale(280), Name = "trkOpacity", AutoSize = false, Height = Dpi.Scale(32) };
        left.Controls.Add(L(Loc.T("Độ đậm hình:", "Graphic opacity:"))); left.Controls.Add(_opacity);

        var tBox = new GroupBox { Text = Loc.T("Thông tin hiển thị", "Displayed information"), Dock = DockStyle.Top, Height = Dpi.Scale(124), Width = Dpi.Scale(410) };
        var tFlow = new FlowLayoutPanel { Dock = DockStyle.Fill };
        _name = new CheckBox { Text = Loc.T("Tên người ký", "Signer name"), AutoSize = true, Name = "chkShowName" };
        _date = new CheckBox { Text = Loc.T("Ngày ký", "Signing date"), AutoSize = true, Name = "chkShowDate" };
        _reason = new CheckBox { Text = Loc.T("Lý do", "Reason"), AutoSize = true, Name = "chkShowReason" };
        _location = new CheckBox { Text = Loc.T("Vị trí", "Location"), AutoSize = true, Name = "chkShowLocation" };
        _contact = new CheckBox { Text = Loc.T("Liên hệ", "Contact"), AutoSize = true, Name = "chkShowContact" };
        _dn = new CheckBox { Text = Loc.T("Tên phân biệt (DN)", "Distinguished name (DN)"), AutoSize = true, Name = "chkShowDN" };
        _platform = new CheckBox { Text = Loc.T("Nền tảng Microads Office", "Microads Office platform"), AutoSize = true, Name = "chkShowPlatform" };
        _labels = new CheckBox { Text = Loc.T("Nhãn (Ký bởi:, Ngày ký:...)", "Labels (Signed by:, Date:...)"), AutoSize = true, Name = "chkShowLabels" };
        tFlow.Controls.AddRange([_name, _date, _reason, _location, _contact, _dn, _platform, _labels]);
        tBox.Controls.Add(tFlow);
        left.Controls.Add(tBox); left.SetColumnSpan(tBox, 2);

        _custom = new TextBox { Width = Dpi.Scale(280), Name = "txtCustomText", PlaceholderText = Loc.T("VD: GIÁM ĐỐC", "e.g. DIRECTOR") };
        left.Controls.Add(L(Loc.T("Dòng chữ thêm:", "Extra text:"))); left.Controls.Add(_custom);
        var fontRow = new FlowLayoutPanel { AutoSize = true, Margin = Padding.Empty };
        _font = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = Dpi.Scale(150), Name = "cboFont" };
        foreach (var f in new[] { "Arial", "Times New Roman", "Tahoma", "Segoe UI", "Calibri", "Verdana" }) _font.Items.Add(f);
        _size = new NumericUpDown { Minimum = 0, Maximum = 48, DecimalPlaces = 1, Increment = 0.5m, Width = Dpi.Scale(56), Name = "numFontSize" };
        _bold = new CheckBox { Text = Loc.T("Đậm", "Bold"), AutoSize = true, Name = "chkBold" };
        fontRow.Controls.AddRange([_font, _size, _bold]);
        left.Controls.Add(L(Loc.T("Phông / cỡ (0=tự động):", "Font / size (0=auto):"))); left.Controls.Add(fontRow);
        var colorRow = new FlowLayoutPanel { AutoSize = true, Margin = Padding.Empty };
        _color = new Button { Text = Loc.T("Màu chữ", "Text colour"), AutoSize = true, Name = "btnTextColor" };
        _color.Click += (_, _) =>
        {
            using var cd = new ColorDialog { Color = Color.FromArgb(Style.TextColorArgb), FullOpen = true };
            if (cd.ShowDialog(this) == DialogResult.OK) { Style.TextColorArgb = cd.Color.ToArgb(); UpdatePreview(); }
        };
        _border = new CheckBox { Text = Loc.T("Viền khung", "Border"), AutoSize = true, Name = "chkBorder" };
        _dateFmt = new TextBox { Width = Dpi.Scale(130), Name = "txtDateFormat" };
        colorRow.Controls.AddRange([_color, _border]);
        left.Controls.Add(L(Loc.T("Màu / viền:", "Colour / border:"))); left.Controls.Add(colorRow);
        left.Controls.Add(L(Loc.T("Định dạng ngày:", "Date format:"))); left.Controls.Add(_dateFmt);

        var right = new Panel { Dock = DockStyle.Fill, Padding = new Padding(Dpi.Scale(12), 0, 0, 0) };
        var pvLabel = new Label { Text = Loc.T("Xem trước", "Preview"), Dock = DockStyle.Top, Font = Theme.UiFontBold, Height = Dpi.Scale(24) };
        _preview = new PictureBox { Dock = DockStyle.Top, Height = Dpi.Scale(150), BorderStyle = BorderStyle.FixedSingle, BackColor = Color.White, SizeMode = PictureBoxSizeMode.Zoom, Name = "picStylePreview" };
        var tip = new Label
        {
            Dock = DockStyle.Top, Height = Dpi.Scale(160), ForeColor = Theme.SubtleInk, Padding = new Padding(0, Dpi.Scale(10), 0, 0),
            Text = Loc.T("Mẹo:\n• Dùng ảnh PNG nền trong suốt cho con dấu để chữ hiển thị rõ.\n• Kiểu \"Hình làm nền\" phù hợp văn bản hành chính: con dấu đỏ nằm dưới thông tin người ký.\n• Kích thước khung chữ ký được điều chỉnh trực tiếp trên trang khi ký.", "Tips:\n• Use a PNG seal with a transparent background so the text stays readable.\n• \"Graphic as background\" suits Vietnamese administrative documents: the red seal sits under the signer details.\n• The size of the signature box is adjusted directly on the page when signing.")
        };
        right.Controls.Add(tip);
        right.Controls.Add(_preview);
        right.Controls.Add(pvLabel);
        Body.Controls.Add(right);
        Body.Controls.Add(left);

        LoadFromStyle();
        foreach (var c in new Control[] { _gNone, _gImage, _gName, _gDrawn, _name, _date, _reason, _location, _dn, _contact, _platform, _labels, _bold, _border })
        {
            if (c is RadioButton rb) rb.CheckedChanged += (_, _) => UpdatePreview();
            if (c is CheckBox cb) cb.CheckedChanged += (_, _) => UpdatePreview();
        }
        _layout.SelectedIndexChanged += (_, _) => UpdatePreview();
        _font.SelectedIndexChanged += (_, _) => UpdatePreview();
        _size.ValueChanged += (_, _) => UpdatePreview();
        _opacity.ValueChanged += (_, _) => UpdatePreview();
        _custom.TextChanged += (_, _) => UpdatePreview();
        _dateFmt.TextChanged += (_, _) => UpdatePreview();
        _title.TextChanged += (_, _) => Style.Title = _title.Text;
        // The graphic radio buttons live in different containers: enforce a single choice manually.
        var radios = new[] { _gNone, _gImage, _gName, _gDrawn };
        foreach (var rb in radios)
            rb.CheckedChanged += (s, _) =>
            {
                if (s is RadioButton { Checked: true } on)
                    foreach (var other in radios) if (other != on && other.Checked) other.Checked = false;
            };
        _loading = false;
        UpdatePreview();

        AddOkCancel(Loc.T("Lưu mẫu", "Save style"), Loc.T("Hủy", "Cancel"), (_, _) =>
        {
            ReadToStyle();
            if (Style.Graphic is SignatureGraphic.Image or SignatureGraphic.Drawn && string.IsNullOrEmpty(Style.ImageFile))
            {
                Msg.Error(this, Loc.T("Bạn đã chọn hình ảnh nhưng chưa chọn ảnh / chưa vẽ chữ ký.", "You chose a graphic but have not picked an image / drawn a signature yet."));
                return;
            }
            if (string.IsNullOrWhiteSpace(Style.Title)) Style.Title = Loc.T("Mẫu chữ ký", "Signature styles");
            DialogResult = DialogResult.OK;
        });
    }

    private void LoadFromStyle()
    {
        _gNone.Checked = Style.Graphic == SignatureGraphic.None;
        _gImage.Checked = Style.Graphic == SignatureGraphic.Image;
        _gName.Checked = Style.Graphic == SignatureGraphic.Name;
        _gDrawn.Checked = Style.Graphic == SignatureGraphic.Drawn;
        _layout.SelectedIndex = (int)Style.Layout;
        _opacity.Value = (int)Math.Clamp(Style.ImageOpacity * 100, 10, 100);
        _name.Checked = Style.ShowName; _date.Checked = Style.ShowDate; _reason.Checked = Style.ShowReason;
        _location.Checked = Style.ShowLocation; _dn.Checked = Style.ShowDistinguishedName; _contact.Checked = Style.ShowContact;
        _platform.Checked = Style.ShowPlatform; _labels.Checked = Style.ShowLabels; _bold.Checked = Style.Bold; _border.Checked = Style.Border;
        _custom.Text = Style.CustomText;
        _font.SelectedItem = _font.Items.Contains(Style.FontFamily) ? Style.FontFamily : "Arial";
        _size.Value = (decimal)Math.Clamp(Style.FontSize, 0, 48);
        _dateFmt.Text = Style.DateFormat;
        _imgLabel.Text = _image != null ? Loc.T($"Ảnh hiện tại: {_image.Width}×{_image.Height}px", $"Current image: {_image.Width}×{_image.Height}px") : "";
    }

    private void ReadToStyle()
    {
        Style.Title = _title.Text.Trim();
        Style.Graphic = _gImage.Checked ? SignatureGraphic.Image : _gName.Checked ? SignatureGraphic.Name : _gDrawn.Checked ? SignatureGraphic.Drawn : SignatureGraphic.None;
        Style.Layout = (SignatureLayout)Math.Max(0, _layout.SelectedIndex);
        Style.ImageOpacity = _opacity.Value / 100f;
        Style.ShowName = _name.Checked; Style.ShowDate = _date.Checked; Style.ShowReason = _reason.Checked;
        Style.ShowLocation = _location.Checked; Style.ShowDistinguishedName = _dn.Checked; Style.ShowContact = _contact.Checked;
        Style.ShowPlatform = _platform.Checked; Style.ShowLabels = _labels.Checked; Style.Bold = _bold.Checked; Style.Border = _border.Checked;
        Style.CustomText = _custom.Text;
        Style.FontFamily = _font.SelectedItem as string ?? "Arial";
        Style.FontSize = (float)_size.Value;
        Style.DateFormat = string.IsNullOrWhiteSpace(_dateFmt.Text) ? "dd/MM/yyyy HH:mm:ss" : _dateFmt.Text;
    }

    private void ChooseImage()
    {
        using var ofd = new OpenFileDialog { Filter = StyleImages.ImageFilter, Title = Loc.T("Chọn ảnh con dấu / chữ ký", "Choose a seal / signature image") };
        if (ofd.ShowDialog(this) != DialogResult.OK) return;
        try
        {
            Style.ImageFile = StyleImages.Import(ofd.FileName);
            _image?.Dispose();
            _image = StyleImages.Load(new SignatureStyle { Graphic = SignatureGraphic.Image, ImageFile = Style.ImageFile });
            _gImage.Checked = true;
            _imgLabel.Text = _image != null ? Loc.T($"Ảnh: {Path.GetFileName(ofd.FileName)} ({_image.Width}×{_image.Height}px)", $"Image: {Path.GetFileName(ofd.FileName)} ({_image.Width}×{_image.Height}px)") : "";
            UpdatePreview();
        }
        catch (Exception ex) { Msg.Error(this, ex.Message); }
    }

    private void DrawSignature()
    {
        using var d = new DrawSignatureDialog();
        if (d.ShowDialog(this) != DialogResult.OK || d.Result == null) return;
        Style.ImageFile = StyleImages.Save(d.Result);
        _image?.Dispose();
        _image = new Bitmap(d.Result);
        _gDrawn.Checked = true;
        _imgLabel.Text = Loc.T("Đã vẽ chữ ký tay.", "Handwritten signature drawn.");
        UpdatePreview();
    }

    private void UpdatePreview()
    {
        if (_loading) return;
        ReadToStyle();
        var info = new SignatureAppearanceInfo
        {
            SignerName = Loc.T("Nguyễn Văn An", "Nguyen Van An"), DistinguishedName = Loc.T("CN=Nguyễn Văn An, O=Công ty ABC, C=VN", "CN=Nguyen Van An, O=ABC Company, C=VN"), Reason = Loc.T("Tôi đồng ý với nội dung văn bản", "I agree with the content of this document"),
            Location = Loc.T("Hà Nội", "Hanoi"), Contact = "an@abc.vn", SigningTime = DateTimeOffset.Now, Graphic = _image
        };
        var old = _preview.Image;
        _preview.Image = AppearanceRenderer.RenderBitmap(Style, info, 240, 90, 2.2f, transparent: false);
        old?.Dispose();
    }

    protected override void Dispose(bool disposing)
    {
        if (disposing) { _image?.Dispose(); _preview.Image?.Dispose(); }
        base.Dispose(disposing);
    }
}

/// <summary>Freehand signature pad.</summary>
public sealed class DrawSignatureDialog : DialogBase
{
    private readonly List<List<PointF>> _strokes = new();
    private List<PointF>? _current;
    private readonly Panel _pad;
    private Color _ink = Color.FromArgb(20, 40, 120);
    private float _width = 3f;
    public Bitmap? Result { get; private set; }

    public DrawSignatureDialog() : base(Loc.T("Vẽ chữ ký tay", "Draw handwritten signature"), 640, 360)
    {
        _pad = new DoubleBufferedPanel { Dock = DockStyle.Fill, BackColor = Color.White, BorderStyle = BorderStyle.FixedSingle, Cursor = Cursors.Cross, Name = "pnlSignaturePad" };
        _pad.MouseDown += (_, e) => { if (e.Button == MouseButtons.Left) { _current = [e.Location]; _strokes.Add(_current); } };
        _pad.MouseMove += (_, e) => { if (_current != null && e.Button == MouseButtons.Left) { _current.Add(e.Location); _pad.Invalidate(); } };
        _pad.MouseUp += (_, _) => _current = null;
        _pad.Paint += (_, e) =>
        {
            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
            using var guide = new Pen(Color.FromArgb(220, 220, 220)) { DashStyle = DashStyle.Dash };
            e.Graphics.DrawLine(guide, 20, _pad.Height * 0.72f, _pad.Width - 20, _pad.Height * 0.72f);
            DrawStrokes(e.Graphics, 1f, PointF.Empty);
            if (_strokes.Count == 0)
                e.Graphics.DrawString(Loc.T("Ký vào đây bằng chuột hoặc bút cảm ứng", "Sign here with the mouse or a pen"), Theme.UiFont, Brushes.Silver, 24, 20);
        };
        var tools = new FlowLayoutPanel { Dock = DockStyle.Top, Height = Dpi.Scale(38) };
        var clear = new Button { Text = Loc.T("Xóa", "Delete"), AutoSize = true, Name = "btnClearPad" };
        clear.Click += (_, _) => { _strokes.Clear(); _pad.Invalidate(); };
        var colors = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = Dpi.Scale(110), Name = "cboInk" };
        colors.Items.AddRange([Loc.T("Xanh đậm", "Dark blue"), Loc.T("Đen", "Black"), Loc.T("Đỏ", "Red")]);
        colors.SelectedIndex = 0;
        colors.SelectedIndexChanged += (_, _) =>
        {
            _ink = colors.SelectedIndex switch { 1 => Color.Black, 2 => Color.FromArgb(200, 20, 30), _ => Color.FromArgb(20, 40, 120) };
            _pad.Invalidate();
        };
        var width = new NumericUpDown { Minimum = 1, Maximum = 10, Value = 3, Width = Dpi.Scale(50), Name = "numInkWidth" };
        width.ValueChanged += (_, _) => { _width = (float)width.Value; _pad.Invalidate(); };
        tools.Controls.AddRange([clear, new Label { Text = Loc.T("Màu mực:", "Ink colour:"), AutoSize = true, Margin = new Padding(8, 7, 0, 0) }, colors,
            new Label { Text = Loc.T("Độ dày:", "Thickness:"), AutoSize = true, Margin = new Padding(8, 7, 0, 0) }, width]);
        Body.Controls.Add(_pad);
        Body.Controls.Add(tools);
        AddOkCancel(Loc.T("Dùng chữ ký này", "Use this signature"), Loc.T("Hủy", "Cancel"), (_, _) =>
        {
            if (_strokes.Sum(s => s.Count) < 3) { Msg.Error(this, Loc.T("Vui lòng vẽ chữ ký trước.", "Please draw a signature first.")); return; }
            Result = Export();
            DialogResult = DialogResult.OK;
        });
    }

    private void DrawStrokes(Graphics g, float scale, PointF offset)
    {
        using var pen = new Pen(_ink, _width * scale) { StartCap = LineCap.Round, EndCap = LineCap.Round, LineJoin = LineJoin.Round };
        foreach (var s in _strokes)
        {
            var pts = s.Select(p => new PointF((p.X - offset.X) * scale, (p.Y - offset.Y) * scale)).ToArray();
            if (pts.Length == 1) g.FillEllipse(new SolidBrush(_ink), pts[0].X - _width * scale / 2, pts[0].Y - _width * scale / 2, _width * scale, _width * scale);
            else if (pts.Length == 2) g.DrawLine(pen, pts[0], pts[1]);
            else g.DrawCurve(pen, pts, 0.4f);
        }
    }

    /// <summary>Renders the strokes at 3x into a tightly cropped transparent bitmap.</summary>
    public Bitmap Export()
    {
        var all = _strokes.SelectMany(s => s).ToList();
        float pad = _width * 2 + 4;
        float minX = all.Min(p => p.X) - pad, minY = all.Min(p => p.Y) - pad, maxX = all.Max(p => p.X) + pad, maxY = all.Max(p => p.Y) + pad;
        const float k = 3f;
        var bmp = new Bitmap(Math.Max(1, (int)((maxX - minX) * k)), Math.Max(1, (int)((maxY - minY) * k)), PixelFormat.Format32bppArgb);
        using var g = Graphics.FromImage(bmp);
        g.Clear(Color.Transparent);
        g.SmoothingMode = SmoothingMode.AntiAlias;
        DrawStrokes(g, k, new PointF(minX, minY));
        return bmp;
    }

    /// <summary>For automated tests: adds a stroke programmatically.</summary>
    public void AddStroke(params PointF[] pts) { _strokes.Add(pts.ToList()); _pad.Invalidate(); }
}

public sealed class DoubleBufferedPanel : Panel
{
    public DoubleBufferedPanel() { DoubleBuffered = true; ResizeRedraw = true; }
}

/// <summary>Create a self-signed digital ID saved as .pfx in the portable Data folder.</summary>
public sealed class CreateIdDialog : DialogBase
{
    public System.Security.Cryptography.X509Certificates.X509Certificate2? Certificate { get; private set; }
    public string? PfxPath { get; private set; }

    public CreateIdDialog() : base(Loc.T("Tạo ID số tự ký mới", "Create a new self-signed digital ID"), 520, 420)
    {
        var g = Grid();
        var name = new TextBox { Width = Dpi.Scale(320), Name = "txtIdName" };
        var org = new TextBox { Width = Dpi.Scale(320), Name = "txtIdOrg" };
        var unit = new TextBox { Width = Dpi.Scale(320), Name = "txtIdUnit" };
        var email = new TextBox { Width = Dpi.Scale(320), Name = "txtIdEmail" };
        var country = new TextBox { Width = Dpi.Scale(60), Text = "VN", Name = "txtIdCountry", MaxLength = 2 };
        var years = new NumericUpDown { Minimum = 1, Maximum = 30, Value = 5, Width = Dpi.Scale(60), Name = "numIdYears" };
        var pw = new TextBox { Width = Dpi.Scale(200), UseSystemPasswordChar = true, Name = "txtIdPassword" };
        var pw2 = new TextBox { Width = Dpi.Scale(200), UseSystemPasswordChar = true, Name = "txtIdPassword2" };
        g.Controls.Add(L(Loc.T("Họ tên / Tên đơn vị *:", "Full name / Organisation name *:"))); g.Controls.Add(name);
        g.Controls.Add(L(Loc.T("Tổ chức:", "Organisation:"))); g.Controls.Add(org);
        g.Controls.Add(L(Loc.T("Phòng ban:", "Department:"))); g.Controls.Add(unit);
        g.Controls.Add(L("Email:")); g.Controls.Add(email);
        g.Controls.Add(L(Loc.T("Quốc gia:", "Country:"))); g.Controls.Add(country);
        g.Controls.Add(L(Loc.T("Thời hạn (năm):", "Validity (years):"))); g.Controls.Add(years);
        g.Controls.Add(L(Loc.T("Mật khẩu bảo vệ *:", "Protection password *:"))); g.Controls.Add(pw);
        g.Controls.Add(L(Loc.T("Nhập lại mật khẩu *:", "Confirm password *:"))); g.Controls.Add(pw2);
        var note = new Label
        {
            AutoSize = true, MaximumSize = new Size(Dpi.Scale(470), 0), ForeColor = Theme.SubtleInk, Margin = new Padding(0, Dpi.Scale(10), 0, 0),
            Text = Loc.T("ID số tự ký dùng cho nội bộ/thử nghiệm: chữ ký toàn vẹn nhưng người nhận sẽ thấy \"danh tính chưa xác thực\". ", "A self-signed ID is for internal use/testing: the signature is intact but recipients will see \"identity not verified\". ") +
                   Loc.T("Văn bản pháp lý nên ký bằng chứng thư số do CA được cấp phép (USB Token).", "Legal documents should be signed with a certificate from a licensed CA (USB token).")
        };
        g.Controls.Add(note); g.SetColumnSpan(note, 2);
        Body.Controls.Add(g);
        AddOkCancel(Loc.T("Tạo ID", "Create ID"), Loc.T("Hủy", "Cancel"), (_, _) =>
        {
            if (string.IsNullOrWhiteSpace(name.Text)) { Msg.Error(this, Loc.T("Vui lòng nhập họ tên.", "Please enter a name.")); return; }
            if (pw.Text.Length < 4) { Msg.Error(this, Loc.T("Mật khẩu tối thiểu 4 ký tự.", "The password needs at least 4 characters.")); return; }
            if (pw.Text != pw2.Text) { Msg.Error(this, Loc.T("Mật khẩu nhập lại không khớp.", "The passwords do not match.")); return; }
            try
            {
                var safe = string.Concat(name.Text.Trim().Select(c => Path.GetInvalidFileNameChars().Contains(c) ? '_' : c));
                PfxPath = Path.Combine(AppSettings.IdsDirectory, $"{safe}_{DateTime.Now:yyyyMMddHHmmss}.pfx");
                Certificate = CertificateService.CreateSelfSigned(name.Text, org.Text, unit.Text, email.Text, country.Text, (int)years.Value, pw.Text, PfxPath);
                DialogResult = DialogResult.OK;
            }
            catch (Exception ex) { Msg.Error(this, Loc.T("Không tạo được ID số: ", "Could not create the digital ID: ") + ex.Message); }
        });
    }
}

/// <summary>Choose an image for an image stamp (seal) and its opacity.</summary>
public sealed class StampDialog : DialogBase
{
    public Image? Image { get; private set; }
    public float StampOpacity => _opacity.Value / 100f;
    private readonly PictureBox _pic;
    private readonly TrackBar _opacity;

    public StampDialog(string? initialPath = null) : base(Loc.T("Chèn dấu hình ảnh", "Insert image stamp"), 480, 400)
    {
        var top = new FlowLayoutPanel { Dock = DockStyle.Top, Height = Dpi.Scale(40) };
        var choose = new Button { Text = Loc.T("Chọn ảnh...", "Choose image..."), AutoSize = true, Name = "btnStampChoose" };
        choose.Click += (_, _) =>
        {
            using var ofd = new OpenFileDialog { Filter = StyleImages.ImageFilter, Title = Loc.T("Chọn ảnh con dấu", "Choose a seal image") };
            if (ofd.ShowDialog(this) == DialogResult.OK) SetImage(StyleImages.LoadFile(ofd.FileName));
        };
        top.Controls.Add(choose);
        top.Controls.Add(new Label { Text = Loc.T("Độ đậm:", "Opacity:"), AutoSize = true, Margin = new Padding(10, 8, 0, 0) });
        _opacity = new TrackBar { Minimum = 10, Maximum = 100, Value = 100, TickFrequency = 10, Width = Dpi.Scale(200), AutoSize = false, Height = Dpi.Scale(32), Name = "trkStampOpacity" };
        top.Controls.Add(_opacity);
        _pic = new PictureBox { Dock = DockStyle.Fill, SizeMode = PictureBoxSizeMode.Zoom, BorderStyle = BorderStyle.FixedSingle, BackColor = Color.White, Name = "picStamp" };
        var hint = new Label { Dock = DockStyle.Bottom, Height = Dpi.Scale(40), ForeColor = Theme.SubtleInk, Text = Loc.T("Sau khi bấm \"Tiếp tục\", hãy kéo chuột vẽ vị trí dấu trên trang, điều chỉnh rồi bấm \"Chèn dấu\".", "After clicking \"Continue\", drag on the page to place the stamp, adjust it, then click \"Insert stamp\".") };
        Body.Controls.Add(_pic);
        Body.Controls.Add(hint);
        Body.Controls.Add(top);
        if (initialPath != null) SetImage(StyleImages.LoadFile(initialPath));
        AddOkCancel(Loc.T("Tiếp tục", "Continue"), Loc.T("Hủy", "Cancel"), (_, _) =>
        {
            if (Image == null) { Msg.Error(this, Loc.T("Vui lòng chọn ảnh.", "Please choose an image.")); return; }
            DialogResult = DialogResult.OK;
        });
    }

    public void SetImage(Image? img)
    {
        if (img == null) return;
        Image?.Dispose();
        Image = img;
        _pic.Image = img;
    }
}
