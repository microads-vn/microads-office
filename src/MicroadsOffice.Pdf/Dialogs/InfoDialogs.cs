using MicroadsOffice.Core.Office;
using System.Diagnostics;
using MicroadsOffice.Core;
using MicroadsOffice.Core.Branding;
using MicroadsOffice.Core.Signing;
using MicroadsOffice.UI;
using MicroadsOffice.Views;

namespace MicroadsOffice.Dialogs;

public sealed class HelpDialog : DialogBase
{
    public HelpDialog() : base(Loc.T("Hướng dẫn sử dụng & phím tắt — Microads Office PDF", "User guide & shortcuts — Microads Office PDF"), 680, 580)
    {
        var platform = AppearanceRenderer.PlatformFullName;
        var box = new RichTextBox
        {
            Dock = DockStyle.Fill, ReadOnly = true, BorderStyle = BorderStyle.None, BackColor = Color.White, Name = "txtHelp", Font = new Font("Segoe UI", 9.5f),
            Text = Loc.T(
                "MỞ TÀI LIỆU\n  • Tệp ▸ Mở (Ctrl+O), kéo-thả tệp PDF vào cửa sổ, hoặc chọn \"Mở bằng\" Microads Office PDF.\n  • Mỗi tài liệu mở trong một thẻ riêng (Ctrl+W để đóng thẻ, Ctrl+Tab chuyển thẻ).\n\n" +
                "ĐỌC & ĐIỀU HƯỚNG\n  • Công cụ Tay (kéo để cuộn) và Chọn văn bản (bôi đen rồi Ctrl+C).\n  • Phóng to/thu nhỏ: Ctrl+ / Ctrl- hoặc Ctrl + lăn chuột; Ctrl+0 vừa trang, Ctrl+1 kích thước thật, Ctrl+2 vừa chiều rộng.\n" +
                "  • Xoay trang: Ctrl+Shift+ (+/-).  Trang đầu/cuối: Home/End;  Trang trước/sau: PageUp/PageDown.\n  • Đi tới trang: Ctrl+G.  Tìm kiếm: Ctrl+F, F3 kết quả tiếp theo.\n  • Toàn màn hình: F11 (Esc để thoát).  Ngăn điều hướng: F4.\n\n" +
                "KÝ SỐ\n  1. Ký & Bảo vệ ▸ Ký số (Ctrl+Shift+S): con trỏ chuyển thành dấu +.\n  2. Kéo chuột vẽ khung chữ ký trên trang (hoặc bấm vào trường chữ ký có sẵn màu xanh).\n" +
                "  3. Chọn chứng thư số (USB Token / Windows / tệp .pfx), mẫu hiển thị, lý do, vị trí, chuẩn ký, dấu thời gian.\n  4. Bấm \"Đặt chữ ký\": có thể kéo di chuyển / kéo các ô vuông để đổi kích thước khung.\n" +
                "  5. Bấm \"Ký và lưu\" (Enter) → chọn nơi lưu tệp đã ký → nhập mã PIN Token nếu được hỏi.\n  • Mẫu chữ ký: chèn ảnh con dấu (PNG nền trong suốt), chữ ký tay vẽ hoặc tên; chọn thông tin hiển thị.\n" +
                "  • Chữ ký ghi nhận nền tảng \"" + platform + "\" trong thuộc tính chữ ký và trên hình chữ ký.\n\n" +
                "XÁC THỰC CHỮ KÝ\n  • Thanh thông báo phía trên trang cho biết tình trạng chữ ký; bấm vào chữ ký để xem chi tiết.\n  • Xanh: hợp lệ. Vàng: toàn vẹn nhưng chưa xác thực được danh tính (VD: chứng thư tự ký). Đỏ: không hợp lệ.\n\n" +
                "IN ẤN\n  • Ctrl+P: chọn máy in, số bản, phạm vi trang (VD: 1-3,5), vừa khổ giấy / kích thước thật, in màu xám, xem trước.\n\n" +
                "CHÈN DẤU HÌNH ẢNH\n  • Ký & Bảo vệ ▸ Chèn dấu ảnh: chọn ảnh, vẽ khung lên trang, điều chỉnh, rồi Lưu (Ctrl+S).\n\n" +
                "CẬP NHẬT & NGÔN NGỮ\n  • Trợ giúp ▸ Kiểm tra cập nhật: tải phiên bản mới nhất từ microads.vn/office.\n  • Nút lá cờ ở góc trên bên phải: đổi Tiếng Việt / English.\n\n" +
                "Microads Office PDF — phát hành bởi Công ty TNHH Công nghệ Microads, luôn luôn miễn phí. microads.vn\n",
                "OPENING DOCUMENTS\n  • File ▸ Open (Ctrl+O), drag & drop PDF files into the window, or choose \"Open with\" Microads Office PDF.\n  • Each document opens in its own tab (Ctrl+W closes a tab, Ctrl+Tab switches tabs).\n\n" +
                "READING & NAVIGATION\n  • Hand tool (drag to scroll) and Select text (highlight, then Ctrl+C).\n  • Zoom: Ctrl+ / Ctrl- or Ctrl + mouse wheel; Ctrl+0 fit page, Ctrl+1 actual size, Ctrl+2 fit width.\n" +
                "  • Rotate: Ctrl+Shift+ (+/-).  First/last page: Home/End;  Previous/next page: PageUp/PageDown.\n  • Go to page: Ctrl+G.  Search: Ctrl+F, F3 for the next result.\n  • Full screen: F11 (Esc to exit).  Navigation pane: F4.\n\n" +
                "DIGITAL SIGNING\n  1. Sign & Protect ▸ Sign (Ctrl+Shift+S): the cursor becomes a +.\n  2. Drag on the page to draw the signature box (or click an existing green signature field).\n" +
                "  3. Choose a certificate (USB token / Windows / .pfx file), appearance, reason, location, standard and timestamp.\n  4. Click \"Place signature\": drag to move the box, drag the handles to resize it.\n" +
                "  5. Click \"Sign and save\" (Enter) → choose where to save the signed file → enter the token PIN if asked.\n  • Signature styles: add a seal image (transparent PNG), a drawn signature or a name; choose the information shown.\n" +
                "  • Signatures record the platform \"" + platform + "\" in the signature properties and on the signature appearance.\n\n" +
                "VALIDATING SIGNATURES\n  • The bar above the page shows the signature status; click a signature for details.\n  • Green: valid. Yellow: intact but the identity is not verified (e.g. self-signed). Red: invalid.\n\n" +
                "PRINTING\n  • Ctrl+P: choose the printer, copies, page range (e.g. 1-3,5), fit to paper / actual size, grayscale, preview.\n\n" +
                "IMAGE STAMPS\n  • Sign & Protect ▸ Image stamp: choose an image, draw its box on the page, adjust, then Save (Ctrl+S).\n\n" +
                "UPDATES & LANGUAGE\n  • Help ▸ Check for updates: downloads the latest version from microads.vn/office.\n  • Flag button at the top right: switch between Tiếng Việt / English.\n\n" +
                "Microads Office PDF — published by Microads Technology Co., Ltd., always free. microads.vn\n")
        };
        Body.Controls.Add(box);
        AddButton(Loc.T("Đóng", "Close"), "btnClose", DialogResult.OK, true);
    }
}

public sealed class DocumentPropertiesDialog : DialogBase
{
    public DocumentPropertiesDialog(DocumentView v) : base(Loc.T("Thuộc tính tài liệu", "Document properties"), 560, 480)
    {
        var pdf = v.Pdf;
        var size = pdf.PageCount > 0 ? pdf.GetPageSize(0) : SizeF.Empty;
        string Mm(float pt) => (pt / 72f * 25.4f).ToString("0.#");
        var perms = pdf.GetPermissions();
        var rows = new List<(string, string)>
        {
            (Loc.T("Tên tệp", "File name"), v.DisplayName),
            (Loc.T("Đường dẫn", "Path"), v.FilePath ?? Loc.T("(chưa lưu)", "(not saved)")),
            (Loc.T("Dung lượng", "Size"), $"{v.Bytes.Length / 1024.0:0.0} KB ({v.Bytes.Length:N0} byte)"),
            (Loc.T("Phiên bản PDF", "PDF version"), pdf.GetFileVersion()),
            (Loc.T("Số trang", "Page number"), pdf.PageCount.ToString()),
            (Loc.T("Khổ trang 1", "Page 1 size"), $"{Mm(size.Width)} × {Mm(size.Height)} mm ({size.Width:0} × {size.Height:0} pt)"),
            (Loc.T("Tiêu đề", "Title"), pdf.GetMeta("Title")),
            (Loc.T("Tác giả", "Author"), pdf.GetMeta("Author")),
            (Loc.T("Chủ đề", "Subject"), pdf.GetMeta("Subject")),
            (Loc.T("Từ khóa", "Keywords"), pdf.GetMeta("Keywords")),
            (Loc.T("Ứng dụng tạo", "Creator"), pdf.GetMeta("Creator")),
            (Loc.T("Trình tạo PDF", "PDF producer"), pdf.GetMeta("Producer")),
            (Loc.T("Ngày tạo", "Created"), FormatDate(pdf.GetMeta("CreationDate"))),
            (Loc.T("Ngày sửa", "Modified"), FormatDate(pdf.GetMeta("ModDate"))),
            (Loc.T("Bảo mật", "Security"), pdf.IsEncrypted ? Loc.T("Có mã hóa / mật khẩu", "Encrypted / password") : Loc.T("Không", "No")),
            (Loc.T("Cho phép in", "Printing allowed"), (perms & 0x4) != 0 ? Loc.T("Có", "Yes") : Loc.T("Không", "No")),
            (Loc.T("Cho phép sao chép", "Copying allowed"), (perms & 0x10) != 0 ? Loc.T("Có", "Yes") : Loc.T("Không", "No")),
            (Loc.T("Chữ ký số", "Digital signatures"), v.Signatures.Count(s => s.IsSigned) + Loc.T(" chữ ký", " signature(s)")),
        };
        var lv = new ListView { Dock = DockStyle.Fill, View = View.Details, FullRowSelect = true, HeaderStyle = ColumnHeaderStyle.None, Name = "lstProperties", BorderStyle = BorderStyle.None };
        lv.Columns.Add("", Dpi.Scale(150));
        lv.Columns.Add("", Dpi.Scale(360));
        foreach (var (k, val) in rows) lv.Items.Add(new ListViewItem([k, string.IsNullOrWhiteSpace(val) ? "—" : val]));
        Body.Controls.Add(lv);
        AddButton(Loc.T("Đóng", "Close"), "btnClose", DialogResult.OK, true);
    }

    private static string FormatDate(string s) => PdfSigner.ParsePdfDate(s) is { } d ? d.ToString("dd/MM/yyyy HH:mm:ss zzz") : s;
}

public sealed class SettingsDialog : DialogBase
{
    private readonly ComboBox _std;
    private readonly CheckBox _ts, _rev;
    private readonly TextBox _tsa, _suffix;
    private readonly AppSettings _s;

    public SettingsDialog(AppSettings s) : base(Loc.T("Thiết lập ký số & xác thực", "Signing & validation settings"), 560, 330)
    {
        _s = s;
        var g = Grid();
        _std = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = Dpi.Scale(330), Name = "cboStandard" };
        _std.Items.AddRange([Loc.T("PAdES (ETSI.CAdES.detached) — khuyến nghị", "PAdES (ETSI.CAdES.detached) — recommended"), Loc.T("PKCS#7 (adbe.pkcs7.detached) — tương thích cũ", "PKCS#7 (adbe.pkcs7.detached) — legacy compatible")]);
        _std.SelectedIndex = s.Standard == SignatureStandard.PAdES ? 0 : 1;
        _ts = new CheckBox { Text = Loc.T("Gắn dấu thời gian (TSA) khi ký", "Add a timestamp (TSA) when signing"), Checked = s.UseTimestamp, AutoSize = true, Name = "chkTimestamp" };
        _tsa = new TextBox { Text = s.TsaUrl, Width = Dpi.Scale(330), Name = "txtTsaUrl" };
        _rev = new CheckBox { Text = Loc.T("Kiểm tra thu hồi chứng thư trực tuyến (CRL/OCSP) khi xác thực", "Check certificate revocation online (CRL/OCSP) when validating"), Checked = s.OnlineRevocationCheck, AutoSize = true, Name = "chkRevocation" };
        _suffix = new TextBox { Text = s.SignedFileSuffix, Width = Dpi.Scale(160), Name = "txtSuffix" };
        g.Controls.Add(L(Loc.T("Chuẩn chữ ký mặc định:", "Default signature standard:"))); g.Controls.Add(_std);
        g.Controls.Add(L("")); g.Controls.Add(_ts);
        g.Controls.Add(L(Loc.T("Máy chủ TSA:", "TSA server:"))); g.Controls.Add(_tsa);
        g.Controls.Add(L("")); g.Controls.Add(_rev);
        g.Controls.Add(L(Loc.T("Hậu tố tên tệp đã ký:", "Signed file name suffix:"))); g.Controls.Add(_suffix);
        var note = new Label
        {
            Text = Loc.T("Gợi ý TSA: http://timestamp.digicert.com, http://timestamp.sectigo.com, hoặc TSA của nhà cung cấp chứng thư số (VNPT, Viettel, FPT...).", "Suggested TSAs: http://timestamp.digicert.com, http://timestamp.sectigo.com, or your certificate provider's TSA (VNPT, Viettel, FPT...)."),
            AutoSize = true, MaximumSize = new Size(Dpi.Scale(500), 0), ForeColor = Theme.SubtleInk, Margin = new Padding(0, Dpi.Scale(10), 0, 0)
        };
        g.Controls.Add(note);
        g.SetColumnSpan(note, 2);
        Body.Controls.Add(g);
        AddOkCancel(Loc.T("Lưu", "Save"), Loc.T("Hủy", "Cancel"), (_, _) =>
        {
            _s.Standard = _std.SelectedIndex == 0 ? SignatureStandard.PAdES : SignatureStandard.Pkcs7;
            _s.UseTimestamp = _ts.Checked;
            _s.TsaUrl = _tsa.Text.Trim();
            _s.OnlineRevocationCheck = _rev.Checked;
            _s.SignedFileSuffix = string.IsNullOrWhiteSpace(_suffix.Text) ? "_signed" : _suffix.Text.Trim();
            _s.Save();
            DialogResult = DialogResult.OK;
        });
    }
}

public sealed class SignaturePropertiesDialog : DialogBase
{
    public event Action<SignatureValidationResult>? ViewSignedVersion;

    public SignaturePropertiesDialog(SignatureValidationResult r) : base(Loc.T("Thuộc tính chữ ký số", "Signature properties"), 620, 560)
    {
        var color = r.Status switch { SignatureStatus.Valid => Theme.OkGreen, SignatureStatus.Unknown => Theme.WarnAmber, SignatureStatus.Invalid => Theme.ErrRed, _ => Theme.SubtleInk };
        var glyph = r.Status switch { SignatureStatus.Valid => Icons.Check, SignatureStatus.Unknown => Icons.Unknown, SignatureStatus.Invalid => Icons.Error, _ => Icons.Sign };
        var head = new Panel { Dock = DockStyle.Top, Height = Dpi.Scale(56) };
        var icon = new PictureBox { Image = Icons.Get(glyph, Dpi.Scale(40), color), Dock = DockStyle.Left, Width = Dpi.Scale(52), SizeMode = PictureBoxSizeMode.CenterImage };
        var status = new Label { Text = r.StatusText, Dock = DockStyle.Fill, Font = Theme.TitleFont, ForeColor = color, TextAlign = ContentAlignment.MiddleLeft, Name = "lblSigStatus" };
        head.Controls.Add(status);
        head.Controls.Add(icon);
        var sb = new System.Text.StringBuilder();
        void Line(string k, string? v) { if (!string.IsNullOrWhiteSpace(v)) sb.AppendLine($"{k}: {v}"); }
        Line(Loc.T("Người ký", "Signer"), r.SignerName);
        Line(Loc.T("Thời gian ký (theo máy ký)", "Signing time (signer's clock)"), r.ClaimedSigningTime?.ToString("dd/MM/yyyy HH:mm:ss zzz"));
        if (r.HasTimestamp) Line(Loc.T("Dấu thời gian (TSA)", "Timestamp (TSA)"), (r.TimestampTime?.ToString("dd/MM/yyyy HH:mm:ss zzz") ?? "?") + (r.TimestampValid ? Loc.T(" — hợp lệ", " — valid") : Loc.T(" — KHÔNG hợp lệ", " — NOT valid")));
        Line(Loc.T("Lý do", "Reason"), r.Reason);
        Line(Loc.T("Vị trí", "Location"), r.Location);
        Line(Loc.T("Liên hệ", "Contact"), r.ContactInfo);
        Line(Loc.T("Nền tảng ký", "Signing platform"), r.Application);
        Line(Loc.T("Chuẩn chữ ký", "Signature standard"), r.SubFilter switch { "ETSI.CAdES.detached" => "PAdES (ETSI.CAdES.detached)", "adbe.pkcs7.detached" => "PKCS#7 (adbe.pkcs7.detached)", _ => r.SubFilter });
        Line(Loc.T("Thuật toán băm", "Hash algorithm"), r.DigestAlgorithm);
        Line(Loc.T("Trường chữ ký", "Signature field"), Loc.T($"{r.FieldName} (trang {r.PageIndex + 1}), phiên bản {r.Revision}", $"{r.FieldName} (page {r.PageIndex + 1}), revision {r.Revision}"));
        sb.AppendLine();
        sb.AppendLine(Loc.T("TÍNH TOÀN VẸN", "INTEGRITY"));
        sb.AppendLine(r.IntegrityValid ? Loc.T("  ✔ Tài liệu không bị sửa đổi kể từ khi chữ ký được áp dụng.", "  ✔ The document has not been modified since the signature was applied.") : Loc.T("  ✘ Tài liệu đã bị sửa đổi hoặc chữ ký bị hỏng.", "  ✘ The document has been modified or the signature is corrupted."));
        sb.AppendLine(r.SignatureValid ? Loc.T("  ✔ Giá trị chữ ký số hợp lệ.", "  ✔ The signature value is valid.") : Loc.T("  ✘ Giá trị chữ ký số không hợp lệ.", "  ✘ The signature value is invalid."));
        sb.AppendLine(r.CoversWholeDocument ? Loc.T("  ✔ Chữ ký bao phủ toàn bộ tài liệu.", "  ✔ The signature covers the whole document.")
            : r.ChangesAfterAreSignaturesOnly ? Loc.T("  ✔ Sau chữ ký chỉ có các chữ ký bổ sung (thay đổi được phép).", "  ✔ Only additional signatures were added after this one (permitted changes).")
            : Loc.T("  ⚠ Tài liệu đã được thay đổi sau khi ký (xem phiên bản đã ký để đối chiếu).", "  ⚠ The document was changed after signing (view the signed version to compare).") + (string.IsNullOrEmpty(r.ModificationDetails) ? "" : Loc.T("\n     Chi tiết: ", "\n     Details: ") + r.ModificationDetails));
        sb.AppendLine();
        sb.AppendLine(Loc.T("DANH TÍNH NGƯỜI KÝ", "SIGNER IDENTITY"));
        sb.AppendLine(r.CertificateTrusted ? Loc.T("  ✔ Chứng thư số hợp lệ và được tin cậy.", "  ✔ The certificate is valid and trusted.") : Loc.T("  ⚠ Danh tính người ký chưa được xác thực.", "  ⚠ The signer's identity could not be verified."));
        if (!string.IsNullOrWhiteSpace(r.TrustDetails)) foreach (var l in r.TrustDetails.Split('\n')) sb.AppendLine("    " + l);
        if (r.Certificate != null)
        {
            sb.AppendLine(Loc.T($"  Chủ thể: {r.Certificate.Subject}", $"  Subject: {r.Certificate.Subject}"));
            sb.AppendLine(Loc.T($"  Tổ chức cấp: {r.Certificate.Issuer}", $"  Issuer: {r.Certificate.Issuer}"));
            sb.AppendLine(Loc.T($"  Hiệu lực: {r.Certificate.NotBefore:dd/MM/yyyy} – {r.Certificate.NotAfter:dd/MM/yyyy}", $"  Valid: {r.Certificate.NotBefore:dd/MM/yyyy} – {r.Certificate.NotAfter:dd/MM/yyyy}"));
            sb.AppendLine(Loc.T($"  Số serial: {r.Certificate.SerialNumber}", $"  Serial number: {r.Certificate.SerialNumber}"));
        }
        if (r.Errors.Count > 0)
        {
            sb.AppendLine();
            sb.AppendLine(Loc.T("CHI TIẾT LỖI", "ERROR DETAILS"));
            foreach (var e in r.Errors.Distinct()) sb.AppendLine("  • " + e);
        }
        var details = new TextBox { Multiline = true, ReadOnly = true, ScrollBars = ScrollBars.Vertical, Dock = DockStyle.Fill, Text = sb.ToString().Replace("\n", "\r\n").Replace("\r\r", "\r"), BackColor = Color.White, Name = "txtSigDetails", Font = new Font("Segoe UI", 9.5f) };
        Body.Controls.Add(details);
        Body.Controls.Add(head);
        AddButton(Loc.T("Đóng", "Close"), "btnClose", DialogResult.OK, true);
        if (r.Certificate != null)
            AddButton(Loc.T("Xem chứng thư...", "View certificate..."), "btnViewCert", DialogResult.None, false, (_, _) => CertificateUi.Show(this, r.Certificate));
        if (r.IsSigned && r.RevisionLength > 0)
            AddButton(Loc.T("Xem phiên bản đã ký", "View signed version"), "btnViewSignedVersion", DialogResult.None, false, (_, _) => { ViewSignedVersion?.Invoke(r); Close(); });
    }
}

public static class CertificateUi
{
    [System.Runtime.InteropServices.DllImport("cryptui.dll", CharSet = System.Runtime.InteropServices.CharSet.Unicode, SetLastError = true)]
    private static extern bool CryptUIDlgViewContext(int dwContextType, IntPtr pvContext, IntPtr hwnd, string? pwszTitle, int dwFlags, IntPtr pvReserved);

    /// <summary>Shows the standard Windows certificate dialog.</summary>
    public static void Show(IWin32Window owner, System.Security.Cryptography.X509Certificates.X509Certificate2 cert)
    {
        try { CryptUIDlgViewContext(1, cert.Handle, owner.Handle, Loc.T("Chứng thư số", "Certificate"), 0, IntPtr.Zero); }
        catch (Exception ex) { Msg.Error(owner, Loc.T("Không mở được hộp thoại chứng thư: ", "Could not open the certificate dialog: ") + ex.Message); }
    }

    public static void OpenUrl(string url)
    {
        try { Process.Start(new ProcessStartInfo(url) { UseShellExecute = true }); } catch { }
    }
}
