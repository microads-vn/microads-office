# Microads Office — Kiến trúc & quy ước bắt buộc

## 1. Cấu trúc giải pháp

| Project | Vai trò |
|---|---|
| `src/MicroadsOffice.Core` | Lõi không giao diện: PDF/PDFium/ký số (có sẵn), `Office/Loc.cs` (song ngữ), `Office/Product.cs` (định danh, tên app, đuôi tệp), `Office/OfficeConfig.cs` (cấu hình chung `Data/office.json`, gần đây, quảng bá), `Office/UpdateService.cs` (cập nhật có chữ ký ECDSA), `Pdfium/PdfiumLoader.cs` (PDFium nhúng trong exe), `Branding/BrandArt.cs` (logo, biểu tượng từng app) |
| `src/MicroadsOffice.UI` | Giao diện dùng chung: `Theme`/`Icons`/`Dpi` (Theme.cs), `Ribbon` (ribbon đơn giản của PDF), **`RibbonBar`** (ribbon nhóm kiểu Office cho Word/Excel), `RibbonButton`, `ColorPickerDropDown`, **`Backstage`** (màn hình Tệp), **`StartScreen`**, `PromoStrip`, `DialogBase`, `Msg` (hộp thoại song ngữ), `OfficeShell` (ngôn ngữ, cờ, cập nhật, giới thiệu, quảng bá), `PdfExporter` |
| `src/MicroadsOffice.Pdf` | Microads Office PDF (đọc, ký số, in) — `PdfApp.CreateMainForm(files)` |
| `src/MicroadsOffice.Word` | Microads Office Word — `WordApp.CreateMainForm(files)` |
| `src/MicroadsOffice.Excel` | Microads Office Excel — `ExcelApp.CreateMainForm(files)` |
| `src/MicroadsOffice` | **Exe duy nhất**: điều phối (`--app=word|excel|pdf`, theo đuôi tệp), Hub, trình cài đặt/gỡ (`Setup/`), một phiên bản/ứng dụng (`SingleInstance`) |
| `src/MicroadsOffice.WebSetup` | **Trình cài đặt trực tuyến** `MicroadsOffice_Setup.exe` (.NET Framework 4.6.2 có sẵn trên Windows 10/11, 80 KB): tải bản mới nhất, kiểm SHA-256 + chữ ký ECDSA, chạy bộ cài. Tệp phát hành **đóng băng** trong `release/web-setup` (mục 4) |
| `tests/MicroadsOffice.Tests` | Kiểm thử PDF (lõi + UI đầu-cuối FlaUI) |
| `tests/WebSetup.E2E` | Kiểm thử đầu-cuối trình cài đặt trực tuyến (không thuộc .sln; `build.ps1 -WebSetupOnly -UiTests`) |
| `tests/MicroadsOffice.Word.Tests`, `tests/MicroadsOffice.Excel.Tests` | Kiểm thử Word / Excel (mô hình, đọc/ghi tệp, công thức, nút bấm trong tiến trình) |

Exe gọi: `WordApp.CreateMainForm(string[] files)` → `Form` phải hiện thực `IOfficeAppWindow`.
Trước đó Program đã gọi `OfficeShell.Init(AppKind.X)` (đặt ngôn ngữ, màu nhấn, Dpi, biểu tượng) và
`OfficeShell.EnsureLanguageChosen`. Sau `Shown`, Program tự gọi `OfficeShell.StartBackgroundServices(form)`
và — nếu mở bằng tệp từ thư mục — `OfficeShell.OnFileOpenedFromShell(form)` (thông báo quảng bá).
Khi phiên bản thứ hai chuyển tệp sang, Program gọi `IOfficeAppWindow.OpenFiles(paths, fromShell: true)`.

## 2. Quy ước bắt buộc cho mọi ứng dụng

1. **Song ngữ**: mọi chữ hiển thị viết `Loc.T("tiếng Việt", "English")` (hoặc `Loc.F` có tham số). Không có chữ cứng.
2. **Tên điều khiển**: mọi nút/ô/menu có `Name` duy nhất, ổn định (vd. `btnBold`, `cboFontName`, `miSaveAs`) để kiểm thử tự động tìm được (UIA AutomationId = `Name`).
3. **Thanh tiêu đề Ribbon (bên phải)**: `ribbon.AddHeaderRight(OfficeShell.CreateHeaderItems(this))` → nút cờ ngôn ngữ, nút kiểm tra cập nhật, liên kết **microads.vn**.
4. **Thẻ Trợ giúp / Help** có đủ: Hướng dẫn (`btnHelp`), Kiểm tra cập nhật (`btnCheckUpdate` → `OfficeShell.CheckForUpdatesInteractive`), Cài đặt & ngôn ngữ (`btnSettings` → `OfficeShell.ShowSettings`), Giới thiệu (`btnAbout` → `OfficeShell.ShowAbout`), microads.vn (`btnVisitWebsite` → `OfficeShell.OpenWebsite()`).
5. **Màn hình Tệp (Backstage)**: Trang chủ (StartScreen), Mới, Mở, Thông tin, Lưu, Lưu thành, In, Xuất PDF, Đóng, Giới thiệu, Tùy chọn/Cài đặt.
6. **Gần đây**: `OfficeConfig.AddRecent(AppKind.X, path)` sau khi mở/lưu; StartScreen tự đọc.
7. **Thay đổi chưa lưu**: hỏi bằng `Msg.AskSave(owner, name)` (nút Lưu / Không lưu / Hủy theo ngôn ngữ app).
8. **Thông báo/lỗi**: `Msg.Info/Warn/Error/Confirm` — không dùng `MessageBox` (nút theo ngôn ngữ Windows).
9. **Tự lưu phục hồi**: `Path.Combine(AppSettings.DataDirectory, "AutoRecover", "<app>")`, mỗi 2 phút khi có thay đổi; mở lại sau sự cố thì đề nghị khôi phục; xóa khi lưu/đóng bình thường.
10. **Cài đặt riêng của app**: `OfficeConfig.GetAppData/SetAppData("word.xxx", json)`.
11. **Không ghi ổ C/D khi phát triển**: build bằng `source .build/env.sh` (mọi cache vào `E:\Office\.build`); tệp tạm của kiểm thử nằm trong `E:\Office\.build\...`.
12. **Giao diện tự nhiên, gọn**: dùng `Theme`, `Icons` (glyph phông Segoe Fluent/MDL2 hoặc `Icons.Painted` vẽ vector), khoảng cách đều, không biểu tượng cảm xúc trong nút, không màu sặc sỡ ngoài màu nhấn của app (Word xanh dương, Excel xanh lá, PDF đỏ Microads).
13. **Ribbon**: nút không nhận focus (không làm mất con trỏ soạn thảo); sau khi chọn trong ô phông/cỡ chữ phải trả focus về tài liệu.

## 3. Hệ thống cập nhật

- Máy chủ: `https://microads.vn/office` (trang tải + trang quản trị), API `https://microads.vn/office/api/manifest.json`.
- Manifest: `latest {version, date, url, size, sha256, signature, notes{vi,en}}`, `forceUpdate {enabled, commandId}`, `promo {...}`.
- Bản phát hành được ký bằng khóa ECDSA P-256 riêng (`.build/secrets/update-signing.pem`, không bao giờ rời máy build); app chỉ cài bản có chữ ký khớp khóa công khai trong `Product.UpdatePublicKey`.
- "Cập nhật toàn bộ": quản trị bật `forceUpdate` với `commandId` mới → mọi máy đang mở app tự tải, kiểm tra chữ ký, thay exe và mời khởi động lại.
- Thông tin quảng bá (`promo`) luôn được làm mới khi app kiểm tra (khởi động + mỗi 2 giờ), lưu đệm để dùng khi ngoại tuyến.

## 4. Phân phối không cảnh báo (trình cài đặt trực tuyến, PowerShell, winget)

- Trình duyệt và SmartScreen chấm danh tiếng tệp không ký theo **SHA-256**, nên tệp khách tải phải **không đổi**:
  `release/web-setup/MicroadsOffice_Setup.exe` + `.sha256`. `build.ps1` dựng lại (build tất định) chỉ để so sánh, không thay tệp;
  thay chỉ khi `-NewWebSetup` (danh tiếng về 0). Máy chủ (`cli.php websetup`) từ chối tệp khác nếu không có `--replace`.
- **Hợp đồng tham số (không được phá):** trình cài đặt trực tuyến chạy bộ cài với
  `--setup --web-setup --lang=vi|en [--silent] [--dir "PATH"]`; winget chạy `--install --silent [--dir "PATH"]`. Mọi phiên bản
  `MicroadsOffice.exe` phải tiếp tục nhận các tham số này. Tham số `--` lạ phải được bỏ qua.
- Manifest `https://microads.vn/office/api/manifest.json` và khóa công khai `Product.UpdatePublicKey` nằm cố định trong trình cài đặt
  trực tuyến: **không đổi đường dẫn API**. Trình cài đặt trực tuyến nhận chữ ký của khóa chính **hoặc khóa dự phòng**
  (`.build/secrets/release-signing-backup.pem`, phải cất bản sao ngoại tuyến) → nếu mất khóa chính vẫn đổi khóa được mà không thay tệp.
- Hợp đồng manifest với tệp đóng băng: `latest.version` dạng `a.b.c` (chỉ chữ số) và **≥ 2.0.1** (chống hạ cấp xuống bản cũ có chữ ký),
  `url` HTTPS, `size` số nguyên ≤ 1 GB, `sha256` hex, `signature` ECDSA P-256 DER trên `MicroadsOffice|version|sha256`.
- `--dir` phải là đường dẫn đầy đủ (`X:\…` hoặc `\\máy\chia sẻ\…`, biến `%…%` được mở rộng); đường dẫn tương đối bị từ chối (mã 6)
  vì sẽ bị hiểu theo thư mục tạm. Bộ cài cần quyền quản trị trong tương lai (lỗi 740) sẽ được mở lại qua shell (hiện UAC).
- Đếm lượt tải nằm ở `counters.json` (khóa độc quyền), không ghi `config.json` → tải xuống không bao giờ ghi đè bản phát hành vừa đăng;
  `json_write` chỉ thay tệp khi ghi đủ (đầy ổ đĩa → giữ tệp cũ).
- Tham số `v=` của manifest: số phiên bản (ứng dụng), `web-setup` (trình cài đặt trực tuyến), `ps-install` (lệnh PowerShell) → thống kê.
- Biến môi trường chỉ dùng cho kiểm thử: `MICROADS_UPDATE_URL` (https hoặc http trên máy này), `MICROADS_WEBSETUP_OPENURL_LOG`,
  `MICROADS_SETUP_SANDBOX`, `MICROADS_DATA_DIR`; trang web: `MO_DATA_DIR`, `MO_BASE_URL`, `MO_DEV_SENDFILE`, `MO_DEV_INSECURE_COOKIES`
  (chỉ máy chủ thử `server/dev/router.php`).
- Lệnh PowerShell: `server/office/inc/install.ps1`, phục vụ tại `/office/install.ps1` qua `install.php` (UTF-8, không BOM).
- winget: `tools/winget` (hồ sơ YAML, README gửi PR); mã gói bật trên trang tải qua quản trị (`wingetId`).
