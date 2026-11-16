# Microads Office 2.0 — Hướng dẫn sử dụng & vận hành

Phát hành bởi Công ty TNHH Công nghệ Microads — luôn luôn miễn phí — **microads.vn**

## A. Dành cho người dùng

### 1. Cài đặt
1. Tại **https://microads.vn/office** bấm **Tải Microads Office** → tải `MicroadsOffice_Setup.exe` (80 KB, Windows 10/11 64-bit).
   Trình cài đặt này tự tải bản mới nhất từ microads.vn, kiểm tra chữ ký số của Microads rồi mở bộ cài.
   (Nút phụ **Bộ cài đầy đủ** tải tệp ~51 MB dùng được khi không có mạng lúc cài.)
2. Nếu Chrome hỏi: mở danh sách tải xuống (Ctrl+J) → **Giữ lại**. Nếu Windows hiện "Windows đã bảo vệ PC của bạn": **Thông tin bổ sung → Vẫn chạy**
   (chỉ hỏi một lần; lý do và lộ trình hết cảnh báo: `docs/PHAT_HANH_KHONG_CANH_BAO.md`).
3. Chọn ngôn ngữ bằng lá cờ 🇻🇳 Tiếng Việt / 🇬🇧 English → **Tiếp tục**.
4. Chọn thư mục (mặc định không cần quyền quản trị), lối tắt Desktop, mở tệp .docx/.xlsx/.csv/.pdf bằng Microads Office → **Cài đặt**.
5. Trang **Cài đặt hoàn tất**: có thể đổi lại ngôn ngữ, bấm **Word / Excel / PDF** để mở ngay.

Cài **không qua trình duyệt** (không có cảnh báo): mở PowerShell, dán `irm https://microads.vn/office/install.ps1 | iex`, nhấn Enter;
hoặc `winget install --id Microads.Office -e` (khi gói đã được Microsoft duyệt). IT cài hàng loạt: `MicroadsOffice_Setup.exe --silent`.

Gỡ cài đặt: Cài đặt Windows › Ứng dụng › Microads Office › Gỡ cài đặt (được hỏi giữ hay xóa dữ liệu cá nhân).

### 2. Ba ứng dụng
| Ứng dụng | Mở bằng | Dùng để |
|---|---|---|
| Microads Office **Word** | Lối tắt "Microads Office Word", hoặc bấm đúp .docx/.rtf/.txt | Soạn văn bản theo trang in thật, mẫu công văn NĐ 30/2020, bảng, ảnh, mục lục, đầu/chân trang, xuất PDF, in |
| Microads Office **Excel** | Lối tắt "Microads Office Excel", hoặc .xlsx/.csv | Bảng tính, >230 hàm (`=DOCSO(A1)` đọc số thành chữ), biểu đồ, lọc/sắp xếp, định dạng có điều kiện, in/PDF |
| Microads Office **PDF** | Lối tắt "Microads Office PDF", hoặc .pdf | Đọc, ký số (USB Token, .pfx), xác thực chữ ký, chèn dấu ảnh, in |
| **Microads Office** (Hub) | Lối tắt "Microads Office" | Chọn ứng dụng, mở tệp bất kỳ, tài liệu gần đây |

Màn hình **Tệp** (nút màu ở góc trái ribbon): Trang chủ, Mới (mẫu), Mở, Thông tin, Lưu, Lưu thành, In, Xuất PDF, Đóng, Giới thiệu, Cài đặt.
Phím tắt giống Microsoft Office (Ctrl+S/O/N/P/Z/Y/C/X/V/B/I/U/F/H…).

### 3. Đổi ngôn ngữ
Bấm nút **lá cờ** ở góc phải trên ribbon → chọn Tiếng Việt / English. Ứng dụng hỏi lưu tài liệu (nếu có), rồi mở lại
bằng ngôn ngữ mới cùng các tài liệu đang mở. Ngôn ngữ dùng chung cho Word, Excel, PDF.

### 4. Cập nhật
- Ứng dụng tự kiểm tra bản mới khi mở (và mỗi 2 giờ) và hiện thông báo nhỏ góc phải dưới.
- Kiểm tra thủ công: nút **↻** góc phải trên, hoặc thẻ **Trợ giúp › Kiểm tra cập nhật**.
- Mọi bản cập nhật đều có chữ ký số của Microads; tệp bị sửa/giả mạo sẽ bị từ chối.

## B. Dành cho quản trị (Microads)

### 5. Trang quản trị — https://microads.vn/office/admin/
| Thẻ | Việc |
|---|---|
| Tổng quan | Phiên bản mới nhất, lượt kiểm tra cập nhật theo ngày & phiên bản, lượt tải; **trình cài đặt trực tuyến** (lượt tải, lượt cài, lượt cài bằng PowerShell, SHA-256); nút **Cập nhật toàn bộ** |
| Phát hành bản mới | Tải lên `MicroadsOffice_Setup_x.y.z.exe` + `….sig.json`, ghi chú VI/EN, tùy chọn "Cập nhật toàn bộ ngay" → **Phát hành** |
| Nội dung quảng bá | Bật/tắt thông báo khi mở tệp, liên kết & chữ hiển thị (microads.vn), tên công ty, tiêu đề, dịch vụ, nút, điện thoại, email — áp dụng ngay, không cần phát hành lại; **Mã gói winget** (hiện lệnh `winget install` trên trang tải); Store ID |
| Lịch sử phiên bản | Tải, đặt làm mới nhất (quay lui), xóa |
| Tài khoản | Đổi mật khẩu (tên đăng nhập `admin`) |

**Cập nhật toàn bộ**: bấm nút → mọi máy dùng bản cũ (khi mở app hoặc tối đa 2 giờ) tự tải, kiểm chữ ký, cài và mời khởi động lại.
Không bấm → khách chỉ thấy gợi ý và tự bấm "Cập nhật".

### 6. Ra phiên bản mới (máy build E:\Office)
1. Sửa `<Version>` trong `src\MicroadsOffice\MicroadsOffice.csproj` (vd. 2.0.1).
2. `powershell -ExecutionPolicy Bypass -File build.ps1` → build, chạy toàn bộ kiểm thử, đóng gói 1 exe, ký số →
   `release\MicroadsOffice_Setup_2.0.1.exe` + `.sig.json`.
3. Phát hành: trang quản trị › Phát hành bản mới (tải 2 tệp lên), **hoặc** `build.ps1 -Publish [-ForceUpdate] -NotesVi a.txt -NotesEn b.txt`
   (máy build đọc mật khẩu SSH từ `.build\secrets\vps_pass`, hoặc biến môi trường `MICROADS_VPS_PASS`; hoặc dùng trang quản trị).
4. Trình cài đặt trực tuyến (`release\web-setup\MicroadsOffice_Setup.exe`) **không đổi** khi ra bản mới; `-Publish` chỉ đưa lên lần đầu.
   Kiểm tra riêng: `build.ps1 -WebSetupOnly [-UiTests] [-Publish]`. Hồ sơ winget của bản mới được tạo trong `tools\winget\manifests` → gửi PR (`tools\winget\README.md`).

Giảm cảnh báo khi tải/cài (miễn phí), ký Authenticode / Microsoft Store (nếu sau này cần): xem `docs/PHAT_HANH_KHONG_CANH_BAO.md`
(`build.ps1 -CodeSignThumbprint <thumbprint>`; `build.ps1 -Store -StoreIdentityName … -StorePublisher …`).

**Khóa ký bản cập nhật** `.build\secrets\update-signing.pem`: chỉ nằm trên máy build, phải **sao lưu an toàn** (USB/két).
Mất khóa = không phát hành được bản cập nhật cho các máy đã cài. Lộ khóa = người khác có thể ký bản giả → báo ngay để đổi khóa.

### 7. Máy chủ
- Mã nguồn: `/var/www/microads.vn/office` (nginx khối `# >>> microads-office >>>` trong `sites-available/microads.vn`).
- Dữ liệu: `/var/www/microads-office-data` (`config.json`, `releases/`, `stats/`, `logs/audit.log`, `admin.json`).
- PHP-FPM pool riêng `/etc/php/8.5/fpm/pool.d/microads-office.conf`.
- Triển khai lại mã trang: `python server\deploy\deploy.py` (sao lưu cũ, kiểm `php -l` và `nginx -t`, tự hoàn tác nếu lỗi).
- Đưa trình cài đặt trực tuyến: `python server\deploy\publish_websetup.py` (máy chủ giữ nguyên nếu trùng, từ chối nếu khác).
- Mật khẩu SSH: `.build\secrets\vps_pass` (chỉ tài khoản Windows của máy build đọc được; biến môi trường `MICROADS_VPS_PASS` được ưu tiên nếu có). Không chép tệp này ra ngoài.
- Thử trang web trên máy (PHP 8.3, không cần VPS): `python -I server\dev\test_site.py <php.exe> <thư mục> <MicroadsOffice_Setup.exe> <bộ cài đầy đủ> <.sig.json>`;
  lệnh PowerShell: `server\dev\test_install_command.ps1`.
