# Microads Office 2.0 — Đối chiếu yêu cầu, vai trò & luồng nghiệp vụ

Phiên bản 2.0.0 · 16/11/2026 · Công ty TNHH Công nghệ Microads — microads.vn

## 1. Vai trò & luồng làm việc

| Vai trò | Trách nhiệm | Bằng chứng |
|---|---|---|
| Khách hàng / Product Owner | Đưa yêu cầu (mục 2) | Yêu cầu ngày 16/11/2026 |
| BA | Nghiên cứu Word/Excel → bảng tính năng đầy đủ trước khi làm | `docs/BANG_TINH_NANG.md` |
| Kiến trúc sư | Tách thư viện dùng chung, hợp đồng giữa các app | `docs/KIEN_TRUC.md` |
| Dev Word / Dev Excel / Dev PDF | 3 luồng song song, mỗi luồng chỉ sửa thư mục của mình | `src/MicroadsOffice.Word`, `.Excel`, `.Pdf` |
| Dev nền tảng & máy chủ | Exe duy nhất, cài đặt, ngôn ngữ, cập nhật, quảng bá, trang `microads.vn/office` | `src/MicroadsOffice`, `src/MicroadsOffice.UI`, `server/` |
| QA tự động | Kiểm thử lõi, kiểm thử từng nút trong tiến trình, kiểm thử đầu-cuối trên exe thật | mục 4 |
| Bảo mật | Chữ ký số ECDSA cho bản cập nhật, quản trị có CSRF/chống dò mật khẩu, máy chủ tách quyền | mục 5 |
| Phát hành / Vận hành | `build.ps1` (build → test → 1 exe → ký → phát hành), kiểm soát ổ C/D | mục 6 |
| Nghiên cứu phân phối (16/11 tối) | Nghiên cứu cơ chế cảnh báo Chrome/SmartScreen, các nguồn chứng thư miễn phí, chọn giải pháp hợp lệ | `docs/PHAT_HANH_KHONG_CANH_BAO.md` mục 1, 3, 7 |

Luồng: BA lập bảng tính năng → Kiến trúc chốt hợp đồng → 3 Dev làm song song → mỗi Dev tự chạy cổng chất lượng
(build 0 lỗi, kiểm thử đạt, chụp ảnh giao diện tự xem lại) → tích hợp → QA chạy toàn bộ → sửa lỗi phát hiện chéo → phát hành.

## 2. Đối chiếu từng yêu cầu

| # | Yêu cầu của khách | Trạng thái | Cách đáp ứng | Bằng chứng |
|---|---|---|---|---|
| 1 | Nghiên cứu, phát triển Microads Office **Word** đầy đủ như MS Word | ✅ | Bộ soạn thảo tự xây: trang in thật, ribbon 7 thẻ + thẻ ngữ cảnh Bảng/Ảnh, .docx đọc/ghi, mẫu công văn NĐ 30/2020, xuất PDF, in | Word tests 37/37 (gọi >120 lệnh ribbon), DOCX đọc lại bằng python-docx |
| 2 | Microads Office **Excel** đầy đủ như MS Excel | ✅ | Bảng tính tự xây: 502 hàm đủ các nhóm Excel 365, mảng động, LET/LAMBDA (có `DOCSO` đọc số tiếng Việt), biểu đồ 7 loại, định dạng có điều kiện, lọc/sắp xếp, .xlsx/.csv, in/PDF | Excel tests 28/28 (298 phép kiểm công thức, 111 nút + 366 mục menu) |
| 3 | Tìm hiểu tính năng trước, **lên bảng tính năng** rồi mới làm | ✅ | Bảng đối chiếu MS Word/Excel ↔ Microads (đầy đủ/cơ bản/không làm + lý do) | `docs/BANG_TINH_NANG.md` |
| 4 | Trọn gói trong **1 bộ Office** | ✅ | Một exe: Word, Excel, PDF + màn hình Microads Office (Hub) | E2E: Hub mở Excel |
| 5 | Đổi tên PDF → **Microads Office PDF** | ✅ | Tiêu đề, trang bắt đầu, Giới thiệu, lối tắt | E2E "Pdf: file opened" |
| 6 | Thông tin dự án có **logo Microads Office** | ✅ | Logo vector (dấu M đỏ/xanh) + biểu tượng riêng DOC/XLS/PDF; Giới thiệu, trang bắt đầu, thông báo, trình cài | Ảnh chụp giao diện |
| 7 | "Phát hành bởi **Công ty TNHH Công nghệ Microads**, luôn luôn miễn phí" | ✅ | Giới thiệu, dải quảng bá, thông báo mở tệp, trình cài, mục gỡ cài đặt của Windows | E2E "About shows publisher, always free" |
| 8 | Link **microads.vn** hiển thị "microads.vn" | ✅ | Nút "microads.vn" góc phải ribbon mọi app, Giới thiệu, thông báo, Hub | E2E "microads.vn link visible" |
| 9 | Giới thiệu dịch vụ: phần mềm doanh nghiệp, bảo mật, CRM, ERP, website… | ✅ | Danh sách dịch vụ trong Giới thiệu/thông báo/trang tải — sửa được từ trang quản trị | Trang quản trị › Nội dung quảng bá |
| 10 | **Nút phát hành** trọn bộ; mọi bản có nút **Kiểm tra cập nhật** | ✅ | Quản trị: "Phát hành"; app: nút ↻ góc phải + thẻ Trợ giúp + menu Tệp | E2E manual update; admin smoke test |
| 11 | **Tiếng Việt / English**, chọn ngay khi vừa cài, đổi sau bằng **lá cờ** 🇻🇳/🇬🇧 | ✅ | Trình cài: trang đầu chọn cờ + trang Hoàn tất có cờ; lần chạy đầu (bản chép tay) hỏi ngôn ngữ; nút cờ trên ribbon mọi app, Cài đặt | E2E first-run, installer, language switch |
| 12 | Ký số xong ghi nền tảng **Microads Office (Microads.vn)** | ✅ | `/Prop_Build /App /Name` + dòng "Nền tảng" trên hình chữ ký | Core test 5, pyHanko |
| 13 | Mở tệp từ thư mục → **thông báo quảng bá có logo, tắt ngay** | ✅ | Hộp thông báo có logo, nút ✕ + "Đóng" bấm được ngay, Esc; bật/tắt từ máy chủ | E2E Word/Excel/PDF "notice closes at once" |
| 14 | Giao diện trực quan, dễ dùng, **không trông như AI làm** | ✅ | Ribbon nhóm kiểu Office, thước kẻ, trang in thật, màu nhấn riêng từng app, không biểu tượng cảm xúc | Ảnh chụp `.build/word-test`, `excel-test`, `pdf-test` |
| 15 | **Chỉ 1 file exe**, cài là dùng, không cần thư viện | ✅ | .NET tự chứa + PDFium nhúng (tự giải nén khi dùng PDF lần đầu); cài theo người dùng, không cần quyền Admin | E2E "PDF engine extracted from the single exe" |
| 16 | Trang cập nhật **microads.vn/office** trên VPS | ✅ | PHP, pool PHP-FPM riêng, dữ liệu ngoài web root, trang tải + API + quản trị | Kiểm tra HTTP 200/404 đúng; site chính & /help không ảnh hưởng |
| 17 | Đổi link/thông tin quảng cáo **không cần phát hành lại** | ✅ | Manifest mang khối `promo`, app làm mới khi khởi động + mỗi 2 giờ | Update client test "promo cached" |
| 18 | **Cập nhật toàn bộ khi tôi ra lệnh**; không ra lệnh thì khách tự bấm Cập nhật | ✅ | Nút "Cập nhật toàn bộ" (forceUpdate + commandId): app tự tải, kiểm chữ ký, thay exe, mời khởi động lại; không lệnh → chỉ gợi ý | E2E forced update; tampered update bị từ chối |
| 19 | Mọi nút qua kiểm thử, không ẩn tính năng, luồng chuẩn | ✅ | Kiểm thử trong tiến trình gọi mọi nút ribbon/menu; E2E trên exe thật qua UI Automation | Mục 4 |
| 20 | Không ghi ổ C/D, kể cả cache | ✅ | SDK/NuGet/pip/TEMP/APPDATA của công cụ trong `E:\Office\.build`; trình duyệt chụp màn hình dùng hồ sơ trên E: rồi xóa | Mục 6 |
| 21 | Tải về/cài đặt **không bị Chrome / Windows Defender cảnh báo** (tải trực tiếp, miễn phí, không mua chứng thư, không Store) | ✅ kênh không cảnh báo · ◐ tải qua trình duyệt (giới hạn kỹ thuật, có lộ trình) | Xem mục 7 | `docs/PHAT_HANH_KHONG_CANH_BAO.md` |

## 3. Luồng nghiệp vụ & tiêu chí nghiệm thu

1. **Cài đặt**: chạy `MicroadsOffice_Setup_2.0.0.exe` → chọn cờ → thư mục, lối tắt, liên kết tệp → Cài đặt → Hoàn tất (chọn lại ngôn ngữ, mở Word/Excel/PDF). *Đạt khi* có lối tắt Start/Desktop, .docx/.xlsx/.pdf có "Mở bằng Microads", có mục gỡ trong Cài đặt Windows.
2. **Mở tệp từ thư mục**: bấm đúp .docx/.xlsx/.pdf → đúng ứng dụng → thông báo Microads (tắt ngay). Mở thêm tệp khi app đang chạy → chuyển vào cửa sổ đang mở (một phiên bản).
3. **Kiểm tra cập nhật (khách tự bấm)**: ↻ → "đang dùng bản mới nhất" hoặc "Cập nhật ngay" → tải → kiểm chữ ký → cài → khởi động lại.
4. **Cập nhật toàn bộ (quản trị ra lệnh)**: quản trị bấm "Cập nhật toàn bộ" → mọi máy (khi mở app hoặc tối đa 2 giờ) tự cập nhật → thông báo "Khởi động lại".
5. **Phát hành**: `build.ps1 -Publish [-ForceUpdate]` hoặc trang quản trị › Phát hành bản mới (tải .exe + .sig.json) → máy chủ kiểm chữ ký → bản mới nhất.
6. **Đổi ngôn ngữ**: bấm cờ → chọn → app lưu tài liệu (hỏi nếu chưa lưu), mở lại bằng ngôn ngữ mới cùng các tài liệu đang mở.
7. **Gỡ cài đặt**: Cài đặt Windows › Ứng dụng › Microads Office › Gỡ → hỏi giữ/xóa dữ liệu cá nhân.

## 4. Kiểm thử

| Bộ | Phạm vi | Kết quả |
|---|---|---|
| PDF core | Đọc, ký PAdES/PKCS#7, xác thực, tấn công sửa sau ký, in | 20/20 |
| Word | Mô hình, bố cục trang, DOCX/RTF/TXT/HTML/PDF, gọi mọi nút ribbon & menu (VI + EN) | 37/37 |
| Excel | 298 phép kiểm hàm/lỗi/vòng tham chiếu, XLSX/CSV/HTML/PDF, 111 nút + 366 mục menu | 28/28 |
| Cập nhật (máy khách) | Chữ ký hợp lệ/giả/sai phiên bản, so phiên bản, manifest thật, lệnh cập nhật toàn bộ chỉ chạy 1 lần, chặn link không phải http(s) | 15/15 |
| Máy chủ | Chữ ký PHP khớp máy build, đăng nhập sai bị chặn, CSRF, các thẻ quản trị | Đạt |
| E2E exe thật | Lần chạy đầu, trình cài, mở tệp từ thư mục ×3, Giới thiệu, Kiểm tra cập nhật, Hub, cập nhật thủ công, bản giả mạo bị từ chối, cập nhật toàn bộ, đổi ngôn ngữ, gỡ cài đặt | xem `.build/e2e/report-e2e.txt` |

## 5. Bảo mật

- Bản cập nhật ký ECDSA P-256; khóa bí mật chỉ ở máy build (`.build/secrets`), app và máy chủ đều kiểm chữ ký → máy chủ bị chiếm cũng không phát tán được chương trình độc.
- Quản trị: mật khẩu băm (password_hash), khóa 15 phút sau 5 lần sai, CSRF, cookie Secure/HttpOnly/SameSite=Strict, CSP, nhật ký `audit.log`.
- PHP-FPM pool riêng: `open_basedir`, cấm `exec/system/...`; bộ cài và cấu hình nằm ngoài web root, tải qua X-Accel-Redirect.
- Thông tin quảng bá chỉ cho phép liên kết http(s).

## 6. Ổ C/D

Công cụ build, NuGet, pip, TEMP, APPDATA của tiến trình build nằm trong `E:\Office\.build`. Hồ sơ trình duyệt dùng để chụp
màn hình trang web đặt ở `E:` và xóa ngay. Các thay đổi dung lượng trên D đến từ Zalo/Chrome/Codeium của máy (đã kiểm tra
danh sách tệp mới ghi). Xuất PDF qua "Microsoft Print to PDF" dùng hàng đợi in của Windows (tệp tạm tự xóa).

## 7. Yêu cầu 16/11/2026 (tối) — tải bằng Chrome và cài không bị cảnh báo, miễn phí

Luồng làm việc: **Nghiên cứu** (tài liệu Microsoft 2026, SmartScreen, Safe Browsing, kiểm tra trạng thái microads.vn) → **Kiến trúc**
(trình cài đặt trực tuyến không đổi + kênh không qua trình duyệt; hợp đồng tham số với exe chính) → **Dev** (WebSetup, máy chủ, lệnh
PowerShell, winget) → **QA** (E2E giao diện, trang web cục bộ, PowerShell 5.1/7, winget, Defender) → **Phát hành** (build.ps1, đóng băng
tệp, publish) → **Đối chiếu** (bảng dưới).

| # | Yêu cầu của khách | Trạng thái | Cách đáp ứng | Bằng chứng |
|---|---|---|---|---|
| 7.1 | Nghiên cứu kỹ, tìm mọi cách | ✅ | Tra tài liệu Microsoft (10–11/2026), bài của Eric Lawrence (SmartScreen), Rick Strahl (6/2026); kiểm tra Defender và Google Safe Browsing cho microads.vn | `PHAT_HANH_KHONG_CANH_BAO.md` mục 1, 3, 5 |
| 7.2 | Tích hợp **thông tin, chứng chỉ miễn phí** | ✅ / ◐ | Thông tin nhà phát hành (công ty, sản phẩm, mô tả, bản quyền, biểu tượng, manifest) nhúng trong trình cài đặt trực tuyến; chữ ký ECDSA của Microads kiểm trước khi chạy. Chứng thư Authenticode miễn phí cho phần mềm đóng nguồn **không tồn tại** (bảng Microsoft); phương án miễn phí duy nhất là SignPath Foundation, cần mở mã nguồn: chờ quyết định | E2E "publisher information embedded"; mục 7 tài liệu |
| 7.3 | Không mua chứng thư, không đăng ký Microsoft Store | ✅ | Không dùng Store, không mua gì; winget là kho cộng đồng miễn phí trên GitHub (không phải Store) | — |
| 7.4 | Tải bằng Chrome **không bị cảnh báo** | ◐ có lộ trình | Nút tải chính là `MicroadsOffice_Setup.exe` **không bao giờ đổi**, nên lượt tải được cộng dồn và cảnh báo biến mất vĩnh viễn khi đủ danh tiếng. Kênh **không cảnh báo ngay**: lệnh PowerShell, winget (sau khi duyệt). Ngày đầu tiên, một tệp chưa ai tải không thể hết cảnh báo bằng cách hợp lệ nào (Microsoft xác nhận, kể cả chứng thư trả phí) | Trang tải, test_site 43/43 |
| 7.5 | Cài **không bị Windows Defender cảnh báo** | ✅ / ◐ | Defender **không phát hiện** gì. Hộp "Windows đã bảo vệ PC" (SmartScreen) chỉ hiện **1 lần** ở trình cài đặt trực tuyến (bộ cài tải về không mang dấu Internet), rồi biến mất khi đủ danh tiếng. Lệnh PowerShell và winget: không hiện | E2E "NO Mark of the Web"; PowerShell 18/18 ×2 |
| 7.6 | Các nút phải qua kiểm thử, bấm được, không ẩn tính năng | ✅ | Trình cài đặt trực tuyến: Thử lại, Tải bộ cài đầy đủ, Đóng/Hủy, English/Tiếng Việt, microads.vn. Trang web: 2 nút tải, 3 nút Sao chép. Quản trị: ô mã winget (lưu/sai/xóa). Lệnh winget **chỉ hiện khi gói đã được duyệt**, tránh hiện lệnh chưa chạy được | E2E 64/64, site 43/43 |
| 7.7 | Luồng chuẩn, móc nối hệ thống | ✅ | Dùng chung manifest `api/manifest.json`, khóa ký ECDSA, thống kê (`web-setup`, `ps-install`), trang quản trị; `build.ps1 -Publish` tự đưa trình cài đặt trực tuyến (chỉ lần đầu) và tạo hồ sơ winget cho mỗi bản; hợp đồng tham số ghi trong `Program.cs` | `build.ps1`, `KIEN_TRUC.md` mục 4 |
| 7.8 | Không ghi ổ C/D | ✅ | Mọi build, test, PHP, PowerShell 7 portable trong `E:\Office\.build`; 2 tệp nhật ký winget tạo trên C: đã xóa ngay; đối chiếu dung lượng C/D với mốc ban đầu | mục 6 |

Kiểm thử bổ sung: Defender quét trình cài đặt trực tuyến và `install.ps1` → không phát hiện; build tất định (dựng lại cho cùng mã băm);
`build.ps1` phát hiện tệp đóng băng bị sửa. Rà soát mã độc lập → sửa 1 lỗi thật + gia cố (bản tối thiểu 2.0.1, khóa dự phòng, `--dir`, hủy giữa chừng, đếm lượt tách khỏi cấu hình)
trước khi đóng băng (SHA-256 `76628b52…188271`). **Đã đưa lên microads.vn 16/11/2026 20:08** và cài thử thật thành công.
**Việc còn chờ Microads**: gửi PR winget (cần tài khoản GitHub), sao lưu ngoại tuyến 2 khóa ký, quyết định có mở mã nguồn để dùng SignPath hay không.
