# Ký số miễn phí qua SignPath Foundation — hồ sơ & hướng dẫn từng bước

Cập nhật 16/11/2026. Kho mã công khai: https://github.com/microads-vn/microads-office (GPL-3.0).
Trang bắt buộc: https://microads.vn/office/code-signing · Chính sách quyền riêng tư: https://microads.vn/office/privacy

## Đã chuẩn bị xong (đáp ứng điều kiện của SignPath Foundation)

| Điều kiện của SignPath | Đáp ứng |
|---|---|
| Giấy phép mã nguồn mở OSI cho mọi thành phần, không cấp phép kép thương mại | GPL-3.0 (`LICENSE`); bên thứ ba: PDFium (BSD), .NET (MIT) — `THIRD_PARTY_NOTICES.md` |
| Không có mã độc quyền, không phần mềm không mong muốn | Toàn bộ mã nguồn công khai; Defender quét sạch |
| Đang được bảo trì, đã phát hành | Bản 2.0.0, 2.0.1 (16/11/2026), 2.0.2; trang tải microads.vn/office |
| Tính năng được mô tả trên trang tải | microads.vn/office + `README.md` |
| Build tự động, kiểm chứng được từ mã nguồn | GitHub Actions `.github/workflows/release.yml` (build + kiểm thử + đóng gói; ký khi gắn thẻ phiên bản) |
| Siêu dữ liệu sản phẩm cố định | ProductName "Microads Office", phiên bản lấy từ `<Version>` (workflow kiểm tra thẻ khớp phiên bản) |
| Mỗi lần ký đều được duyệt thủ công | Chính sách ký "release-signing" của SignPath (bước duyệt) |
| Trang "Code signing policy" có câu bắt buộc, vai trò, quyền riêng tư; được liên kết ở trang chủ & trang tải | microads.vn/office/code-signing + liên kết ở chân trang và mục Cài đặt an toàn; `CODE_SIGNING_POLICY.md` trong kho |
| Gửi dữ liệu tới hệ thống người dùng không chỉ định → mô tả trong chính sách, **hiện khi cài**, **cho tắt khi cài** | Bản 2.0.2: trình cài có ô "Tự động kiểm tra bản cập nhật…" + liên kết chính sách; tắt thì ứng dụng không tự kết nối Internet (đã kiểm thử); `--no-auto-update` cho cài im lặng |
| Có gỡ cài đặt | Có (Cài đặt Windows › Ứng dụng; `--uninstall`) |
| Xác thực nhiều lớp (MFA) cho GitHub và SignPath | **Anh/chị bật — bước 1 dưới đây** |

## Các bước anh/chị làm (mỗi bước vài phút)

### Bước 1 — Bật xác thực 2 lớp cho GitHub (bắt buộc)
1. Cài **Google Authenticator** (hoặc Microsoft Authenticator) trên điện thoại.
2. Mở https://github.com/settings/security → **Enable two-factor authentication**.
3. Chọn **Authenticator app** → quét mã QR bằng điện thoại → nhập 6 số → **Continue**.
4. **Tải về / in mã khôi phục (recovery codes)** và cất cùng chỗ với 2 khóa ký. → **Done**.

### Bước 2 — Gửi đơn cho SignPath Foundation
1. Mở https://signpath.org/apply → làm theo biểu mẫu (đăng nhập/đăng ký tài khoản SignPath bằng Gmail công ty, **bật MFA** khi được hỏi).
2. Điền theo bảng sau (bằng tiếng Anh, sao chép nguyên văn):

Biểu mẫu (HubSpot, kiểm tra 16/11/2026) có đúng các ô sau, theo thứ tự:

| Ô trên biểu mẫu | Nội dung |
|---|---|
| Project Name* | `Microads Office` |
| Repository URL* | `https://github.com/microads-vn/microads-office` |
| Homepage URL* | `https://microads.vn/office/` |
| Download URL | `https://microads.vn/office/` (trang này có nút tải **và** câu nhắc SignPath Foundation trong mục "Cài đặt an toàn") |
| Privacy Policy URL | `https://microads.vn/office/privacy` |
| Wikipedia URL (optional) | để trống (dự án chưa có bài Wikipedia) |
| Tagline* | xem dưới |
| Category* (chọn) | `Productivity` |
| Description* | xem dưới |
| Reputation* | xem dưới (cuối đoạn có link Code signing policy vì biểu mẫu không có ô riêng) |
| Maintainer Type (chọn) | `For-profit company or corporate-backed project` |
| Build System* (chọn) | `GitHub Actions` |
| First Name* / Last Name* | `Thị Bé Bi` / `Phạm` (người đại diện theo pháp luật) |
| Email* | Gmail công ty |
| Company Name | `Microads Technology Company Limited` |
| Primary Discovery Channel* (chọn) | `AI / LLM tools`; ô nguồn cụ thể: `Claude (Anthropic)` |
| Ô đồng ý | Tích **Code of Conduct** (bắt buộc) và **store and process my personal data** (bắt buộc); để trống ô nhận thư quảng cáo |

Biểu mẫu **không** có ô chọn chứng thư hay thời hạn: chứng thư là của SignPath Foundation (đứng tên họ, họ tự gia hạn); mỗi chữ ký
có dấu thời gian nên tệp đã ký vẫn hợp lệ sau khi chứng thư hết hạn.

**Tagline** (một câu):
> Free, open source office suite for Windows: word processor, spreadsheet and PDF reader with digital signatures.

**Description** (một đoạn, không liệt kê tính năng theo phiên bản hay thư viện phụ thuộc):
> Microads Office is a free and open source (GPL-3.0) office suite for Windows. It brings together a word processor, a spreadsheet
> application and a PDF reader that can digitally sign and verify documents with USB tokens or certificate files, in a single installer
> with a Vietnamese and English user interface. Its goal is to give individuals, small businesses and public administration staff in
> Vietnam a free, lightweight tool for everyday document work, including the document templates and digital signatures used in
> Vietnamese administration. It is developed by Microads Technology Company Limited, Ho Chi Minh City, Vietnam.

**Reputation** (trung thực — SignPath kiểm tra được; không phóng đại số người dùng):
> Microads Office is a young project: its source code was published on GitHub and its first releases were distributed from
> https://microads.vn/office/ in November 2026, so its user base is still small and growing. It is developed and maintained by
> Microads Technology Company Limited, a registered company in Ho Chi Minh City, Vietnam (tax code 0319785225, https://microads.vn),
> which takes legal responsibility for the project. Releases are built by GitHub Actions from the public repository, are additionally
> signed with our own ECDSA release key that the apps and the update server verify, and are reported clean by Microsoft Defender and
> Google Safe Browsing. Because the installer is unsigned, Windows SmartScreen and browsers warn new users; signing through SignPath
> Foundation would let them verify that the software they download is exactly what was built from our public source code.
> Code signing policy: https://microads.vn/office/code-signing

Chỉ tích ô Code of Conduct khi đã bật MFA (Bước 1) và đã đọc https://signpath.org/terms.

3. Gửi đơn (**Submit**, làm reCAPTCHA nếu được hỏi). SignPath Foundation xét duyệt (thường vài ngày đến vài tuần) và trả lời qua email.

### Bước 3 — Sau khi được duyệt (em làm tiếp cùng anh/chị)
SignPath gửi thông tin đăng nhập. Trong SignPath (app.signpath.io):
1. **Project** `microads-office` → *Repository URL* như trên → *Trusted build system*: **GitHub.com** (liên kết kho).
2. **Artifact configuration**: dán nội dung `.signpath/artifact-configuration.xml` (đặt slug `initial`).
3. **Signing policy** `release-signing` (chứng thư của SignPath Foundation), **Approvers**: tài khoản của anh/chị.
4. **Users › tạo API token** cho "CI user" (quyền *Submitter*).
5. Gửi em: *Organization ID*, *API token* (đặt vào biến môi trường như token GitHub). Em sẽ cấu hình GitHub (secret `SIGNPATH_API_TOKEN`,
   biến `SIGNPATH_ORGANIZATION_ID`, `SIGNPATH_PROJECT_SLUG=microads-office`, `SIGNPATH_SIGNING_POLICY_SLUG=release-signing`,
   `SIGNPATH_ARTIFACT_CONFIGURATION_SLUG=initial`).

### Bước 4 — Mỗi lần phát hành có chữ ký
1. Em (hoặc người phụ trách) tăng `<Version>`, đẩy mã, gắn thẻ `vX.Y.Z` → GitHub Actions build, kiểm thử, gửi SignPath.
2. Anh/chị nhận email của SignPath → bấm **Approve** (duyệt ký).
3. Tệp đã ký nằm trong *GitHub Releases* (bản nháp). Máy build tải về, thêm chữ ký ECDSA của Microads và đăng lên microads.vn
   (`build.ps1 -Publish`). Trình cài đặt trực tuyến **có chữ ký** thay cho bản chưa ký (`-NewWebSetup`): danh tiếng gắn với chứng thư
   của SignPath Foundation (đã được nhiều dự án dùng) nên cảnh báo của Windows/Chrome hết nhanh hơn nhiều.

## Lưu ý
- Ô *Download URL* yêu cầu trang tải **có nhắc tới SignPath Foundation**. Từ 16/11/2026 trang chủ `microads.vn/office/` (có nút tải) mang câu
  "Free code signing provided by SignPath.io, certificate by SignPath Foundation" trong mục "Cài đặt an toàn" (`id="signpath-notice"`,
  `server/dev/test_site.py` kiểm tra), ngoài trang `code-signing`. Đừng xóa câu này khi sửa trang chủ.
- "Reputation" là điểm yếu: kho GitHub tạo ngày 16/11/2026, 0 sao; mới khoảng 15 lượt tải. SignPath có thể từ chối hoặc hẹn nộp lại;
  khi đó tăng độ phổ biến (winget, người dùng thật, sao GitHub) rồi nộp lại.
- Tên nhà phát hành trên chữ ký là **SignPath Foundation** (không phải Microads) — đây là điều kiện của chương trình miễn phí.
- SignPath có quyền từ chối; nếu cần bổ sung, họ hỏi qua email — chuyển cho em để trả lời.
- Không bao giờ đưa khóa ký `.build\secrets\*.pem`, mật khẩu VPS hay token lên GitHub (`.gitignore` đã chặn thư mục `.build`).

## Trạng thái (16/11/2026 tối)

- Xong: kho mã công khai github.com/microads-vn/microads-office (GPL-3.0), build tự động trên GitHub chạy đạt, trang
  microads.vn/office/code-signing, bản 2.0.2 có tùy chọn quyền riêng tư (đã build và kiểm thử, chờ phát hành).
- Xong: xác thực 2 lớp GitHub (anh/chị đã bật); trang chủ có câu SignPath Foundation (đã triển khai); đơn tại signpath.org/apply đã được
  điền sẵn theo bảng ở Bước 2.
- Chờ anh/chị: xem lại, bấm **Submit** (và reCAPTCHA); sau đó chờ email của SignPath Foundation.
- Sau khi SignPath duyệt: làm Bước 3 rồi Bước 4 (gắn thẻ v2.0.2 → duyệt ký → phát hành bản có chữ ký, thay trình cài đặt trực tuyến bằng bản có chữ ký).
