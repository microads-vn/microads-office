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

| Trường | Nội dung |
|---|---|
| Project name | `Microads Office` |
| Repository URL | `https://github.com/microads-vn/microads-office` |
| Project website / download page | `https://microads.vn/office/` |
| Code signing policy URL | `https://microads.vn/office/code-signing` |
| Privacy policy URL | `https://microads.vn/office/privacy` |
| License | `GPL-3.0-only` |
| Build system | `GitHub Actions (.github/workflows/release.yml)` |
| Artifacts to sign | `Windows executables: MicroadsOffice.exe (suite, distributed as MicroadsOffice_Setup_<version>.exe) and MicroadsOffice_Setup.exe (web installer)` |
| Organization / maintainer | `Microads Technology Company Limited (Công ty TNHH Công nghệ Microads), Ho Chi Minh City, Vietnam — tax code 0319785225 — https://microads.vn` |
| Contact | Gmail công ty; người đại diện: Phạm Thị Bé Bi (Legal representative) |

**Project description** (dán vào ô mô tả):
> Microads Office is a free, open source (GPL-3.0) office suite for Windows published by Microads Technology Company Limited (Vietnam):
> a word processor (.docx), a spreadsheet application (.xlsx/.csv, ~230 functions, charts) and a PDF reader with PAdES digital
> signatures (USB tokens / .pfx), in one self-contained executable with a per-user installer and uninstaller, Vietnamese and English UI.
> It is used by Vietnamese offices and public administration staff who need free document and digital-signature tools.
> Releases have been published since November 2026 at https://microads.vn/office/ and are additionally signed with our own ECDSA release key,
> verified by the apps and the update server. We need Authenticode signing because Windows SmartScreen and Chrome warn our users about the
> unsigned installer. Builds are produced by GitHub Actions from the public repository; every signing request will be approved manually by
> the project owner; all team members use MFA. The only automatic network access (update check) is described in our privacy policy,
> shown during installation and can be switched off there.

3. Gửi đơn. SignPath Foundation xét duyệt (thường vài ngày đến vài tuần) và trả lời qua email.

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
- Tên nhà phát hành trên chữ ký là **SignPath Foundation** (không phải Microads) — đây là điều kiện của chương trình miễn phí.
- SignPath có quyền từ chối; nếu cần bổ sung, họ hỏi qua email — chuyển cho em để trả lời.
- Không bao giờ đưa khóa ký `.build\secrets\*.pem`, mật khẩu VPS hay token lên GitHub (`.gitignore` đã chặn thư mục `.build`).
