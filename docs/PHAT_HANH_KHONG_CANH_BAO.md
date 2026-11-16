# Tải và cài Microads Office không bị cảnh báo — giải pháp miễn phí

Cập nhật 16/11/2026 · áp dụng cho bản 2.0.1 trở đi · Công ty TNHH Công nghệ Microads

## 1. Kết luận

| Câu hỏi | Trả lời (đã kiểm chứng) |
|---|---|
| Bộ cài có virus không? | **Không.** Microsoft Defender (chữ ký 1.463.324.0, bảo vệ đám mây bật) quét bộ cài, trình cài đặt trực tuyến và lệnh PowerShell: *found no threats*. Google Safe Browsing đánh giá tên miền microads.vn: *không có nội dung nguy hiểm*. |
| Vì sao Chrome và Windows vẫn cảnh báo? | Hai cảnh báo này dựa trên **danh tiếng**, không phải quét virus. Một tệp `.exe` chưa ký số được nhận diện bằng **mã băm SHA-256**. Mỗi phiên bản mới có mã băm mới nên danh tiếng **bắt đầu lại từ 0** sau mỗi lần phát hành. |
| Có chứng thư ký mã miễn phí không? | Với phần mềm **đóng nguồn** của doanh nghiệp: **không có**. Bảng so sánh chính thức của Microsoft (10/2026) chỉ có: Microsoft Store (miễn phí — anh/chị không dùng), Azure Artifact Signing (9,99 USD/tháng, chưa mở cho tổ chức ở Việt Nam), OV (150–300 USD/năm), EV (400+ USD/năm). Chứng thư tự ký bị coi như **không ký**. Nguồn miễn phí duy nhất là **SignPath Foundation**, chỉ dành cho dự án **mã nguồn mở** (mục 7). |
| Mua chứng thư có hết cảnh báo ngay không? | **Không.** Từ 2024, kể cả chứng thư EV cũng phải tích lũy danh tiếng: tệp mới vẫn bị cảnh báo vài tuần đầu. |
| Vậy làm gì? | (1) Cho khách tải một tệp **không bao giờ thay đổi** để danh tiếng được **cộng dồn mãi mãi** thay vì mất sau mỗi bản. (2) Mở thêm kênh cài **không qua trình duyệt** nên **không có cảnh báo ngay từ hôm nay**. Cả hai đều đã làm xong, miễn phí. |

## 2. Đã làm

| Hạng mục | Kết quả |
|---|---|
| **Trình cài đặt trực tuyến** `MicroadsOffice_Setup.exe` (80 KB) — **đã lên microads.vn** | Nút tải chính trên microads.vn/office. Tệp này **không bao giờ đổi** giữa các phiên bản. Nó tải bản mới nhất từ microads.vn, kiểm tra SHA-256 và **chữ ký số ECDSA của Microads** (cùng khóa với hệ thống cập nhật), rồi mới chạy. Tệp bị sửa, sai chữ ký hoặc liên kết không phải HTTPS đều bị từ chối và xóa. Bộ cài tải về **không mang dấu "tải từ Internet"** nên Windows chỉ hỏi **một lần** (ở trình cài đặt trực tuyến). |
| Thông tin nhà phát hành nhúng trong tệp | Công ty: *Công ty TNHH Công nghệ Microads (microads.vn)*, sản phẩm *Microads Office*, mô tả, bản quyền, biểu tượng bộ Office, manifest (chạy không cần quyền quản trị, hỗ trợ Windows 10/11). |
| Đóng băng tệp | Tệp chuẩn nằm ở `release\web-setup\MicroadsOffice_Setup.exe` (SHA-256 `76628b52a75a851756b69ac72853a8cd2256872a72e001253f306af246188271`, lưu trong `.sha256`). `build.ps1` dựng lại để kiểm tra nhưng **không thay** tệp đã phát hành. Máy chủ cũng **từ chối** một trình cài đặt trực tuyến khác, trừ khi chủ động dùng `--replace`. |
| **Cài bằng lệnh PowerShell** (không cảnh báo) | `irm https://microads.vn/office/install.ps1 \| iex`: tải trực tiếp, kiểm SHA-256 và chữ ký Microads, rồi mở trình cài. Không đi qua bước kiểm tra tải xuống của trình duyệt và SmartScreen. Chạy được trên Windows PowerShell 5.1 và PowerShell 7. |
| **winget** (kho gói miễn phí của Microsoft, không phải Store) | Bộ hồ sơ `tools\winget\manifests\m\Microads\Office\2.0.1` đã qua `winget validate`. Sau khi Microsoft duyệt: `winget install --id Microads.Office -e`, cài không cảnh báo. Xem `tools\winget\README.md`. |
| Gia cố theo kết quả rà soát mã độc lập | Chỉ cài bản **≥ 2.0.1** (máy chủ bị chiếm cũng không ép cài bản cũ có lỗi); nhận thêm **khóa ký dự phòng** (`.build\secrets\release-signing-backup.pem`); `--dir` phải là đường dẫn đầy đủ; hủy giữa chừng không thể làm bộ cài vẫn chạy; ổ đầy báo đúng lỗi; chờ phần mềm diệt virus nhả tệp; tự xin quyền nếu bộ cài sau này cần; đường lui TLS. Máy chủ: đếm lượt tải sang tệp riêng (không ghi đè cấu hình khi nhiều người tải cùng lúc), ghi cấu hình an toàn khi đầy ổ. |
| Cài hàng loạt cho IT | `MicroadsOffice_Setup.exe --silent [--dir "D:\Apps\Microads Office"]`: không cửa sổ, mã thoát 0 = thành công, 4 = mất mạng, 5 = sai chữ ký. |
| Trang tải microads.vn/office | Nút chính: trình cài đặt trực tuyến. Nút phụ: bộ cài đầy đủ. Mục *Cài đặt an toàn* nói đúng sự thật, có hướng dẫn Chrome *Giữ lại* và Windows *Thông tin bổ sung → Vẫn chạy*, SHA-256 của cả hai tệp. Mục *Cài bằng lệnh* (PowerShell, winget, IT) có nút **Sao chép** và lời nhắc an toàn chống lừa đảo dán lệnh. |
| Trang quản trị | *Tổng quan*: lượt tải trình cài đặt trực tuyến, lượt cài qua nó và qua PowerShell, mã SHA-256. *Nội dung quảng bá*: ô **Mã gói winget** để bật hoặc tắt lệnh winget trên trang tải. |
| Đã có từ 2.0.1 | Cài xong, mở Word/Excel/PDF không bị hỏi lại; cập nhật trong ứng dụng tải bằng chương trình nên **không bao giờ có cảnh báo**. |

## 3. Vì sao trình cài đặt trực tuyến làm cảnh báo biến mất

```
Trước:  2.0.0.exe (hash A) → 0 lượt  ┐
        2.0.1.exe (hash B) → 0 lượt  ├─ mỗi bản mới lại bị cảnh báo từ đầu
        2.0.2.exe (hash C) → 0 lượt  ┘

Nay:    MicroadsOffice_Setup.exe (hash cố định) → lượt tải của MỌI phiên bản dồn vào một tệp
                     └─ tự tải bản mới nhất, kiểm chữ ký Microads, chạy (không có dấu Internet → không hỏi lại)
```

- Microsoft (*SmartScreen reputation for Windows app developers*, 2026): tệp không ký phải tích lũy danh tiếng **cho từng phiên bản**; cảnh báo hết khi mã băm có đủ lịch sử tải sạch, *"có thể mất vài tuần và hàng trăm lượt cài từ nhiều người dùng"*.
- Eric Lawrence (cựu kỹ sư SmartScreen, *Best Practices for SmartScreen AppRep*): một trình cài đặt nhỏ, hiếm khi thay đổi sẽ giữ danh tiếng **nhiều năm qua mọi phiên bản**, với điều kiện tải qua HTTPS và **kiểm chữ ký trước khi chạy**. Cả hai điều kiện đều đã làm.
- Chrome Safe Browsing cũng đánh giá tệp không ký theo mã băm, nên áp dụng cùng nguyên lý.

**Diễn biến dự kiến:** những tuần đầu, trình cài đặt trực tuyến vẫn bị hỏi một lần như trước. Khi đủ lượt tải, cảnh báo của Chrome và Windows cho tệp này biến mất và **không quay lại khi phát hành bản mới**. Muốn nhanh: luôn chia sẻ đúng **một** liên kết `https://microads.vn/office/download/MicroadsOffice_Setup.exe` (Zalo, email, website), đừng gửi bộ cài đầy đủ, để mọi lượt tải dồn vào một mã băm.

## 4. Việc Microads cần làm

1. ~~Đưa lên máy chủ~~ — **đã xong 16/11/2026 20:08**: trang mới, `install.ps1`, trình cài đặt trực tuyến (máy chủ xác nhận cùng SHA-256).
   Đã cài thử thật từ microads.vn bằng cả trình cài đặt trực tuyến và lệnh PowerShell (vào thư mục thử, đã gỡ).
   **Sao lưu ngoại tuyến** (USB/két) hai khóa: `.build\secrets\update-signing.pem` và `.build\secrets\release-signing-backup.pem`.
2. **Gửi hồ sơ winget**: làm theo `tools\winget\README.md` (tài khoản GitHub, miễn phí, Microsoft duyệt 1–7 ngày), rồi nhập `Microads.Office` vào trang quản trị.
3. **Google Search Console — kênh chính thức để Chrome bỏ cảnh báo sớm** (miễn phí, cần tài khoản Google của công ty):
   1. https://search.google.com/search-console → *Thêm thuộc tính* → **Tiền tố URL** `https://microads.vn/office/` → phương thức **Thẻ HTML** → sao chép thẻ.
   2. Trang quản trị microads.vn/office/admin › *Nội dung quảng bá* › **Google Search Console (Chrome)** → dán thẻ → *Lưu* → quay lại Google bấm **Xác minh**.
   3. *Bảo mật và thao tác thủ công › Vấn đề bảo mật*: nếu có mục **Tệp tải xuống không phổ biến** (Uncommon downloads) → **Yêu cầu xem xét**.
      Google: *"cảnh báo được gỡ tự động khi Safe Browsing xác minh tệp an toàn"*; xem xét mất vài ngày đến vài tuần. Nếu cảnh báo quay lại, gửi xem xét lần nữa.
   (Nếu công ty đã có thuộc tính microads.vn trong Search Console thì vào thẳng bước 3.)
4. Chỉ khi có **phát hiện mã độc thật** (Chrome báo đỏ *"nguy hiểm"*, hoặc Defender báo tên như *Trojan:…*):
   - Google: https://safebrowsing.google.com/safebrowsing/report_error/
   - Microsoft: https://www.microsoft.com/en-us/wdsi/filesubmission → *Software developer* → *Incorrectly detected as malware*.
   - **Đính chính tài liệu cũ:** theo Microsoft (2026), *không có* kênh gửi xét duyệt danh tiếng SmartScreen cho máy người dùng phổ thông. Danh tiếng chỉ tăng theo lượt tải. Không cần gửi mỗi bản.
5. **Mỗi lần ra bản mới**: `build.ps1 -Publish`. Trình cài đặt trực tuyến giữ nguyên; hồ sơ winget cho phiên bản mới được tạo tự động, gửi thêm một PR cập nhật.
6. **Không bao giờ** dựng lại và thay trình cài đặt trực tuyến chỉ vì đổi giao diện. Chỉ dùng `-NewWebSetup` khi bắt buộc (khi đó danh tiếng bắt đầu lại). Không đổi khóa ký bản phát hành (khóa công khai nằm trong trình cài đặt trực tuyến).

## 5. Những cách KHÔNG dùng (và lý do)

| Cách | Vì sao không |
|---|---|
| Chứng thư tự ký | Microsoft xếp ngang "không ký"; khách vẫn thấy *Unknown publisher*. |
| Nén có mật khẩu, đổi đuôi tệp, đóng gói ISO/VHD, tách tệp, xóa "dấu Internet" | Đây là thủ thuật né kiểm tra của mã độc. Tên miền microads.vn có thể bị Google/Microsoft gắn cờ nguy hiểm thật. |
| Bảo khách tắt SmartScreen/Defender | Hại cho an toàn của khách và uy tín của Microads. |
| Microsoft Store | Anh/chị không muốn đăng ký Store (gói MSIX vẫn còn sẵn trong `tools\store` nếu sau này đổi ý). |

## 5b. Chrome đang báo gì (kiểm tra thực tế 16/11/2026, Chrome 157, hồ sơ mới)

- Cả `MicroadsOffice_Setup.exe` (80 KB) và bộ cài đầy đủ: **"Chrome đã chặn tệp tải xuống này vì tệp này ít được tải xuống và có thể nguy hiểm"**
  — mã `UNCOMMON_CONTENT` = *tệp lạ*, **không phải** phát hiện mã độc (khi có mã độc Chrome báo `DANGEROUS_CONTENT`, chữ đỏ).
- Khách tải tiếp (đã chạy thử đúng các bước): **Ctrl+J → bấm ⋮ cạnh tệp → "Tải tệp đáng ngờ xuống"** → tệp được lưu, không hỏi thêm.
  Trang tải đã ghi đúng từng chữ này (tiếng Anh: *Download suspicious file*).
- Nút "Quét tìm phần mềm độc hại" có trong giao diện nhưng bị ẩn với hồ sơ Chrome thường (chỉ dành cho chế độ bảo vệ nâng cao) → không hướng dẫn khách dùng.
- Trình cài đặt nhỏ vừa phát hành nên chưa có lượt tải nào trong hệ thống của Google: cảnh báo lúc này là **dự kiến**. Nó hết khi đủ lượt tải
  hoặc khi Google xem xét qua Search Console (mục 4.3). Tệp không đổi nên một khi hết là hết hẳn.

## 6. Giới hạn còn lại (nói rõ để không hứa sai)

- **Ngày đầu tiên**, tệp mới chưa ai tải thì *mọi* phương án tải qua trình duyệt đều có thể bị cảnh báo, kể cả có mua chứng thư. Kênh PowerShell và winget không bị.
- **Smart App Control** (Windows 11, chỉ bật trên một số máy cài mới) chặn hẳn ứng dụng *không ký* chưa có danh tiếng, không có nút *Vẫn chạy*. Chỉ chữ ký Authenticode hoặc danh tiếng mới giải quyết được. Đa số máy đang ở chế độ *Tắt* hoặc *Đánh giá*.
- Bộ cài đầy đủ (nút phụ) vẫn đổi mã băm mỗi bản nên vẫn bị cảnh báo như trước. Vì thế nút chính là trình cài đặt trực tuyến.

## 7. Nếu muốn có chữ ký số thật mà không tốn tiền: SignPath Foundation

- Miễn phí cho dự án **mã nguồn mở** (giấy phép OSI, mã nguồn công khai trên GitHub, build bằng GitHub Actions, có trang chính sách ký mã). SignPath Foundation duyệt hồ sơ trong vài tuần.
- Kết quả: tệp có chữ ký hợp lệ với nhà phát hành là *SignPath Foundation*. Danh tiếng gắn với chứng thư nên mọi bản sau được thừa hưởng, và Smart App Control cũng chấp nhận.
- Ràng buộc: phải công khai **toàn bộ** mã nguồn Microads Office. Đây là quyết định kinh doanh của công ty. Nếu đồng ý, có thể chuẩn bị quy trình GitHub Actions và trang chính sách ký mã.

## 8. Kiểm thử đã chạy (16/11/2026)

| Bộ | Kết quả |
|---|---|
| E2E trình cài đặt trực tuyến (UI Automation, máy chủ giả lập): tải → kiểm chữ ký → trình cài → cài xong; lỗi mạng → English → *Tải bộ cài đầy đủ* → *Thử lại*; *Hủy* khi đang tải; liên kết microads.vn; tệp bị sửa; chữ ký sai phiên bản; liên kết không HTTPS; chưa có bản phát hành; cài im lặng; mất mạng khi im lặng; `--dir` tương đối bị từ chối; bản cũ 2.0.0 có chữ ký thật bị từ chối; khóa dự phòng được chấp nhận; `--dir` có `%BIẾN%`; dọn tệp tạm; gỡ cài | **64/64** (`.build\e2e-web\report-websetup.txt`) |
| Trang web chạy cục bộ (PHP 8.3, `server\dev\test_site.py`): trang tải VI/EN, tải 2 tệp đúng từng byte, đếm lượt, `/download/latest`, chặn đường dẫn lạ, `install.ps1`, CLI chống thay trình cài đặt trực tuyến, quản trị (đăng nhập, thống kê, mã winget hợp lệ/sai/xóa, CSRF, các thẻ) | **43/43** (thêm: tải xuống không ghi `config.json`, đếm lượt bộ cài đầy đủ, chặn tên tệp có ký tự xuống dòng) |
| Lệnh PowerShell + trình cài đặt trực tuyến với trang PHP thật: Windows PowerShell 5.1 **18/18**, PowerShell 7.7 **18/18** (tiếng Việt hiển thị đúng, chữ ký đúng, thư mục `%BIẾN%` được mở rộng, thư mục tương đối bị từ chối, tệp bị sửa bị từ chối, chế độ tương tác mở trình cài và *Hủy* hoạt động, dọn tệp tạm) | **36/36** |
| winget: `winget validate` (v1.29) đạt; giả lập đúng lệnh winget chạy (tên tệp không có chữ "setup", `--silent --dir … --install`), khớp ProductCode và phiên bản, gỡ im lặng | Đạt |
| Kiểm tra chữ ký trên .NET Framework (máy khách thật): chữ ký đúng ✓, sai phiên bản ✗, sai mã băm ✗, chữ ký rác ✗; tệp bị khóa ghi trong lúc kiểm và chạy | Đạt |
| Dựng lại từ đầu cho ra **đúng cùng mã băm** (build tất định); `build.ps1` phát hiện tệp đóng băng bị sửa | Đạt |
| Microsoft Defender quét trình cài đặt trực tuyến và `install.ps1` | Không phát hiện mối đe dọa |
| Rà soát mã độc lập (bảo mật + luồng): không có cách nào khiến tệp không có chữ ký Microads được chạy; 1 lỗi thật (`--dir` tương đối) và các điểm gia cố đã sửa trước khi đóng băng | Đã xử lý |
| **Trên máy chủ thật** sau khi đưa lên: trang tải (nút chính = trình cài đặt trực tuyến 80 KB), tải về đúng từng byte, đếm lượt, máy chủ từ chối thay trình cài đặt trực tuyến khác, cài thật qua HTTPS bằng trình cài đặt trực tuyến (8 giây) và bằng lệnh PowerShell, trang quản trị (đăng nhập, CSRF, các thẻ, thẻ trình cài đặt trực tuyến, ô winget) | Đạt |

Nguồn: Microsoft Learn — *SmartScreen reputation for Windows app developers*, *Code signing options for Windows app developers* (2026); Eric Lawrence — *Best Practices for SmartScreen AppRep* (2024); Rick Strahl — *Dealing with Windows SmartScreen on Installation* (6/2026).
