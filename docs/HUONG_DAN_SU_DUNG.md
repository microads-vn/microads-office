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
| Microads Office **Excel** | Lối tắt "Microads Office Excel", hoặc .xlsx/.csv | Bảng tính, 502 hàm (đủ các nhóm như Excel 365, `=DOCSO(A1)` đọc số thành chữ), biểu đồ, lọc/sắp xếp, định dạng có điều kiện, in/PDF |
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

### 4b. Trợ lý AI (từ bản 2.0.3)
Nút **Trợ lý AI** (ngôi sao tím) ở cuối thẻ **Trang chủ** của Word và Excel, và trên thẻ Trang chủ của PDF, mở khung trợ lý bên phải.

**Kết nối (chọn một cách, chỉ làm một lần, dùng chung cho Word/Excel/PDF):**

| Bạn có | Làm thế này |
|---|---|
| Tài khoản **ChatGPT Plus / Pro** | Khung trợ lý › **Continue with ChatGPT** → trình duyệt mở trang đăng nhập ChatGPT → đăng nhập và **đồng ý cho Microads Office dùng gói ChatGPT** → quay lại ứng dụng. Không cần khóa API. Xem/giới hạn mức dùng: dòng *Đang dùng gói ChatGPT · Quản lý hạn mức* (mở chatgpt.com/settings/usage). |
| Gói **Claude Pro / Max** | Cài ứng dụng **Claude Desktop** (claude.ai/download) và đăng nhập ở đó. Trong Microads: Trợ lý AI › **Dùng Claude** › **Kết nối với Claude Desktop** → khởi động lại Claude Desktop. Rồi trong Claude Desktop yêu cầu, ví dụ: *"Đọc văn bản đang mở trong Microads Word và viết lại phần kết luận"*; Claude hỏi bạn trước mỗi thao tác. (Anthropic không cho phần mềm khác đăng nhập tài khoản Claude, nên đây là cách chính thức để dùng gói Claude.) |
| **Khóa API** Claude hoặc OpenAI | Trợ lý AI › Cài đặt › dán khóa (bấm *Lấy khóa* để mở trang tạo khóa) › **Kiểm tra kết nối** › chọn mô hình › **Lưu**. Trả phí theo mức dùng cho Anthropic/OpenAI. |
| Máy chủ AI riêng / **Ollama** (miễn phí, chạy trên máy) | Cài Ollama và tải một mô hình; Trợ lý AI › Cài đặt › *Máy chủ tương thích OpenAI* (mặc định `http://127.0.0.1:11434/v1`) › Kiểm tra kết nối › chọn mô hình › Lưu. Dữ liệu không ra khỏi máy/mạng nội bộ. |

**Sử dụng:**
- Gõ yêu cầu rồi **Enter** (Shift+Enter để xuống dòng), hoặc bấm các nút nhanh: Word — *Tóm tắt, Viết lại, Sửa chính tả, Dịch sang tiếng Anh/Việt, Rút gọn, Lập dàn ý, Soạn công văn*; Excel — *Giải thích công thức, Viết công thức, Phân tích dữ liệu, Bảng tổng hợp, Vẽ biểu đồ, Làm sạch dữ liệu, Tạo bảng mẫu*; PDF — *Tóm tắt, Điểm cần lưu ý, Dịch trang đang xem, Giải thích đoạn chọn, Trích bảng số liệu*.
- Muốn sửa một đoạn: **bôi đen** đoạn đó trước rồi yêu cầu ("viết lại trang trọng hơn", "dịch sang tiếng Anh").
- Khi bạn yêu cầu soạn/sửa, trợ lý **ghi thẳng vào văn bản** (đúng kiểu Tiêu đề 1–3, căn giữa quốc hiệu, danh sách, bảng) hoặc **vào trang tính** (giá trị, công thức, định dạng, biểu đồ, trang mới) và hiện danh sách **Các bước**. Mỗi lần ghi là **một bước Hoàn tác** — bấm **Hoàn tác** dưới câu trả lời hoặc Ctrl+Z.
- Bỏ chọn *Tự ghi kết quả vào văn bản* nếu muốn xem trước rồi tự bấm **Ghi vào văn bản**. Câu trả lời dạng chữ có nút **Chèn vào văn bản / Thay vùng chọn / Ghi vào ô đang chọn / Sao chép**.
- Excel: công thức do trợ lý viết được **kiểm tra trước** (cú pháp, tên hàm có trong Microads Excel); nếu có lỗi thì **không ghi gì** và nêu rõ lỗi.
- **Cuộc trò chuyện mới** (dấu +) xóa lịch sử; **Dừng** để ngắt câu trả lời đang chạy.
- *Cài đặt › Hướng dẫn luôn áp dụng*: ví dụ "văn phong hành chính, công ty tôi là …".

**Quyền riêng tư:** trợ lý chỉ kết nối khi bạn bấm Gửi; nội dung được gửi thẳng tới dịch vụ bạn chọn (không qua máy chủ Microads). Bỏ chọn
*Gửi kèm nội dung văn bản đang mở* nếu chỉ muốn hỏi chung. Khóa và phiên đăng nhập được mã hóa trên máy (Windows DPAPI). Đăng xuất / Xóa khóa /
Ngắt kết nối Claude Desktop trong Cài đặt Trợ lý AI. Chi tiết: https://microads.vn/office/privacy#ai-assistant

### 4c. Dán (Ctrl+V) và kéo thả từ chương trình khác (từ bản 2.0.4)

| Bạn làm | Word | Excel | PDF |
|---|---|---|---|
| Chụp màn hình (PrtSc, Snipping Tool, Zalo…) rồi Ctrl+V | Chèn ảnh tại con trỏ | Đặt ảnh tại ô đang chọn (kể cả khi trước đó đã sao chép ô) | Ảnh thành **dấu ảnh**: bấm vào trang để đặt, Ctrl+S để lưu |
| Chọn ảnh rồi Ctrl+C / Ctrl+X | Ảnh lên bộ nhớ tạm dạng ảnh (dán được sang Zalo, Paint, Excel…) | Như Word; dán lại trong Excel giữ nguyên kích thước | — |
| Sao chép tệp / thư mục trong Explorer rồi Ctrl+V | Ảnh được chèn; tệp khác và thư mục thành **liên kết** (Ctrl+bấm mở thư mục / hiện tệp trong Explorer, không bao giờ chạy chương trình) | Ảnh được đặt; tệp khác và thư mục ghi vào các ô bên dưới, có liên kết | Ảnh đầu tiên thành dấu ảnh |
| Kéo thả từ chương trình khác (chữ, ảnh từ trình duyệt / Zalo / Paint, tệp) | Chèn đúng chỗ thả chuột | Chèn vào ô dưới chuột | PDF được mở; ảnh thành dấu ảnh |
| Kéo thả tài liệu .docx / .xlsx / .pdf | Mở tài liệu | Mở tài liệu | Mở tệp PDF |

Mỗi lần dán là một bước Hoàn tác (Ctrl+Z). Ảnh từ trang web chỉ lấy khi trình duyệt đã đưa sẵn ảnh vào bộ nhớ tạm hoặc khi kéo thả; Microads Office không tự tải ảnh từ Internet.

### 4d. Hàm và mảng động trong Excel (từ bản 2.0.4)

Microads Excel có **502 hàm** — đủ các nhóm của Excel 365: Toán học, Thống kê (cả các hàm phân phối NORM.DIST, T.TEST…), Logic,
Tra cứu, Văn bản (cả REGEXTEST/REGEXEXTRACT/REGEXREPLACE, TEXTSPLIT), Ngày giờ, Tài chính (cả trái phiếu PRICE, YIELD, XIRR…),
Kỹ thuật (đổi cơ số, số phức, CONVERT…), Cơ sở dữ liệu (DSUM, DCOUNT…), Web (ENCODEURL, FILTERXML) và nhóm Tương thích
(tên hàm cũ như NORMDIST, STDEVP). Xem tất cả: thẻ **Công thức › Chèn hàm** (Shift+F3) hoặc nút **Hàm khác**.

- **Mảng động (tràn ô)**: công thức trả về nhiều giá trị chỉ cần gõ ở **một ô**, kết quả tự tràn xuống/sang phải, ví dụ
  `=FILTER(A2:C100, C2:C100>1000000)`, `=SORT(UNIQUE(B2:B100))`, `=SEQUENCE(12)`. Vùng tràn có viền xanh; bấm ô trong vùng tràn,
  thanh công thức hiện công thức gốc màu xám. Nếu vùng tràn bị ô có dữ liệu chắn, ô công thức báo **#SPILL!** — xóa ô chắn là xong.
  Tham chiếu cả vùng tràn bằng dấu `#`: `=SUM(E2#)`.
- **Tính theo từng phần tử**: `=SUMPRODUCT(--(LEFT(A2:A100,2)="HN"))`, `=LEN(A2:A10)`, `=SUM(COUNTIF(B:B,{"Đạt","Khá"}))` chạy như Excel 365.
- **LET và LAMBDA**: đặt tên cho phần tính trung gian `=LET(thue, B2*10%, B2+thue)`; tạo hàm riêng trong **Trình quản lý tên**
  (ví dụ tên `VAT` = `=LAMBDA(x, x*10%)` rồi dùng `=VAT(B2)`), cùng MAP, REDUCE, SCAN, BYROW, BYCOL, MAKEARRAY.
- **Tổng hợp nhanh**: `=GROUPBY(A2:A100, C2:C100, SUM)` và `=PIVOTBY(A2:A100, B2:B100, C2:C100, SUM)` (bảng tổng hợp bằng công thức).
- Tệp .xlsx lưu từ Microads mở đúng trong Microsoft Excel 365/2021 (công thức mảng động, LET, LAMBDA giữ nguyên). Tệp từ Excel đời cũ
  (2016 trở về trước) vẫn tính như cũ: công thức cũ không tự tràn ô.

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
