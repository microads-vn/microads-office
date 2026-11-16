# Đưa Microads Office lên Microsoft Store — hướng dẫn từng bước

Cập nhật 17/11/2026. SignPath đã từ chối; Microsoft Store là kênh phân phối chính từ bản 2.0.3. Bản Store do **Microsoft ký số**
(miễn phí): khách cài không còn cảnh báo của Chrome / SmartScreen, Store tự cập nhật ứng dụng.

## Định danh sản phẩm (Partner Center › Microads Office › Product identity)

| Mục | Giá trị |
|---|---|
| Package/Identity/Name | `MicroadsTech.MicroadsOffice` |
| Package/Identity/Publisher | `CN=1BE10500-500A-4011-8C02-028DAF5EF854` (đã kiểm chứng bằng mã băm PFN) |
| Package/Properties/PublisherDisplayName | `Microads Tech` |
| Package Family Name | `MicroadsTech.MicroadsOffice_p0kmfkpdegte6` |
| Store ID | `9PHM0WX16WDR` — https://apps.microsoft.com/detail/9PHM0WX16WDR |
| Loại sản phẩm | **MSIX or PWA app** (đúng) |

Đóng gói (máy build): `build.ps1 -Store -StoreIdentityName MicroadsTech.MicroadsOffice -StorePublisher "CN=1BE10500-500A-4011-8C02-028DAF5EF854"`
→ `release\store\MicroadsOffice_<phiên bản>.0_x64.msix` (tên hiển thị nhà phát hành mặc định "Microads Tech", phải khớp Partner Center).

## Bước 1 — Giữ tên ✅ (đã xong 17/11/2026)

## Bước 2 — Tạo hồ sơ nộp (Start submission)

Partner Center › **Apps and games** › **Microads Office** › **Application overview** › **Start submission**. Có 6 mục; làm mục nào
xong mục đó bấm **Save draft**.

### 2.1 Pricing and availability
- **Markets**: để mặc định (tất cả thị trường).
- **Discoverability**: *Make this product available and discoverable in the Microsoft Store*.
- **Schedule**: *Release as soon as possible*.
- **Base price**: **Free**. Free trial: *No free trial*. Các mục còn lại để mặc định → **Save draft**.

### 2.2 Properties
- **Category**: **Productivity** (không chọn subcategory).
- **Privacy policy URL**: `https://microads.vn/office/privacy`
- **Website**: `https://microads.vn/office/`
- **Support contact info**: `https://microads.vn/office/` (hoặc email hỗ trợ của công ty).
- **Contact details** (bắt buộc với tài khoản công ty): tên, địa chỉ, điện thoại, email của công ty.
- **Product declarations**: tích *Customers can install this app to alternate drives or removable storage* và *Windows can include
  this app's data in automatic backups*. Nếu có ô khai báo **generative AI / AI-generated content** thì **tích** (chính sách 11.16).
  Không tích ô "accessibility guidelines" (chưa kiểm định).
- **System requirements**: Minimum — *Memory 4 GB*; Architecture — *x64*. → **Save draft**.

### 2.3 Age ratings
Bấm **Start questionnaire**, chọn loại **"All other app types"** (không phải game), trả lời trung thực:
- Bạo lực, tình dục, ngôn ngữ thô tục, ma túy, cờ bạc: **No**.
- Người dùng nói chuyện/chia sẻ nội dung với **người khác** trong ứng dụng: **No** (trợ lý AI không phải kênh trò chuyện giữa người với người).
- Chia sẻ vị trí: **No**. Mua bán trong ứng dụng: **No**. Trình duyệt web không giới hạn: **No**.
- Nếu có câu hỏi về **nội dung do AI tạo ra (generative AI)**: **Yes** — trợ lý AI tạo văn bản theo yêu cầu, có nút **Báo cáo**.
Kết quả dự kiến: mọi lứa tuổi / 3+ (có thể kèm nhãn "AI-generated content"). → **Save draft**.

### 2.4 Packages
Kéo thả tệp **`release\store\MicroadsOffice_2.0.3.0_x64.msix`** (tôi đóng gói và gửi đường dẫn khi xong) → chờ chữ **Validated**.
Device family: chỉ **Windows 10/11 Desktop** → **Save draft**.

### 2.5 Store listings
Bấm **Add/remove languages** → chọn **Vietnamese (Vietnam)** và **English (United States)**. Với từng ngôn ngữ, dán các mục dưới đây
(mục "Store listing — Tiếng Việt / English"). Ảnh chụp màn hình: tải các tệp trong **`release\store\screenshots\`** (1920×1080, tối
thiểu 1, nên ≥ 4). Store logos: tệp **`release\store\logos\StoreLogo-300.png`** (1:1, 300×300) nếu được hỏi. → **Save draft**.

### 2.6 Submission options
- **Restricted capabilities** — ô giải trình cho `runFullTrust` (dán nguyên văn):
  > Microads Office is a classic Win32 desktop office suite (word processor, spreadsheet, PDF reader with digital signatures) packaged
  > with the Desktop Bridge. runFullTrust is required to run the desktop executable, to open and save documents anywhere the user
  > chooses, to use the Windows certificate store and USB token drivers for PDF digital signatures, and to print. It does not install
  > drivers or services and does not update itself outside the Store.
- **Notes for certification** (dán nguyên văn):
  > No account or sign-in is needed to test Microads Office. Start "Microads Office" (hub), then Word, Excel or PDF from the Start menu.
  > Live generative AI (Store policy 11.16): the optional "AI assistant" (purple sparkle button on the Home tab of Word, Excel and PDF)
  > sends a request only when the user clicks Send, to the AI service the user connects (their own ChatGPT account via OpenAI's
  > "Sign in with ChatGPT", their own Claude/OpenAI API key, or their own OpenAI-compatible server such as Ollama). Every answer has a
  > "Report" (Báo cáo) button that sends the reported content to the developer (microads.vn) for review. The assistant is not
  > connected by default, so certification can test the rest of the app without any AI account. The Store edition never downloads
  > or installs updates by itself (updates come only through the Store). Interface language: Vietnamese (default) or English
  > (flag button, top right).
- Publishing hold: để mặc định (*Publish as soon as it passes certification*). → **Save draft**.

## Bước 3 — Gửi kiểm duyệt
**Application overview** › **Submit for certification**. Microsoft thường duyệt trong 1–3 ngày làm việc (có thể lâu hơn lần đầu).
Nếu bị trả về: chụp màn hình báo cáo lỗi (Certification report) gửi cho tôi.

## Bước 4 — Sau khi lên Store
1. Trang quản trị https://microads.vn/office/admin › **Nội dung quảng bá** › **Microsoft Store** › nhập `9PHM0WX16WDR` › **Lưu liên kết**:
   nút tải chính trên microads.vn/office chuyển sang Microsoft Store.
2. Khách đang dùng bản cài từ web (2.0.1): giữ nguyên, họ vẫn dùng được; có thể cài bản Store (dữ liệu cài đặt tách biệt).
3. Mỗi bản mới: tăng `<Version>` → `build.ps1 -Store …` → Partner Center › **Update** (tạo hồ sơ mới) › Packages › tải .msix mới › Submit.

---

## Store listing — Tiếng Việt

**Tên sản phẩm**: Microads Office

**Mô tả** (Description):
```
Microads Office — bộ ứng dụng văn phòng miễn phí cho Windows: soạn thảo văn bản (Word), bảng tính (Excel) và đọc, ký số PDF trong một lần cài. Giao diện tiếng Việt và tiếng Anh, không cần tài khoản, không quảng cáo theo dõi.

• Microads Office Word: soạn văn bản theo đúng trang in, kiểu Tiêu đề, bảng, hình ảnh, đầu/chân trang, mục lục, mẫu văn bản hành chính theo Nghị định 30/2020/NĐ-CP; mở và lưu .docx, .rtf, .txt; xuất PDF; in.
• Microads Office Excel: hơn 230 hàm (có DOCSO đọc số thành chữ), biểu đồ, định dạng có điều kiện, sắp xếp và lọc; mở và lưu .xlsx, .csv; in và xuất PDF.
• Microads Office PDF: đọc PDF nhanh, ký số bằng USB Token hoặc tệp .pfx (PAdES), kiểm tra chữ ký, chèn dấu ảnh, in.

Trợ lý AI (tùy chọn) — dùng AI tạo sinh: soạn, viết lại, dịch, sửa chính tả ngay trong văn bản; viết công thức, tạo bảng và biểu đồ trong trang tính; tóm tắt và hỏi đáp tệp PDF. Mỗi thay đổi đều hoàn tác được (Ctrl+Z). Trợ lý chỉ hoạt động khi bạn tự kết nối một dịch vụ AI của mình: tài khoản ChatGPT Plus/Pro (Continue with ChatGPT), khóa API Claude hoặc OpenAI, máy chủ AI riêng như Ollama, hoặc dùng gói Claude Pro/Max qua ứng dụng Claude Desktop. Nội dung chỉ được gửi tới dịch vụ bạn chọn khi bạn bấm Gửi; câu trả lời do AI tạo có thể chưa chính xác — hãy kiểm tra lại. Mỗi câu trả lời có nút Báo cáo để gửi nội dung không phù hợp cho Microads.

Phát hành miễn phí bởi Công ty TNHH Công nghệ Microads (microads.vn). Mã nguồn mở GPL-3.0: github.com/microads-vn/microads-office
```

**Có gì mới** (What's new in this version):
```
Phiên bản đầu tiên trên Microsoft Store: Trợ lý AI trong Word, Excel và PDF; cài đặt và cập nhật qua Microsoft Store, không còn cảnh báo khi tải.
```

**Tính năng** (App features, mỗi dòng một mục):
```
Soạn thảo văn bản theo trang in thật, mở/lưu .docx
Mẫu văn bản hành chính theo Nghị định 30/2020/NĐ-CP
Bảng tính hơn 230 hàm, biểu đồ, lọc và sắp xếp, mở/lưu .xlsx
Hàm DOCSO đọc số thành chữ tiếng Việt
Đọc PDF, ký số bằng USB Token hoặc .pfx, kiểm tra chữ ký
Trợ lý AI: soạn, viết lại, dịch, sửa chính tả ngay trong văn bản
Trợ lý AI cho Excel: viết công thức, tạo bảng và biểu đồ
Kết nối ChatGPT, Claude, khóa API hoặc Ollama (tùy chọn)
Giao diện tiếng Việt và tiếng Anh
Miễn phí, không cần tài khoản
```

**Từ khóa** (Keywords, tối đa 7): `văn phòng`, `soạn thảo văn bản`, `bảng tính`, `ký số PDF`, `docx xlsx`, `trợ lý AI`, `miễn phí`

**Bản quyền** (Copyright and trademark info): `© 2026 Công ty TNHH Công nghệ Microads`

**Phát triển bởi** (Developed by): `Công ty TNHH Công nghệ Microads`

## Store listing — English

**Product name**: Microads Office

**Description**:
```
Microads Office is a free office suite for Windows: a word processor (Word), spreadsheets (Excel) and a PDF reader with digital signatures, in one install. Vietnamese and English interface, no account, no tracking ads.

• Microads Office Word: page-accurate word processing with heading styles, tables, pictures, headers/footers, table of contents and Vietnamese administrative templates (Decree 30/2020); open and save .docx, .rtf, .txt; export to PDF; print.
• Microads Office Excel: 230+ functions (including DOCSO, numbers to Vietnamese words), charts, conditional formatting, sort and filter; open and save .xlsx, .csv; print and export to PDF.
• Microads Office PDF: fast PDF viewing, digital signatures with USB tokens or .pfx files (PAdES), signature validation, image stamps, printing.

AI assistant (optional) — uses generative AI: drafts, rewrites, translates and proofreads right in the document; writes formulas, tables and charts in the sheet; summarises and answers questions about PDFs. Every change can be undone (Ctrl+Z). The assistant works only after you connect an AI service of your own: your ChatGPT Plus/Pro account (Continue with ChatGPT), a Claude or OpenAI API key, your own AI server such as Ollama, or your Claude Pro/Max plan through the Claude Desktop app. Content is sent only to the service you chose, when you click Send; AI-generated answers can be inaccurate — please review them. Every answer has a Report button to send inappropriate content to Microads.

Published free of charge by Microads Technology Company Limited (microads.vn). Open source (GPL-3.0): github.com/microads-vn/microads-office
```

**What's new in this version**:
```
First Microsoft Store release: AI assistant in Word, Excel and PDF; installed and updated through the Microsoft Store, no more download warnings.
```

**App features**:
```
Page-accurate word processing, open/save .docx
Vietnamese administrative document templates (Decree 30/2020)
Spreadsheets with 230+ functions, charts, sort and filter, open/save .xlsx
DOCSO function: numbers to Vietnamese words
PDF viewing and digital signatures with USB tokens or .pfx
AI assistant: draft, rewrite, translate and proofread in the document
AI assistant for Excel: formulas, tables and charts
Connect ChatGPT, Claude, an API key or Ollama (optional)
Vietnamese and English interface
Free, no account needed
```

**Keywords**: `office suite`, `word processor`, `spreadsheet`, `PDF signature`, `docx xlsx`, `AI assistant`, `free`

**Copyright**: `© 2026 Microads Technology Company Limited` · **Developed by**: `Microads Technology Company Limited`
