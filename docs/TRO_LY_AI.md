# Trợ lý AI trong Microads Office — yêu cầu, thiết kế, đối chiếu

Bắt đầu 16/11/2026. Tài liệu này là **danh sách chống quên**: mỗi yêu cầu của khách có một dòng, cột "Trạng thái" chỉ được ghi "Đạt"
khi đã có kiểm thử tự động hoặc kiểm tra thực tế ghi ở cột "Bằng chứng".

## 1. Yêu cầu của khách (nguyên văn ý)

| # | Yêu cầu | Trạng thái | Bằng chứng |
|---|---|---|---|
| K1 | Tích hợp sẵn công cụ AI (GPT, Claude) vào Word và Excel giống Microsoft Word/Excel | **Đạt** (thêm cả PDF) | Nút "Trợ lý AI" thẻ Trang chủ Word/Excel/PDF, khung bên phải, ghi thẳng vào văn bản/trang tính kèm "Các bước": W01–W11, X01–X07, P01–P03 |
| K2 | Người dùng **đăng nhập tài khoản ChatGPT / Claude trả phí** (ưu tiên, vì ít người biết khóa API) | **Đạt** — ChatGPT đăng nhập trực tiếp; Claude qua Claude Desktop (Anthropic cấm phần mềm khác đăng nhập Claude — mục 2) | C04–C11, C17, W13 (nút Continue with ChatGPT qua trình duyệt), C21–C22, W14, E01–E04 (exe thật) |
| K3 | Vẫn cho **nhập khóa API** | **Đạt** — Claude, OpenAI, máy chủ riêng/Ollama | C14–C16, W12 |
| K4 | Mọi nút bấm được, tương tác được, làm việc được, **không ẩn tính năng** | **Đạt** — nút không dùng được thì hiện thông báo/giải thích, không ẩn (vd. nút cần vùng chọn, PDF chỉ đọc ghi rõ) | W01–W02, W05, W10 (bấm **mọi** nút nhanh), P01–P03, Word UI11/Excel test bấm mọi nút ribbon |
| K5 | Mọi nút qua kiểm thử | **Đạt** — 53 bài AI + các bộ cũ (Word 37, Excel 28, PDF 20) | `.build\ai-test\report-ai-*.txt`, `build.ps1`, GitHub Actions |
| K6 | Luồng làm việc chuẩn, móc nối logic với hệ thống (văn bản đang mở, vùng chọn, hoàn tác, cài đặt, quyền riêng tư) | **Đạt** | Ngữ cảnh = văn bản + vùng chọn (W03, W04, X04, P02); 1 bước Hoàn tác (W03, X02); cài đặt dùng chung 3 ứng dụng; Cài đặt chung › Trợ lý AI (W15); chính sách quyền riêng tư cập nhật |
| K7 | Tự tạo vai trò và luồng làm việc, kiểm tra chéo | **Đạt** | Mục 3 và mục 6 |
| K8 | Không ghi gì vào ổ C/D khi làm việc; công cụ trên C/D chạy xong phải dọn | **Đạt** — SDK, NuGet, bộ đệm, dữ liệu kiểm thử, hồ sơ Edge đều ở `E:\Office\.build`; kiểm thử Claude Desktop dùng thư mục giả (`MICROADS_CLAUDE_CONFIG_ROOT`) | Rà soát C: lúc 22:45: chỉ có thay đổi của chính người dùng/Windows (jump list vài KB do Windows tự ghi khi mở cửa sổ) |

## 2. Ràng buộc bên ngoài (đã kiểm chứng 16/11/2026)

- **ChatGPT**: OpenAI có chương trình chính thức "Sign in with ChatGPT — ChatGPT plan usage" cho ứng dụng **mã nguồn mở** (Microads
  Office là GPL-3.0): OAuth PKCE, `client_id=dynamic_agent_client` lần đầu, callback `http://127.0.0.1:<port>/auth/callback`, scope
  `openid profile email offline_access resource.invoke chatgpt.tokens.use.direct`, gọi `POST https://api.openai.com/v1/responses`
  với `store:false, stream:true`. Nút phải ghi **"Continue with ChatGPT"**. Nguồn: developers.openai.com/siwc.
- **Claude**: Anthropic **cấm** phần mềm bên thứ ba cho đăng nhập tài khoản Claude (Free/Pro/Max) hoặc chuyển yêu cầu qua gói của
  người dùng (code.claude.com/docs/en/legal-and-compliance, mục Authentication and credential use). Add-in "Claude for Word/Excel"
  trong Microsoft Office là **của chính Anthropic** nên mới được đăng nhập. Cách hợp lệ để khách dùng gói Claude trả phí với
  Microads Office: **kết nối Microads Office vào ứng dụng Claude Desktop** (chuẩn MCP) — đăng nhập diễn ra trong ứng dụng của
  Anthropic. Ngoài ra: khóa API Claude (Claude Console).
- **SignPath / quyền riêng tư**: dữ liệu gửi tới hệ thống người dùng không chỉ định phải được mô tả và cho tắt. Trợ lý AI chỉ kết nối
  khi người dùng **chủ động** chọn nhà cung cấp và bấm gửi; nội dung văn bản chỉ gửi tới nhà cung cấp người dùng đã chọn.

## 3. Vai trò (người dùng giả định để kiểm tra luồng)

| Vai trò | Có gì | Cần gì từ trợ lý | Luồng phải chạy được |
|---|---|---|---|
| V1 Nhân viên văn phòng | Tài khoản ChatGPT Plus | Soạn công văn, tóm tắt, sửa chính tả, dịch | Cài đặt › Trợ lý AI › **Continue with ChatGPT** → trình duyệt → quay lại ứng dụng đã đăng nhập → Word › Trợ lý AI → hỏi → **Chèn / Thay vùng chọn** → Hoàn tác được |
| V2 Kế toán | Gói Claude Pro + ứng dụng Claude Desktop | Đọc bảng tính, viết công thức, tổng hợp số liệu | Cài đặt › Trợ lý AI › **Kết nối với Claude Desktop** → mở Claude Desktop → ra lệnh "đọc bảng đang mở trong Microads Excel" → Claude đọc/ghi qua Microads Office |
| V3 Người dùng kỹ thuật | Khóa API OpenAI hoặc Claude | Dùng mô hình mình chọn | Cài đặt › Trợ lý AI › dán khóa → **Kiểm tra kết nối** → dùng trong Word/Excel/PDF |
| V4 Cơ quan không cho gửi dữ liệu ra ngoài | Máy chủ AI nội bộ / Ollama | AI chạy trong mạng nội bộ | Cài đặt › Trợ lý AI › "Máy chủ tương thích OpenAI" (địa chỉ 127.0.0.1 hoặc nội bộ) |
| V5 Người không dùng AI | — | Không bị làm phiền, không kết nối mạng | Trợ lý mặc định **chưa kết nối**; khung trợ lý chỉ hiện hướng dẫn; không gửi gì khi chưa chọn nhà cung cấp |
| V6 Quản trị Microads (người phát hành) | Mã nguồn, máy chủ | Không giữ khóa của khách, đúng điều khoản OpenAI/Anthropic/SignPath | Khóa và mã thông báo chỉ lưu trên máy khách, mã hóa DPAPI; chính sách quyền riêng tư mô tả đúng |

## 4. Thiết kế (đã làm)

| Phần | Tệp | Ghi chú |
|---|---|---|
| Cài đặt, khóa bí mật | `src/MicroadsOffice.AI/AiSettings.cs` | `office.json` › `AppData.ai` (không chứa bí mật); khóa/phiên trong `Data\ai\*.bin` mã hóa DPAPI (người dùng Windows hiện tại) |
| Đăng nhập ChatGPT | `ChatGptAuth.cs` | `dynamic_agent_client` → client id được cấp; PKCE S256, state, nonce; cổng 127.0.0.1:1455 (hoặc cổng trống); kiểm chữ ký ID token bằng JWKS, iss/aud/exp/nonce; kiểm scope `chatgpt.tokens.use.direct`; gia hạn trước hạn 2 phút, khóa liên tiến trình (semaphore) vì Word/Excel/PDF là 3 tiến trình; đăng xuất thu hồi refresh token |
| Nhà cung cấp | `Providers/*.cs` | ChatGPT + khóa OpenAI: Responses API (`store:false, stream:true`, `instructions`, không gửi trường bị cấm); Claude: SDK chính thức `Anthropic` 13.8.0, `claude-opus-5-5`, effort medium, fallbacks "default"; Ollama/máy chủ riêng: Chat Completions |
| Khung trợ lý | `UI/AiPane.cs` | Trạng thái tài khoản, "Đang dùng gói ChatGPT · Quản lý hạn mức", nút nhanh, hội thoại có "Các bước", Gửi/Dừng, Chèn/Thay vùng chọn/Ghi vào ô/Sao chép/Hoàn tác/Thử lại |
| Cài đặt AI | `UI/AiSettingsDialog.cs` | Continue with ChatGPT / Đăng xuất / tài khoản khác / hạn mức; Claude Desktop; khóa Claude/OpenAI (Lấy khóa, Xóa khóa); máy chủ riêng; mô hình + Kiểm tra kết nối; hướng dẫn luôn áp dụng; quyền riêng tư |
| Giao thức ghi | `AiHost.cs` | Khối ```` ```microads-document mode=… ```` (Word) và ```` ```microads-sheet ```` JSON (Excel) — hoạt động với mọi nhà cung cấp, không phụ thuộc "tool calling" |
| Word | `Word/AI/WordMarkdown.cs`, `WordMainForm.Ai.cs` | Markdown ↔ mô hình Word (Tiêu đề 1–3, căn giữa `:::center`, danh sách, bảng có hàng tiêu đề, ngắt trang); 1 lần ghi = 1 bước Hoàn tác |
| Excel | `Excel/AI/SheetActions.cs`, `ExcelForm.Ai.cs` | Kiểm tra toàn bộ trước (địa chỉ, cú pháp, tên hàm có trong Microads Excel), lỗi thì không ghi gì; giá trị/công thức/định dạng/biểu đồ/sắp xếp/cố định/độ rộng/trang mới; 1 bước Hoàn tác |
| PDF | `Pdf/MainForm.Ai.cs` | Chỉ đọc: văn bản từng trang, vùng chọn; trả lời trong khung |
| Claude Desktop | `Mcp/McpServer.cs`, `Mcp/AiBridge.cs`, `Mcp/ClaudeDesktop.cs` | `MicroadsOffice.exe --mcp` (MCP qua stdio); named pipe chỉ cho người dùng Windows hiện tại; từ chối mọi yêu cầu nếu người dùng chưa bật; ghi chạy trên luồng giao diện; sửa `claude_desktop_config.json` giữ nguyên mục khác + bản sao lưu |
| Cài đặt chung | `UI/Shell/ShellDialogs.cs` | Nút "Trợ lý AI (ChatGPT, Claude)…"; sửa lỗi cũ: ô "Tự cài bản cập nhật…" trước đây không được lưu |
| Quyền riêng tư | `server/office/privacy.php`, `code-signing.php`, `README.md`, `CODE_SIGNING_POLICY.md` | Mục "Trợ lý AI" (vi/en) — phù hợp điều kiện SignPath |

## 5. Kiểm thử (tests/MicroadsOffice.AI.Tests — máy chủ giả trên 127.0.0.1, không cần tài khoản thật)

| Chế độ | Nội dung |
|---|---|
| `core` (24) | DPAPI; PKCE theo RFC 7636; URL đăng nhập; đăng nhập trọn luồng; gia hạn token; đăng nhập lại dùng client id cũ; hủy; state giả mạo; ID token sai chữ ký/nonce/aud/iss; thiếu quyền gói; refresh bị thu hồi; Responses API (trường gửi đi); mã lỗi gói ChatGPT; danh sách mô hình; Claude qua SDK (header, fallbacks, effort); máy chủ tương thích; đăng xuất thu hồi; Markdown; khối hành động; MCP; cầu nối Claude Desktop (từ chối khi chưa bật, đọc/ghi/hoàn tác, chạy trên luồng giao diện); cấu hình Claude Desktop; cài đặt không chứa khóa; chỉ dẫn chống prompt injection |
| `word` (15) | Nút ribbon, thẻ kết nối, mọi nút thiết lập, ghi nội dung có kiểu Tiêu đề + 1 bước Hoàn tác + đọc ngược sạch, Viết lại vùng chọn, nút cần vùng chọn, Chèn/Thay/Sao chép (trả lại clipboard), văn bản hành chính (quốc hiệu căn giữa, bảng), lỗi + Thử lại + mở cài đặt, Dừng, Enter, **mọi nút nhanh**, Cuộc trò chuyện mới, Đóng, bỏ "gửi kèm văn bản"/"tự ghi", hộp thoại cài đặt (máy chủ riêng, khóa, xóa khóa, **Continue with ChatGPT + Đăng xuất qua trình duyệt giả lập**), Claude Desktop (hỏi trước, ngắt), Cài đặt chung |
| `excel` (7) | Nút ribbon + nút nhanh, bảng có giá trị/công thức/định dạng/tự vừa cột/biểu đồ + 1 Hoàn tác, hàm/địa chỉ sai không ghi gì, ngữ cảnh gửi đi (giá trị + công thức + vùng chọn + danh sách hàm), chèn bảng/công thức từ câu trả lời, trang mới, công thức sai khi chèn |
| `pdf` (3) | Nút ribbon, khung chỉ đọc, văn bản theo trang gửi đi, chỉ có Sao chép, nút cần vùng chọn, Đóng |
| `e2e` (4) | **Exe thật**: Claude Desktop (tiến trình `--mcp`) mở Word, ghi + đọc + hoàn tác; ghi Excel có công thức, từ chối công thức sai; bị từ chối sau khi ngắt kết nối |

Chạy: `build.ps1` (core/word/excel/pdf), GitHub Actions (như trên), `.build\scripts\run_ai_ui_when_idle.ps1` (word/excel/pdf/e2e khi máy rảnh ≥ 2 phút).

## 6. Đối chiếu theo vai trò (luồng đã chạy thử)

| Vai trò | Luồng | Kết quả | Bằng chứng |
|---|---|---|---|
| V1 Nhân viên (ChatGPT Plus) | Trợ lý AI › Continue with ChatGPT → trình duyệt → đồng ý → về ứng dụng → hỏi → nội dung ghi vào văn bản → Hoàn tác | Đạt (với máy chủ đăng nhập giả lập đúng giao thức OpenAI) | W13, C04–C12. **Chưa thử với tài khoản ChatGPT thật** — cần anh/chị bấm thử một lần sau khi cài 2.0.3 |
| V2 Kế toán (Claude Pro + Claude Desktop) | Kết nối với Claude Desktop → Claude gọi `MicroadsOffice.exe --mcp` → đọc bảng, ghi công thức → Hoàn tác | Đạt (tiến trình `--mcp` thật đóng vai Claude Desktop) | E01–E04, C20–C22, W14. **Chưa thử trong ứng dụng Claude Desktop thật** |
| V3 Người dùng kỹ thuật (khóa API) | Dán khóa → Kiểm tra kết nối → chọn mô hình → Lưu → dùng | Đạt (máy chủ giả đúng giao thức; khóa Claude gửi qua `Authorization: Bearer` theo API 2026-10-21) | C14, C15, W12 |
| V4 Cơ quan nội bộ (Ollama) | Máy chủ tương thích → địa chỉ nội bộ → danh sách mô hình → dùng | Đạt | C16, W12 |
| V5 Không dùng AI | Không bấm gì → không kết nối mạng; mở khung thì chỉ thấy thẻ hướng dẫn | Đạt | W01–W02, C23 (Provider None không bao giờ "sẵn sàng") |
| V6 Microads (phát hành) | Không giữ khóa của khách; tuân điều khoản OpenAI/Anthropic/SignPath; chính sách mô tả đúng | Đạt | C01, C23; privacy.php mục "Trợ lý AI"; CODE_SIGNING_POLICY.md |

## 7. Việc còn lại (cần người thật)

- Thử "Continue with ChatGPT" bằng **tài khoản ChatGPT Plus/Pro thật** (máy chủ OpenAI thật chỉ chấp nhận người thật đăng nhập).
- Thử **Claude Desktop thật**: bấm Kết nối, khởi động lại Claude Desktop, yêu cầu "Đọc văn bản đang mở trong Microads Word".
- (Tùy chọn) dán một khóa API Claude/OpenAI thật để thử.

## 8. Trạng thái 17/11/2026 07:00 (bàn giao)

- Mã nguồn 2.0.3 hoàn chỉnh (commit 299525a). Kiểm thử lõi 31/31; kiểm thử giao diện chạy lại trên bản cuối — xem `.build\ai-test\report-ai-*.txt`.
- Rà soát độc lập: 11 vấn đề, đã sửa hết, mỗi vấn đề có kiểm thử (C08, C27–C31, W17–W19, X08).
- Chưa làm: build phát hành một lần có `-Store` (MSIX định danh MicroadsTech.MicroadsOffice), thử bản Store trên máy,
  ảnh chụp Store (chế độ `shots-word` / `shots-excel`), nộp hồ sơ theo `docs/MICROSOFT_STORE.md`.
- Cần người dùng xác nhận: kết nối Claude Desktop thật (trạng thái "Đang hoạt động"), đăng nhập ChatGPT thật.

## 9. Rà soát lại luồng đăng nhập ChatGPT / Claude (17/11/2026, không cần người dùng thử)

Cách làm: (a) một tác tử độc lập đối chiếu mã với tài liệu chính thức hiện hành (OpenAI Sign in with ChatGPT, Anthropic, MCP,
Microsoft MSIX); (b) tự đọc lại mã giao diện theo từng vai trò; (c) kiểm tra với **máy chủ thật** ở các bước không cần mật khẩu;
(d) kiểm tra (chỉ đọc) cấu hình Claude Desktop thật trên máy.

**Kiểm tra với máy chủ thật (không cần đăng nhập):**

| Kiểm tra | Kết quả |
|---|---|
| `auth.openai.com/.well-known/openid-configuration` | issuer, authorize, token, jwks, revoke, PKCE S256 **khớp đúng** mã |
| URL đăng nhập do ứng dụng tạo (dynamic_agent_client, 127.0.0.1:1455/auth/callback, scope có `chatgpt.tokens.use.direct`, resource, ext_agent_host_id, agent_name_hint) | OpenAI **chấp nhận** → chuyển tới trang đăng nhập; đối chứng client sai → 400 `invalid_authorize_client` |
| Cổng dự phòng (khi 1455 bận) | OpenAI chấp nhận |
| Anthropic `/v1/models` với khóa giả, header như SDK (Bearer, API 2026-10-21) | 401 `authentication_error` → ứng dụng báo "Khóa API Claude không đúng" |
| Claude Desktop thật (bản Store 5.x) | đọc cấu hình ở `%LOCALAPPDATA%\Claude-Data\claude_pzs8sxrjxfjjc` (đúng thư mục ứng dụng tìm); tệp chưa có mục `microads-office` vì lần mở cuối (06:38) là bản cũ — bản mới tự ghi khi mở Word/Excel/PDF (C26, C30) |

**Lỗi/thiếu sót tìm thấy và đã sửa (mỗi mục có kiểm thử):**

| # | Vấn đề | Sửa | Kiểm thử |
|---|---|---|---|
| R1 | Test W02 bấm "Continue with ChatGPT" trên thẻ thiết lập có thể mở **trình duyệt thật** của người dùng | Bộ test chặn mở trình duyệt ở mọi chế độ | Program.cs |
| R2 | Hộp thoại bị đóng trước sự kiện Shown vẫn bắt đầu đăng nhập ngầm (giữ cổng 1455 tới 10 phút) | Không bắt đầu khi hộp thoại đã đóng | W02, W20 |
| R3 | Đang chờ trình duyệt (tới 10 phút) không có nút hủy | Nút **"Hủy đăng nhập"**; nhả cổng ngay | W20, C36 |
| R4 | Lưu "khóa API Claude/OpenAI" khi chưa nhập khóa → trợ lý im lặng "chưa kết nối" | Cảnh báo, không đóng hộp thoại | W21 |
| R5 | Đã kết nối Claude Desktop nhưng khung trợ lý chỉ ghi "Chưa kết nối AI" | Thẻ và dòng trạng thái nói rõ "hãy hỏi trong Claude Desktop" | W22 |
| R6 | Bản Store: trạng thái nói "Đã ghi cấu hình" trong khi người dùng phải tự dán | Câu hướng dẫn riêng cho bản Store | W22 |
| R7 | Bấm "Continue with ChatGPT" trên thẻ thiết lập nhưng hộp thoại vẫn chọn cách cũ | Tự chọn ChatGPT | — |
| R8 | Lỗi `chatpass_v2_*` (cấu hình quyền) và `subscription_sharing_invalid_user` bắt đăng nhập lại — OpenAI nói không nên | Thông báo đúng, không bắt đăng nhập lại; kèm **mã yêu cầu** để chẩn đoán | C32 |
| R9 | 503 (tạm không khả dụng) không tự thử lại; hết hạn mức vẫn tiếp tục gửi | Thử lại có giới hạn (1 s, 3 s); tạm dừng 1 phút sau `usage_limit_exceeded`; hiện `error.param`; xử lý `route_not_supported` | C32 |
| R10 | Danh sách mô hình gói ChatGPT nhận cả mục không có `visibility` | Chỉ nhận `visibility = "list"` | C32 |
| R11 | Lần đăng ký đầu lỗi ở bước đổi mã → mất client id đã cấp, lần sau đăng ký lại | Giữ client id, lần sau dùng lại | C35 |
| R12 | Claude: `effort` gửi cho mọi mô hình → lỗi 400 khi chọn Sonnet 4.5 / Haiku 4.5; danh sách mô hình chỉ đọc trang đầu | Chỉ gửi effort/fallback cho mô hình hỗ trợ; đọc mọi trang | C33 |
| R13 | Máy chủ MCP chỉ hiểu giao thức cũ (`initialize`); MCP 2026-07-28 bắt buộc `server/discover` | Hỗ trợ cả hai thế hệ: `server/discover`, phiên bản theo từng yêu cầu, lỗi −32022, ghi "bắt tay" cả với client mới | C34 |
| R14 | Sau `await`, mã giao diện có lúc chạy trên luồng nền (khi thiếu SynchronizationContext) → nút "Hủy đăng nhập" treo hộp thoại | `UiThread.On`: phần sau mọi `await` trong khung/hộp thoại AI luôn quay về luồng giao diện qua `BeginInvoke` | W20 (có ghi vết luồng) |
| R15 | (Phát hiện khi xem ảnh Store) Dán/ghi nội dung có **bảng ở giữa**: các đoạn sau bảng bị nhét vào ô đầu của bảng — lỗi chung của Word (cả copy/paste), không riêng Trợ lý | `Editing.PasteBlocks` tiếp tục dán **sau** bảng | Word M08b, AI W23 |
| R16 | (Phát hiện khi xem ảnh Store) Excel "tự vừa cột" đo cột công thức trước khi tính → cột Tổng hiện `####` | Tính lại công thức trước khi đo | AI X09 |
| R17 | Gói MSIX không tạo được: chú thích trong manifest chứa `--` (XML cấm) | Sửa chú thích; gói `MicroadsOffice_2.0.3.0_x64.msix` tạo được, định danh khớp Partner Center | make_msix.ps1 |

Kết quả sau rà soát (17/11/2026 07:26): **lõi 36/36, Word 22/22, Excel 8/8, PDF 3/3, E2E 4/4** — 73 bài AI (sau R15–R16: thêm W23, X09, Word M08b).

**Bản Store chạy thử trên máy (đăng ký tạm từ ổ E, đã gỡ, ổ C trở về đúng dung lượng ban đầu):** PFN `MicroadsTech.MicroadsOffice_p0kmfkpdegte6`; Claude Desktop gọi `MicroadsOffice.exe --mcp` **qua execution alias** được (giao thức cũ và 2026-07-28); Word bản Store mở được; tiến trình `--mcp` thấy Word qua named pipe và từ chối khi chưa cho phép.

Còn lại (ghi nhận, không sửa): chỉ lưu **một** tài khoản ChatGPT (đổi tài khoản = đăng xuất + thu hồi rồi đăng nhập
mới — cố ý, đơn giản cho người dùng văn phòng).
