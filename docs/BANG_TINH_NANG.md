# Bảng tính năng Microads Office Word & Microads Office Excel

Đối chiếu với Microsoft Word / Excel (bản Microsoft 365, giao diện Ribbon). Mỗi dòng là một tính năng
của Microsoft Office; cột **Microads** cho biết phạm vi triển khai:

- ✅ **Đầy đủ**: dùng được như Office cho nhu cầu văn phòng thông thường.
- ◐ **Cơ bản**: có tính năng, ít tùy chọn hơn Office.
- ✗ **Không làm**: kèm lý do (thường là tính năng đám mây/cộng tác trả phí hoặc ít dùng ở văn phòng VN).

Bảng này là "hợp đồng" kiểm thử: mỗi dòng ✅/◐ phải có nút bấm được và ca kiểm thử tương ứng
(xem `docs/DOI_CHIEU_YEU_CAU.md`).

---

## A. Dùng chung cho cả bộ (Word · Excel · PDF)

| Nhóm | Tính năng | Microads | Ghi chú |
|---|---|---|---|
| Cài đặt | 1 tệp `.exe` duy nhất, cài là dùng, không cần thư viện ngoài | ✅ | .NET tự chứa; PDFium nhúng trong exe |
| Cài đặt | Trình cài đặt: chọn thư mục, lối tắt Desktop/Start, liên kết tệp, gỡ cài đặt | ✅ | Cài theo người dùng, không cần quyền Admin |
| Cài đặt | Trình cài đặt trực tuyến 80 KB không đổi giữa các bản (danh tiếng Chrome/SmartScreen cộng dồn), kiểm chữ ký Microads trước khi chạy; nút Thử lại / Tải bộ cài đầy đủ / Hủy / English | ✅ | Thay cho việc mua chứng thư ký mã |
| Cài đặt | Cài không qua trình duyệt: lệnh PowerShell, winget, cài im lặng `--silent` cho IT | ✅ | winget: sau khi Microsoft duyệt hồ sơ |
| Ngôn ngữ | Tiếng Việt / English, chọn ngay khi vừa cài; đổi sau bằng lá cờ 🇻🇳 / 🇬🇧 | ✅ | Áp dụng cho cả 3 ứng dụng |
| Cập nhật | Nút **Kiểm tra cập nhật** trong mọi ứng dụng | ✅ | Máy chủ `microads.vn/office` |
| Cập nhật | Cập nhật toàn bộ theo lệnh từ máy chủ (tự động tải & cài) | ✅ | Có chữ ký số ECDSA chống giả mạo bản cập nhật |
| Cập nhật | Thông tin quảng bá / liên kết đổi từ máy chủ không cần phát hành lại | ✅ | |
| Thương hiệu | Logo Microads Office, "Phát hành bởi Công ty TNHH Công nghệ Microads — Luôn luôn miễn phí", liên kết **microads.vn** | ✅ | Giới thiệu, trang bắt đầu, thông báo khi mở tệp |
| Quảng bá | Thông báo khi mở tệp từ thư mục, có nút tắt ngay | ✅ | Không bắt chờ giây nào |
| Giao diện | Trang bắt đầu kiểu Office (Mới · Mở · Gần đây · Mẫu) | ✅ | |
| Giao diện | Ribbon theo nhóm (nút lớn/nhỏ, ô chọn phông, bảng màu), tự thu gọn khi cửa sổ hẹp, thu gọn Ribbon (Ctrl+F1) | ✅ | |
| Giao diện | Thanh truy cập nhanh: Lưu · Hoàn tác · Làm lại | ✅ | |
| Tệp | Mới, Mở, Lưu, Lưu thành, Gần đây, Đóng, Thuộc tính tài liệu | ✅ | |
| Tệp | Tự lưu bản phục hồi (AutoRecover) và khôi phục khi ứng dụng bị tắt đột ngột | ✅ | |
| Tệp | Xuất PDF | ✅ | |
| Tệp | In: chọn máy in, số bản, phạm vi trang, xem trước | ✅ | |
| Tệp | Lưu lên OneDrive/SharePoint, đồng tác giả thời gian thực | ✗ | Dịch vụ đám mây của Microsoft |
| Khác | Copilot / AI, Macro VBA, Add-in Store | ✗ | Không thuộc phạm vi phần mềm miễn phí ngoại tuyến |

## B. Microads Office Word

### B1. Tệp & định dạng

| Tính năng Word | Microads | Ghi chú |
|---|---|---|
| Mở/Lưu `.docx` (định dạng Office Open XML) | ✅ | Đoạn, run, kiểu, bảng, ảnh, danh sách, đầu/chân trang, ngắt trang, khổ giấy, lề |
| Mở `.rtf`, `.txt`; Lưu `.txt`, `.html`, `.rtf` | ✅ | |
| Mở `.doc` (Word 97-2003 nhị phân) | ✗ | Báo hướng dẫn lưu lại thành .docx |
| Mở `.odt` | ✗ | |
| Xuất PDF giữ nguyên chữ (tìm/chọn được trong PDF) | ✅ | |
| Mẫu văn bản: trống, công văn hành chính (NĐ 30/2020), thư, sơ yếu lý lịch, báo cáo | ✅ | |

### B2. Thẻ Trang chủ (Home)

| Nhóm | Tính năng | Microads |
|---|---|---|
| Bộ nhớ tạm | Cắt, Sao chép, Dán, Dán chỉ văn bản, Sao định dạng (Format Painter) | ✅ |
| Phông chữ | Tên phông, cỡ chữ, tăng/giảm cỡ | ✅ |
| Phông chữ | Đậm, Nghiêng, Gạch chân (đơn/đôi/chấm), Gạch ngang, Chỉ số dưới/trên | ✅ |
| Phông chữ | Đổi kiểu chữ HOA/thường/Hoa Đầu Từ | ✅ |
| Phông chữ | Màu chữ, Tô sáng (highlight), Xóa định dạng | ✅ |
| Phông chữ | Hộp thoại Phông chữ (giãn ký tự, hiệu ứng nâng cao) | ◐ |
| Đoạn văn | Dấu đầu dòng (•, ◦, ▪, –, ✓), Đánh số (1. a. i. A. I.), danh sách nhiều cấp | ✅ |
| Đoạn văn | Tăng/giảm thụt lề, thụt dòng đầu, thụt treo | ✅ |
| Đoạn văn | Căn trái/giữa/phải/đều hai bên | ✅ |
| Đoạn văn | Giãn dòng (1.0, 1.15, 1.5, 2.0, chính xác), khoảng cách trước/sau đoạn | ✅ |
| Đoạn văn | Tô nền đoạn, Viền đoạn | ✅ |
| Đoạn văn | Hiện/ẩn ký tự định dạng ¶ | ✅ |
| Đoạn văn | Sắp xếp đoạn A→Z | ✅ |
| Kiểu (Styles) | Normal, Tiêu đề 1–3, Title, Subtitle, Trích dẫn, Không giãn cách | ✅ |
| Kiểu | Tạo/sửa kiểu tùy chỉnh | ◐ |
| Chỉnh sửa | Tìm, Thay thế (phân biệt hoa thường, nguyên từ), Đi tới trang, Chọn tất cả | ✅ |

### B3. Thẻ Chèn (Insert)

| Tính năng | Microads | Ghi chú |
|---|---|---|
| Ngắt trang | ✅ | Ctrl+Enter |
| Bảng: chọn nhanh lưới, hộp thoại số hàng/cột | ✅ | |
| Ảnh từ tệp (PNG/JPG/BMP/GIF), dán ảnh từ clipboard, đổi kích thước bằng tay nắm | ✅ | Ảnh nằm cùng dòng (inline) |
| Ảnh trôi/bao quanh chữ (wrap square/tight) | ✗ | Chỉ hỗ trợ ảnh cùng dòng |
| Hình vẽ (Shapes), SmartArt, Biểu đồ trong Word, 3D model | ✗ | Dùng Excel cho biểu đồ, chèn dạng ảnh |
| Siêu liên kết (Ctrl+K), mở bằng Ctrl+Click | ✅ | |
| Đầu trang, Chân trang, Số trang (Trang X / Y) | ✅ | |
| Ngày giờ, Ký hiệu đặc biệt (Ω, ©, ™, →, ✓, ký tự tiếng Việt) | ✅ | |
| Đường kẻ ngang | ✅ | |
| Hộp văn bản (Text box), WordArt, Phương trình | ✗ | |

### B4. Thẻ Bố cục (Layout) & Thiết kế

| Tính năng | Microads |
|---|---|
| Lề: Thường, Hẹp, Vừa, Rộng, Văn bản hành chính VN (trên 20, dưới 20, trái 30, phải 15 mm), Tùy chỉnh | ✅ |
| Hướng giấy: Dọc / Ngang | ✅ |
| Khổ giấy: A4, A5, A3, Letter, Legal, B5 | ✅ |
| Ngắt trang / ngắt dòng | ✅ |
| Thụt lề & giãn cách đoạn trên Ribbon | ✅ |
| Màu nền trang, Hình mờ (Watermark) chữ: BẢN NHÁP, MẬT, MẪU… | ✅ |
| Chia cột báo chí (Columns), Đánh số dòng, Ngắt đoạn (section) nhiều khổ giấy | ✗ |

### B5. Thẻ Tham chiếu / Xem lại / Xem

| Tính năng | Microads | Ghi chú |
|---|---|---|
| Mục lục tự động từ Tiêu đề 1–3, cập nhật mục lục | ✅ | |
| Chú thích cuối trang (footnote), trích dẫn, chỉ mục | ✗ | |
| Đếm từ (từ, ký tự, đoạn, trang) | ✅ | Hiện cả ở thanh trạng thái |
| Kiểm tra chính tả (gạch đỏ, gợi ý) | ◐ | Dùng bộ kiểm tra chính tả của Windows theo ngôn ngữ đã cài |
| Ghi chú (Comment) | ✗ | |
| Theo dõi thay đổi (Track Changes), So sánh tài liệu | ✗ | |
| Bố cục In (Print Layout) theo trang thật, thước kẻ (Ruler) | ✅ | |
| Thu phóng: thanh trượt, 100%, Một trang, Vừa chiều rộng | ✅ | |
| Ngăn điều hướng theo tiêu đề | ✅ | |
| Toàn màn hình / chế độ đọc | ✅ | |

### B6. Công cụ bảng & ảnh

| Tính năng | Microads |
|---|---|
| Chèn hàng trên/dưới, cột trái/phải; xóa hàng/cột/bảng | ✅ |
| Gộp ô, Tô nền ô, Viền bảng (tất cả/không viền) | ✅ |
| Tab trong ô chuyển ô, Tab ở ô cuối thêm hàng | ✅ |
| Đổi độ rộng cột bằng kéo | ◐ |
| Ảnh: đổi kích thước giữ tỉ lệ, đặt lại kích thước gốc | ✅ |

### B7. Soạn thảo & bàn phím

| Tính năng | Microads |
|---|---|
| Gõ tiếng Việt (Unikey, EVKey, bàn phím Telex của Windows) | ✅ |
| Hoàn tác / Làm lại nhiều cấp (Ctrl+Z / Ctrl+Y) | ✅ |
| Chọn bằng chuột: kéo, bấm đúp chọn từ, bấm ba lần chọn đoạn; Shift+mũi tên, Ctrl+mũi tên | ✅ |
| Phím tắt chuẩn Word: Ctrl+B/I/U, Ctrl+L/E/R/J, Ctrl+]/[, Ctrl+Shift+=, Ctrl+=, Ctrl+F/H/G, Ctrl+P, Ctrl+S | ✅ |
| Kéo-thả văn bản đã chọn | ✗ |

## C. Microads Office Excel

### C1. Tệp

| Tính năng | Microads | Ghi chú |
|---|---|---|
| Mở/Lưu `.xlsx` | ✅ | Giá trị, công thức, kiểu ô, gộp ô, độ rộng cột/cao hàng, cố định ô, nhiều trang tính, tên định nghĩa, định dạng có điều kiện, xác thực dữ liệu, ghi chú, siêu liên kết, biểu đồ cơ bản |
| Mở/Lưu `.csv`, `.txt` (tab); chọn mã hóa UTF-8 | ✅ | |
| Mở `.xls` (Excel 97-2003 nhị phân), `.ods` | ✗ | |
| Xuất PDF, lưu HTML | ✅ | |
| Mẫu: Trống, Ngân sách cá nhân, Hóa đơn/Báo giá, Bảng chấm công, Danh sách | ✅ | |

### C2. Lưới & thao tác ô

| Tính năng | Microads |
|---|---|
| 1.048.576 hàng × 16.384 cột, cuộn mượt, thu phóng 10–400% | ✅ |
| Thanh công thức + Hộp tên (Name Box: gõ A1/B2:D9/tên để nhảy tới) | ✅ |
| Nhập/sửa trong ô (F2, gõ trực tiếp, Alt+Enter xuống dòng trong ô) | ✅ |
| Chọn vùng bằng chuột/Shift, nhiều vùng bằng Ctrl, chọn cả hàng/cột/bảng | ✅ |
| Kéo tay nắm điền (Fill handle): chép, chuỗi số, ngày, thứ, tháng, công thức tự đổi tham chiếu | ✅ |
| Bấm chọn ô khi đang gõ công thức để chèn tham chiếu, tô màu tham chiếu | ✅ |
| Gợi ý tên hàm khi gõ công thức | ✅ |
| Cắt/Chép/Dán, Dán đặc biệt (giá trị, công thức, định dạng, chuyển vị) | ✅ |
| Chèn/Xóa hàng, cột, ô; Ẩn/Hiện hàng/cột; kéo đổi độ rộng; bấm đúp tự vừa | ✅ |
| Trang tính: thêm, đổi tên, xóa, nhân bản, di chuyển, màu thẻ | ✅ |
| Hoàn tác / Làm lại nhiều cấp | ✅ |

### C3. Thẻ Trang chủ (Home)

| Nhóm | Tính năng | Microads |
|---|---|---|
| Phông chữ | Phông, cỡ, Đậm/Nghiêng/Gạch chân/Gạch ngang, màu chữ, màu nền ô, Viền (dưới, trên, trái, phải, tất cả, ngoài, dày, không) | ✅ |
| Căn lề | Trên/Giữa/Dưới, Trái/Giữa/Phải, Xuống dòng tự động, Gộp & căn giữa / Gộp ngang / Bỏ gộp, Thụt lề | ✅ |
| Căn lề | Xoay chữ (orientation) | ✗ |
| Số | Chung, Số, Tiền tệ (₫, $), Kế toán, Ngày ngắn/dài, Giờ, Phần trăm, Phân số, Khoa học, Văn bản, Tùy chỉnh mã định dạng; Tăng/giảm chữ số thập phân; Dấu phân cách hàng nghìn | ✅ |
| Kiểu | Định dạng có điều kiện: lớn hơn, nhỏ hơn, giữa, bằng, chứa văn bản, trùng lặp, Top 10, thanh dữ liệu, thang màu; Xóa quy tắc | ✅ |
| Kiểu | Định dạng như Bảng (bảng có sọc), Kiểu ô (Tốt/Xấu/Trung tính/Tiêu đề/Tổng) | ✅ |
| Ô | Chèn/Xóa, Định dạng: cao hàng, rộng cột, tự vừa, ẩn/hiện, đổi tên trang tính, Định dạng ô (hộp thoại nhiều thẻ) | ✅ |
| Chỉnh sửa | Tự động Tính (Tổng, Trung bình, Đếm, Lớn nhất, Nhỏ nhất), Điền xuống/phải/chuỗi, Xóa (tất cả/định dạng/nội dung), Sắp xếp & Lọc, Tìm & Thay thế, Đi tới | ✅ |

### C4. Thẻ Chèn / Bố trí trang / Công thức / Dữ liệu / Xem lại / Xem

| Thẻ | Tính năng | Microads | Ghi chú |
|---|---|---|---|
| Chèn | Biểu đồ: Cột, Thanh, Đường, Tròn, Vành khuyên, Vùng, Phân tán; tiêu đề, chú giải, đổi loại, kéo/đổi kích thước | ✅ | |
| Chèn | Ảnh, Siêu liên kết, Ghi chú (Note), Ký hiệu | ✅ | |
| Chèn | PivotTable, Sparkline, Slicer, Bản đồ 3D | ✗ | PivotTable: dự kiến phiên bản sau |
| Bố trí trang | Lề, Hướng, Khổ giấy, Vùng in, In tiêu đề hàng lặp lại, Co vừa trang (rộng 1 trang), In đường lưới/tiêu đề | ✅ | |
| Công thức | Chèn hàm (hộp thoại theo nhóm, mô tả tham số), thư viện hàm theo nhóm | ✅ | |
| Công thức | Trình quản lý Tên (Name Manager), Hiện công thức, Tính tự động/thủ công, Tính ngay (F9) | ✅ | |
| Công thức | Truy vết phụ thuộc (mũi tên), Watch Window, Evaluate Formula | ✗ | |
| Dữ liệu | Sắp xếp A→Z/Z→A, Sắp xếp tùy chỉnh nhiều cấp | ✅ | |
| Dữ liệu | Bộ lọc tự động (danh sách giá trị có ô chọn, lọc số/văn bản), Xóa lọc | ✅ | |
| Dữ liệu | Loại bỏ trùng lặp, Văn bản thành cột, Xác thực dữ liệu (danh sách thả xuống, số, ngày, độ dài) | ✅ | |
| Dữ liệu | Power Query, kết nối CSDL, What-If, Solver | ✗ | |
| Xem lại | Ghi chú ô (tam giác đỏ, di chuột xem), Bảo vệ trang tính (khóa sửa, mật khẩu) | ✅ | |
| Xem | Cố định ô (hàng đầu, cột đầu, tại ô chọn), Đường lưới, Tiêu đề hàng/cột, Thanh công thức, Thu phóng, Vừa vùng chọn | ✅ | |
| Thanh trạng thái | Tổng / Trung bình / Đếm vùng chọn, thanh trượt thu phóng | ✅ | |

### C5. Bộ máy công thức

| Hạng mục | Microads |
|---|---|
| Toán tử `+ - * / ^ & % = <> < > <= >=`, tham chiếu tương đối/tuyệt đối `$A$1`, vùng `A1:B9`, cả cột `A:A`, trang khác `'Trang 1'!A1`, tên định nghĩa, hằng mảng `{1,2,3}` | ✅ |
| Mã lỗi `#DIV/0! #N/A #NAME? #REF! #VALUE! #NUM! #NULL! #CALC! #SPILL!`, phát hiện tham chiếu vòng | ✅ |
| **502 hàm** — đủ 12 nhóm của Excel 365 (trừ nhóm Cube/OLAP và các hàm cần dịch vụ trực tuyến: WEBSERVICE, STOCKHISTORY, IMAGE, TRANSLATE, COPILOT, PY, FORECAST.ETS…) | ✅ |
| Tự tính lại khi ô phụ thuộc thay đổi; chèn/xóa hàng cột tự sửa tham chiếu | ✅ |
| Hàm Toán học (SUM, SUMIF(S), SUMPRODUCT, ROUND*, INT, MOD, ABS, SQRT, POWER, EXP, LN, LOG, PI, RAND, RANDBETWEEN, CEILING, FLOOR, TRUNC, PRODUCT, FACT, COMBIN, GCD, LCM, MROUND, QUOTIENT, SUBTOTAL, lượng giác…) | ✅ |
| Hàm Thống kê (AVERAGE(IF/IFS), COUNT, COUNTA, COUNTBLANK, COUNTIF(S), MAX/MIN(IFS), MEDIAN, MODE, STDEV, VAR, LARGE, SMALL, RANK, PERCENTILE, QUARTILE, CORREL…) | ✅ |
| Hàm Logic (IF, IFS, IFERROR, IFNA, AND, OR, NOT, XOR, SWITCH, CHOOSE) | ✅ |
| Hàm Tra cứu (VLOOKUP, HLOOKUP, LOOKUP, XLOOKUP, INDEX, MATCH, XMATCH, OFFSET, INDIRECT, ROW(S), COLUMN(S), ADDRESS) | ✅ |
| Hàm Văn bản (LEFT, RIGHT, MID, LEN, LOWER, UPPER, PROPER, TRIM, CONCAT, CONCATENATE, TEXTJOIN, SUBSTITUTE, REPLACE, FIND, SEARCH, TEXT, VALUE, REPT, EXACT, CHAR, CODE, CLEAN, FIXED…) | ✅ |
| Hàm Ngày giờ (DATE, TIME, TODAY, NOW, YEAR, MONTH, DAY, HOUR, MINUTE, SECOND, WEEKDAY, WEEKNUM, EDATE, EOMONTH, DATEDIF, DAYS, NETWORKDAYS, WORKDAY, DATEVALUE, YEARFRAC…) | ✅ |
| Hàm Tài chính (PMT, FV, PV, NPV, IRR, XIRR, XNPV, MIRR, RATE, NPER, IPMT, PPMT, CUMIPMT, CUMPRINC, SLN, SYD, DB, DDB, VDB, EFFECT, NOMINAL, trái phiếu PRICE/YIELD/DURATION/ACCRINT/COUP*/ODDF*/ODDL*, TBILL*, AMORLINC…) | ✅ |
| Hàm Kỹ thuật (BIN2DEC, DEC2HEX…, BITAND…, CONVERT, ERF, BESSEL*, số phức COMPLEX/IM*) | ✅ |
| Hàm Cơ sở dữ liệu (DSUM, DAVERAGE, DCOUNT, DCOUNTA, DGET, DMAX, DMIN, DPRODUCT, DSTDEV(P), DVAR(P)) | ✅ |
| Hàm Thống kê nâng cao (NORM/T/CHISQ/F/BINOM/POISSON/GAMMA/BETA/LOGNORM/WEIBULL… .DIST/.INV, T.TEST, Z.TEST, CONFIDENCE, LINEST, LOGEST, TREND, GROWTH, FREQUENCY, SKEW, KURT, COVARIANCE…) và các tên cũ (NORMDIST, TTEST…) | ✅ |
| Hàm Web (ENCODEURL, FILTERXML) | ✅ |
| Hàm Thông tin (ISBLANK, ISNUMBER, ISTEXT, ISERROR, ISNA, ISEVEN, ISODD, NA, TYPE, N, T) | ✅ |
| Hàm tiếng Việt: `DOCSO(số)` đọc số thành chữ ("Một trăm hai mươi nghìn đồng") | ✅ |
| Mảng động tràn ô (FILTER, SORT, SORTBY, UNIQUE, SEQUENCE, VSTACK/HSTACK, TAKE/DROP, TOCOL/TOROW, WRAPROWS…, #SPILL!, tham chiếu `A1#`), tính theo từng phần tử (`LEN(A1:A9)`) | ✅ |
| LET, LAMBDA (cả LAMBDA đặt tên, đệ quy, tham số tùy chọn `[x]`), MAP, REDUCE, SCAN, BYROW, BYCOL, MAKEARRAY, GROUPBY, PIVOTBY | ✅ |
| Lưu .xlsx đúng chuẩn Excel 365 (tiền tố `_xlfn.`, mảng động `cm`/metadata) | ✅ |
