<?php
// Privacy policy of Microads Office (required for the Microsoft Store listing).
declare(strict_types=1);
require __DIR__ . '/inc/bootstrap.php';
$LANG = lang();
$c = load_config();
$company = $c['promo']['company'][$LANG] ?? 'Công ty TNHH Công nghệ Microads';
header('X-Content-Type-Options: nosniff');
?><!doctype html>
<html lang="<?= $LANG ?>">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title><?= t('Chính sách quyền riêng tư — Microads Office', 'Privacy policy — Microads Office') ?></title>
<link rel="canonical" href="https://microads.vn/office/privacy">
<link rel="icon" href="assets/office.png">
<link rel="stylesheet" href="assets/style.css">
</head>
<body>
<header class="bar">
  <a class="brand" href="./"><img src="assets/mark.svg" alt="Microads" width="34" height="27"> MICROADS <span>Office</span></a>
  <nav>
    <a href="?lang=vi" class="flag <?= $LANG === 'vi' ? 'on' : '' ?>"><img src="assets/flag-vi.svg" alt="" width="24" height="16"> VI</a>
    <a href="?lang=en" class="flag <?= $LANG === 'en' ? 'on' : '' ?>"><img src="assets/flag-en.svg" alt="" width="24" height="16"> EN</a>
  </nav>
</header>
<main class="wrap">
<section class="card">
<?php if ($LANG === 'en'): ?>
  <h1>Microads Office privacy policy</h1>
  <p>Effective 16 November 2026 (updated for version 2.0.3: AI assistant). Microads Office (Word, Excel, PDF) is published free of charge by <?= h($company) ?> (microads.vn).</p>
  <h2>Your documents stay on your computer</h2>
  <p>Documents, spreadsheets, PDF files, signatures and certificates you open or create are processed only on your device. Microads Office does not upload, read or collect their content. The only exception is the optional AI assistant described below, and only for what you send to the AI service you chose.</p>
  <h2>What is sent over the Internet</h2>
  <ul>
    <li><b>Update check</b>: when an app starts (and every 2 hours while it runs) it requests <code>https://microads.vn/office/api/manifest.json</code> and sends only the installed version number. The server counts checks per day and per version; IP addresses are not stored in these statistics.</li>
    <li><b>Downloads</b>: when you choose to update, the new version is downloaded from microads.vn and verified with Microads' digital signature.</li>
    <li><b>Timestamp / revocation services</b> (PDF signing, optional): only when you enable them, the PDF app contacts the time-stamping authority or the certificate authority you configured.</li>
    <li><b>AI assistant</b> (optional, not connected until you set it up): see the next section.</li>
  </ul>
  <p>No account, no advertising identifiers, no tracking cookies, no analytics of your usage. Settings are stored locally in the app's Data folder.</p>
  <h2 id="ai-assistant">AI assistant (Word, Excel, PDF)</h2>
  <ul>
    <li><b>Off until you choose a service.</b> The assistant does nothing and connects nowhere until you connect it in <b>AI assistant › Settings</b>: your ChatGPT account (OpenAI "Sign in with ChatGPT"), a Claude or OpenAI API key, your own OpenAI-compatible server (for example Ollama on your computer), or Claude Desktop.</li>
    <li><b>Connects only when you ask.</b> A request is sent only when you click <b>Send</b> or a quick action (or, if you connected Claude Desktop, when you ask Claude there and approve its action). It never runs in the background.</li>
    <li><b>What is sent, and to whom.</b> Your question, the conversation so far and — when <b>"Include the open document"</b> is ticked (you can untick it) — the text of the open document or sheet and your selection are sent <b>directly</b> to the service you chose, under that service's own terms and privacy policy (OpenAI, Anthropic, or your server). Nothing goes through Microads servers and Microads never sees it.</li>
    <li><b>Your sign-in and keys stay on your computer.</b> API keys and the ChatGPT sign-in are encrypted with Windows (DPAPI) for your Windows account and stored only in the app's Data folder. Signing out revokes the ChatGPT session; you can also disconnect Microads Office in ChatGPT › Settings, and review ChatGPT plan usage at chatgpt.com/settings/usage.</li>
    <li><b>Claude Desktop.</b> Only if you click <b>Connect to Claude Desktop</b>, Microads Office adds itself to Claude Desktop's configuration so Claude can read and edit the documents open in Microads Office when you ask it; Claude Desktop asks you before each action. <b>Disconnect</b> removes it.</li>
    <li><b>Reporting an answer.</b> Under every answer, <b>Report</b> lets you tell Microads about inappropriate AI content. Only when you click <b>Send report</b>, the reason, your comment and — if you leave the box ticked — that question and the full answer (including what the assistant wrote into the document) are sent to microads.vn so Microads can review them. No API key, account or other part of your document is sent, and no IP address is stored.</li>
  </ul>
  <h2>Turning automatic connections off</h2>
  <p>The update check is the only connection the apps make on their own. It is shown during installation: untick <b>"Check microads.vn for updates automatically"</b> in the setup (or install with <code>--no-auto-update</code>), or later untick it in <b>Help › Settings &amp; language</b>. When it is off, Microads Office never connects to the Internet by itself; <b>Check for updates</b> still works when you click it. The same check also refreshes the short Microads notice shown in the apps.</p>
  <p>The web installer (<code>MicroadsOffice_Setup.exe</code>) and the PowerShell install command contact microads.vn only to download the version you asked to install.</p>
  <h2>Contact</h2>
  <p><?= h($company) ?> — <a href="https://microads.vn">microads.vn</a><?= $c['promo']['phone'] ? ' — ' . h($c['promo']['phone']) : '' ?></p>
<?php else: ?>
  <h1>Chính sách quyền riêng tư của Microads Office</h1>
  <p>Hiệu lực từ 16/11/2026 (cập nhật cho phiên bản 2.0.3: Trợ lý AI). Microads Office (Word, Excel, PDF) được <?= h($company) ?> (microads.vn) phát hành miễn phí.</p>
  <h2>Tài liệu của bạn luôn nằm trên máy của bạn</h2>
  <p>Văn bản, bảng tính, tệp PDF, chữ ký và chứng thư số mà bạn mở hoặc tạo chỉ được xử lý trên máy tính của bạn. Microads Office không tải lên, không đọc và không thu thập nội dung đó. Ngoại lệ duy nhất là Trợ lý AI (tùy chọn) mô tả bên dưới, và chỉ với những gì bạn gửi tới dịch vụ AI do bạn chọn.</p>
  <h2>Những gì được gửi qua Internet</h2>
  <ul>
    <li><b>Kiểm tra cập nhật</b>: khi mở ứng dụng (và mỗi 2 giờ khi đang chạy), ứng dụng truy cập <code>https://microads.vn/office/api/manifest.json</code> và chỉ gửi số phiên bản đang dùng. Máy chủ đếm số lượt kiểm tra theo ngày và theo phiên bản; thống kê này không lưu địa chỉ IP.</li>
    <li><b>Tải bản cập nhật</b>: khi bạn đồng ý cập nhật, bản mới được tải từ microads.vn và được kiểm tra chữ ký số của Microads trước khi cài.</li>
    <li><b>Dấu thời gian / kiểm tra thu hồi</b> (ký số PDF, tùy chọn): chỉ khi bạn bật, ứng dụng PDF mới liên hệ máy chủ cấp dấu thời gian hoặc nhà cung cấp chứng thư số mà bạn đã cấu hình.</li>
    <li><b>Trợ lý AI</b> (tùy chọn, chưa kết nối cho tới khi bạn thiết lập): xem mục tiếp theo.</li>
  </ul>
  <p>Không cần tài khoản, không có mã định danh quảng cáo, không cookie theo dõi, không phân tích hành vi sử dụng. Cài đặt được lưu cục bộ trong thư mục dữ liệu của ứng dụng.</p>
  <h2 id="ai-assistant">Trợ lý AI (Word, Excel, PDF)</h2>
  <ul>
    <li><b>Tắt cho tới khi bạn chọn dịch vụ.</b> Trợ lý không làm gì và không kết nối đi đâu cho tới khi bạn kết nối trong <b>Trợ lý AI › Cài đặt</b>: tài khoản ChatGPT của bạn (chương trình "Sign in with ChatGPT" của OpenAI), khóa API Claude hoặc OpenAI, máy chủ tương thích OpenAI của riêng bạn (ví dụ Ollama trên máy), hoặc Claude Desktop.</li>
    <li><b>Chỉ kết nối khi bạn yêu cầu.</b> Yêu cầu chỉ được gửi khi bạn bấm <b>Gửi</b> hoặc một nút yêu cầu nhanh (hoặc, nếu đã kết nối Claude Desktop, khi bạn ra lệnh cho Claude ở đó và đồng ý thao tác). Trợ lý không bao giờ chạy ngầm.</li>
    <li><b>Gửi gì, gửi cho ai.</b> Câu hỏi, đoạn hội thoại và — khi ô <b>"Gửi kèm nội dung văn bản đang mở"</b> được chọn (bạn có thể bỏ chọn) — nội dung văn bản/trang tính đang mở và vùng chọn được gửi <b>thẳng</b> tới dịch vụ bạn đã chọn, theo điều khoản và chính sách riêng tư của dịch vụ đó (OpenAI, Anthropic hoặc máy chủ của bạn). Không có gì đi qua máy chủ Microads và Microads không xem được.</li>
    <li><b>Đăng nhập và khóa nằm trên máy bạn.</b> Khóa API và phiên đăng nhập ChatGPT được mã hóa bằng Windows (DPAPI) cho tài khoản Windows của bạn và chỉ lưu trong thư mục dữ liệu của ứng dụng. Đăng xuất sẽ thu hồi phiên ChatGPT; bạn cũng có thể ngắt kết nối Microads Office trong ChatGPT › Settings và xem mức dùng gói ChatGPT tại chatgpt.com/settings/usage.</li>
    <li><b>Claude Desktop.</b> Chỉ khi bạn bấm <b>Kết nối với Claude Desktop</b>, Microads Office mới tự thêm vào cấu hình của Claude Desktop để Claude đọc và chỉnh văn bản đang mở trong Microads Office khi bạn yêu cầu; Claude Desktop hỏi bạn trước mỗi thao tác. Bấm <b>Ngắt kết nối</b> để gỡ.</li>
    <li><b>Báo cáo câu trả lời.</b> Dưới mỗi câu trả lời có nút <b>Báo cáo</b> để báo cho Microads nội dung AI không phù hợp. Chỉ khi bạn bấm <b>Gửi báo cáo</b>, lý do, ghi chú của bạn và — nếu bạn để nguyên ô chọn — câu hỏi và toàn bộ câu trả lời đó (kể cả nội dung trợ lý đã soạn vào văn bản) mới được gửi tới microads.vn để Microads xem xét. Không gửi khóa API, tài khoản hay phần nào khác của văn bản; không lưu địa chỉ IP.</li>
  </ul>
  <h2>Tắt kết nối tự động</h2>
  <p>Kiểm tra cập nhật là kết nối duy nhất ứng dụng tự thực hiện. Tùy chọn này hiện ngay khi cài đặt: bỏ chọn <b>"Tự động kiểm tra bản cập nhật từ microads.vn"</b> trong trình cài (hoặc cài với <code>--no-auto-update</code>), hay bỏ chọn sau này trong <b>Trợ giúp › Cài đặt &amp; ngôn ngữ</b>. Khi tắt, Microads Office không bao giờ tự kết nối Internet; nút <b>Kiểm tra cập nhật</b> vẫn hoạt động khi bạn bấm. Lần kiểm tra này cũng làm mới thông báo ngắn của Microads hiển thị trong ứng dụng.</p>
  <p>Trình cài đặt trực tuyến (<code>MicroadsOffice_Setup.exe</code>) và lệnh cài bằng PowerShell chỉ kết nối microads.vn để tải phiên bản bạn yêu cầu cài.</p>
  <h2>Liên hệ</h2>
  <p><?= h($company) ?> — <a href="https://microads.vn">microads.vn</a><?= $c['promo']['phone'] ? ' — ' . h($c['promo']['phone']) : '' ?></p>
<?php endif; ?>
</section>
</main>
<footer class="foot wrap">© <?= date('Y') ?> <?= h($company) ?> · <a href="./"><?= t('Tải Microads Office', 'Download Microads Office') ?></a></footer>
</body>
</html>
