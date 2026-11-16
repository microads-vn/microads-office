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
  <p>Effective 16 November 2026. Microads Office (Word, Excel, PDF) is published free of charge by <?= h($company) ?> (microads.vn).</p>
  <h2>Your documents stay on your computer</h2>
  <p>Documents, spreadsheets, PDF files, signatures and certificates you open or create are processed only on your device. Microads Office does not upload, read or collect their content.</p>
  <h2>What is sent over the Internet</h2>
  <ul>
    <li><b>Update check</b>: when an app starts (and every 2 hours while it runs) it requests <code>https://microads.vn/office/api/manifest.json</code> and sends only the installed version number. The server counts checks per day and per version; IP addresses are not stored in these statistics.</li>
    <li><b>Downloads</b>: when you choose to update, the new version is downloaded from microads.vn and verified with Microads' digital signature.</li>
    <li><b>Timestamp / revocation services</b> (PDF signing, optional): only when you enable them, the PDF app contacts the time-stamping authority or the certificate authority you configured.</li>
  </ul>
  <p>No account, no advertising identifiers, no tracking cookies, no analytics of your usage. Settings are stored locally in the app's Data folder.</p>
  <h2>Turning automatic connections off</h2>
  <p>The update check is the only connection the apps make on their own. It is shown during installation: untick <b>"Check microads.vn for updates automatically"</b> in the setup (or install with <code>--no-auto-update</code>), or later untick it in <b>Help › Settings &amp; language</b>. When it is off, Microads Office never connects to the Internet by itself; <b>Check for updates</b> still works when you click it. The same check also refreshes the short Microads notice shown in the apps.</p>
  <p>The web installer (<code>MicroadsOffice_Setup.exe</code>) and the PowerShell install command contact microads.vn only to download the version you asked to install.</p>
  <h2>Contact</h2>
  <p><?= h($company) ?> — <a href="https://microads.vn">microads.vn</a><?= $c['promo']['phone'] ? ' — ' . h($c['promo']['phone']) : '' ?></p>
<?php else: ?>
  <h1>Chính sách quyền riêng tư của Microads Office</h1>
  <p>Hiệu lực từ 16/11/2026. Microads Office (Word, Excel, PDF) được <?= h($company) ?> (microads.vn) phát hành miễn phí.</p>
  <h2>Tài liệu của bạn luôn nằm trên máy của bạn</h2>
  <p>Văn bản, bảng tính, tệp PDF, chữ ký và chứng thư số mà bạn mở hoặc tạo chỉ được xử lý trên máy tính của bạn. Microads Office không tải lên, không đọc và không thu thập nội dung đó.</p>
  <h2>Những gì được gửi qua Internet</h2>
  <ul>
    <li><b>Kiểm tra cập nhật</b>: khi mở ứng dụng (và mỗi 2 giờ khi đang chạy), ứng dụng truy cập <code>https://microads.vn/office/api/manifest.json</code> và chỉ gửi số phiên bản đang dùng. Máy chủ đếm số lượt kiểm tra theo ngày và theo phiên bản; thống kê này không lưu địa chỉ IP.</li>
    <li><b>Tải bản cập nhật</b>: khi bạn đồng ý cập nhật, bản mới được tải từ microads.vn và được kiểm tra chữ ký số của Microads trước khi cài.</li>
    <li><b>Dấu thời gian / kiểm tra thu hồi</b> (ký số PDF, tùy chọn): chỉ khi bạn bật, ứng dụng PDF mới liên hệ máy chủ cấp dấu thời gian hoặc nhà cung cấp chứng thư số mà bạn đã cấu hình.</li>
  </ul>
  <p>Không cần tài khoản, không có mã định danh quảng cáo, không cookie theo dõi, không phân tích hành vi sử dụng. Cài đặt được lưu cục bộ trong thư mục dữ liệu của ứng dụng.</p>
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
