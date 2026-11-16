<?php
// Code signing policy of Microads Office (how releases are signed and verified; linked from the home and download page).
declare(strict_types=1);
require __DIR__ . '/inc/bootstrap.php';
$LANG = lang();
$c = load_config();
const MO_REPO = 'https://github.com/microads-vn/microads-office';
header('X-Content-Type-Options: nosniff');
?><!doctype html>
<html lang="<?= $LANG ?>">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Code signing policy — Microads Office</title>
<link rel="canonical" href="https://microads.vn/office/code-signing">
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
<section class="card" id="code-signing-policy">
  <h1>Code signing policy</h1>
  <p><b><?= t('Mỗi bản phát hành được ký bằng khóa phát hành ECDSA P-256 của Microads; ứng dụng, trình cài đặt trực tuyến và máy chủ cập nhật đều kiểm tra chữ ký này trước khi cài. Bản trên Microsoft Store do Microsoft ký số.', 'Every release is signed with the Microads ECDSA P-256 release key; the apps, the web installer and the update server verify that signature before installing anything. The Microsoft Store edition is signed by Microsoft.') ?></b></p>
  <?php if ($LANG === 'en'): ?>
  <p>Microads Office is free and open source software (<a href="<?= MO_REPO ?>/blob/main/LICENSE">GNU GPL v3</a>), published by Microads Technology Company Limited, Ho Chi Minh City, Vietnam. Source code: <a href="<?= MO_REPO ?>"><?= h(substr(MO_REPO, 8)) ?></a>.</p>
  <h2>What is signed</h2>
  <p>Only binaries built from the public source code (same version tag in the repository) are signed and published: <code>MicroadsOffice_Setup_&lt;version&gt;.exe</code> (the suite) and <code>MicroadsOffice_Setup.exe</code> (the web installer). Each release is approved by the Microads release owner before it is signed and published.</p>
  <h2>Team roles</h2>
  <ul>
    <li><b>Committers and reviewers</b>: members of the <a href="https://github.com/microads-vn">microads-vn</a> GitHub account (Microads Technology Company Limited).</li>
    <li><b>Approvers</b>: owner of the <a href="https://github.com/microads-vn">microads-vn</a> GitHub account.</li>
  </ul>
  <p>All team members use multi-factor authentication for GitHub and Microsoft Partner Center.</p>
  <h2>Privacy</h2>
  <p>This program will not transfer any information to other networked systems unless specifically requested by the user or the person installing or operating it, except for the automatic update check, which is shown during installation and can be switched off there or in <i>Help › Settings</i>. The optional AI assistant sends requests only when the user clicks <i>Send</i>, directly to the AI service the user chose and connected (never through Microads servers). See the <a href="privacy?lang=en#ai-assistant">privacy policy</a>.</p>
  <?php else: ?>
  <p>Microads Office là phần mềm miễn phí, mã nguồn mở (<a href="<?= MO_REPO ?>/blob/main/LICENSE">GNU GPL v3</a>), do Công ty TNHH Công nghệ Microads (TP. Hồ Chí Minh) phát hành. Mã nguồn: <a href="<?= MO_REPO ?>"><?= h(substr(MO_REPO, 8)) ?></a>.</p>
  <h2>Những gì được ký số</h2>
  <p>Chỉ các tệp được dựng từ mã nguồn công khai (đúng phiên bản trong kho mã) mới được ký và phát hành: <code>MicroadsOffice_Setup_&lt;phiên bản&gt;.exe</code> (bộ ứng dụng) và <code>MicroadsOffice_Setup.exe</code> (trình cài đặt trực tuyến). Mỗi bản phát hành đều được người phụ trách phát hành của Microads duyệt trước khi ký và đăng.</p>
  <h2>Vai trò</h2>
  <ul>
    <li><b>Người đóng góp mã và người duyệt mã</b> (committers, reviewers): thành viên tài khoản GitHub <a href="https://github.com/microads-vn">microads-vn</a> của Công ty TNHH Công nghệ Microads.</li>
    <li><b>Người phê duyệt ký số</b> (approvers): chủ tài khoản GitHub <a href="https://github.com/microads-vn">microads-vn</a>.</li>
  </ul>
  <p>Mọi thành viên dùng xác thực nhiều lớp cho GitHub và Microsoft Partner Center.</p>
  <h2>Quyền riêng tư</h2>
  <p>Chương trình không gửi thông tin tới hệ thống mạng khác trừ khi người dùng hoặc người cài đặt yêu cầu, ngoại trừ việc tự động kiểm tra cập nhật: tùy chọn này hiện khi cài đặt và tắt được ngay tại đó hoặc trong <i>Trợ giúp › Cài đặt</i>. Trợ lý AI (tùy chọn) chỉ gửi yêu cầu khi người dùng bấm <i>Gửi</i>, thẳng tới dịch vụ AI người dùng đã chọn và kết nối (không qua máy chủ Microads). Xem <a href="privacy#ai-assistant">chính sách quyền riêng tư</a>.</p>
  <?php endif; ?>
</section>
</main>
<footer class="foot wrap">© <?= date('Y') ?> <?= h($c['promo']['company'][$LANG] ?? 'Microads') ?> · <a href="./"><?= t('Tải Microads Office', 'Download Microads Office') ?></a> · <a href="privacy"><?= t('Chính sách quyền riêng tư', 'Privacy policy') ?></a></footer>
</body>
</html>
