<?php
// Code signing policy of Microads Office (required by SignPath Foundation: linked from the home and download page).
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
  <p><b>Free code signing provided by <a href="https://about.signpath.io/">SignPath.io</a>, certificate by <a href="https://signpath.org/">SignPath Foundation</a>.</b></p>
  <?php if ($LANG === 'en'): ?>
  <p>Microads Office is free and open source software (<a href="<?= MO_REPO ?>/blob/main/LICENSE">GNU GPL v3</a>), published by Microads Technology Company Limited, Ho Chi Minh City, Vietnam. Source code: <a href="<?= MO_REPO ?>"><?= h(substr(MO_REPO, 8)) ?></a>.</p>
  <h2>What is signed</h2>
  <p>Only binaries built from the public repository by its GitHub Actions workflow are signed: <code>MicroadsOffice_Setup_&lt;version&gt;.exe</code> (the suite) and <code>MicroadsOffice_Setup.exe</code> (the web installer). Every signing request is approved manually.</p>
  <h2>Team roles</h2>
  <ul>
    <li><b>Committers and reviewers</b>: members of the <a href="https://github.com/microads-vn">microads-vn</a> GitHub account (Microads Technology Company Limited).</li>
    <li><b>Approvers</b>: owner of the <a href="https://github.com/microads-vn">microads-vn</a> GitHub account.</li>
  </ul>
  <p>All team members use multi-factor authentication for GitHub and SignPath.</p>
  <h2>Privacy</h2>
  <p>This program will not transfer any information to other networked systems unless specifically requested by the user or the person installing or operating it, except for the automatic update check, which is shown during installation and can be switched off there or in <i>Help › Settings</i>. See the <a href="privacy?lang=en">privacy policy</a>.</p>
  <?php else: ?>
  <p>Microads Office là phần mềm miễn phí, mã nguồn mở (<a href="<?= MO_REPO ?>/blob/main/LICENSE">GNU GPL v3</a>), do Công ty TNHH Công nghệ Microads (TP. Hồ Chí Minh) phát hành. Mã nguồn: <a href="<?= MO_REPO ?>"><?= h(substr(MO_REPO, 8)) ?></a>.</p>
  <h2>Những gì được ký số</h2>
  <p>Chỉ các tệp được GitHub Actions dựng từ kho mã công khai mới được ký: <code>MicroadsOffice_Setup_&lt;phiên bản&gt;.exe</code> (bộ ứng dụng) và <code>MicroadsOffice_Setup.exe</code> (trình cài đặt trực tuyến). Mỗi lần ký đều được người phê duyệt xác nhận thủ công.</p>
  <h2>Vai trò</h2>
  <ul>
    <li><b>Người đóng góp mã và người duyệt mã</b> (committers, reviewers): thành viên tài khoản GitHub <a href="https://github.com/microads-vn">microads-vn</a> của Công ty TNHH Công nghệ Microads.</li>
    <li><b>Người phê duyệt ký số</b> (approvers): chủ tài khoản GitHub <a href="https://github.com/microads-vn">microads-vn</a>.</li>
  </ul>
  <p>Mọi thành viên dùng xác thực nhiều lớp cho GitHub và SignPath.</p>
  <h2>Quyền riêng tư</h2>
  <p>Chương trình không gửi thông tin tới hệ thống mạng khác trừ khi người dùng hoặc người cài đặt yêu cầu, ngoại trừ việc tự động kiểm tra cập nhật: tùy chọn này hiện khi cài đặt và tắt được ngay tại đó hoặc trong <i>Trợ giúp › Cài đặt</i>. Xem <a href="privacy">chính sách quyền riêng tư</a>.</p>
  <?php endif; ?>
</section>
</main>
<footer class="foot wrap">© <?= date('Y') ?> <?= h($c['promo']['company'][$LANG] ?? 'Microads') ?> · <a href="./"><?= t('Tải Microads Office', 'Download Microads Office') ?></a> · <a href="privacy"><?= t('Chính sách quyền riêng tư', 'Privacy policy') ?></a></footer>
</body>
</html>
