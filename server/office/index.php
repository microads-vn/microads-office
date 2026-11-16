<?php
// Public page: microads.vn/office — latest version, download, release notes.
declare(strict_types=1);
require __DIR__ . '/inc/bootstrap.php';
$LANG = lang();
$c = load_config();
$p = $c['promo'];
$latest = release_by_version($c, $c['latest']);
$store = store_links($c);
$web = websetup_info($c);           // web installer: small, never changes, always installs the latest signed release
$winget = winget_id($c);
$L = fn(array $x) => $x[$LANG] ?? ($x['vi'] ?? '');
header('X-Content-Type-Options: nosniff');
header('Referrer-Policy: strict-origin-when-cross-origin');
?><!doctype html>
<html lang="<?= $LANG ?>">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Microads Office — <?= t('Word, Excel, PDF miễn phí · Tải & cập nhật', 'Free Word, Excel, PDF · Download & updates') ?></title>
<meta name="description" content="<?= h(t('Microads Office: bộ ứng dụng văn phòng Word, Excel, PDF (đọc & ký số) — phát hành bởi Công ty TNHH Công nghệ Microads, luôn luôn miễn phí.', 'Microads Office: Word, Excel and PDF (read & sign) office suite — published by Microads Technology Co., Ltd., always free.')) ?>">
<link rel="canonical" href="https://microads.vn/office/">
<?php if (($c['gscVerification'] ?? '') !== ''): ?><meta name="google-site-verification" content="<?= h($c['gscVerification']) ?>">
<?php endif; ?>
<link rel="icon" href="assets/office.png">
<link rel="stylesheet" href="assets/style.css">
</head>
<body>
<header class="bar">
  <a class="brand" href="https://microads.vn"><img src="assets/mark.svg" alt="Microads" width="34" height="27"> MICROADS <span>Office</span></a>
  <nav>
    <a href="?lang=vi" class="flag <?= $LANG === 'vi' ? 'on' : '' ?>" title="Tiếng Việt"><img src="assets/flag-vi.svg" alt="" width="24" height="16"> VI</a>
    <a href="?lang=en" class="flag <?= $LANG === 'en' ? 'on' : '' ?>" title="English"><img src="assets/flag-en.svg" alt="" width="24" height="16"> EN</a>
    <a href="https://microads.vn" class="site">microads.vn</a>
  </nav>
</header>

<section class="hero">
  <div class="wrap hero-in">
    <div class="hero-text">
      <span class="badge"><?= t('Luôn luôn miễn phí', 'Always free') ?></span>
      <h1>Microads Office</h1>
      <p class="lead"><?= t('Soạn thảo văn bản, bảng tính và PDF (đọc, ký số, in) trong một bộ cài duy nhất — không giới hạn tính năng, không cần tài khoản.', 'Documents, spreadsheets and PDF (read, sign, print) in a single installer — no hidden limits, no account needed.') ?></p>
      <?php if ($latest): ?>
        <?php if ($store): ?>
          <a class="btn primary big" href="<?= h($store['installer']) ?>"><?= t('Tải Microads Office', 'Download Microads Office') ?></a>
          <a class="btn big ghost" href="download/<?= h(rawurlencode($web['file'] ?? $latest['file'])) ?>"><?= t('Tải bộ cài .exe', 'Download .exe installer') ?></a>
        <?php elseif ($web): ?>
          <a class="btn primary big" id="dl-main" href="download/<?= h(rawurlencode($web['file'])) ?>"><?= t('Tải Microads Office', 'Download Microads Office') ?></a>
          <a class="btn big ghost" id="dl-full" href="download/<?= h(rawurlencode($latest['file'])) ?>"><?= t('Bộ cài đầy đủ', 'Full installer') ?> <?= h($latest['version']) ?></a>
        <?php else: ?>
          <a class="btn primary big" id="dl-full" href="download/<?= h(rawurlencode($latest['file'])) ?>"><?= t('Tải Microads Office', 'Download Microads Office') ?> <?= h($latest['version']) ?></a>
        <?php endif; ?>
        <?php if ($web && !$store): ?>
        <p class="meta"><?= t('Windows 10/11 64-bit · trình cài đặt', 'Windows 10/11 64-bit · installer') ?> <?= fmt_size((int)$web['size']) ?> <?= t('luôn cài bản mới nhất', 'always installs the latest version') ?> <b><?= h($latest['version']) ?></b> (<?= t('phát hành', 'released') ?> <?= h(date('d/m/Y', strtotime($latest['date']))) ?>) · <?= t('bộ cài đầy đủ', 'full installer') ?> <?= fmt_size((int)$latest['size']) ?></p>
        <?php else: ?>
        <p class="meta"><?= t('Windows 10/11 64-bit', 'Windows 10/11 64-bit') ?> · <?= fmt_size((int)$latest['size']) ?> · <?= t('phát hành', 'released') ?> <?= h(date('d/m/Y', strtotime($latest['date']))) ?></p>
        <?php endif; ?>
      <?php else: ?>
        <p class="meta"><?= t('Phiên bản đầu tiên sắp được phát hành.', 'The first version will be released soon.') ?></p>
      <?php endif; ?>
    </div>
    <div class="apps">
      <div class="app word"><img src="assets/word.png" alt="" width="64" height="64"><b>Word</b><span><?= t('Soạn thảo văn bản, .docx', 'Word processing, .docx') ?></span></div>
      <div class="app excel"><img src="assets/excel.png" alt="" width="64" height="64"><b>Excel</b><span><?= t('Bảng tính, công thức, biểu đồ', 'Spreadsheets, formulas, charts') ?></span></div>
      <div class="app pdf"><img src="assets/pdf.png" alt="" width="64" height="64"><b>PDF</b><span><?= t('Đọc, ký số, in PDF', 'Read, sign, print PDF') ?></span></div>
    </div>
  </div>
</section>

<main class="wrap">
  <?php if ($latest): ?>
  <section class="card" id="cai-dat-an-toan">
    <h2><?= t('Cài đặt an toàn', 'Safe installation') ?></h2>
    <?php if ($store): ?>
    <p><?= t('Nút <b>Tải Microads Office</b> tải bộ cài chính thức của <b>Microsoft Store</b> (do Microsoft ký số): cài đặt không có cảnh báo, tự cập nhật. Xem trên <a href="' . h($store['page']) . '">Microsoft Store</a>.', 'The <b>Download</b> button gets the official <b>Microsoft Store</b> installer (signed by Microsoft): no warnings, automatic updates. See it on <a href="' . h($store['page']) . '">Microsoft Store</a>.') ?></p>
    <?php endif; ?>
    <p><?= t('Microads Office là phần mềm miễn phí, không mua chứng thư ký mã trả phí. Với tệp .exe mới, Chrome có thể báo <i>"tệp không thường được tải xuống / đáng ngờ"</i> và Windows có thể hiện <i>"Windows đã bảo vệ PC của bạn"</i>. Đây là cảnh báo về <b>độ phổ biến</b> của tệp, không phải phát hiện virus: Microsoft Defender không tìm thấy mối đe dọa nào trong Microads Office.',
        'Microads Office is free software and does not buy a paid code-signing certificate. For a new .exe, Chrome may say the file is <i>"not commonly downloaded / suspicious"</i> and Windows may show <i>"Windows protected your PC"</i>. These warnings are about how <b>common</b> the file is, not a virus detection: Microsoft Defender finds no threat in Microads Office.') ?></p>
    <ol class="steps">
      <li><?= t('Chỉ tải từ đúng địa chỉ <b>microads.vn/office</b>.', 'Only download from <b>microads.vn/office</b>.') ?></li>
      <li><?= t('<b>Chrome</b> báo <i>"Chrome đã chặn tệp tải xuống này vì tệp này ít được tải xuống"</i>: nhấn <kbd>Ctrl</kbd>+<kbd>J</kbd> → bấm <b>⋮</b> bên cạnh tệp → <b>Tải tệp đáng ngờ xuống</b>. <b>Edge</b>: bấm <b>…</b> → <b>Giữ lại</b>.', '<b>Chrome</b> says the file <i>"isn\'t commonly downloaded"</i>: press <kbd>Ctrl</kbd>+<kbd>J</kbd> → click <b>⋮</b> next to the file → <b>Download suspicious file</b>. <b>Edge</b>: click <b>…</b> → <b>Keep</b>.') ?></li>
      <li><?= t('<b>Windows</b>: bấm <b>Thông tin bổ sung</b> → <b>Vẫn chạy</b>. Windows chỉ hỏi một lần: Word, Excel, PDF sau khi cài và mọi bản cập nhật về sau đều không hỏi lại.', '<b>Windows</b>: click <b>More info</b> → <b>Run anyway</b>. Windows asks only once: Word, Excel, PDF after installing and every later update never ask again.') ?></li>
    </ol>
    <?php if ($web && !$store): ?>
    <p><?= t('<b>Vì sao nên dùng trình cài đặt trực tuyến?</b> Tệp này nhỏ và <b>không bao giờ thay đổi</b> giữa các phiên bản, nên mỗi lượt tải đều giúp Chrome và Windows nhận biết nó — cảnh báo giảm dần rồi hết hẳn, kể cả khi Microads phát hành bản mới. Trình cài đặt tải bản mới nhất từ microads.vn và chỉ chạy khi tệp có <b>chữ ký số hợp lệ của Microads</b>.',
        '<b>Why the web installer?</b> It is small and <b>never changes</b> between versions, so every download helps Chrome and Windows recognise it — the warnings fade and disappear, even when Microads releases a new version. It downloads the latest version from microads.vn and only runs it when the file carries a <b>valid Microads digital signature</b>.') ?></p>
    <?php endif; ?>
    <p class="muted" id="signpath-notice">Free code signing provided by <a href="https://about.signpath.io/">SignPath.io</a>, certificate by <a href="https://signpath.org/">SignPath Foundation</a>.</p>
    <p class="muted"><a href="code-signing">Code signing policy</a> · <?= t('mã nguồn công khai tại', 'source code at') ?> <a href="https://github.com/microads-vn/microads-office">github.com/microads-vn/microads-office</a></p>
    <p class="muted"><?= t('Kiểm tra tính toàn vẹn (PowerShell):', 'Verify integrity (PowerShell):') ?> <code>Get-FileHash &lt;<?= t('tệp', 'file') ?>&gt;</code><br>
      <?php if ($web): ?><?= h($web['file']) ?> — SHA-256: <code class="hash" id="sha-web"><?= h($web['sha256']) ?></code><br><?php endif; ?>
      <?= h($latest['file']) ?> — SHA-256: <code class="hash" id="sha-full"><?= h($latest['sha256']) ?></code></p>
  </section>

  <section class="card" id="cai-bang-lenh">
    <h2><?= t('Cài bằng lệnh — không có cảnh báo của trình duyệt', 'Install with a command — no browser warnings') ?></h2>
    <p><?= t('Dành cho người dùng quen máy tính và bộ phận IT. Lệnh tải bản mới nhất trực tiếp từ microads.vn, kiểm tra SHA-256 và chữ ký số của Microads rồi mới cài, nên không đi qua bước cảnh báo tải xuống của trình duyệt.',
        'For experienced users and IT departments. The command downloads the latest version directly from microads.vn, checks the SHA-256 and the Microads digital signature before installing, so the browser download warning does not apply.') ?></p>
    <?php if ($winget): ?>
    <h3>Windows Package Manager (winget)</h3>
    <div class="cmd"><code id="cmd-winget">winget install --id <?= h($winget) ?> -e</code><button type="button" class="btn sm copy" data-copy="cmd-winget"><?= t('Sao chép', 'Copy') ?></button></div>
    <?php endif; ?>
    <h3>PowerShell</h3>
    <p class="muted"><?= t('Mở <b>PowerShell</b> từ menu Start (không cần quyền quản trị), dán lệnh rồi nhấn Enter:', 'Open <b>PowerShell</b> from the Start menu (no administrator rights needed), paste the command and press Enter:') ?></p>
    <div class="cmd"><code id="cmd-ps">irm <?= h(MO_BASE_URL) ?>/install.ps1 | iex</code><button type="button" class="btn sm copy" data-copy="cmd-ps"><?= t('Sao chép', 'Copy') ?></button></div>
    <?php if ($web && !$store): ?>
    <h3><?= t('Cài hàng loạt (IT)', 'Mass deployment (IT)') ?></h3>
    <div class="cmd"><code id="cmd-silent">MicroadsOffice_Setup.exe --silent</code><button type="button" class="btn sm copy" data-copy="cmd-silent"><?= t('Sao chép', 'Copy') ?></button></div>
    <p class="muted"><?= t('Cài im lặng cho người dùng hiện tại, không cần quyền quản trị; thêm <code>--dir "D:\\Apps\\Microads Office"</code> để chọn thư mục. Mã thoát 0 = thành công.', 'Silent install for the current user, no administrator rights; add <code>--dir "D:\\Apps\\Microads Office"</code> to choose the folder. Exit code 0 = success.') ?></p>
    <?php endif; ?>
    <p class="muted"><?= t('An toàn: chỉ dán lệnh lấy từ đúng trang <b>microads.vn/office</b>. Microads không bao giờ yêu cầu bạn dán lệnh vào hộp thoại Run (Win+R).', 'Safety: only paste commands taken from <b>microads.vn/office</b> itself. Microads never asks you to paste a command into the Run box (Win+R).') ?></p>
  </section>
  <?php endif; ?>

  <section class="card">
    <h2><?= t('Cập nhật phần mềm', 'Software updates') ?></h2>
    <p><?= t('Ứng dụng tự kiểm tra phiên bản mới mỗi khi mở. Bạn cũng có thể bấm <b>Kiểm tra cập nhật</b> (biểu tượng ↻ ở góc trên bên phải, hoặc thẻ Trợ giúp) trong Word, Excel hoặc PDF. Mọi bản cập nhật đều có chữ ký số của Microads và được kiểm tra trước khi cài.',
        'The apps check for a new version every time they start. You can also click <b>Check for updates</b> (the ↻ icon at the top right, or the Help tab) in Word, Excel or PDF. Every update is digitally signed by Microads and verified before it is installed.') ?></p>
    <?php if ($latest && trim($L($latest['notes'])) !== ''): ?>
      <h3><?= t('Có gì mới trong', "What's new in") ?> <?= h($latest['version']) ?></h3>
      <div class="notes"><?= nl2br(h($L($latest['notes']))) ?></div>
    <?php endif; ?>
  </section>

  <?php if (count($c['releases']) > 1): ?>
  <section class="card">
    <h2><?= t('Các phiên bản', 'Versions') ?></h2>
    <table class="tbl"><thead><tr><th><?= t('Phiên bản', 'Version') ?></th><th><?= t('Ngày', 'Date') ?></th><th><?= t('Ghi chú', 'Notes') ?></th></tr></thead><tbody>
    <?php foreach ($c['releases'] as $r): ?>
      <tr><td><b><?= h($r['version']) ?></b></td><td><?= h(date('d/m/Y', strtotime($r['date']))) ?></td><td class="muted"><?= nl2br(h($L($r['notes']))) ?></td></tr>
    <?php endforeach; ?>
    </tbody></table>
  </section>
  <?php endif; ?>

  <section class="card publisher">
    <img src="assets/mark.svg" alt="" width="70" height="56">
    <div>
      <h2><?= h($L($p['company'])) ?></h2>
      <p><?= h($L($p['headline'])) ?></p>
      <p class="muted"><?= h($L($p['tagline'])) ?>:</p>
      <ul class="services"><?php foreach ($p['services'][$LANG] ?? $p['services']['vi'] as $s): ?><li><?= h($s) ?></li><?php endforeach; ?></ul>
      <p><a class="btn primary" href="<?= h($p['website']) ?>"><?= h($L($p['callToAction'])) ?> →</a>
        <?php if ($p['phone']): ?><span class="contact">☎ <?= h($p['phone']) ?></span><?php endif; ?>
        <?php if ($p['email']): ?><span class="contact">✉ <?= h($p['email']) ?></span><?php endif; ?></p>
    </div>
  </section>
</main>
<footer class="foot wrap">© <?= date('Y') ?> <?= h($L($p['company'])) ?> · <a href="https://microads.vn">microads.vn</a> · <a href="privacy"><?= t('Chính sách quyền riêng tư', 'Privacy policy') ?></a> · <a href="code-signing" id="lnk-code-signing">Code signing policy</a> · <a href="https://github.com/microads-vn/microads-office" id="lnk-source"><?= t('Mã nguồn mở (GPL-3.0)', 'Open source (GPL-3.0)') ?></a></footer>
<script>
document.querySelectorAll('button.copy').forEach(function (b) {
  b.addEventListener('click', function () {
    var el = document.getElementById(b.getAttribute('data-copy')), text = el.textContent, label = b.textContent;
    function done() { b.textContent = '<?= t('Đã sao chép ✓', 'Copied ✓') ?>'; setTimeout(function () { b.textContent = label; }, 1800); }
    function fallback() { var r = document.createRange(); r.selectNodeContents(el); var s = getSelection(); s.removeAllRanges(); s.addRange(r); try { document.execCommand('copy'); done(); } catch (e) {} }
    if (navigator.clipboard && window.isSecureContext) { navigator.clipboard.writeText(text).then(done, fallback); } else { fallback(); }
  });
});
</script>
</body>
</html>
