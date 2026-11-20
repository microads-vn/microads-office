<?php
// Public page: microads.vn/office — latest version, download, install guide, versions.
declare(strict_types=1);
require __DIR__ . '/inc/bootstrap.php';
require __DIR__ . '/inc/blog.php';
require __DIR__ . '/inc/layout.php';
// Language: /office/ = Vietnamese, /office/en/ = English (nginx passes hl=en). Each language has its own URL for search engines.
// Visitors who chose English (cookie) or whose browser does not ask for Vietnamese are sent to /office/en/ (crawlers send no
// Accept-Language and always get the Vietnamese page).
$setLang = function (string $l): void { if (PHP_SAPI !== 'cli') setcookie('mo_lang', $l, ['expires' => time() + 31536000, 'path' => '/office', 'secure' => true, 'samesite' => 'Lax']); };
if (($_GET['hl'] ?? '') === 'en') { $LANG = 'en'; }
elseif (($_GET['lang'] ?? '') === 'en') { $setLang('en'); header('Location: /office/en/', true, 301); exit; }
elseif (($_GET['lang'] ?? '') === 'vi') { $LANG = 'vi'; $setLang('vi'); }
else {
    $al = strtolower($_SERVER['HTTP_ACCEPT_LANGUAGE'] ?? '');
    $cookie = $_COOKIE['mo_lang'] ?? '';
    header('Vary: Accept-Language, Cookie');
    if ($cookie === 'en' || ($cookie === '' && $al !== '' && !str_contains($al, 'vi'))) { header('Location: /office/en/', true, 302); exit; }
    $LANG = 'vi';
}
$canonical = 'https://microads.vn' . home_path($LANG);
$posts = blog_list($LANG);
track_view(home_path($LANG));
$c = load_config();
$p = $c['promo'];
$latest = release_by_version($c, $c['latest']);
$store = store_links($c);
$web = websetup_info($c);           // web installer: small, never changes, always installs the latest signed release
$L = fn(array $x) => $x[$LANG] ?? ($x['vi'] ?? '');
// Generic shopping-bag icon for the Microsoft Store button (no Microsoft logo).
$storeIcon = '<svg aria-hidden="true" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M5 8h14l-1 12H6L5 8z"/><path d="M9 8V6a3 3 0 0 1 6 0v2"/></svg>';
// Release notes: lines starting with "-" or "•" become a list, other lines paragraphs.
$notes = function (string $s): string {
    $out = ''; $items = [];
    $flush = function () use (&$out, &$items) { if ($items) { $out .= '<ul>' . implode('', $items) . '</ul>'; $items = []; } };
    foreach (preg_split('/\R/u', trim($s)) as $line) {
        $line = trim($line);
        if ($line === '') continue;
        if (preg_match('/^[-•]\s*(.+)$/u', $line, $m)) { $items[] = '<li>' . h($m[1]) . '</li>'; continue; }
        $flush(); $out .= '<p>' . h($line) . '</p>';
    }
    $flush();
    return $out;
};
header('X-Content-Type-Options: nosniff');
header('Referrer-Policy: strict-origin-when-cross-origin');
?><!doctype html>
<html lang="<?= $LANG ?>">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title><?= t('Microads Office — Word, Excel, PDF miễn phí, không cần tài khoản', 'Microads Office — Free Word, Excel & PDF for Windows, no account') ?></title>
<meta name="description" content="<?= h(t('Tải Microads Office miễn phí vĩnh viễn: soạn thảo văn bản .docx, bảng tính .xlsx với 502 hàm Excel, đọc và ký số PDF bằng USB token. Không cần tài khoản, dành cho Windows 10/11.', 'Download Microads Office free forever: a .docx word processor, .xlsx spreadsheets with 502 Excel functions, and a PDF reader that signs with digital certificates. No account needed. Windows 10/11.')) ?>">
<link rel="canonical" href="<?= h($canonical) ?>">
<link rel="alternate" hreflang="vi" href="https://microads.vn/office/">
<link rel="alternate" hreflang="en" href="https://microads.vn/office/en/">
<link rel="alternate" hreflang="x-default" href="https://microads.vn/office/en/">
<link rel="alternate" type="application/rss+xml" title="Microads Office Blog" href="<?= h(blog_base($LANG)) ?>feed.xml">
<meta property="og:type" content="website">
<meta property="og:site_name" content="Microads Office">
<meta property="og:title" content="<?= h(t('Microads Office — Word, Excel, PDF miễn phí vĩnh viễn', 'Microads Office — free Word, Excel & PDF, forever')) ?>">
<meta property="og:description" content="<?= h(t('Bộ ứng dụng văn phòng miễn phí cho Windows. Không cần tài khoản.', 'A free office suite for Windows. No account needed.')) ?>">
<meta property="og:url" content="<?= h($canonical) ?>">
<meta property="og:image" content="https://microads.vn/office/assets/blog/free-office-<?= $LANG ?>.jpg">
<meta property="og:locale" content="<?= $LANG === 'en' ? 'en_US' : 'vi_VN' ?>">
<meta name="twitter:card" content="summary_large_image">
<script type="application/ld+json"><?= json_encode([
    '@context' => 'https://schema.org',
    '@graph' => [
        ['@type' => 'Organization', '@id' => 'https://microads.vn/#org', 'name' => $LANG === 'en' ? 'Microads Technology Co., Ltd.' : 'Công ty TNHH Công nghệ Microads', 'url' => 'https://microads.vn/', 'logo' => 'https://microads.vn/office/assets/office.png'],
        ['@type' => 'SoftwareApplication', 'name' => 'Microads Office', 'operatingSystem' => 'Windows 10, Windows 11', 'applicationCategory' => 'BusinessApplication',
            'inLanguage' => ['vi', 'en'], 'url' => $canonical, 'downloadUrl' => 'https://microads.vn/office/download/MicroadsOffice_Setup.exe',
            'softwareVersion' => $latest['version'] ?? null, 'license' => 'https://www.gnu.org/licenses/gpl-3.0.html', 'isAccessibleForFree' => true,
            'offers' => ['@type' => 'Offer', 'price' => '0', 'priceCurrency' => 'USD'], 'publisher' => ['@id' => 'https://microads.vn/#org'],
            'screenshot' => 'https://microads.vn/office/assets/blog/shot-excel-explain.jpg'],
    ]], JSON_UNESCAPED_SLASHES | JSON_UNESCAPED_UNICODE) ?></script>
<?php if (($c['gscVerification'] ?? '') !== ''): ?><meta name="google-site-verification" content="<?= h($c['gscVerification']) ?>">
<?php endif; ?>
<link rel="icon" href="/office/assets/office.png">
<link rel="stylesheet" href="/office/assets/style.css?v=7">
</head>
<body>
<?php layout_header($LANG, ['vi' => '/office/?lang=vi', 'en' => '/office/en/']); ?>

<section class="hero">
  <div class="wrap hero-in">
    <div class="hero-text">
      <span class="badge"><?= t('Luôn luôn miễn phí', 'Always free') ?></span>
      <h1>Microads Office</h1>
      <p class="lead"><?= t('Soạn thảo văn bản, bảng tính và PDF (đọc, ký số, in) trong một bộ cài duy nhất — không giới hạn tính năng, không cần tài khoản.', 'Documents, spreadsheets and PDF (read, sign, print) in a single installer — no hidden limits, no account needed.') ?></p>
      <?php if ($latest): ?>
        <div class="dl">
        <?php if ($web): ?>
          <a class="btn primary big" id="dl-main" href="/office/download/<?= h(rawurlencode($web['file'])) ?>" data-warn><?= t('Tải Microads Office', 'Download Microads Office') ?></a>
          <a class="btn big ghost" id="dl-full" href="/office/download/<?= h(rawurlencode($latest['file'])) ?>" data-warn><?= t('Bộ cài đầy đủ', 'Full installer') ?> <?= h($latest['version']) ?></a>
        <?php else: ?>
          <a class="btn primary big" id="dl-full" href="/office/download/<?= h(rawurlencode($latest['file'])) ?>" data-warn><?= t('Tải Microads Office', 'Download Microads Office') ?> <?= h($latest['version']) ?></a>
        <?php endif; ?>
        <?php if ($store): ?>
          <a class="btn big store" id="dl-store" href="<?= h($store['page']) ?>"><?= $storeIcon ?><?= t('Tải từ Microsoft Store', 'Get it from Microsoft Store') ?></a>
        <?php else: ?>
          <button type="button" class="btn big store" id="dl-store" data-wait><?= $storeIcon ?><?= t('Tải từ Microsoft Store', 'Get it from Microsoft Store') ?></button>
        <?php endif; ?>
        </div>
        <?php if ($web): ?>
        <p class="meta"><?= t('Windows 10/11 64-bit · trình cài đặt', 'Windows 10/11 64-bit · installer') ?> <?= fmt_size((int)$web['size']) ?> <?= t('luôn cài bản mới nhất', 'always installs the latest version') ?> <b><?= h($latest['version']) ?></b> (<?= t('phát hành', 'released') ?> <?= h(date('d/m/Y', strtotime($latest['date']))) ?>) · <?= t('bộ cài đầy đủ', 'full installer') ?> <?= fmt_size((int)$latest['size']) ?></p>
        <?php else: ?>
        <p class="meta"><?= t('Windows 10/11 64-bit', 'Windows 10/11 64-bit') ?> · <?= fmt_size((int)$latest['size']) ?> · <?= t('phát hành', 'released') ?> <?= h(date('d/m/Y', strtotime($latest['date']))) ?></p>
        <?php endif; ?>

        <details class="guide" id="cai-dat-an-toan">
          <summary><?= t('Hướng dẫn cài đặt', 'Installation guide') ?></summary>
          <div class="guide-body">
            <p><?= t('Microads Office là phần mềm miễn phí và an toàn. Vì là ứng dụng mới phát hành, còn ít lượt tải, Chrome có thể báo <i>"tệp ít được tải xuống / đáng ngờ"</i> và Windows có thể hiện màn hình xanh <i>"Windows đã bảo vệ PC của bạn"</i>. Đây là cảnh báo về <b>độ phổ biến</b> của tệp, không phải phát hiện virus: Microsoft Defender không tìm thấy mối đe dọa nào trong Microads Office.',
                'Microads Office is free and safe software. Because it is a newly released app with few downloads so far, Chrome may say the file is <i>"not commonly downloaded / suspicious"</i> and Windows may show the blue <i>"Windows protected your PC"</i> screen. These warnings are about how <b>common</b> the file is, not a virus detection: Microsoft Defender finds no threat in Microads Office.') ?></p>
            <ol class="steps">
              <li><?= t('Chỉ tải từ đúng địa chỉ <b>microads.vn/office</b>.', 'Only download from <b>microads.vn/office</b>.') ?></li>
              <li><?= t('<b>Chrome</b> báo <i>"Chrome đã chặn tệp tải xuống này vì tệp này ít được tải xuống"</i>: nhấn <kbd>Ctrl</kbd>+<kbd>J</kbd> → bấm <b>⋮</b> bên cạnh tệp → <b>Tải tệp đáng ngờ xuống</b>. <b>Edge</b>: bấm <b>…</b> → <b>Giữ lại</b>.', '<b>Chrome</b> says the file <i>"isn\'t commonly downloaded"</i>: press <kbd>Ctrl</kbd>+<kbd>J</kbd> → click <b>⋮</b> next to the file → <b>Download suspicious file</b>. <b>Edge</b>: click <b>…</b> → <b>Keep</b>.') ?></li>
              <li><?= t('<b>Windows</b> hiện màn hình xanh: bấm <b>Thông tin bổ sung</b> → <b>Vẫn chạy</b>. Windows chỉ hỏi một lần: Word, Excel, PDF sau khi cài và mọi bản cập nhật về sau đều không hỏi lại.', '<b>Windows</b>: click <b>More info</b> → <b>Run anyway</b>. Windows asks only once: Word, Excel, PDF after installing and every later update never ask again.') ?></li>
              <li><?= t('Ứng dụng tự kiểm tra phiên bản mới; mọi bản cập nhật đều có chữ ký số của Microads và được kiểm tra trước khi cài.', 'The apps check for new versions themselves; every update is digitally signed by Microads and verified before it is installed.') ?></li>
            </ol>
            <?php if ($web): ?>
            <p><?= t('<b>Nên chọn nút nào?</b> <b>Tải Microads Office</b> là trình cài đặt trực tuyến nhỏ gọn, tự tải bản mới nhất từ microads.vn và chỉ chạy khi tệp có <b>chữ ký số hợp lệ của Microads</b>. <b>Bộ cài đầy đủ</b> dùng khi cần cài trên máy không có Internet. <b>Microsoft Store</b> cài không có cảnh báo và tự cập nhật.',
                '<b>Which button?</b> <b>Download Microads Office</b> is a small web installer: it downloads the latest version from microads.vn and only runs it when the file carries a <b>valid Microads digital signature</b>. Use the <b>Full installer</b> for computers without Internet access. <b>Microsoft Store</b> installs without warnings and updates automatically.') ?></p>
            <?php endif; ?>
            <p class="muted small"><a href="/office/code-signing">Code signing policy</a> · <?= t('mã nguồn công khai tại', 'source code at') ?> <a href="https://github.com/microads-vn/microads-office">github.com/microads-vn/microads-office</a><br>
              <?= t('Kiểm tra tính toàn vẹn (PowerShell):', 'Verify integrity (PowerShell):') ?> <code>Get-FileHash &lt;<?= t('tệp', 'file') ?>&gt;</code><br>
              <?php if ($web): ?><?= h($web['file']) ?> — SHA-256: <code class="hash" id="sha-web"><?= h($web['sha256']) ?></code><br><?php endif; ?>
              <?= h($latest['file']) ?> — SHA-256: <code class="hash" id="sha-full"><?= h($latest['sha256']) ?></code></p>
          </div>
        </details>
      <?php else: ?>
        <p class="meta"><?= t('Phiên bản đầu tiên sắp được phát hành.', 'The first version will be released soon.') ?></p>
      <?php endif; ?>
    </div>
    <div class="apps">
      <div class="app word"><img src="/office/assets/word.png" alt="" width="64" height="64"><b>Word</b><span><?= t('Soạn thảo văn bản, .docx', 'Word processing, .docx') ?></span></div>
      <div class="app excel"><img src="/office/assets/excel.png" alt="" width="64" height="64"><b>Excel</b><span><?= t('Bảng tính, công thức, biểu đồ', 'Spreadsheets, formulas, charts') ?></span></div>
      <div class="app pdf"><img src="/office/assets/pdf.png" alt="" width="64" height="64"><b>PDF</b><span><?= t('Đọc, ký số, in PDF', 'Read, sign, print PDF') ?></span></div>
    </div>
  </div>
</section>

<main class="wrap">
  <?php if ($c['releases']): ?>
  <section class="card" id="cac-phien-ban">
    <h2><?= t('Các phiên bản', 'Versions') ?></h2>
    <table class="tbl versions"><thead><tr><th><?= t('Phiên bản', 'Version') ?></th><th><?= t('Ngày', 'Date') ?></th><th><?= t('Ghi chú', 'Notes') ?></th></tr></thead><tbody>
    <?php foreach ($c['releases'] as $r): ?>
      <tr><td class="v"><b><?= h($r['version']) ?></b></td><td class="d"><?= h(date('d/m/Y', strtotime($r['date']))) ?></td><td class="n"><?= $notes($L($r['notes'])) ?></td></tr>
    <?php endforeach; ?>
    </tbody></table>
  </section>
  <?php endif; ?>

  <?php if ($posts): ?>
  <section class="card" id="bai-viet">
    <h2><?= t('Bài viết mới', 'Latest articles') ?></h2>
    <div class="post-grid">
    <?php foreach (array_slice($posts, 0, 3) as $bp): ?>
      <a class="post-card" href="<?= h(blog_base($LANG) . $bp['slug']) ?>"><img src="/office/assets/blog/<?= h($bp['pillar']) ?>-<?= $LANG ?>.jpg" alt="" width="600" height="315" loading="lazy"><b><?= h($bp['title']) ?></b><span><?= h($bp['excerpt'] ?? '') ?></span></a>
    <?php endforeach; ?>
    </div>
    <p><a href="<?= h(blog_base($LANG)) ?>"><?= t('Xem tất cả bài viết', 'See all articles') ?> →</a></p>
  </section>
  <?php endif; ?>

  <section class="card tool-promo">
    <div>
      <h2><?= t('Tạo công thức Excel bằng AI — miễn phí', 'Free AI Excel formula generator') ?></h2>
      <p><?= t('Mô tả việc cần tính bằng lời, nhận ngay công thức kèm giải thích. Dùng được trong Microads Office Excel và Microsoft Excel.', 'Describe the calculation in plain words and get a ready-to-paste formula with an explanation. Works in Microads Office Excel and Microsoft Excel.') ?></p>
    </div>
    <a class="btn primary" href="<?= h(tool_path($LANG)) ?>"><?= t('Thử ngay', 'Try it') ?> →</a>
  </section>

  <section class="card publisher">
    <img src="/office/assets/mark.svg" alt="" width="70" height="56">
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
<?php layout_footer($LANG, $p); ?>
<?php if ($latest) layout_download_notice($LANG, $store); ?>
<script>
// Old links to #cai-dat-an-toan open the installation guide.
(function () { var d = document.getElementById('cai-dat-an-toan'); if (d && location.hash === '#cai-dat-an-toan') d.open = true; })();
</script>
</body>
</html>
