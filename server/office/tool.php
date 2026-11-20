<?php
// Free AI Excel formula generator: /office/en/excel-formula-generator (EN) · /office/cong-cu/tao-cong-thuc-excel (VI).
// The page posts to /office/api/formula (tool-api.php), which calls Claude with per-IP and daily limits.
declare(strict_types=1);
require __DIR__ . '/inc/bootstrap.php';
require __DIR__ . '/inc/blog.php';
require __DIR__ . '/inc/layout.php';
require __DIR__ . '/inc/tool-lib.php';
$LANG = ($_GET['hl'] ?? 'vi') === 'en' ? 'en' : 'vi';
$c = load_config();
$store = store_links($c);
$T = fn($vi, $en) => $LANG === 'en' ? $en : $vi;
$canonical = 'https://microads.vn' . tool_path($LANG);
track_view(tool_path($LANG));
$title = $T('Tạo công thức Excel bằng AI miễn phí — viết công thức từ mô tả tiếng Việt', 'Free AI Excel Formula Generator — plain English to formula');
$desc = $T('Mô tả phép tính bằng tiếng Việt, AI viết công thức Excel kèm giải thích từng bước. Miễn phí, không cần đăng ký. Dùng được với Excel 365 và Microads Office Excel.',
    'Describe what you need in plain English and get a working Excel formula with a step-by-step explanation. Free, no sign-up. Works in Excel 365 and Microads Office Excel.');
$examples = $LANG === 'en'
    ? ['Sum column C when column A is "North" and column B is in 2026', 'Look up the price for the product code in E2 from the table A2:C200, show "Not found" if missing', 'Count unique customers in A2:A500', 'Working days between the dates in B2 and C2, excluding the holidays in H2:H20', 'Extract the domain from the email address in A2']
    : ['Tính tổng cột C khi cột A là "Hà Nội" và tháng ở cột B là tháng 11', 'Tra giá theo mã hàng ở E2 trong bảng A2:C200, không có thì ghi "Không tìm thấy"', 'Đếm số khách hàng không trùng trong A2:A500', 'Đọc số tiền ở ô D10 thành chữ tiếng Việt', 'Tính số ngày công giữa B2 và C2, trừ ngày lễ ở H2:H20'];
$faq = $LANG === 'en' ? [
    ['Is the formula generator really free?', 'Yes. It is free to use without an account. To keep it fast for everyone there is a small daily limit per visitor.'],
    ['Which Excel versions do the formulas work in?', 'Formulas are written for Excel 365 / Excel 2021 and newer. Dynamic-array functions such as FILTER, UNIQUE or XLOOKUP need Excel 2021 or 365. Every formula also works in the free Microads Office Excel, which supports 502 Excel functions.'],
    ['Do you store what I type?', 'Your description is sent to the AI model to write the formula and is not used to build a profile of you. Do not paste confidential data — describe the columns instead (for example “dates in column B”).'],
    ['Can it write formulas for Google Sheets?', 'Most generated formulas also work in Google Sheets, but they are written and checked for Excel syntax.'],
] : [
    ['Công cụ có thật sự miễn phí không?', 'Có. Dùng miễn phí, không cần tài khoản. Để công cụ luôn nhanh cho mọi người, mỗi người dùng có giới hạn số lượt mỗi ngày.'],
    ['Công thức dùng được trong Excel bản nào?', 'Công thức viết cho Excel 365 / Excel 2021 trở lên. Các hàm mảng động như FILTER, UNIQUE, XLOOKUP cần Excel 2021 hoặc 365. Tất cả đều dùng được trong Microads Office Excel miễn phí (hỗ trợ 502 hàm), kể cả hàm DOCSO đọc số thành chữ tiếng Việt.'],
    ['Nội dung tôi nhập có bị lưu lại không?', 'Mô tả của bạn được gửi tới mô hình AI để viết công thức, không dùng để lập hồ sơ người dùng. Đừng dán dữ liệu bí mật — hãy mô tả cột (ví dụ “ngày ở cột B”).'],
    ['Dấu phân cách trong công thức là dấu phẩy hay chấm phẩy?', 'Công thức hiển thị với dấu phẩy (,) như Excel tiếng Anh. Nếu Windows của bạn dùng định dạng Việt Nam và Excel báo lỗi, hãy thay dấu phẩy giữa các đối số bằng dấu chấm phẩy (;).'],
];
$ld = ['@context' => 'https://schema.org', '@graph' => [
    ['@type' => 'WebApplication', 'name' => $T('Tạo công thức Excel bằng AI', 'AI Excel Formula Generator'), 'url' => $canonical, 'applicationCategory' => 'BusinessApplication',
        'operatingSystem' => 'Any (web browser)', 'isAccessibleForFree' => true, 'offers' => ['@type' => 'Offer', 'price' => '0', 'priceCurrency' => 'USD'],
        'publisher' => ['@type' => 'Organization', 'name' => 'Microads', 'url' => 'https://microads.vn/'], 'inLanguage' => $LANG],
]];
header('X-Content-Type-Options: nosniff');
header('Referrer-Policy: strict-origin-when-cross-origin');
?><!doctype html>
<html lang="<?= $LANG ?>">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title><?= h($title) ?></title>
<meta name="description" content="<?= h($desc) ?>">
<link rel="canonical" href="<?= h($canonical) ?>">
<link rel="alternate" hreflang="vi" href="https://microads.vn<?= tool_path('vi') ?>">
<link rel="alternate" hreflang="en" href="https://microads.vn<?= tool_path('en') ?>">
<link rel="alternate" hreflang="x-default" href="https://microads.vn<?= tool_path('en') ?>">
<meta property="og:type" content="website"><meta property="og:site_name" content="Microads Office">
<meta property="og:title" content="<?= h($title) ?>"><meta property="og:description" content="<?= h($desc) ?>"><meta property="og:url" content="<?= h($canonical) ?>">
<meta property="og:image" content="https://microads.vn/office/assets/blog/excel-formulas-<?= $LANG ?>.jpg">
<meta name="twitter:card" content="summary_large_image">
<script type="application/ld+json"><?= json_encode($ld, JSON_UNESCAPED_SLASHES | JSON_UNESCAPED_UNICODE | JSON_HEX_TAG) ?></script>
<link rel="icon" href="/office/assets/office.png">
<link rel="stylesheet" href="/office/assets/style.css?v=11">
</head>
<body class="blog">
<?php layout_header($LANG, ['vi' => tool_path('vi'), 'en' => tool_path('en')]); ?>
<main class="wrap blog-main">
  <header class="blog-head">
    <p class="kicker"><?= $T('Công cụ miễn phí · Không cần tài khoản', 'Free tool · No account needed') ?></p>
    <h1><?= $T('Tạo công thức Excel bằng AI', 'AI Excel Formula Generator') ?></h1>
    <p class="lead-dark"><?= $T('Viết điều bạn muốn tính bằng lời thường. AI trả về công thức sẵn sàng dán vào ô, kèm giải thích từng phần.', 'Write what you want to calculate in plain words. The AI returns a formula ready to paste, with an explanation of each part.') ?></p>
  </header>
  <section class="card tool">
    <form id="fx-form" autocomplete="off">
      <label for="fx-q"><?= $T('Bạn muốn công thức làm gì?', 'What should the formula do?') ?></label>
      <textarea id="fx-q" name="q" rows="4" maxlength="600" required placeholder="<?= h($examples[0]) ?>"></textarea>
      <label for="fx-ctx" class="opt"><?= $T('Cấu trúc bảng (không bắt buộc)', 'Your columns (optional)') ?></label>
      <input id="fx-ctx" name="ctx" maxlength="300" placeholder="<?= h($T('Ví dụ: A = ngày, B = chi nhánh, C = doanh thu, dữ liệu từ hàng 2 đến 500', 'e.g. A = date, B = region, C = sales, data in rows 2 to 500')) ?>">
      <input type="text" name="website" class="hp" tabindex="-1" autocomplete="off" aria-hidden="true">
      <input type="hidden" name="t" value="<?= h(tool_token()) ?>">
      <div class="tool-actions"><button class="btn primary big" id="fx-go" type="submit"><?= $T('Tạo công thức', 'Generate formula') ?></button><span class="muted small" id="fx-status" role="status"></span></div>
      <p class="muted small"><?= $T('Gợi ý:', 'Try:') ?> <?php foreach ($examples as $i => $e): ?><button type="button" class="chip" data-ex="<?= h($e) ?>"><?= h(mb_strimwidth($e, 0, 48, '…')) ?></button><?php endforeach; ?></p>
    </form>
    <div id="fx-out" class="fx-out" hidden>
      <div class="fx-formula"><code id="fx-formula"></code><button type="button" class="btn sm" id="fx-copy"><?= $T('Sao chép', 'Copy') ?></button></div>
      <div id="fx-expl" class="fx-expl"></div>
    </div>
  </section>

  <section class="card prose">
    <h2><?= $T('Cách dùng công cụ', 'How to use it') ?></h2>
    <ol>
      <li><?= $T('Mô tả kết quả bạn cần, càng cụ thể càng tốt: cột nào chứa gì, điều kiện ra sao.', 'Describe the result you need as precisely as you can: what each column holds and which conditions apply.') ?></li>
      <li><?= $T('Bấm <b>Tạo công thức</b>, chép công thức và dán vào ô cần tính.', 'Click <b>Generate formula</b>, copy the formula and paste it into the cell.') ?></li>
      <li><?= $T('Đọc phần giải thích để hiểu và tự sửa khi bảng của bạn thay đổi.', 'Read the explanation so you can adapt the formula when your sheet changes.') ?></li>
    </ol>
    <h2><?= $T('Mẹo để có công thức đúng ngay lần đầu', 'Tips for a correct formula on the first try') ?></h2>
    <ul>
      <li><?= $T('Ghi rõ vùng dữ liệu (ví dụ A2:A500) thay vì “cả cột”.', 'Name the exact ranges (for example A2:A500) instead of “the whole column”.') ?></li>
      <li><?= $T('Nói rõ khi không có kết quả thì hiển thị gì (0, ô trống, hay “Không tìm thấy”).', 'Say what to show when nothing matches (0, blank or “Not found”).') ?></li>
      <li><?= $T('Với ngày tháng, cho biết ô chứa ngày thật hay văn bản dạng ngày.', 'For dates, say whether the cells contain real dates or text that looks like dates.') ?></li>
    </ul>
    <p><?= $T('Chưa có Excel? <a href="/office/">Microads Office</a> có Excel miễn phí với 502 hàm — gồm XLOOKUP, FILTER, LET, LAMBDA và hàm DOCSO đọc số thành chữ — không cần tài khoản.',
        'No Excel? <a href="/office/en/">Microads Office</a> includes a free Excel-compatible spreadsheet with 502 functions — including XLOOKUP, FILTER, LET and LAMBDA — and needs no account.') ?></p>
    <h2><?= $T('Câu hỏi thường gặp', 'Frequently asked questions') ?></h2>
    <?php foreach ($faq as [$q, $a]): ?><details><summary><?= h($q) ?></summary><p><?= h($a) ?></p></details><?php endforeach; ?>
  </section>
  <?= blog_cta($LANG, 'end') ?>
</main>
<?php layout_footer($LANG, $c['promo']); ?>
<?php layout_download_notice($LANG, $store); ?>
<script>
(function () {
  var f = document.getElementById('fx-form'), q = document.getElementById('fx-q'), go = document.getElementById('fx-go'), st = document.getElementById('fx-status');
  var out = document.getElementById('fx-out'), fx = document.getElementById('fx-formula'), ex = document.getElementById('fx-expl');
  var L = <?= json_encode(['wait' => $T('AI đang viết công thức…', 'Writing your formula…'), 'err' => $T('Có lỗi, vui lòng thử lại.', 'Something went wrong, please try again.'), 'copied' => $T('Đã sao chép ✓', 'Copied ✓'), 'example' => $T('Ví dụ', 'Example')], JSON_UNESCAPED_UNICODE) ?>;
  document.querySelectorAll('.chip').forEach(function (b) { b.addEventListener('click', function () { q.value = b.getAttribute('data-ex'); q.focus(); }); });
  function para(t) { var p = document.createElement('p'); p.textContent = t; return p; }
  f.addEventListener('submit', function (e) {
    e.preventDefault();
    if (!q.value.trim()) return;
    go.disabled = true; st.textContent = L.wait;
    fetch('/office/api/formula', { method: 'POST', headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ q: q.value, ctx: f.ctx.value, website: f.website.value, t: f.t.value, lang: <?= json_encode($LANG) ?> }) })
      .then(function (r) { return r.json(); })
      .then(function (d) {
        go.disabled = false; st.textContent = '';
        if (!d.ok) { st.textContent = d.message || L.err; return; }
        out.hidden = false; fx.textContent = d.formula; ex.textContent = '';
        ex.appendChild(para(d.explanation));
        if (d.steps && d.steps.length) { var ul = document.createElement('ul'); d.steps.forEach(function (s) { var li = document.createElement('li'); li.textContent = s; ul.appendChild(li); }); ex.appendChild(ul); }
        if (d.example) ex.appendChild(para(L.example + ': ' + d.example));
        if (d.note) { var n = para(d.note); n.className = 'muted small'; ex.appendChild(n); }
      })
      .catch(function () { go.disabled = false; st.textContent = L.err; });
  });
  document.getElementById('fx-copy').addEventListener('click', function () {
    var b = this, t = fx.textContent;
    (navigator.clipboard ? navigator.clipboard.writeText(t) : Promise.reject()).then(function () { b.textContent = L.copied; }, function () {});
  });
})();
</script>
</body>
</html>
