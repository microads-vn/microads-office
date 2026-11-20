<?php
// Shared page parts: header with language switch, footer, download notice dialog (home, blog, tool pages).
declare(strict_types=1);

// Expected Microsoft Store release (Vietnam time). Shown as a countdown until the Store ID is set in admin.
const MO_STORE_ETA = '2026-11-21T00:00:00+07:00';

/**
 * Top bar. $alt = ['vi' => url, 'en' => url] (the same page in the other language, or that language's home/blog).
 */
function layout_header(string $lang, array $alt): void {
    $blog = blog_base($lang);
    ?>
<header class="bar">
  <a class="brand" href="<?= h(home_path($lang)) ?>"><img src="/office/assets/mark.svg" alt="Microads" width="34" height="27"> MICROADS <span>Office</span></a>
  <nav>
    <a href="<?= h($blog) ?>" class="navlink">Blog</a>
    <a href="<?= h(tool_path($lang)) ?>" class="navlink tool"><?= $lang === 'en' ? 'Formula generator' : 'Tạo công thức' ?></a>
    <a href="<?= h($alt['vi']) ?>" class="flag <?= $lang === 'vi' ? 'on' : '' ?>" title="Tiếng Việt" hreflang="vi"><img src="/office/assets/flag-vi.svg" alt="" width="24" height="16"> VI</a>
    <a href="<?= h($alt['en']) ?>" class="flag <?= $lang === 'en' ? 'on' : '' ?>" title="English" hreflang="en"><img src="/office/assets/flag-en.svg" alt="" width="24" height="16"> EN</a>
    <a href="https://microads.vn" class="site">microads.vn</a>
  </nav>
</header>
<?php
}

function layout_footer(string $lang, array $promo): void {
    $L = fn(array $x) => $x[$lang] ?? ($x['vi'] ?? '');
    $t = fn($vi, $en) => $lang === 'en' ? $en : $vi;
    ?>
<footer class="foot wrap">© <?= date('Y') ?> <?= h($L($promo['company'])) ?> · <a href="https://microads.vn">microads.vn</a> · <a href="<?= h(blog_base($lang)) ?>">Blog</a> · <a href="<?= h(tool_path($lang)) ?>"><?= $t('Tạo công thức Excel', 'Excel formula generator') ?></a> · <a href="/office/privacy"><?= $t('Chính sách quyền riêng tư', 'Privacy policy') ?></a> · <a href="/office/code-signing" id="lnk-code-signing">Code signing policy</a> · <a href="https://github.com/microads-vn/microads-office" id="lnk-source"><?= $t('Mã nguồn mở (GPL-3.0)', 'Open source (GPL-3.0)') ?></a></footer>
<?php
}

/** The "before you download" notice + the Microsoft Store waiting notice, with their scripts. */
function layout_download_notice(string $lang, ?array $store): void {
    $t = fn($vi, $en) => $lang === 'en' ? $en : $vi;
    ?>
<dialog class="warn" id="dl-warn" aria-labelledby="dl-warn-title">
  <form method="dialog"><button class="x" value="cancel" aria-label="<?= $t('Đóng', 'Close') ?>">×</button></form>
  <h2 id="dl-warn-title"><?= $t('Lưu ý trước khi tải', 'Before you download') ?></h2>
  <p><?= $t('Microads Office là phần mềm miễn phí, an toàn và <b>không chứa virus</b>. Vì đây là ứng dụng mới phát hành, chưa có nhiều lượt tải nên chưa được trình duyệt và Windows “quen mặt”, bạn có thể gặp cảnh báo. Cứ yên tâm tiếp tục:',
      'Microads Office is free, safe and <b>contains no viruses</b>. Because it is a newly released app with few downloads so far, your browser and Windows do not “know” it yet and may show a warning. You can safely continue:') ?></p>
  <ol class="steps">
    <li><?= $t('<b>Chrome chặn tải về</b> (“tệp này ít được tải xuống”): nhấn <kbd>Ctrl</kbd>+<kbd>J</kbd> → bấm <b>⋮</b> cạnh tệp → <b>Tải tệp đáng ngờ xuống</b>. Trên <b>Edge</b>: bấm <b>…</b> → <b>Giữ lại</b>.', '<b>Chrome blocks the download</b> (“this file isn’t commonly downloaded”): press <kbd>Ctrl</kbd>+<kbd>J</kbd> → click <b>⋮</b> next to the file → <b>Download suspicious file</b>. In <b>Edge</b>: click <b>…</b> → <b>Keep</b>.') ?></li>
    <li><?= $t('<b>Khi cài, Windows hiện màn hình xanh</b> “Windows đã bảo vệ PC của bạn”: bấm <b>Thông tin bổ sung</b> → <b>Vẫn chạy</b>.', '<b>While installing, Windows shows a blue screen</b> “Windows protected your PC”: click <b>More info</b> → <b>Run anyway</b>.') ?></li>
  </ol>
  <p class="muted"><?= $store
      ? $t('Không muốn gặp cảnh báo? Hãy <a href="' . h($store['page']) . '">tải từ Microsoft Store</a>.', 'Prefer no warnings? <a href="' . h($store['page']) . '">Get it from Microsoft Store</a>.')
      : $t('Không muốn gặp cảnh báo? Microads Office sắp có trên Microsoft Store (đang chờ Microsoft duyệt).', 'Prefer no warnings? Microads Office is coming to Microsoft Store (awaiting Microsoft review).') ?></p>
  <div class="warn-actions">
    <a class="btn primary" id="dl-warn-go" href="#"><?= $t('Tiếp tục tải', 'Continue download') ?></a>
    <form method="dialog"><button class="btn" value="cancel"><?= $t('Hủy', 'Cancel') ?></button></form>
  </div>
</dialog>
<?php if (!$store): ?>
<dialog class="warn" id="store-wait" aria-labelledby="store-wait-title">
  <form method="dialog"><button class="x" value="cancel" aria-label="<?= $t('Đóng', 'Close') ?>">×</button></form>
  <h2 id="store-wait-title"><?= $t('Sắp có trên Microsoft Store', 'Coming to Microsoft Store') ?></h2>
  <p id="store-wait-text"><?= $t('Microads Office đang chờ Microsoft duyệt, dự kiến có mặt trên Microsoft Store ngày <b>' . date('d/m/Y', strtotime(MO_STORE_ETA)) . '</b>.',
      'Microads Office is awaiting Microsoft review and is expected on Microsoft Store on <b>' . date('j F Y', strtotime(MO_STORE_ETA)) . '</b>.') ?></p>
  <div class="countdown" id="store-eta" data-eta="<?= h(MO_STORE_ETA) ?>" hidden>
    <div><b data-u="d">0</b><span><?= $t('ngày', 'days') ?></span></div>
    <div><b data-u="h">00</b><span><?= $t('giờ', 'hours') ?></span></div>
    <div><b data-u="m">00</b><span><?= $t('phút', 'min') ?></span></div>
    <div><b data-u="s">00</b><span><?= $t('giây', 'sec') ?></span></div>
  </div>
  <p class="muted small countdown-note" id="store-eta-note" hidden><?= $t('Đếm ngược theo giờ Việt Nam (GMT+7).', 'Countdown in Vietnam time (GMT+7).') ?></p>
  <p class="muted small" id="store-eta-done" hidden><?= $t('Microsoft đang hoàn tất phát hành, bạn vui lòng quay lại sau ít giờ.', 'Microsoft is finishing the release, please check back in a few hours.') ?></p>
  <p><?= $t('Trong lúc chờ, bạn có thể tải bộ cài trực tiếp từ microads.vn bằng nút <b>Tải Microads Office</b>.',
      'Meanwhile, you can download the installer directly from microads.vn with the <b>Download Microads Office</b> button.') ?></p>
  <div class="warn-actions">
    <form method="dialog"><button class="btn primary" value="ok" autofocus><?= $t('Đã hiểu', 'Got it') ?></button></form>
  </div>
</dialog>
<?php endif; ?>
<script>
// Microsoft Store button while the Store listing is not live yet: explain that it is awaiting review.
(function () {
  var b = document.getElementById('dl-store'), d = document.getElementById('store-wait');
  if (!b || !d || !b.hasAttribute('data-wait')) return;
  // Countdown to the expected Store date (fixed instant in Vietnam time, so it is right in any time zone).
  var box = document.getElementById('store-eta'), eta = Date.parse(box.getAttribute('data-eta')), timer = null;
  var note = document.getElementById('store-eta-note'), done = document.getElementById('store-eta-done');
  function pad(n) { return (n < 10 ? '0' : '') + n; }
  function tick() {
    var left = Math.floor((eta - Date.now()) / 1000);
    if (isNaN(eta) || left <= 0) { box.hidden = note.hidden = true; done.hidden = isNaN(eta); if (timer) clearInterval(timer); return; }
    box.hidden = note.hidden = false; done.hidden = true;
    var v = { d: Math.floor(left / 86400), h: pad(Math.floor(left % 86400 / 3600)), m: pad(Math.floor(left % 3600 / 60)), s: pad(left % 60) };
    box.querySelectorAll('b[data-u]').forEach(function (el) { el.textContent = v[el.getAttribute('data-u')]; });
  }
  tick();
  b.addEventListener('click', function () {
    tick(); if (!timer) timer = setInterval(tick, 1000);
    if (typeof d.showModal === 'function') d.showModal(); else alert(document.getElementById('store-wait-text').textContent);
  });
  d.addEventListener('close', function () { if (timer) { clearInterval(timer); timer = null; } });
  d.addEventListener('click', function (e) { if (e.target === d) d.close(); });
})();
// Download buttons first show a short notice (Chrome may block the new file); without <dialog> support they download directly.
(function () {
  var d = document.getElementById('dl-warn'), go = document.getElementById('dl-warn-go');
  if (!d || typeof d.showModal !== 'function') return;
  document.querySelectorAll('a[data-warn]').forEach(function (a) {
    a.addEventListener('click', function (e) {
      if (e.ctrlKey || e.metaKey || e.shiftKey || e.button !== 0) return;
      e.preventDefault(); go.href = a.href; d.showModal(); go.focus();
    });
  });
  go.addEventListener('click', function () { setTimeout(function () { d.close(); }, 0); });
  d.addEventListener('click', function (e) { if (e.target === d) d.close(); });
})();
</script>
<?php
}
