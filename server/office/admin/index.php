<?php
// Microads Office — trang quản trị cập nhật: microads.vn/office/admin/
declare(strict_types=1);
require dirname(__DIR__) . '/inc/bootstrap.php';
require dirname(__DIR__) . '/inc/blog.php';
require dirname(__DIR__) . '/inc/ai.php';

session_name('mo_admin');
// HTTPS-only cookie; MO_DEV_INSECURE_COOKIES is set only by the local test server (plain http on 127.0.0.1).
session_set_cookie_params(['lifetime' => 0, 'path' => '/office/admin', 'secure' => !getenv('MO_DEV_INSECURE_COOKIES'), 'httponly' => true, 'samesite' => 'Strict']);
session_start();
header('X-Frame-Options: DENY');
header('X-Content-Type-Options: nosniff');
header('Referrer-Policy: no-referrer');
header("Content-Security-Policy: default-src 'self'; style-src 'self' 'unsafe-inline'; img-src 'self' data:; script-src 'self' 'unsafe-inline'; frame-ancestors 'none'");
header('Cache-Control: no-store');

$admin = json_read(MO_ADMIN, []);
$flash = [];

function csrf(): string { return $_SESSION['csrf'] ??= bin2hex(random_bytes(16)); }
function csrf_ok(): bool { return hash_equals($_SESSION['csrf'] ?? '', (string)($_POST['csrf'] ?? '')); }
function client_ip(): string { return $_SERVER['REMOTE_ADDR'] ?? '?'; }

function throttle_check(): int {
    $t = json_read(MO_LOGS . '/login-throttle.json', []);
    $e = $t[client_ip()] ?? null;
    if ($e && $e['count'] >= 5 && time() - $e['last'] < 900) return 900 - (time() - $e['last']);
    return 0;
}
function throttle_fail(): void {
    $t = json_read(MO_LOGS . '/login-throttle.json', []);
    $e = $t[client_ip()] ?? ['count' => 0, 'last' => 0];
    if (time() - $e['last'] > 900) $e['count'] = 0;
    $e['count']++; $e['last'] = time();
    $t[client_ip()] = $e;
    foreach ($t as $ip => $v) if (time() - $v['last'] > 86400) unset($t[$ip]);
    json_write(MO_LOGS . '/login-throttle.json', $t);
}
function throttle_reset(): void {
    $t = json_read(MO_LOGS . '/login-throttle.json', []);
    unset($t[client_ip()]);
    json_write(MO_LOGS . '/login-throttle.json', $t);
}

// ---------------------------------------------------------------- login / logout
$action = $_POST['action'] ?? '';
if ($action === 'login') {
    $wait = throttle_check();
    if ($wait > 0) $flash[] = ['err', 'Đăng nhập sai quá nhiều lần. Thử lại sau ' . ceil($wait / 60) . ' phút.'];
    elseif (csrf_ok() && hash_equals($admin['user'] ?? '', (string)($_POST['user'] ?? '')) && password_verify((string)($_POST['pass'] ?? ''), $admin['hash'] ?? '')) {
        session_regenerate_id(true);
        $_SESSION['mo_user'] = $admin['user'];
        $_SESSION['csrf'] = bin2hex(random_bytes(16));
        throttle_reset();
        audit('login');
        header('Location: ./'); exit;
    } else { throttle_fail(); audit('login-failed', (string)($_POST['user'] ?? '')); $flash[] = ['err', 'Sai tên đăng nhập hoặc mật khẩu.']; }
}
if (($_GET['logout'] ?? '') === '1') { audit('logout'); $_SESSION = []; session_destroy(); header('Location: ./'); exit; }
$logged = isset($_SESSION['mo_user']);

// ---------------------------------------------------------------- actions
if ($logged && $_SERVER['REQUEST_METHOD'] === 'POST' && $action !== 'login') {
    if (!csrf_ok()) $flash[] = ['err', 'Phiên làm việc hết hạn, vui lòng thử lại.'];
    else {
        $c = load_config();
        switch ($action) {
            case 'publish':
                $exe = $_FILES['exe'] ?? null;
                $sigFile = $_FILES['sig'] ?? null;
                $version = trim((string)($_POST['version'] ?? ''));
                $sigB64 = trim((string)($_POST['signature'] ?? ''));
                if ($sigFile && $sigFile['error'] === UPLOAD_ERR_OK) {
                    $j = json_decode((string)file_get_contents($sigFile['tmp_name']), true);
                    if (is_array($j)) { $sigB64 = $j['signature'] ?? $sigB64; $version = $version ?: ($j['version'] ?? ''); }
                }
                if (!$exe || $exe['error'] !== UPLOAD_ERR_OK) { $flash[] = ['err', 'Chưa chọn tệp cài đặt (.exe) hoặc tải lên lỗi (mã ' . ($exe['error'] ?? '-') . ').']; break; }
                [$ok, $msg] = publish_release($exe['tmp_name'], $version, $sigB64, ['vi' => (string)($_POST['notes_vi'] ?? ''), 'en' => (string)($_POST['notes_en'] ?? '')],
                    isset($_POST['latest']), isset($_POST['force']), true);
                $flash[] = [$ok ? 'ok' : 'err', $msg . ($ok && isset($_POST['force']) ? ' Đã phát lệnh cập nhật toàn bộ.' : '')];
                break;
            case 'force_on':
                $c['forceUpdate'] = ['enabled' => true, 'commandId' => bin2hex(random_bytes(8)), 'issuedAt' => date('c')];
                save_config($c); audit('force', 'on ' . $c['latest']);
                $flash[] = ['ok', 'Đã phát lệnh CẬP NHẬT TOÀN BỘ lên phiên bản ' . h($c['latest']) . '. Mọi máy đang dùng bản cũ sẽ tự tải và cài trong lần kiểm tra kế tiếp (khi mở ứng dụng hoặc tối đa 2 giờ).'];
                break;
            case 'force_off':
                $c['forceUpdate']['enabled'] = false;
                save_config($c); audit('force', 'off');
                $flash[] = ['ok', 'Đã tắt lệnh cập nhật toàn bộ. Khách tự cập nhật bằng nút "Kiểm tra cập nhật".'];
                break;
            case 'set_latest':
                $v = (string)($_POST['version'] ?? '');
                if (release_by_version($c, $v)) { $c['latest'] = $v; save_config($c); audit('set-latest', $v); $flash[] = ['ok', "Phiên bản mới nhất hiện là $v."]; }
                break;
            case 'delete_release':
                $v = (string)($_POST['version'] ?? '');
                if ($v === $c['latest']) { $flash[] = ['err', 'Không thể xóa phiên bản đang là bản mới nhất.']; break; }
                if ($r = release_by_version($c, $v)) {
                    @unlink(MO_RELEASES . '/' . $r['file']);
                    $c['releases'] = array_values(array_filter($c['releases'], fn($x) => $x['version'] !== $v));
                    save_config($c); audit('delete', $v); $flash[] = ['ok', "Đã xóa phiên bản $v."];
                }
                break;
            case 'store':
                $u = strtoupper(trim((string)($_POST['storeId'] ?? '')));
                if (preg_match('#(?:^|/)([0-9A-Z]{12,14})(?:[/?\#].*)?$#', $u, $mm)) $u = $mm[1]; // a full Store link is accepted too
                if ($u !== '' && !preg_match('/^[0-9A-Z]{12,14}$/', $u)) { $flash[] = ['err', 'Store ID gồm 12–14 ký tự (ví dụ 9NXXXXXXXXXX) — xem Partner Center › Product identity.']; break; }
                $c['storeId'] = $u; save_config($c); audit('store-id', $u);
                $flash[] = ['ok', $u === '' ? 'Đã tắt: trang tải dùng lại bộ cài .exe.' : 'Nút tải chính trên microads.vn/office giờ tải bộ cài do Microsoft ký (không cảnh báo).'];
                break;
            case 'winget':
                $w = trim((string)($_POST['wingetId'] ?? ''));
                if ($w !== '' && !preg_match('/^[A-Za-z0-9][A-Za-z0-9.\-]{2,80}$/D', $w)) { $flash[] = ['err', 'Mã gói winget dạng Microads.Office (chữ, số, dấu chấm, gạch ngang).']; break; }
                $c['wingetId'] = $w; save_config($c); audit('winget-id', $w);
                $flash[] = ['ok', $w === '' ? 'Đã ẩn lệnh winget trên trang tải.' : "Trang tải hiện lệnh: winget install --id $w -e"];
                break;
            case 'gsc':
                // Google Search Console › Add property › URL prefix https://microads.vn/office/ › HTML tag: the whole <meta> tag or only its content is accepted.
                $g = trim((string)($_POST['gscVerification'] ?? ''));
                if (preg_match('/content\s*=\s*["\']([^"\']+)["\']/i', $g, $mm)) $g = $mm[1];
                if ($g !== '' && !preg_match('/^[A-Za-z0-9_\-]{10,100}$/D', $g)) { $flash[] = ['err', 'Mã xác minh không hợp lệ — dán thẻ <meta name="google-site-verification" …> hoặc phần content của nó.']; break; }
                $c['gscVerification'] = $g; save_config($c); audit('gsc', $g);
                $flash[] = ['ok', $g === '' ? 'Đã gỡ mã xác minh Google.' : 'Đã thêm mã xác minh vào trang microads.vn/office/ — quay lại Google Search Console và bấm Xác minh.'];
                break;
            case 'promo':
                $lines = fn(string $k) => array_values(array_filter(array_map('trim', preg_split('/\R/', (string)($_POST[$k] ?? ''))), fn($x) => $x !== ''));
                $site = trim((string)($_POST['website'] ?? ''));
                if (!preg_match('#^https?://#i', $site)) { $flash[] = ['err', 'Liên kết website phải bắt đầu bằng https://']; break; }
                $c['promo'] = [
                    'enabled' => isset($_POST['enabled']), 'showOnFileOpen' => isset($_POST['showOnFileOpen']), 'showOnStart' => isset($_POST['showOnStart']),
                    'website' => $site, 'websiteLabel' => trim((string)$_POST['websiteLabel']),
                    'company' => ['vi' => trim((string)$_POST['company_vi']), 'en' => trim((string)$_POST['company_en'])],
                    'headline' => ['vi' => trim((string)$_POST['headline_vi']), 'en' => trim((string)$_POST['headline_en'])],
                    'tagline' => ['vi' => trim((string)$_POST['tagline_vi']), 'en' => trim((string)$_POST['tagline_en'])],
                    'services' => ['vi' => $lines('services_vi'), 'en' => $lines('services_en')],
                    'callToAction' => ['vi' => trim((string)$_POST['cta_vi']), 'en' => trim((string)$_POST['cta_en'])],
                    'phone' => trim((string)$_POST['phone']), 'email' => trim((string)$_POST['email']),
                ];
                save_config($c); audit('promo');
                $flash[] = ['ok', 'Đã lưu nội dung quảng bá. Ứng dụng của khách nhận nội dung mới ở lần kiểm tra kế tiếp (không cần phát hành lại).'];
                break;
            case 'seo_settings':
                $st = blog_settings();
                $st['enabled'] = isset($_POST['enabled']);
                $st['minPerDay'] = max(0, min(10, (int)$_POST['minPerDay']));
                $st['maxPerDay'] = max($st['minPerDay'], min(10, (int)$_POST['maxPerDay']));
                $st['dailyBudgetUsd'] = max(0, min(50, (float)$_POST['dailyBudgetUsd']));
                $st['toolDailyBudgetUsd'] = max(0, min($st['dailyBudgetUsd'], (float)$_POST['toolDailyBudgetUsd']));
                $st['toolDailyCap'] = max(0, min(5000, (int)$_POST['toolDailyCap']));
                $st['toolPerIp'] = max(1, min(100, (int)$_POST['toolPerIp']));
                blog_save_settings($st); audit('seo-settings');
                $flash[] = ['ok', 'Đã lưu cài đặt viết bài & SEO. Số bài mỗi ngày áp dụng từ ngày mai.'];
                break;
            case 'seo_key':
                $k = trim((string)($_POST['key'] ?? '')); $exp = trim((string)($_POST['expires'] ?? ''));
                if (!preg_match('/^sk-ant-[A-Za-z0-9_\-]{20,}$/', $k)) { $flash[] = ['err', 'Khóa API phải có dạng sk-ant-…']; break; }
                if ($exp !== '' && !preg_match('/^\d{4}-\d{2}-\d{2}$/', $exp)) { $flash[] = ['err', 'Ngày hết hạn dạng YYYY-MM-DD.']; break; }
                ai_save_secret($k, $exp); audit('seo-key', substr($k, 0, 12) . '…');
                $flash[] = ['ok', 'Đã lưu khóa Claude API mới (chỉ lưu trên máy chủ, ngoài thư mục web).'];
                break;
            case 'seo_gsc':
                $j = json_decode((string)($_POST['gscJson'] ?? ''), true);
                if (($_POST['gscJson'] ?? '') === '') { @unlink(MO_DATA . '/secrets/gsc.json'); $flash[] = ['ok', 'Đã gỡ kết nối Google Search Console.']; break; }
                if (!is_array($j) || ($j['type'] ?? '') !== 'service_account' || empty($j['client_email']) || empty($j['private_key'])) { $flash[] = ['err', 'Hãy dán nguyên nội dung tệp JSON khóa của service account (type = service_account).']; break; }
                if (!is_dir(MO_DATA . '/secrets')) mkdir(MO_DATA . '/secrets', 0750, true);
                json_write(MO_DATA . '/secrets/gsc.json', $j); @chmod(MO_DATA . '/secrets/gsc.json', 0640); audit('seo-gsc', $j['client_email']);
                $flash[] = ['ok', 'Đã lưu. Nhớ thêm ' . h($j['client_email']) . ' làm người dùng (Owner) của thuộc tính https://microads.vn/office/ trong Google Search Console.'];
                break;
            case 'seo_run':
                $what = (string)($_POST['what'] ?? '');
                $req = $what === 'audit' ? ['action' => 'audit'] : ['lang' => $what === 'en' ? 'en' : 'vi'];
                if (!is_dir(MO_DATA . '/seo')) mkdir(MO_DATA . '/seo', 0750, true);
                json_write(MO_DATA . '/seo/run-now.json', $req + ['at' => date('c')]); audit('seo-run', $what);
                $flash[] = ['ok', 'Đã xếp lịch: hệ thống chạy trong vòng 10 phút (viết một bài mất khoảng 3–8 phút).'];
                break;
            case 'seo_post':
                $pl = ($_POST['lang'] ?? '') === 'en' ? 'en' : 'vi'; $ps = (string)($_POST['slug'] ?? '');
                $post = blog_get($pl, $ps);
                if (!$post) { $flash[] = ['err', 'Không tìm thấy bài viết.']; break; }
                $op = (string)($_POST['op'] ?? '');
                if ($op === 'delete') { @unlink(blog_post_path($pl, $ps)); blog_index_rebuild(); audit('seo-post-delete', "$pl/$ps"); $flash[] = ['ok', 'Đã xóa bài viết (địa chỉ trả về 404).']; break; }
                $post['status'] = $op === 'unpublish' ? 'unpublished' : 'published';
                blog_save($post); audit('seo-post-' . $op, "$pl/$ps");
                $flash[] = ['ok', $op === 'unpublish' ? 'Đã gỡ bài khỏi website.' : 'Đã đăng lại bài.'];
                break;
            case 'password':
                $old = (string)($_POST['old'] ?? ''); $new = (string)($_POST['new'] ?? '');
                if (!password_verify($old, $admin['hash'])) { $flash[] = ['err', 'Mật khẩu hiện tại không đúng.']; break; }
                if (strlen($new) < 10) { $flash[] = ['err', 'Mật khẩu mới cần ít nhất 10 ký tự.']; break; }
                $admin['hash'] = password_hash($new, PASSWORD_DEFAULT);
                json_write(MO_ADMIN, $admin); audit('password');
                $flash[] = ['ok', 'Đã đổi mật khẩu.'];
                break;
        }
    }
}

$c = load_config();
$tab = $_GET['tab'] ?? 'overview';
$latest = release_by_version($c, $c['latest']);
?><!doctype html>
<html lang="vi">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="robots" content="noindex, nofollow">
<title>Quản trị cập nhật · Microads Office</title>
<link rel="icon" href="../assets/office.png">
<link rel="stylesheet" href="../assets/style.css">
</head>
<body class="admin">
<header class="bar">
  <a class="brand" href="./"><img src="../assets/office.png" alt="" width="28" height="28"> Microads Office <span>· Quản trị cập nhật</span></a>
  <?php if ($logged): ?><nav><a href="../" target="_blank">Trang tải ↗</a><a href="?logout=1">Đăng xuất</a></nav><?php endif; ?>
</header>
<main class="wrap">
<?php foreach ($flash as [$k, $m]): ?><div class="flash <?= $k ?>"><?= $m ?></div><?php endforeach; ?>

<?php if (!$logged): ?>
  <form method="post" class="card login" autocomplete="off">
    <h1>Đăng nhập</h1>
    <input type="hidden" name="csrf" value="<?= h(csrf()) ?>"><input type="hidden" name="action" value="login">
    <label>Tên đăng nhập<input name="user" required autofocus></label>
    <label>Mật khẩu<input name="pass" type="password" required></label>
    <button class="btn primary">Đăng nhập</button>
  </form>
<?php else: ?>
  <nav class="tabs">
    <?php foreach (['overview' => 'Tổng quan', 'publish' => 'Phát hành bản mới', 'promo' => 'Nội dung quảng bá', 'releases' => 'Lịch sử phiên bản', 'seo' => 'Bài viết & SEO', 'reports' => 'Báo cáo AI', 'account' => 'Tài khoản'] as $k => $v): ?>
      <a href="?tab=<?= $k ?>" class="<?= $tab === $k ? 'on' : '' ?>"><?= $v ?></a>
    <?php endforeach; ?>
  </nav>

  <?php $kd = ai_key_days_left(); if ($kd !== null && $kd <= 7): ?><div class="flash err">Khóa Claude API <?= $kd < 0 ? 'đã hết hạn' : "hết hạn sau $kd ngày" ?> — tạo khóa mới tại platform.claude.com › API keys rồi dán vào thẻ <a href="?tab=seo#key">Bài viết &amp; SEO</a>, nếu không hệ thống ngừng viết bài.</div><?php endif; ?>
  <?php if ($tab === 'overview'): $stats = stats_last_days(14); $today = end($stats); ?>
  <section class="grid3">
    <div class="card stat"><small>Phiên bản mới nhất</small><b><?= h($c['latest'] ?? '—') ?></b><span><?= $latest ? h($latest['date']) . ' · ' . fmt_size((int)$latest['size']) : 'Chưa phát hành' ?></span></div>
    <div class="card stat"><small>Lượt kiểm tra cập nhật hôm nay</small><b><?= array_sum($today ?: []) ?></b><span>Mỗi lần khách mở ứng dụng</span></div>
    <div class="card stat"><small>Lượt tải bản mới nhất</small><b><?= download_count($latest) ?></b><span>Qua trang microads.vn/office</span></div>
  </section>

  <?php $web = websetup_info($c); ?>
  <section class="card" id="web-setup">
    <h2>Trình cài đặt trực tuyến (nút tải chính)</h2>
    <?php if ($web): ?>
      <p><b><?= h($web['file']) ?></b> · <?= fmt_size((int)$web['size']) ?> · <b><?= (int)$web['downloads'] ?></b> lượt tải<?= $web['publishedAt'] ? ' từ ' . h(date('d/m/Y', strtotime((string)$web['publishedAt']))) : '' ?> ·
        <?= (int)(array_sum(array_map(fn($d) => (int)($d['web-setup'] ?? 0), $stats))) ?> lượt cài trong 14 ngày ·
        <?= (int)(array_sum(array_map(fn($d) => (int)($d['ps-install'] ?? 0), $stats))) ?> lượt cài bằng lệnh PowerShell</p>
      <p class="muted">SHA-256 <code class="hash"><?= h($web['sha256']) ?></code><br>Tệp này <b>không thay đổi</b> khi phát hành bản mới (nó luôn tải bản mới nhất và kiểm tra chữ ký Microads), nên danh tiếng với Chrome / Windows SmartScreen được cộng dồn qua mọi phiên bản. Đừng thay tệp này nếu không thật cần.</p>
    <?php else: ?>
      <p class="muted">Chưa có trình cài đặt trực tuyến — trang tải đang dùng bộ cài đầy đủ. Phát hành bằng <code>server/deploy/publish_websetup.py</code> (máy build).</p>
    <?php endif; ?>
  </section>

  <section class="card force <?= $c['forceUpdate']['enabled'] ? 'active' : '' ?>">
    <div>
      <h2>Cập nhật toàn bộ</h2>
      <?php if ($c['forceUpdate']['enabled']): ?>
        <p><b>Đang BẬT</b> từ <?= h(date('d/m/Y H:i', strtotime((string)$c['forceUpdate']['issuedAt']))) ?>: mọi máy dùng bản cũ hơn <b><?= h($c['latest']) ?></b> sẽ tự động tải, kiểm tra chữ ký và cài bản mới, rồi mời khách khởi động lại.</p>
      <?php else: ?>
        <p>Đang tắt: khách tự cập nhật khi bấm <b>Kiểm tra cập nhật</b> (ứng dụng chỉ hiện gợi ý có bản mới). Bấm nút bên cạnh để ra lệnh cho <b>toàn bộ</b> máy tự cập nhật lên <b><?= h($c['latest'] ?? '—') ?></b>.</p>
      <?php endif; ?>
    </div>
    <form method="post" onsubmit="return confirm('<?= $c['forceUpdate']['enabled'] ? 'Tắt lệnh cập nhật toàn bộ?' : 'Ra lệnh cho TẤT CẢ máy tự cập nhật lên ' . h($c['latest'] ?? '') . '?' ?>')">
      <input type="hidden" name="csrf" value="<?= h(csrf()) ?>">
      <?php if ($c['forceUpdate']['enabled']): ?>
        <button class="btn" name="action" value="force_off">Tắt lệnh</button>
        <button class="btn primary" name="action" value="force_on" title="Phát lại lệnh cho cả những máy đã nhận lệnh trước">Phát lại lệnh</button>
      <?php else: ?>
        <button class="btn primary big" name="action" value="force_on" <?= $latest ? '' : 'disabled' ?>>Cập nhật toàn bộ</button>
      <?php endif; ?>
    </form>
  </section>

  <section class="card">
    <h2>Máy đang kiểm tra cập nhật (14 ngày)</h2>
    <table class="tbl"><thead><tr><th>Ngày</th><th>Tổng</th><th>Theo phiên bản</th></tr></thead><tbody>
    <?php foreach (array_reverse($stats, true) as $d => $s): ?>
      <tr><td><?= h(date('d/m/Y', strtotime($d))) ?></td><td><?= array_sum($s) ?></td><td><?php foreach ($s as $v => $n) echo '<span class="pill">' . h($v) . ': ' . (int)$n . '</span> '; ?></td></tr>
    <?php endforeach; ?>
    </tbody></table>
  </section>

  <?php elseif ($tab === 'publish'): ?>
  <form method="post" enctype="multipart/form-data" class="card">
    <h2>Phát hành phiên bản mới (trọn bộ Word · Excel · PDF)</h2>
    <p class="muted">Tải lên tệp <code>MicroadsOffice_Setup_x.y.z.exe</code> và tệp chữ ký <code>.sig.json</code> do lệnh build tạo ra. Máy chủ kiểm tra chữ ký số trước khi chấp nhận — tệp không có chữ ký hợp lệ của Microads sẽ bị từ chối.</p>
    <input type="hidden" name="csrf" value="<?= h(csrf()) ?>"><input type="hidden" name="action" value="publish">
    <div class="row2">
      <label>Tệp cài đặt (.exe)<input type="file" name="exe" accept=".exe" required></label>
      <label>Tệp chữ ký (.sig.json)<input type="file" name="sig" accept=".json"></label>
    </div>
    <div class="row2">
      <label>Số phiên bản (để trống = lấy từ tệp chữ ký)<input name="version" placeholder="2.0.1" pattern="\d+\.\d+\.\d+"></label>
      <label>Hoặc dán chữ ký (base64)<input name="signature" placeholder="MEUCIQ..."></label>
    </div>
    <div class="row2">
      <label>Ghi chú phát hành (Tiếng Việt)<textarea name="notes_vi" rows="5" placeholder="- Thêm ...&#10;- Sửa lỗi ..."></textarea></label>
      <label>Release notes (English)<textarea name="notes_en" rows="5"></textarea></label>
    </div>
    <label class="chk"><input type="checkbox" name="latest" checked> Đặt làm phiên bản mới nhất</label>
    <label class="chk"><input type="checkbox" name="force"> Cập nhật toàn bộ ngay (mọi máy tự cập nhật)</label>
    <button class="btn primary big">Phát hành</button>
  </form>

  <?php elseif ($tab === 'promo'): $p = $c['promo']; ?>
  <form method="post" class="card">
    <h2>Microsoft Store</h2>
    <p class="muted">Bản trên Microsoft Store được Microsoft ký số (miễn phí) nên khách <b>không gặp cảnh báo</b> của Chrome/Windows. Sau khi ứng dụng được duyệt, nhập <b>Store ID</b> (12 ký tự, Partner Center › Product identity) — nút tải chính trên trang microads.vn/office sẽ tải <i>bộ cài web do Microsoft ký</i>; bộ cài .exe vẫn còn ở liên kết phụ. Để trống để tắt.</p>
    <input type="hidden" name="csrf" value="<?= h(csrf()) ?>"><input type="hidden" name="action" value="store">
    <label>Store ID<input name="storeId" value="<?= h($c['storeId'] ?? '') ?>" placeholder="9NXXXXXXXXXX" maxlength="80"></label>
    <button class="btn primary">Lưu liên kết</button>
  </form>
  <form method="post" class="card" id="gsc">
    <h2>Google Search Console (Chrome)</h2>
    <p class="muted">Để Google xem xét cảnh báo <i>"tệp ít được tải xuống"</i> của Chrome: vào <b>search.google.com/search-console</b> › Thêm thuộc tính › <b>Tiền tố URL</b> <code>https://microads.vn/office/</code> › phương thức <b>Thẻ HTML</b> › dán thẻ vào đây › Lưu › quay lại bấm <b>Xác minh</b>. Sau đó mở <b>Bảo mật và thao tác thủ công › Vấn đề bảo mật</b>; nếu có mục <i>Tệp tải xuống không phổ biến</i> thì bấm <b>Yêu cầu xem xét</b> (miễn phí, vài ngày đến vài tuần).</p>
    <input type="hidden" name="csrf" value="<?= h(csrf()) ?>"><input type="hidden" name="action" value="gsc">
    <label>Thẻ / mã xác minh<input name="gscVerification" value="<?= h($c['gscVerification'] ?? '') ?>" placeholder='&lt;meta name="google-site-verification" content="..."&gt;' maxlength="300"></label>
    <button class="btn primary">Lưu</button>
  </form>
  <form method="post" class="card" id="winget">
    <h2>Windows Package Manager (winget)</h2>
    <p class="muted">Kho gói cộng đồng của Microsoft (miễn phí, không phải Microsoft Store). Khi gói Microads Office được duyệt trên github.com/microsoft/winget-pkgs (hồ sơ có sẵn trong <code>tools/winget</code>), nhập mã gói để trang tải hiện lệnh <code>winget install</code> — cài không có cảnh báo của trình duyệt/SmartScreen. Để trống để ẩn.</p>
    <input type="hidden" name="csrf" value="<?= h(csrf()) ?>"><input type="hidden" name="action" value="winget">
    <label>Mã gói winget<input name="wingetId" value="<?= h($c['wingetId'] ?? '') ?>" placeholder="Microads.Office" maxlength="80"></label>
    <button class="btn primary">Lưu</button>
  </form>
  <form method="post" class="card">
    <h2>Nội dung quảng bá trong ứng dụng</h2>
    <p class="muted">Hiển thị ở thông báo khi khách mở tệp, trang Giới thiệu, trang bắt đầu. Thay đổi có hiệu lực ngay ở lần kiểm tra kế tiếp — không cần phát hành lại.</p>
    <input type="hidden" name="csrf" value="<?= h(csrf()) ?>"><input type="hidden" name="action" value="promo">
    <label class="chk"><input type="checkbox" name="enabled" <?= $p['enabled'] ? 'checked' : '' ?>> Bật nội dung quảng bá</label>
    <label class="chk"><input type="checkbox" name="showOnFileOpen" <?= $p['showOnFileOpen'] ? 'checked' : '' ?>> Hiện thông báo khi khách mở tệp bằng Word / Excel / PDF</label>
    <label class="chk"><input type="checkbox" name="showOnStart" <?= ($p['showOnStart'] ?? true) ? 'checked' : '' ?>> Hiện thông báo mỗi khi khách mở Word / Excel / PDF (cần bản 2.0.2 trở lên)</label>
    <div class="row2">
      <label>Liên kết website<input name="website" value="<?= h($p['website']) ?>" required></label>
      <label>Chữ hiển thị của liên kết<input name="websiteLabel" value="<?= h($p['websiteLabel']) ?>" required></label>
    </div>
    <?php foreach (['company' => 'Tên công ty', 'headline' => 'Tiêu đề thông báo', 'tagline' => 'Dòng giới thiệu dịch vụ', 'callToAction' => 'Chữ trên nút'] as $k => $label): $kk = $k === 'callToAction' ? 'cta' : $k; ?>
    <div class="row2">
      <label><?= $label ?> (VI)<input name="<?= $kk ?>_vi" value="<?= h($p[$k]['vi'] ?? '') ?>"></label>
      <label><?= $label ?> (EN)<input name="<?= $kk ?>_en" value="<?= h($p[$k]['en'] ?? '') ?>"></label>
    </div>
    <?php endforeach; ?>
    <div class="row2">
      <label>Danh sách dịch vụ (VI) — mỗi dòng một mục<textarea name="services_vi" rows="6"><?= h(implode("\n", $p['services']['vi'] ?? [])) ?></textarea></label>
      <label>Services (EN) — one per line<textarea name="services_en" rows="6"><?= h(implode("\n", $p['services']['en'] ?? [])) ?></textarea></label>
    </div>
    <div class="row2">
      <label>Điện thoại<input name="phone" value="<?= h($p['phone']) ?>"></label>
      <label>Email<input name="email" value="<?= h($p['email']) ?>"></label>
    </div>
    <button class="btn primary">Lưu nội dung</button>
  </form>

  <?php elseif ($tab === 'releases'): ?>
  <section class="card">
    <h2>Lịch sử phiên bản</h2>
    <table class="tbl"><thead><tr><th>Phiên bản</th><th>Ngày</th><th>Dung lượng</th><th>Lượt tải</th><th>SHA-256</th><th></th></tr></thead><tbody>
    <?php foreach ($c['releases'] as $r): ?>
      <tr>
        <td><b><?= h($r['version']) ?></b><?= $r['version'] === $c['latest'] ? ' <span class="pill ok">mới nhất</span>' : '' ?></td>
        <td><?= h(date('d/m/Y', strtotime($r['date']))) ?></td><td><?= fmt_size((int)$r['size']) ?></td><td><?= download_count($r) ?></td>
        <td><code title="<?= h($r['sha256']) ?>"><?= h(substr($r['sha256'], 0, 16)) ?>…</code></td>
        <td class="actions">
          <a class="btn sm" href="../download/<?= h($r['file']) ?>">Tải</a>
          <?php if ($r['version'] !== $c['latest']): ?>
          <form method="post"><input type="hidden" name="csrf" value="<?= h(csrf()) ?>"><input type="hidden" name="version" value="<?= h($r['version']) ?>">
            <button class="btn sm" name="action" value="set_latest">Đặt làm mới nhất</button>
            <button class="btn sm danger" name="action" value="delete_release" onclick="return confirm('Xóa phiên bản <?= h($r['version']) ?>?')">Xóa</button></form>
          <?php endif; ?>
        </td>
      </tr>
    <?php endforeach; ?>
    <?php if (!$c['releases']): ?><tr><td colspan="6" class="muted">Chưa có phiên bản nào.</td></tr><?php endif; ?>
    </tbody></table>
  </section>

  <?php elseif ($tab === 'reports'):
    $reports = [];
    foreach (array_reverse(glob(MO_DATA . '/reports/*.jsonl') ?: []) as $f) {
        foreach (array_reverse(file($f, FILE_IGNORE_NEW_LINES | FILE_SKIP_EMPTY_LINES) ?: []) as $l) { $r = json_decode($l, true); if (is_array($r)) $reports[] = $r; if (count($reports) >= 200) break 2; }
    }
    $reasonVi = ['harmful' => 'Có hại / nguy hiểm', 'offensive' => 'Xúc phạm / không phù hợp', 'wrong' => 'Sai sự thật', 'privacy' => 'Lộ thông tin riêng tư', 'other' => 'Khác'];
  ?>
  <section class="card">
    <h2>Báo cáo nội dung AI (<?= count($reports) ?>)</h2>
    <p class="muted">Người dùng bấm <b>Báo cáo</b> dưới câu trả lời của Trợ lý AI (yêu cầu của Microsoft Store, chính sách 11.16). Xem và xử lý: nếu lặp lại theo một kiểu, bổ sung hướng dẫn an toàn cho trợ lý trong phiên bản sau. Không lưu địa chỉ IP.</p>
    <table><thead><tr><th>Thời gian</th><th>Ứng dụng</th><th>Lý do</th><th>Nội dung</th></tr></thead><tbody>
    <?php foreach ($reports as $r): ?>
      <tr>
        <td><?= h(substr($r['time'] ?? '', 0, 16)) ?><br><span class="muted"><?= h($r['id'] ?? '') ?></span></td>
        <td><?= h(($r['app'] ?? '') . ' ' . ($r['version'] ?? '')) ?><br><span class="muted"><?= h(($r['provider'] ?? '') . ' · ' . ($r['model'] ?? '')) ?></span></td>
        <td><?= h($reasonVi[$r['reason'] ?? 'other'] ?? 'Khác') ?></td>
        <td><?php if (($r['comment'] ?? '') !== ''): ?><b><?= h($r['comment']) ?></b><br><?php endif; ?>
          <details><summary>Câu hỏi và câu trả lời</summary><p class="muted"><?= nl2br(h($r['question'] ?? '')) ?></p><hr><p><?= nl2br(h($r['answer'] ?? '')) ?></p></details></td>
      </tr>
    <?php endforeach; ?>
    <?php if (!$reports): ?><tr><td colspan="4" class="muted">Chưa có báo cáo nào.</td></tr><?php endif; ?>
    </tbody></table>
  </section>

  <?php elseif ($tab === 'seo'):
    $st = blog_settings();
    $all = blog_index_rebuild();
    $sched = json_read(MO_DATA . '/blog/schedule.json', []);
    $repFiles = array_reverse(glob(MO_DATA . '/seo/reports/*.json') ?: []);
    $rep = $repFiles ? json_read($repFiles[0], []) : [];
    $logF = MO_DATA . '/seo/logs/' . date('Y-m') . '.log'; $logLines = is_file($logF) ? array_slice(file($logF, FILE_IGNORE_NEW_LINES | FILE_SKIP_EMPTY_LINES) ?: [], -40) : [];
    $sec = ai_secret();
    $topics = json_read(MO_DATA . '/blog/topics.json', []);
  ?>
  <section class="grid3">
    <div class="card stat"><small>Bài đã đăng</small><b><?= count(array_filter($all, fn($p) => ($p['status'] ?? '') === 'published')) ?></b><span>VI <?= count(array_filter($all, fn($p) => $p['lang'] === 'vi')) ?> · EN <?= count(array_filter($all, fn($p) => $p['lang'] === 'en')) ?></span></div>
    <div class="card stat"><small>Chi phí Claude tháng này</small><b>$<?= number_format(ai_spend(date('Y-m')), 2) ?></b><span>Hôm nay $<?= number_format(ai_spend(date('Y-m-d')), 2) ?> / $<?= number_format((float)$st['dailyBudgetUsd'], 2) ?> · trung bình $<?= number_format(ai_article_cost(), 2) ?>/bài</span></div>
    <div class="card stat"><small>Kế hoạch hôm nay</small><b><?= count($sched['slots'] ?? []) ?> bài</b><span><?php foreach ($sched['slots'] ?? [] as $sl): ?><?= h(strtoupper($sl['lang']) . ' ' . substr($sl['at'], 11, 5)) ?><?= $sl['done'] ? ' ✓' : '' ?> · <?php endforeach; ?></span></div>
  </section>

  <?php if ($rep): $an = $rep['analysis'] ?? []; ?>
  <section class="card">
    <h2>Báo cáo SEO ngày <?= h($rep['date']) ?> <span class="pill <?= ($an['health'] ?? '') === 'good' ? 'ok' : '' ?>"><?= h(['good' => 'Tốt', 'warning' => 'Cần chú ý', 'bad' => 'Có vấn đề'][$an['health'] ?? ''] ?? '—') ?></span></h2>
    <p><?= nl2br(h($an['summaryVi'] ?? '')) ?></p>
    <?php if (!empty($an['issues'])): ?><table class="tbl"><thead><tr><th>Mức</th><th>Trang</th><th>Vấn đề</th><th>Cách xử lý</th></tr></thead><tbody>
      <?php foreach ($an['issues'] as $is): ?><tr><td><?= h(['high' => 'Cao', 'medium' => 'Vừa', 'low' => 'Thấp'][$is['severity']] ?? '') ?></td><td><?= h(str_replace('https://microads.vn', '', $is['url'])) ?></td><td><?= h($is['problemVi']) ?></td><td><?= h($is['fixVi']) ?></td></tr><?php endforeach; ?>
    </tbody></table><?php endif; ?>
    <?php if (!empty($rep['actions'])): ?><p class="muted">Đã tự xử lý: <?php foreach ($rep['actions'] as $ac): ?><?= h($ac['type'] . ' ' . $ac['lang'] . '/' . $ac['slug'] . ($ac['ok'] ? ' ✓' : ' ✗')) ?>; <?php endforeach; ?></p><?php endif; ?>
    <?php if (!empty($rep['performance'])): ?><table class="tbl"><thead><tr><th>Tốc độ (điện thoại)</th><th>Hiệu năng</th><th>SEO</th><th>LCP</th><th>CLS</th></tr></thead><tbody>
      <?php foreach ($rep['performance'] as $pf): ?><tr><td><?= h(str_replace('https://microads.vn', '', $pf['url'])) ?></td><?php if (isset($pf['error'])): ?><td colspan="4" class="muted"><?= h($pf['error']) ?></td><?php else: ?><td><?= (int)$pf['performance'] ?></td><td><?= (int)$pf['seo'] ?></td><td><?= number_format(($pf['lcpMs'] ?? 0) / 1000, 1) ?> s</td><td><?= number_format((float)($pf['cls'] ?? 0), 3) ?></td><?php endif; ?></tr><?php endforeach; ?>
    </tbody></table><?php endif; ?>
    <p class="muted">Hôm qua: <?= (int)($rep['traffic']['views'] ?? 0) ?> lượt xem · <?= array_sum($rep['traffic']['downloadsSinceLastAudit'] ?? []) ?> lượt tải ·
      hoạt động <?= isset($rep['uptime']['checks']) ? (int)$rep['uptime']['checks'] - (int)$rep['uptime']['fails'] . '/' . (int)$rep['uptime']['checks'] . ' lần kiểm tra đạt' : '—' ?> ·
      SSL còn <?= h((string)($rep['security']['certDaysLeft'] ?? '—')) ?> ngày</p>
    <p class="muted">Google Search Console: <?= $rep['gsc'] === null ? 'chưa kết nối (xem hướng dẫn bên dưới)' : (isset($rep['gsc']['error']) ? 'lỗi — ' . h($rep['gsc']['error']) : count($rep['gsc']['inspections'] ?? []) . ' trang đã kiểm tra index') ?> ·
      Sitemap: <?= (int)($rep['site']['sitemapUrls'] ?? 0) ?> địa chỉ · Link hỏng: <?= count($rep['site']['brokenLinks'] ?? []) ?></p>
    <?php if (!empty($rep['gsc']['inspections'])): ?><details><summary>Trạng thái index từng trang</summary><table class="tbl"><thead><tr><th>Trang</th><th>Kết luận</th><th>Trạng thái Google</th><th>Lần thu thập</th></tr></thead><tbody>
      <?php foreach ($rep['gsc']['inspections'] as $u => $in): ?><tr><td><?= h(str_replace('https://microads.vn', '', $u)) ?></td><td><?= h($in['verdict']) ?></td><td><?= h($in['coverage']) ?></td><td><?= h(substr($in['lastCrawl'], 0, 10)) ?></td></tr><?php endforeach; ?>
    </tbody></table></details><?php endif; ?>
  </section>
  <?php endif; ?>

  <section class="card">
    <h2>Bài viết (<?= count($all) ?>)</h2>
    <form method="post" class="actions"><input type="hidden" name="csrf" value="<?= h(csrf()) ?>"><input type="hidden" name="action" value="seo_run">
      <button class="btn" name="what" value="vi">Viết 1 bài tiếng Việt ngay</button> <button class="btn" name="what" value="en">Viết 1 bài tiếng Anh ngay</button> <button class="btn" name="what" value="audit">Chạy rà soát SEO ngay</button></form>
    <table class="tbl"><thead><tr><th>Ngày</th><th>Bài</th><th>Điểm</th><th>Chi phí</th><th></th></tr></thead><tbody>
    <?php foreach ($all as $bp): ?><tr>
      <td><?= h(substr($bp['publishedAt'], 0, 10)) ?><br><span class="pill"><?= h(strtoupper($bp['lang'])) ?></span></td>
      <td><a href="<?= h(blog_base($bp['lang']) . $bp['slug']) ?>" target="_blank"><?= h($bp['title']) ?></a><br><span class="muted"><?= h(($bp['pillar'] ?? '') . ' · ' . ($bp['format'] ?? '') . ' · ' . ($bp['wordCount'] ?? 0) . ' từ') ?><?= ($bp['status'] ?? '') !== 'published' ? ' · ĐÃ GỠ' : '' ?></span></td>
      <td><?= (int)($bp['review']['score'] ?? 0) ?></td><td>$<?= number_format((float)($bp['costUsd'] ?? 0), 2) ?></td>
      <td><form method="post" class="actions"><input type="hidden" name="csrf" value="<?= h(csrf()) ?>"><input type="hidden" name="action" value="seo_post"><input type="hidden" name="lang" value="<?= h($bp['lang']) ?>"><input type="hidden" name="slug" value="<?= h($bp['slug']) ?>">
        <?php if (($bp['status'] ?? '') === 'published'): ?><button class="btn sm" name="op" value="unpublish">Gỡ</button><?php else: ?><button class="btn sm" name="op" value="publish">Đăng lại</button><?php endif; ?>
        <button class="btn sm danger" name="op" value="delete" onclick="return confirm('Xóa hẳn bài này?')">Xóa</button></form></td>
    </tr><?php endforeach; ?>
    <?php if (!$all): ?><tr><td colspan="5" class="muted">Chưa có bài nào.</td></tr><?php endif; ?>
    </tbody></table>
    <p class="muted">Chủ đề đang chờ: VI <?= count($topics['vi'] ?? []) ?> · EN <?= count($topics['en'] ?? []) ?> (tự nghiên cứu thêm khi còn dưới 4).</p>
  </section>

  <form method="post" class="card">
    <h2>Cài đặt viết bài</h2>
    <input type="hidden" name="csrf" value="<?= h(csrf()) ?>"><input type="hidden" name="action" value="seo_settings">
    <label class="chk"><input type="checkbox" name="enabled" <?= $st['enabled'] ? 'checked' : '' ?>> Tự động viết và đăng bài mỗi ngày</label>
    <div class="row2"><label>Ít nhất (bài/ngày)<input type="number" name="minPerDay" min="0" max="10" value="<?= (int)$st['minPerDay'] ?>"></label><label>Nhiều nhất (bài/ngày)<input type="number" name="maxPerDay" min="0" max="10" value="<?= (int)$st['maxPerDay'] ?>"></label></div>
    <div class="row2"><label>Ngân sách Claude mỗi ngày (USD, gồm mọi việc)<input type="number" name="dailyBudgetUsd" min="0" step="0.1" value="<?= h((string)$st['dailyBudgetUsd']) ?>"></label><label>Trong đó cho công cụ tạo công thức (USD/ngày)<input type="number" name="toolDailyBudgetUsd" min="0" step="0.05" value="<?= h((string)$st['toolDailyBudgetUsd']) ?>"></label></div>
    <label>Công cụ tạo công thức: tối đa lượt/ngày (toàn trang)<input type="number" name="toolDailyCap" min="0" value="<?= (int)$st['toolDailyCap'] ?>"></label>
    <label>Công cụ tạo công thức: lượt/ngày cho mỗi người<input type="number" name="toolPerIp" min="1" max="100" value="<?= (int)$st['toolPerIp'] ?>"></label>
    <button class="btn primary">Lưu</button>
  </form>

  <form method="post" class="card" id="key">
    <h2>Khóa Claude API</h2>
    <p class="muted">Đang dùng: <code><?= h($sec['key'] !== '' ? substr($sec['key'], 0, 14) . '…' . substr($sec['key'], -4) : 'chưa có') ?></code> · hết hạn <?= h($sec['expires'] ?: '—') ?> (mô hình claude-opus-5-5). Khóa chỉ lưu trên máy chủ, ngoài thư mục web, không nằm trong mã nguồn.</p>
    <input type="hidden" name="csrf" value="<?= h(csrf()) ?>"><input type="hidden" name="action" value="seo_key">
    <div class="row2"><label>Khóa mới<input name="key" autocomplete="off" placeholder="sk-ant-…"></label><label>Ngày hết hạn (YYYY-MM-DD)<input name="expires" placeholder="2027-01-31"></label></div>
    <button class="btn primary">Lưu khóa mới</button>
  </form>

  <form method="post" class="card">
    <h2>Kết nối Google Search Console (để biết trang nào chưa được index và vì sao)</h2>
    <ol class="muted">
      <li>Vào console.cloud.google.com → tạo dự án → bật <b>Google Search Console API</b>.</li>
      <li>IAM &amp; Admin → Service accounts → tạo tài khoản → Keys → Add key → JSON → tải tệp về.</li>
      <li>Search Console → thuộc tính <b>https://microads.vn/office/</b> → Cài đặt → Người dùng và quyền → Thêm người dùng: email của service account, quyền <b>Chủ sở hữu</b>.</li>
      <li>Dán nội dung tệp JSON vào ô dưới và lưu.</li>
    </ol>
    <p class="muted">Trạng thái: <?= is_file(MO_DATA . '/secrets/gsc.json') ? 'đã kết nối (' . h(json_read(MO_DATA . '/secrets/gsc.json', [])['client_email'] ?? '') . ')' : 'chưa kết nối' ?></p>
    <input type="hidden" name="csrf" value="<?= h(csrf()) ?>"><input type="hidden" name="action" value="seo_gsc">
    <label>Tệp JSON của service account (để trống rồi lưu = gỡ kết nối)<textarea name="gscJson" rows="4" autocomplete="off"></textarea></label>
    <button class="btn primary">Lưu</button>
  </form>

  <section class="card"><h2>Nhật ký</h2><pre class="log"><?= h(implode("\n", $logLines)) ?></pre></section>

  <?php elseif ($tab === 'account'): ?>
  <form method="post" class="card narrow">
    <h2>Đổi mật khẩu quản trị</h2>
    <input type="hidden" name="csrf" value="<?= h(csrf()) ?>"><input type="hidden" name="action" value="password">
    <label>Mật khẩu hiện tại<input type="password" name="old" required></label>
    <label>Mật khẩu mới (≥ 10 ký tự)<input type="password" name="new" minlength="10" required></label>
    <button class="btn primary">Đổi mật khẩu</button>
  </form>
  <?php endif; ?>
<?php endif; ?>
</main>
</body>
</html>
