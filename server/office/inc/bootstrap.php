<?php
// Microads Office update service — shared helpers.
// Code: /var/www/microads.vn/office   (nginx denies /office/inc/)
// Data: /var/www/microads-office-data (outside the web root)
declare(strict_types=1);
date_default_timezone_set('Asia/Ho_Chi_Minh');

// MO_DATA_DIR / MO_BASE_URL are only set by the local test server (server/dev); production uses the defaults.
define('MO_DATA', getenv('MO_DATA_DIR') ?: '/var/www/microads-office-data');
define('MO_RELEASES', MO_DATA . '/releases');
define('MO_CONFIG', MO_DATA . '/config.json');
define('MO_ADMIN', MO_DATA . '/admin.json');
define('MO_STATS', MO_DATA . '/stats');
define('MO_LOGS', MO_DATA . '/logs');
define('MO_BASE_URL', getenv('MO_BASE_URL') ?: 'https://microads.vn/office');
// Web installer: a small program that NEVER changes (its SHA-256 keeps the browser / SmartScreen reputation it earns)
// and always installs the latest signed release. Stored next to the releases.
const MO_WEBSETUP = 'MicroadsOffice_Setup.exe';
// Public half of the release-signing key (ECDSA P-256). Must match Product.UpdatePublicKey in the app.
const MO_PUBKEY_B64 = 'MFkwEwYHKoZIzj0CAQYIKoZIzj0DAQcDQgAEXSjzXUOQ/LgAupMzJCQNytd+BZaRwafJixpx9FX1LW4jyHov1tOAxAi86oetJjZFXeVbutcMCa52lXxcX42CXw==';

function h(?string $s): string { return htmlspecialchars((string)$s, ENT_QUOTES | ENT_SUBSTITUTE, 'UTF-8'); }

function json_read(string $path, $default) {
    if (!is_file($path)) return $default;
    $fp = fopen($path, 'rb');
    if (!$fp) return $default;
    flock($fp, LOCK_SH);
    $raw = stream_get_contents($fp);
    flock($fp, LOCK_UN);
    fclose($fp);
    $v = json_decode((string)$raw, true);
    return is_array($v) ? $v : $default;
}

/** Atomic write: the old file is only replaced when the new content was written completely (full disk → old file kept). */
function json_write(string $path, $data): void {
    $json = json_encode($data, JSON_PRETTY_PRINT | JSON_UNESCAPED_UNICODE | JSON_UNESCAPED_SLASHES);
    if ($json === false) throw new RuntimeException('json_encode failed for ' . basename($path));
    $tmp = $path . '.' . getmypid() . '.tmp';
    $n = @file_put_contents($tmp, $json, LOCK_EX);
    if ($n !== strlen($json) || !@rename($tmp, $path)) {
        @unlink($tmp);
        throw new RuntimeException('could not write ' . basename($path));
    }
}

/**
 * Download counters live in their own file (counters.json), updated under an exclusive lock, so a download never
 * rewrites config.json (a download running while a release is published could otherwise put the old config back).
 */
function bump_counter(string $key): void {
    $fp = @fopen(MO_DATA . '/counters.json', 'c+');
    if (!$fp) return;
    flock($fp, LOCK_EX);
    $d = json_decode((string)stream_get_contents($fp), true) ?: [];
    $d[$key] = (int)($d[$key] ?? 0) + 1;
    $json = json_encode($d);
    rewind($fp);
    if (fwrite($fp, $json) === strlen($json)) { fflush($fp); ftruncate($fp, strlen($json)); }
    flock($fp, LOCK_UN);
    fclose($fp);
}

function counters(): array { static $c = null; return $c ??= json_read(MO_DATA . '/counters.json', []); }

/** Downloads of a release file: older counts kept in config.json + the counter file. */
function download_count(?array $r): int { return $r ? (int)($r['downloads'] ?? 0) + (int)(counters()[$r['file']] ?? 0) : 0; }

function default_promo(): array {
    return [
        'enabled' => true,
        'showOnFileOpen' => true,
        'showOnStart' => true,
        'website' => 'https://microads.vn',
        'websiteLabel' => 'microads.vn',
        'company' => ['vi' => 'Công ty TNHH Công nghệ Microads', 'en' => 'Microads Technology Co., Ltd.'],
        'headline' => [
            'vi' => 'Microads Office — phát hành bởi Công ty TNHH Công nghệ Microads, luôn luôn miễn phí',
            'en' => 'Microads Office — published by Microads Technology Co., Ltd., always free'],
        'tagline' => ['vi' => 'Chuyên các giải pháp công nghệ cho doanh nghiệp', 'en' => 'Technology solutions for businesses'],
        'services' => [
            'vi' => ['Lập trình phần mềm doanh nghiệp theo yêu cầu', 'Bảo mật hệ thống, an toàn thông tin', 'Giải pháp CRM quản lý khách hàng', 'Hệ thống ERP quản trị doanh nghiệp', 'Thiết kế website chuyên nghiệp, chuẩn SEO'],
            'en' => ['Custom enterprise software development', 'System security & information safety', 'CRM customer management solutions', 'ERP business management systems', 'Professional, SEO-ready website design']],
        'callToAction' => ['vi' => 'Truy cập microads.vn', 'en' => 'Visit microads.vn'],
        'phone' => '0919 788 815',
        'email' => '',
    ];
}

function load_config(): array {
    $c = json_read(MO_CONFIG, []);
    $c += ['latest' => null, 'releases' => [], 'forceUpdate' => ['enabled' => false, 'commandId' => '', 'issuedAt' => null], 'promo' => default_promo(), 'storeId' => '',
        'webSetup' => ['sha256' => '', 'size' => 0, 'publishedAt' => null], 'wingetId' => '', 'gscVerification' => ''];
    $c['promo'] = array_replace(default_promo(), $c['promo'] ?? []);
    return $c;
}

function save_config(array $c): void { json_write(MO_CONFIG, $c); }

function release_by_version(array $c, ?string $v): ?array {
    foreach ($c['releases'] as $r) if ($r['version'] === $v) return $r;
    return null;
}

function public_release(array $r): array {
    return [
        'version' => $r['version'],
        'date' => $r['date'],
        'url' => MO_BASE_URL . '/download/' . rawurlencode($r['file']),
        'size' => (int)$r['size'],
        'sha256' => $r['sha256'],
        'signature' => $r['signature'],
        'notes' => ['vi' => $r['notes']['vi'] ?? '', 'en' => $r['notes']['en'] ?? ''],
    ];
}

/** Manifest consumed by every Microads Office installation. */
function build_manifest(array $c): array {
    $p = $c['promo'];
    $latest = release_by_version($c, $c['latest']);
    return [
        'schema' => 1,
        'latest' => $latest ? public_release($latest) : null,
        'forceUpdate' => ['enabled' => (bool)$c['forceUpdate']['enabled'], 'commandId' => (string)$c['forceUpdate']['commandId']],
        'promo' => [
            'enabled' => (bool)$p['enabled'],
            'showOnFileOpen' => (bool)$p['showOnFileOpen'],
            'showOnStart' => (bool)($p['showOnStart'] ?? true),
            'website' => $p['website'],
            'websiteLabel' => $p['websiteLabel'],
            'company' => $p['company'],
            'headline' => $p['headline'],
            'tagline' => $p['tagline'],
            'services' => $p['services'],
            'callToAction' => $p['callToAction'],
            'phone' => $p['phone'],
            'email' => $p['email'],
        ],
    ];
}

function pubkey_pem(): string {
    return "-----BEGIN PUBLIC KEY-----\n" . chunk_split(MO_PUBKEY_B64, 64, "\n") . "-----END PUBLIC KEY-----\n";
}

/** Verifies the ECDSA signature produced by the build machine for "MicroadsOffice|version|sha256". */
function verify_release_signature(string $version, string $sha256, string $sigB64): bool {
    $sig = base64_decode($sigB64, true);
    if ($sig === false) return false;
    $ok = openssl_verify("MicroadsOffice|$version|$sha256", $sig, pubkey_pem(), OPENSSL_ALGO_SHA256);
    return $ok === 1;
}

function valid_version(string $v): bool { return (bool)preg_match('/^[0-9]{1,4}\.[0-9]{1,4}\.[0-9]{1,6}$/D', $v); }

/**
 * Registers a release from a file already on disk (upload or CLI). Returns [ok, message].
 */
function publish_release(string $srcPath, string $version, string $sigB64, array $notes, bool $makeLatest, bool $force, bool $move): array {
    if (!valid_version($version)) return [false, 'Số phiên bản không hợp lệ (dạng 2.0.1).'];
    if (!is_file($srcPath)) return [false, 'Không tìm thấy tệp cài đặt.'];
    $head = file_get_contents($srcPath, false, null, 0, 2);
    if ($head !== 'MZ') return [false, 'Tệp không phải chương trình Windows (.exe).'];
    $sha = hash_file('sha256', $srcPath);
    if (!verify_release_signature($version, $sha, trim($sigB64)))
        return [false, 'Chữ ký phát hành KHÔNG hợp lệ — tệp không do máy build của Microads ký, hoặc sai số phiên bản. Đã từ chối.'];
    if (!is_dir(MO_RELEASES)) mkdir(MO_RELEASES, 0750, true);
    $file = "MicroadsOffice_Setup_$version.exe";
    $dest = MO_RELEASES . '/' . $file;
    if ($move) { if (!rename($srcPath, $dest)) return [false, 'Không lưu được tệp.']; }
    else { if (!copy($srcPath, $dest)) return [false, 'Không lưu được tệp.']; }
    chmod($dest, 0640);
    $c = load_config();
    $c['releases'] = array_values(array_filter($c['releases'], fn($r) => $r['version'] !== $version));
    $c['releases'][] = [
        'version' => $version, 'file' => $file, 'size' => filesize($dest), 'sha256' => $sha, 'signature' => trim($sigB64),
        'date' => date('Y-m-d'), 'publishedAt' => date('c'), 'notes' => ['vi' => trim($notes['vi'] ?? ''), 'en' => trim($notes['en'] ?? '')], 'downloads' => 0,
    ];
    usort($c['releases'], fn($a, $b) => version_compare($b['version'], $a['version']));
    if ($makeLatest || $c['latest'] === null) $c['latest'] = $version;
    if ($force) $c['forceUpdate'] = ['enabled' => true, 'commandId' => bin2hex(random_bytes(8)), 'issuedAt' => date('c')];
    save_config($c);
    audit('publish', "$version force=" . ($force ? 1 : 0));
    return [true, "Đã phát hành phiên bản $version."];
}

function audit(string $action, string $detail = ''): void {
    if (!is_dir(MO_LOGS)) @mkdir(MO_LOGS, 0750, true);
    $who = $_SESSION['mo_user'] ?? (PHP_SAPI === 'cli' ? 'cli' : '-');
    $ip = $_SERVER['REMOTE_ADDR'] ?? 'local';
    @file_put_contents(MO_LOGS . '/audit.log', date('c') . "\t$ip\t$who\t$action\t$detail\n", FILE_APPEND | LOCK_EX);
}

/** Counts update checks per day and per installed version (anonymous; no IP stored). */
function record_check(string $version): void {
    if (!is_dir(MO_STATS)) @mkdir(MO_STATS, 0750, true);
    // "web-setup" = the web installer asking for the latest version, "ps-install" = the PowerShell install command
    $version = valid_version($version) || in_array($version, ['web-setup', 'ps-install'], true) ? $version : 'unknown';
    $path = MO_STATS . '/' . date('Y-m-d') . '.json';
    $fp = @fopen($path, 'c+');
    if (!$fp) return;
    flock($fp, LOCK_EX);
    $data = json_decode((string)stream_get_contents($fp), true) ?: [];
    $data[$version] = ($data[$version] ?? 0) + 1;
    ftruncate($fp, 0); rewind($fp);
    fwrite($fp, json_encode($data));
    flock($fp, LOCK_UN);
    fclose($fp);
}

function stats_last_days(int $days): array {
    $out = [];
    for ($i = $days - 1; $i >= 0; $i--) {
        $d = date('Y-m-d', strtotime("-$i day"));
        $out[$d] = json_read(MO_STATS . "/$d.json", []);
    }
    return $out;
}

function lang(): string {
    $l = $_GET['lang'] ?? ($_COOKIE['mo_lang'] ?? '');
    if (!in_array($l, ['vi', 'en'], true)) {
        $al = strtolower($_SERVER['HTTP_ACCEPT_LANGUAGE'] ?? 'vi');
        $l = str_starts_with($al, 'vi') ? 'vi' : (str_contains($al, 'vi') ? 'vi' : 'en');
    }
    if (isset($_GET['lang']) && PHP_SAPI !== 'cli') setcookie('mo_lang', $l, ['expires' => time() + 31536000, 'path' => '/office', 'secure' => true, 'samesite' => 'Lax']);
    return $l;
}

function t(string $vi, string $en): string { global $LANG; return ($LANG ?? 'vi') === 'en' ? $en : $vi; }

/** Store links built from the 12-character Store ID (Partner Center > Product identity). */
function store_links(array $c): ?array {
    $id = (string)($c['storeId'] ?? '');
    if (!preg_match('/^[0-9A-Z]{12,14}$/', $id)) return null;
    return [
        // Microsoft-signed "Store Web Installer": no browser / SmartScreen warnings, installs without opening the Store.
        'installer' => 'https://get.microsoft.com/installer/download/' . $id . '?cid=microads_office_web',
        'page' => 'https://apps.microsoft.com/detail/' . $id,
    ];
}

/** The published web installer, or null when none was uploaded yet. */
function websetup_info(array $c): ?array {
    $path = MO_RELEASES . '/' . MO_WEBSETUP;
    if (!is_file($path)) return null;
    $w = $c['webSetup'] ?? [];
    return ['file' => MO_WEBSETUP, 'url' => MO_BASE_URL . '/download/' . MO_WEBSETUP, 'size' => filesize($path),
        'sha256' => ($w['sha256'] ?? '') ?: hash_file('sha256', $path), 'downloads' => (int)(counters()[MO_WEBSETUP] ?? 0), 'publishedAt' => $w['publishedAt'] ?? null];
}

/**
 * Publishes the web installer. Its reputation with Chrome / SmartScreen is tied to its SHA-256, so an existing,
 * different web installer is only replaced with $replace (a new file starts again from zero reputation).
 */
function publish_websetup(string $srcPath, bool $replace): array {
    if (!is_file($srcPath)) return [false, 'Không tìm thấy tệp.'];
    if (file_get_contents($srcPath, false, null, 0, 2) !== 'MZ') return [false, 'Tệp không phải chương trình Windows (.exe).'];
    $size = filesize($srcPath);
    if ($size > 5 * 1048576) return [false, 'Trình cài đặt trực tuyến phải nhỏ (< 5 MB) — tệp này có vẻ là bộ cài đầy đủ.'];
    $sha = hash_file('sha256', $srcPath);
    $dest = MO_RELEASES . '/' . MO_WEBSETUP;
    $c = load_config();
    if (is_file($dest)) {
        $old = hash_file('sha256', $dest);
        if ($old === $sha) return [true, "Trình cài đặt trực tuyến đã có sẵn (SHA-256 $sha) — giữ nguyên."];
        if (!$replace) return [false, "Đã có trình cài đặt trực tuyến khác (SHA-256 $old). Danh tiếng với Chrome/Windows gắn với mã băm này: thay tệp là bắt đầu lại từ đầu. Dùng --replace nếu thật sự cần."];
    }
    if (!is_dir(MO_RELEASES)) mkdir(MO_RELEASES, 0750, true);
    if (!copy($srcPath, $dest . '.tmp') || !rename($dest . '.tmp', $dest)) return [false, 'Không lưu được tệp.'];
    chmod($dest, 0640);
    $c['webSetup'] = ['sha256' => $sha, 'size' => $size, 'publishedAt' => date('c')];
    save_config($c);
    audit('websetup', $sha);
    return [true, "Đã phát hành trình cài đặt trực tuyến (SHA-256 $sha)."];
}

/** Sends a file from the releases folder (nginx X-Accel-Redirect; the local test server streams it itself). */
function send_release_file(string $file): void {
    header('Content-Type: application/vnd.microsoft.portable-executable');
    header('Content-Disposition: attachment; filename="' . $file . '"');
    header('X-Content-Type-Options: nosniff');
    if (getenv('MO_DEV_SENDFILE')) { header('Content-Length: ' . filesize(MO_RELEASES . '/' . $file)); readfile(MO_RELEASES . '/' . $file); return; }
    header('X-Accel-Redirect: /office-files-internal/' . rawurlencode($file));
}

/** "winget install" command once the package is accepted in the Windows Package Manager community repository. */
function winget_id(array $c): ?string {
    $id = (string)($c['wingetId'] ?? '');
    return preg_match('/^[A-Za-z0-9][A-Za-z0-9.\-]{2,80}$/D', $id) ? $id : null;
}

function fmt_size(int $b): string { return $b >= 1048576 ? number_format($b / 1048576, 1) . ' MB' : number_format($b / 1024, 0) . ' KB'; }
