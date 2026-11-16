<?php
// GET /office/download/<file>  → counts the download and lets nginx stream the file (X-Accel-Redirect).
//   MicroadsOffice_Setup.exe          web installer (never changes; installs the latest signed release)
//   MicroadsOffice_Setup_x.y.z.exe    full (offline) installer of one version
//   latest                            → redirect to the full installer of the latest version
declare(strict_types=1);
require __DIR__ . '/inc/bootstrap.php';
$f = basename((string)($_GET['f'] ?? ''));
if ($f === 'latest') {
    $c = load_config();
    $r = release_by_version($c, $c['latest']);
    if (!$r) { http_response_code(404); exit('No release yet'); }
    header('Location: ' . MO_BASE_URL . '/download/' . rawurlencode($r['file']), true, 302);
    exit;
}
if ($f === MO_WEBSETUP) {
    if (!websetup_info(load_config())) { http_response_code(404); exit('Not found'); }
    bump_counter($f);
    send_release_file($f);
    exit;
}
if (!preg_match('/^MicroadsOffice_Setup_[0-9]{1,4}\.[0-9]{1,4}\.[0-9]{1,6}\.exe$/D', $f) || !is_file(MO_RELEASES . '/' . $f)) { http_response_code(404); exit('Not found'); }
bump_counter($f);
send_release_file($f);
