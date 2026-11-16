<?php
// POST /office/report.php — a user reports inappropriate AI assistant content (Microsoft Store policy 11.16).
// Sent only when the user clicks "Báo cáo / Report" and confirms; stored for the Microads admin (admin › Báo cáo AI).
// No IP address is stored: only a salted daily hash, used to limit each sender to 30 reports per day.
declare(strict_types=1);
require __DIR__ . '/inc/bootstrap.php';
header('Content-Type: application/json; charset=utf-8');
header('Cache-Control: no-store');
header('X-Content-Type-Options: nosniff');

function fail(int $code, string $msg): never { http_response_code($code); echo json_encode(['ok' => false, 'error' => $msg]); exit; }

if (($_SERVER['REQUEST_METHOD'] ?? '') !== 'POST') fail(405, 'POST only');
// JSON only: a plain HTML form on another web site cannot send this content type without a CORS preflight (which is refused)
if (stripos((string)($_SERVER['CONTENT_TYPE'] ?? ''), 'application/json') !== 0) fail(415, 'application/json required');
$raw = file_get_contents('php://input', false, null, 0, 65537);
if ($raw === false || strlen($raw) > 65536) fail(413, 'too large');
$in = json_decode($raw, true);
if (!is_array($in)) fail(400, 'invalid JSON');

$s = static fn(string $k, int $max): string => mb_substr(trim((string)($in[$k] ?? '')), 0, $max);
$reasons = ['harmful', 'offensive', 'wrong', 'privacy', 'other'];
$reason = in_array($in['reason'] ?? '', $reasons, true) ? $in['reason'] : 'other';
$report = [
    'id' => bin2hex(random_bytes(6)),
    'time' => date('c'),
    'app' => preg_replace('/[^a-z]/', '', strtolower($s('app', 10))),
    'version' => valid_version($s('version', 20)) ? $s('version', 20) : 'unknown',
    'lang' => $s('lang', 5),
    'provider' => $s('provider', 40),
    'model' => $s('model', 80),
    'reason' => $reason,
    'comment' => $s('comment', 2000),
    'question' => $s('question', 4000),
    'answer' => $s('answer', 8000),
];
if ($report['answer'] === '' && $report['comment'] === '') fail(400, 'empty report');

$dir = MO_DATA . '/reports';
if (!is_dir($dir)) @mkdir($dir, 0750, true);

// rate limit: 30 per sender per day (salted hash of the address, rotated daily, never stored in clear)
$salt = date('Y-m-d') . '|' . (json_read(MO_ADMIN, [])['hash'] ?? 'microads');
$addr = (string)($_SERVER['REMOTE_ADDR'] ?? '');
// IPv6: one person usually owns a whole /64, so count per /64 (otherwise rotating addresses would bypass the limit)
if (str_contains($addr, ':') && ($bin = @inet_pton($addr)) !== false) $addr = bin2hex(substr($bin, 0, 8)) . '::/64';
$who = substr(hash('sha256', $salt . '|' . $addr), 0, 16);
$ratePath = $dir . '/rate-' . date('Y-m-d') . '.json';
$fp = @fopen($ratePath, 'c+');
if (!$fp) fail(500, 'storage unavailable');
flock($fp, LOCK_EX);
$rates = json_decode((string)stream_get_contents($fp), true) ?: [];
// per sender: 30 a day; for everyone together: 3000 a day (bounded disk use even under a distributed flood)
if (($rates[$who] ?? 0) >= 30 || ($rates['_total'] ?? 0) >= 3000 || count($rates) > 20000) { flock($fp, LOCK_UN); fclose($fp); fail(429, 'too many reports today'); }
$rates[$who] = ($rates[$who] ?? 0) + 1;
$rates['_total'] = ($rates['_total'] ?? 0) + 1;
ftruncate($fp, 0); rewind($fp); fwrite($fp, json_encode($rates));
flock($fp, LOCK_UN); fclose($fp);
foreach (glob($dir . '/rate-*.json') ?: [] as $old) if ($old !== $ratePath && filemtime($old) < time() - 3 * 86400) @unlink($old);

$line = json_encode($report, JSON_UNESCAPED_UNICODE | JSON_UNESCAPED_SLASHES) . "\n";
if (@file_put_contents($dir . '/' . date('Y-m') . '.jsonl', $line, FILE_APPEND | LOCK_EX) === false) fail(500, 'could not store');
echo json_encode(['ok' => true, 'id' => $report['id']]);
