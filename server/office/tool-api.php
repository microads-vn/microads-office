<?php
// POST /office/api/formula {q, ctx, t, website, lang} → {ok, formula, explanation, steps, example, note} or {ok:false, message}.
declare(strict_types=1);
require __DIR__ . '/inc/bootstrap.php';
require __DIR__ . '/inc/blog.php';
require __DIR__ . '/inc/ai.php';
require __DIR__ . '/inc/tool-lib.php';
header('Content-Type: application/json; charset=utf-8');
header('X-Content-Type-Options: nosniff');
header('Cache-Control: no-store');
$in = json_decode((string)file_get_contents('php://input', false, null, 0, 8192), true);
$lang = (($in['lang'] ?? 'vi') === 'en') ? 'en' : 'vi';
$T = fn($vi, $en) => $lang === 'en' ? $en : $vi;
$fail = function (int $code, string $msg) { http_response_code($code); echo json_encode(['ok' => false, 'message' => $msg], JSON_UNESCAPED_UNICODE); exit; };
if ($_SERVER['REQUEST_METHOD'] !== 'POST' || !is_array($in)) $fail(405, 'POST JSON only');
if (($in['website'] ?? '') !== '' || !tool_token_ok((string)($in['t'] ?? ''))) $fail(403, $T('Phiên làm việc đã hết hạn, vui lòng tải lại trang.', 'Your session expired, please reload the page.'));
$q = trim((string)($in['q'] ?? ''));
$ctx = trim((string)($in['ctx'] ?? ''));
if (mb_strlen($q) < 5 || mb_strlen($q) > 600 || mb_strlen($ctx) > 300) $fail(400, $T('Mô tả cần từ 5 đến 600 ký tự.', 'Please describe the formula in 5 to 600 characters.'));
$why = tool_take_quota($_SERVER['REMOTE_ADDR'] ?? '', $lang);
if ($why === 'ip') $fail(429, $T('Bạn đã dùng hết lượt hôm nay. Hãy quay lại ngày mai, hoặc dùng Trợ lý AI ngay trong Microads Office Excel (miễn phí).', 'You have used today’s free generations. Come back tomorrow, or use the AI assistant inside Microads Office Excel (free).'));
if ($why !== null) $fail(503, $T('Công cụ đang quá tải, vui lòng thử lại sau.', 'The tool is busy right now, please try again later.'));
try {
    $r = tool_generate($q, $ctx, $lang);
} catch (Throwable $e) {
    error_log('formula tool: ' . $e->getMessage());
    $fail(502, $T('AI đang bận, vui lòng thử lại sau ít phút.', 'The AI is busy, please try again in a few minutes.'));
}
if (!$r['ok'] || !str_starts_with(ltrim($r['formula']), '=')) $fail(200, $r['note'] ?: $T('Mình chỉ viết được công thức bảng tính.', 'I can only write spreadsheet formulas.'));
echo json_encode(['ok' => true, 'formula' => trim($r['formula']), 'explanation' => $r['explanation'], 'steps' => array_slice($r['steps'], 0, 8),
    'example' => $r['example'], 'note' => $r['note']], JSON_UNESCAPED_UNICODE);
