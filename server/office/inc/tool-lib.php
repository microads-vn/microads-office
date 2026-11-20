<?php
// Excel formula generator: anti-abuse token, limits and the Claude call.
declare(strict_types=1);

function tool_secret(): string {
    $s = blog_settings();
    if (empty($s['toolSecret'])) { $s['toolSecret'] = bin2hex(random_bytes(24)); blog_save_settings($s); }
    return $s['toolSecret'];
}

/** Form token valid for the day it was issued and the next (stops scripted calls that never load the page). */
function tool_token(?string $day = null): string { return substr(hash_hmac('sha256', 'fx|' . ($day ?? date('Y-m-d')), tool_secret()), 0, 32); }

function tool_token_ok(string $t): bool {
    return $t !== '' && (hash_equals(tool_token(), $t) || hash_equals(tool_token(date('Y-m-d', time() - 86400)), $t));
}

/** Counts a use for this visitor; returns null when allowed, or the reason it is refused. */
function tool_take_quota(string $ip, string $lang): ?string {
    $s = blog_settings();
    $dir = MO_DATA . '/seo/tool';
    if (!is_dir($dir)) mkdir($dir, 0750, true);
    $fp = fopen($dir . '/' . date('Y-m-d') . '.json', 'c+');
    if (!$fp) return 'busy';
    flock($fp, LOCK_EX);
    $d = json_decode((string)stream_get_contents($fp), true) ?: ['total' => 0, 'ip' => []];
    $who = substr(hash_hmac('sha256', $ip, tool_secret() . date('Y-m-d')), 0, 16); // no raw IP stored
    $reason = null;
    if ($d['total'] >= (int)$s['toolDailyCap']) $reason = 'daily';
    elseif (($d['ip'][$who] ?? 0) >= (int)$s['toolPerIp']) $reason = 'ip';
    elseif (ai_spend(date('Y-m-d'), 'tool') >= (float)$s['toolDailyBudgetUsd'] || ai_spend(date('Y-m-d')) >= (float)$s['dailyBudgetUsd']) $reason = 'budget';
    if ($reason === null) { $d['total']++; $d['ip'][$who] = ($d['ip'][$who] ?? 0) + 1; $d['lang'][$lang] = ($d['lang'][$lang] ?? 0) + 1; }
    $json = json_encode($d);
    rewind($fp); ftruncate($fp, 0); fwrite($fp, $json); fflush($fp);
    flock($fp, LOCK_UN); fclose($fp);
    return $reason;
}

function tool_generate(string $q, string $ctx, string $lang): array {
    $system = <<<TXT
You write Microsoft Excel formulas for people who describe a calculation in plain language (English or Vietnamese).
Rules:
- Return ONE formula that starts with "=", using English function names and commas between arguments (Excel en-US syntax).
- Prefer functions available in Excel 365 that are also supported by Microads Office Excel (it supports 502 Excel functions incl. XLOOKUP,
  FILTER, UNIQUE, SORT, LET, LAMBDA, TEXTSPLIT, DATEDIF, NETWORKDAYS, and the Vietnamese function DOCSO(number, [unit]) that writes a
  number in Vietnamese words, already capitalised and ending in "đồng": DOCSO(120000) = "Một trăm hai mươi nghìn đồng";
  DOCSO(35,"") = "Ba mươi lăm" without unit; never add "đồng" or capitalise it yourself. DOCSO exists only in Microads Office Excel,
  say so in "note"). Do NOT use: PivotTables, VBA, Power Query, WEBSERVICE, STOCKHISTORY, IMAGE, TRANSLATE, COPILOT, PY, FORECAST.ETS.
- Use the exact cell references the user gives; if none, assume a sensible layout and say so in "note".
- Explanation, steps, example and note must be in the user's language ({LANG}). Keep them short and concrete.
- If the request is not about spreadsheet formulas (or asks for anything harmful), set ok=false and explain briefly in "note".
TXT;
    $system = str_replace('{LANG}', $lang === 'en' ? 'English' : 'Vietnamese', $system);
    $schema = ['type' => 'object', 'additionalProperties' => false, 'required' => ['ok', 'formula', 'explanation', 'steps', 'example', 'note'],
        'properties' => ['ok' => ['type' => 'boolean'], 'formula' => ['type' => 'string'], 'explanation' => ['type' => 'string'],
            'steps' => ['type' => 'array', 'items' => ['type' => 'string']], 'example' => ['type' => 'string'], 'note' => ['type' => 'string']]];
    $user = "Request: " . $q . ($ctx !== '' ? "\nSheet layout: " . $ctx : '');
    return ai_call('tool', $system, $user, $schema, 'low', 4000, 0, 60);
}
