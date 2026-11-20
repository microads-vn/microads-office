<?php
// Claude API access for the blog writer, the daily SEO care and the Excel formula tool.
// The key lives in MO_DATA/secrets/anthropic.json (outside the web root, never in git): {"key": "...", "expires": "YYYY-MM-DD"}.
// The official Anthropic PHP SDK is installed with Composer in MO_DATA/lib (see server/deploy/setup_seo.py).
declare(strict_types=1);

const AI_MODEL = 'claude-opus-5-5';
// USD per million tokens for claude-opus-5-5 (input, output, cache read, cache write 5 min) + web search per request.
const AI_PRICE = ['in' => 4.0, 'out' => 20.0, 'cache_read' => 0.20, 'cache_write' => 5.0, 'search' => 0.01];

function ai_secret_path(): string { return MO_DATA . '/secrets/anthropic.json'; }

function ai_secret(): array { return json_read(ai_secret_path(), []) + ['key' => '', 'expires' => '']; }

function ai_save_secret(string $key, string $expires): void {
    if (!is_dir(MO_DATA . '/secrets')) mkdir(MO_DATA . '/secrets', 0750, true);
    json_write(ai_secret_path(), ['key' => trim($key), 'expires' => trim($expires), 'savedAt' => date('c')]);
    @chmod(ai_secret_path(), 0640);
}

/** Days until the key expires (null when unknown). */
function ai_key_days_left(): ?int {
    $e = ai_secret()['expires'];
    if (!preg_match('/^\d{4}-\d{2}-\d{2}$/', $e)) return null;
    return (int)floor((strtotime($e . ' 23:59:59') - time()) / 86400);
}

function ai_client(): \Anthropic\Client {
    static $client = null;
    if ($client) return $client;
    $autoload = MO_DATA . '/lib/vendor/autoload.php';
    if (!is_file($autoload)) throw new RuntimeException('Anthropic SDK not installed (' . $autoload . ')');
    require_once $autoload;
    $key = ai_secret()['key'];
    if ($key === '') throw new RuntimeException('No Claude API key configured (admin › Bài viết & SEO)');
    return $client = new \Anthropic\Client(apiKey: $key);
}

function ai_enum($v): string { return $v instanceof \BackedEnum ? (string)$v->value : (string)$v; }

/** All text blocks of a response joined. */
function ai_text($message): string {
    $out = '';
    foreach ($message->content as $block) if (ai_enum($block->type) === 'text') $out .= $block->text;
    return $out;
}

/** Records the cost of a response in MO_DATA/seo/usage.json (per day) and returns it in USD. */
function ai_track($message, string $purpose): float {
    $u = $message->usage;
    $searches = $u->serverToolUse->webSearchRequests ?? 0;
    $usd = ($u->inputTokens * AI_PRICE['in'] + $u->outputTokens * AI_PRICE['out'] + (int)$u->cacheReadInputTokens * AI_PRICE['cache_read']
        + (int)$u->cacheCreationInputTokens * AI_PRICE['cache_write']) / 1e6 + $searches * AI_PRICE['search'];
    $path = MO_DATA . '/seo/usage.json';
    if (!is_dir(dirname($path))) mkdir(dirname($path), 0750, true);
    $fp = fopen($path, 'c+');
    if ($fp) {
        flock($fp, LOCK_EX);
        $d = json_decode((string)stream_get_contents($fp), true) ?: [];
        $day = date('Y-m-d');
        $d[$day][$purpose]['usd'] = round(($d[$day][$purpose]['usd'] ?? 0) + $usd, 4);
        $d[$day][$purpose]['calls'] = ($d[$day][$purpose]['calls'] ?? 0) + 1;
        $d = array_slice($d, -120, null, true);
        $json = json_encode($d, JSON_PRETTY_PRINT);
        rewind($fp); ftruncate($fp, 0); fwrite($fp, $json); fflush($fp);
        flock($fp, LOCK_UN); fclose($fp);
    }
    return $usd;
}

/** Average cost of one published article (last 10), used to decide whether today's budget allows another one. */
function ai_article_cost(): float {
    $costs = [];
    foreach (json_read(MO_DATA . '/blog/index.json', []) as $p) if (isset($p['costUsd']) && $p['costUsd'] > 0) $costs[] = (float)$p['costUsd'];
    $costs = array_slice($costs, 0, 10);
    return $costs ? max(0.15, array_sum($costs) / count($costs)) : 0.5;
}

/** Spend (USD) for a day or a month prefix ("2026-11"), optionally for one purpose prefix. */
function ai_spend(string $prefix, string $purpose = ''): float {
    $d = json_read(MO_DATA . '/seo/usage.json', []);
    $sum = 0.0;
    foreach ($d as $day => $byPurpose) {
        if (!str_starts_with($day, $prefix)) continue;
        foreach ($byPurpose as $p => $v) if ($purpose === '' || str_starts_with($p, $purpose)) $sum += (float)$v['usd'];
    }
    return $sum;
}

/**
 * One Claude call. $schema (JSON schema) → returns the decoded object; otherwise returns the text.
 * $search > 0 enables the web search server tool (research calls only, never together with a schema).
 */
function ai_call(string $purpose, string $system, string $user, ?array $schema = null, string $effort = 'high', int $maxTokens = 32000, int $search = 0, float $timeout = 900): array|string {
    $client = ai_client();
    $outputConfig = ['effort' => $effort];
    if ($schema) $outputConfig['format'] = ['type' => 'json_schema', 'schema' => $schema];
    $tools = $search > 0 ? [['type' => 'web_search_20260209', 'name' => 'web_search', 'maxUses' => $search]] : null;
    $attempt = 0;
    while (true) {
        try {
            $m = $client->messages->create(
                model: AI_MODEL,
                maxTokens: $maxTokens,
                system: [['type' => 'text', 'text' => $system, 'cacheControl' => ['type' => 'ephemeral']]],
                messages: [['role' => 'user', 'content' => $user]],
                outputConfig: $outputConfig,
                tools: $tools,
                requestOptions: ['timeout' => $timeout],
            );
            break;
        } catch (\Anthropic\Core\Exceptions\APIStatusException $e) {
            $type = $e->type?->value ?? '';
            if (++$attempt < 3 && in_array($type, ['rate_limit_error', 'overloaded_error', 'api_error'], true)) { sleep(20 * $attempt); continue; }
            throw new RuntimeException('Claude API: ' . ($type ?: 'error') . ' — ' . $e->getMessage(), 0, $e);
        }
    }
    ai_track($m, $purpose);
    $stop = ai_enum($m->stopReason);
    if ($stop === 'refusal') throw new RuntimeException('Claude declined the request (' . $purpose . ')');
    if ($stop === 'max_tokens') throw new RuntimeException('Claude response cut off at max_tokens (' . $purpose . ')');
    $text = ai_text($m);
    if (!$schema) return $text;
    $data = json_decode($text, true);
    if (!is_array($data)) throw new RuntimeException('Claude returned invalid JSON (' . $purpose . ')');
    return $data;
}
