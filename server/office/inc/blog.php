<?php
// Blog storage and rendering helpers (posts are JSON files in MO_DATA/blog/posts/<lang>/<slug>.json).
declare(strict_types=1);

define('BLOG_DIR', MO_DATA . '/blog');

/** Topic pillars (cover image = assets/blog/<pillar>-<lang>.jpg). */
const BLOG_PILLARS = [
    'excel-formulas' => ['vi' => 'Công thức Excel', 'en' => 'Excel formulas'],
    'free-office' => ['vi' => 'Phần mềm văn phòng miễn phí', 'en' => 'Free office software'],
    'pdf-signing' => ['vi' => 'PDF & chữ ký số', 'en' => 'PDF & digital signatures'],
    'documents' => ['vi' => 'Soạn thảo văn bản', 'en' => 'Documents'],
    'ai-office' => ['vi' => 'Trợ lý AI', 'en' => 'AI in the office'],
    'office-tips' => ['vi' => 'Mẹo văn phòng', 'en' => 'Office tips'],
];

/** Product screenshots that can illustrate a post (assets/blog/shot-<key>.jpg; Vietnamese interface). */
const BLOG_SHOTS = [
    'excel-explain' => ['vi' => 'Microads Office Excel: bảng doanh thu, biểu đồ cột và Trợ lý AI giải thích công thức', 'en' => 'Microads Office Excel with a revenue table, a column chart and the AI assistant explaining a formula (Vietnamese interface)'],
    'excel-ai' => ['vi' => 'Trợ lý AI tạo bảng và công thức trong Microads Office Excel', 'en' => 'The AI assistant building a table with formulas in Microads Office Excel (Vietnamese interface)'],
    'word-ai' => ['vi' => 'Trợ lý AI soạn văn bản trong Microads Office Word', 'en' => 'The AI assistant drafting a document in Microads Office Word (Vietnamese interface)'],
    'ai-settings' => ['vi' => 'Chọn dịch vụ AI: ChatGPT, Claude, Ollama trong Microads Office', 'en' => 'Choosing the AI service (ChatGPT, Claude, Ollama) in Microads Office (Vietnamese interface)'],
];

function blog_base(string $lang): string { return $lang === 'en' ? '/office/en/blog/' : '/office/blog/'; }
function blog_url(string $lang, string $slug = ''): string { return 'https://microads.vn' . blog_base($lang) . $slug; }
function home_path(string $lang): string { return $lang === 'en' ? '/office/en/' : '/office/'; }
function tool_path(string $lang): string { return $lang === 'en' ? '/office/en/excel-formula-generator' : '/office/cong-cu/tao-cong-thuc-excel'; }

function blog_valid_slug(string $s): bool { return (bool)preg_match('/^[a-z0-9]+(?:-[a-z0-9]+){1,12}$/D', $s) && strlen($s) <= 90; }

function blog_post_path(string $lang, string $slug): string { return BLOG_DIR . "/posts/$lang/$slug.json"; }

function blog_get(string $lang, string $slug): ?array {
    if (!in_array($lang, ['vi', 'en'], true) || !blog_valid_slug($slug)) return null;
    $p = json_read(blog_post_path($lang, $slug), []);
    return $p ?: null;
}

function blog_save(array $post): void {
    $dir = BLOG_DIR . '/posts/' . $post['lang'];
    if (!is_dir($dir)) mkdir($dir, 0750, true);
    json_write(blog_post_path($post['lang'], $post['slug']), $post);
    blog_index_rebuild();
}

/** Small index (no bodies) of all posts, newest first: MO_DATA/blog/index.json. */
function blog_index_rebuild(): array {
    $all = [];
    foreach (['vi', 'en'] as $lang) {
        foreach (glob(BLOG_DIR . "/posts/$lang/*.json") ?: [] as $f) {
            $p = json_read($f, []);
            if (!$p) continue;
            unset($p['bodyHtml'], $p['faq'], $p['research'], $p['history']);
            $all[] = $p;
        }
    }
    usort($all, fn($a, $b) => strcmp($b['publishedAt'] ?? '', $a['publishedAt'] ?? ''));
    if (!is_dir(BLOG_DIR)) mkdir(BLOG_DIR, 0750, true);
    json_write(BLOG_DIR . '/index.json', $all);
    return $all;
}

/** Published posts (index entries), optionally one language. */
function blog_list(?string $lang = null, bool $publishedOnly = true): array {
    static $idx = null;
    $idx ??= json_read(BLOG_DIR . '/index.json', []);
    return array_values(array_filter($idx, fn($p) => ($lang === null || $p['lang'] === $lang) && (!$publishedOnly || ($p['status'] ?? '') === 'published')));
}

/** Word count that works for Vietnamese (syllables separated by spaces). */
function blog_word_count(string $html): int {
    $t = trim(preg_replace('/\s+/u', ' ', html_entity_decode(strip_tags($html), ENT_QUOTES | ENT_HTML5, 'UTF-8')));
    return $t === '' ? 0 : count(explode(' ', $t));
}

function blog_reading_minutes(int $words): int { return max(2, (int)round($words / 220)); }

/**
 * Keeps only safe, simple markup from AI-written HTML: headings, paragraphs, lists, tables, code, quotes, links.
 * Headings get ids (for the table of contents); external links open in a new tab with rel="noopener".
 * Returns [html, headings[[id, text]], links[href]].
 */
function blog_sanitize(string $html): array {
    $allowed = ['h2' => [], 'h3' => [], 'p' => [], 'ul' => [], 'ol' => [], 'li' => [], 'strong' => [], 'em' => [], 'b' => [], 'i' => [],
        'code' => [], 'pre' => [], 'kbd' => [], 'blockquote' => [], 'table' => [], 'thead' => [], 'tbody' => [], 'tr' => [],
        'th' => ['colspan', 'rowspan', 'scope'], 'td' => ['colspan', 'rowspan'], 'a' => ['href', 'title'], 'br' => [], 'hr' => [], 'caption' => []];
    $doc = new DOMDocument('1.0', 'UTF-8');
    libxml_use_internal_errors(true);
    $doc->loadHTML('<?xml encoding="UTF-8"><div id="root">' . $html . '</div>', LIBXML_HTML_NOIMPLIED | LIBXML_HTML_NODEFDTD | LIBXML_NONET);
    libxml_clear_errors();
    $root = $doc->getElementById('root') ?? $doc->getElementsByTagName('div')->item(0);
    $headings = []; $links = []; $ids = [];
    $walk = function (DOMNode $node) use (&$walk, $allowed, $doc, &$headings, &$links, &$ids) {
        foreach (iterator_to_array($node->childNodes) as $child) {
            if ($child instanceof DOMComment || $child instanceof DOMProcessingInstruction) { $node->removeChild($child); continue; }
            if (!$child instanceof DOMElement) continue;
            $tag = strtolower($child->nodeName);
            if ($tag === 'h1') { $h = $doc->createElement('h2'); while ($child->firstChild) $h->appendChild($child->firstChild); $node->replaceChild($h, $child); $child = $h; $tag = 'h2'; }
            if (in_array($tag, ['script', 'style', 'iframe', 'object', 'embed', 'form', 'input', 'button', 'img', 'svg', 'video', 'audio'], true)) { $node->removeChild($child); continue; }
            if (!isset($allowed[$tag])) { // unwrap unknown tags (div, span, section…) keeping their text
                $walk($child);
                while ($child->firstChild) $node->insertBefore($child->firstChild, $child);
                $node->removeChild($child);
                continue;
            }
            foreach (iterator_to_array($child->attributes) as $attr) if (!in_array($attr->name, $allowed[$tag], true)) $child->removeAttribute($attr->name);
            if ($tag === 'a') {
                $href = trim($child->getAttribute('href'));
                $internal = (bool)preg_match('#^(/office/|https://microads\.vn/)#', $href);
                if (!$internal && !preg_match('#^https://[a-z0-9.-]+\.[a-z]{2,}(/|$)#i', $href)) { // unsafe or relative link → plain text
                    while ($child->firstChild) $node->insertBefore($child->firstChild, $child);
                    $node->removeChild($child);
                    continue;
                }
                if (!$internal) { $child->setAttribute('target', '_blank'); $child->setAttribute('rel', 'noopener'); }
                $links[] = $href;
            }
            if ($tag === 'h2' || $tag === 'h3') {
                $text = trim($child->textContent);
                $id = trim(preg_replace('/[^a-z0-9]+/', '-', strtolower(blog_ascii($text))), '-') ?: 'muc';
                $id = substr($id, 0, 60); $base = $id; $n = 2;
                while (isset($ids[$id])) $id = $base . '-' . $n++;
                $ids[$id] = true;
                $child->setAttribute('id', $id);
                if ($tag === 'h2') $headings[] = [$id, $text];
            }
            $walk($child);
        }
    };
    $walk($root);
    $out = '';
    foreach ($root->childNodes as $c) $out .= $doc->saveHTML($c);
    return [trim($out), $headings, array_values(array_unique($links))];
}

/** Vietnamese → ASCII (for slugs and anchor ids). */
function blog_ascii(string $s): string {
    $map = ['à'=>'a','á'=>'a','ạ'=>'a','ả'=>'a','ã'=>'a','â'=>'a','ầ'=>'a','ấ'=>'a','ậ'=>'a','ẩ'=>'a','ẫ'=>'a','ă'=>'a','ằ'=>'a','ắ'=>'a','ặ'=>'a','ẳ'=>'a','ẵ'=>'a',
        'è'=>'e','é'=>'e','ẹ'=>'e','ẻ'=>'e','ẽ'=>'e','ê'=>'e','ề'=>'e','ế'=>'e','ệ'=>'e','ể'=>'e','ễ'=>'e','ì'=>'i','í'=>'i','ị'=>'i','ỉ'=>'i','ĩ'=>'i',
        'ò'=>'o','ó'=>'o','ọ'=>'o','ỏ'=>'o','õ'=>'o','ô'=>'o','ồ'=>'o','ố'=>'o','ộ'=>'o','ổ'=>'o','ỗ'=>'o','ơ'=>'o','ờ'=>'o','ớ'=>'o','ợ'=>'o','ở'=>'o','ỡ'=>'o',
        'ù'=>'u','ú'=>'u','ụ'=>'u','ủ'=>'u','ũ'=>'u','ư'=>'u','ừ'=>'u','ứ'=>'u','ự'=>'u','ử'=>'u','ữ'=>'u','ỳ'=>'y','ý'=>'y','ỵ'=>'y','ỷ'=>'y','ỹ'=>'y','đ'=>'d'];
    return strtr(mb_strtolower($s, 'UTF-8'), $map);
}

function blog_slugify(string $s): string {
    $s = trim(preg_replace('/[^a-z0-9]+/', '-', blog_ascii($s)), '-');
    $parts = array_slice(explode('-', $s), 0, 10);
    return implode('-', $parts);
}

/** Word 3-gram shingles for near-duplicate detection. */
function blog_shingles(string $html): array {
    $t = blog_ascii(html_entity_decode(strip_tags($html), ENT_QUOTES | ENT_HTML5, 'UTF-8'));
    $w = preg_split('/[^a-z0-9]+/', $t, -1, PREG_SPLIT_NO_EMPTY);
    $s = [];
    for ($i = 0; $i + 2 < count($w); $i++) $s[$w[$i] . ' ' . $w[$i + 1] . ' ' . $w[$i + 2]] = true;
    return $s;
}

function blog_similarity(array $a, array $b): float {
    if (!$a || !$b) return 0.0;
    $inter = count(array_intersect_key($a, $b));
    return $inter / max(1, min(count($a), count($b)));
}

/** Download call-to-action box (the buttons open the same download notice as the home page). */
function blog_cta(string $lang, string $variant = 'end'): string {
    $web = 'download/MicroadsOffice_Setup.exe';
    $vi = $lang !== 'en';
    $title = $vi ? 'Tải Microads Office — miễn phí vĩnh viễn' : 'Get Microads Office — free forever';
    $text = $vi
        ? 'Word, Excel và PDF trong một bộ cài cho Windows 10/11. <b>Không cần tài khoản</b>, không đăng ký, không dùng thử có thời hạn, không giới hạn tính năng.'
        : 'Word, Excel and PDF in one installer for Windows 10/11. <b>No account</b>, no sign-up, no trial period, no feature limits.';
    $btn = $vi ? 'Tải miễn phí ngay' : 'Download free now';
    $more = $vi ? 'Xem chi tiết' : 'Learn more';
    $ticks = $vi ? ['Miễn phí 100%, mãi mãi', 'Không cần tài khoản', 'Mở và lưu .docx, .xlsx, ký số PDF'] : ['100% free, forever', 'No account required', 'Opens and saves .docx and .xlsx, signs PDFs'];
    if (store_links(load_config())) $ticks[] = $vi ? 'Đã qua kiểm duyệt của Microsoft, có trên Microsoft Store' : 'Passed Microsoft’s certification, on the Microsoft Store';
    $li = implode('', array_map(fn($x) => '<li>' . h($x) . '</li>', $ticks));
    static $store = false;
    if ($store === false) $store = store_links(load_config());
    // Microsoft Store first (no warnings, automatic updates); the .exe stays available as a secondary link.
    $actions = $store
        ? '<a class="store-badge" href="' . h($store['page']) . '" rel="noopener"><img src="/office/assets/ms-store-badge-' . ($vi ? 'vi' : 'en') . '.svg" alt="' . h($vi ? 'Tải xuống từ Microsoft Store' : 'Download from the Microsoft Store') . '" width="242" height="66" loading="lazy"></a>'
          . '<a class="cta-more" href="/office/' . $web . '" data-warn>' . h($vi ? 'hoặc tải trình cài đặt' : 'or download the installer') . ' →</a>'
        : '<a class="btn primary big" href="/office/' . $web . '" data-warn>' . h($btn) . '</a><a class="cta-more" href="' . h(home_path($lang)) . '">' . h($more) . ' →</a>';
    return '<aside class="cta cta-' . $variant . '" aria-label="' . h($title) . '"><div class="cta-text"><p class="cta-title">' . h($title) . '</p><p>' . $text . '</p><ul class="ticks">' . $li . '</ul></div>'
        . '<div class="cta-actions">' . $actions . '</div></aside>';
}

/** Inserts the middle call-to-action before the third H2 (or at the end of a short post). */
function blog_body_with_cta(string $html, string $lang): string {
    $n = 0;
    $pos = null;
    $offset = 0;
    while (($p = strpos($html, '<h2', $offset)) !== false) { if (++$n === 3) { $pos = $p; break; } $offset = $p + 3; }
    return $pos === null ? $html : substr($html, 0, $pos) . blog_cta($lang, 'mid') . substr($html, $pos);
}

/** Related posts: same language, same pillar first, then newest. */
function blog_related(array $post, int $n = 4): array {
    $all = array_filter(blog_list($post['lang']), fn($p) => $p['slug'] !== $post['slug']);
    usort($all, fn($a, $b) => [($b['pillar'] ?? '') === $post['pillar'], $b['publishedAt'] ?? ''] <=> [($a['pillar'] ?? '') === $post['pillar'], $a['publishedAt'] ?? '']);
    return array_slice(array_values($all), 0, $n);
}

function blog_date(string $iso, string $lang): string {
    $t = strtotime($iso);
    return $lang === 'en' ? date('F j, Y', $t) : date('d/m/Y', $t);
}

function blog_settings(): array {
    // Budget: everything (articles, nightly audit, formula tool) fits in dailyBudgetUsd; the audit and the tool have their own share.
    return json_read(MO_DATA . '/seo/settings.json', []) + ['enabled' => true, 'minPerDay' => 2, 'maxPerDay' => 5, 'dailyBudgetUsd' => 1.5,
        'auditReserveUsd' => 0.15, 'toolDailyBudgetUsd' => 0.2, 'toolDailyCap' => 150, 'toolPerIp' => 10, 'indexnowKey' => '', 'refreshPerDay' => 1];
}

function blog_save_settings(array $s): void {
    if (!is_dir(MO_DATA . '/seo')) mkdir(MO_DATA . '/seo', 0750, true);
    json_write(MO_DATA . '/seo/settings.json', $s);
}

/**
 * Privacy-friendly page-view counter (no cookies, no IP, no third party): per-day counts per page in MO_DATA/seo/views/<date>.json.
 * Crawlers and monitoring tools are not counted. Used by the daily audit (views → downloads conversion, engagement).
 */
function track_view(string $page): void {
    if (PHP_SAPI === 'cli' || ($_SERVER['REQUEST_METHOD'] ?? 'GET') !== 'GET') return;
    $ua = $_SERVER['HTTP_USER_AGENT'] ?? '';
    if ($ua === '' || preg_match('/bot|crawl|spider|slurp|preview|facebookexternalhit|curl|wget|python|java\/|go-http|headless|lighthouse|pagespeed|monitor|MicroadsOffice/i', $ua)) return;
    $dir = MO_DATA . '/seo/views';
    if (!is_dir($dir) && !@mkdir($dir, 0750, true)) return;
    $fp = @fopen($dir . '/' . date('Y-m-d') . '.json', 'c+');
    if (!$fp) return;
    flock($fp, LOCK_EX);
    $d = json_decode((string)stream_get_contents($fp), true) ?: [];
    $d[$page] = ($d[$page] ?? 0) + 1;
    $ref = parse_url($_SERVER['HTTP_REFERER'] ?? '', PHP_URL_HOST) ?: '';
    if ($ref !== '' && $ref !== 'microads.vn') $d['_ref:' . substr($ref, 0, 60)] = ($d['_ref:' . substr($ref, 0, 60)] ?? 0) + 1;
    $json = json_encode($d);
    rewind($fp); ftruncate($fp, 0); fwrite($fp, $json); fflush($fp);
    flock($fp, LOCK_UN); fclose($fp);
}

/** IndexNow key (Bing, Yandex, Seznam, Naver…), created once and served at /office/<key>.txt. */
function blog_indexnow_key(): string {
    $s = blog_settings();
    if (!preg_match('/^[0-9a-f]{32}$/', $s['indexnowKey'])) { $s['indexnowKey'] = bin2hex(random_bytes(16)); blog_save_settings($s); }
    return $s['indexnowKey'];
}
