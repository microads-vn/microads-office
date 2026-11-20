<?php
// Daily content + SEO care for microads.vn/office (run by inc/seo-cron.php from cron as www-data).
//   topics:  research real search demand (web search) → backlog of distinct topics per language (MO_DATA/blog/topics.json)
//   write:   research → write → independent review (score ≥ 80, no factual errors) → publish → IndexNow
//   audit:   crawl own pages, sitemap/robots/canonical/noindex checks, Google Search Console (if connected), Claude analysis,
//            limited automatic fixes (meta rewrite, content refresh), daily report (MO_DATA/seo/reports/<date>.json)
declare(strict_types=1);

const SEO_FORMATS = [
    'how-to' => 'Step-by-step tutorial: short direct answer first, then numbered steps with the exact values/formulas, then variations and common mistakes.',
    'troubleshooting' => 'Problem → causes → fixes: open with the symptom the reader sees (error message, wrong result), list causes from most to least likely, give a fix for each, end with how to prevent it.',
    'comparison' => 'Fair comparison for a decision: criteria table, honest strengths AND limitations of each option (use the facts file for Microads Office limitations), who should pick which.',
    'checklist' => 'Practical checklist the reader can follow at work: grouped items, each with one line of why/how, a short printable summary at the end.',
    'worked-example' => 'One realistic scenario solved end to end (e.g. a payroll sheet, a quotation, an official letter): sample data table, the full solution, then how to adapt it.',
    'explainer' => 'Concept explained simply: plain-language definition, how it works, when to use it and when not to, concrete examples, misconceptions.',
    'use-case' => 'Written for one job role (accountant, teacher, HR, sales admin, small shop owner, student…): the 4–6 tasks they actually do and how to do each faster.',
    'formula-cookbook' => 'Collection of 6–10 short recipes around one need (each: goal, formula, sample input/output, one-line explanation).',
];

const SEO_CLICHES = "English: \"In today's fast-paced world\", \"In today's digital age\", \"delve\", \"unlock the power\", \"game-changer\", \"seamless(ly)\", \"elevate\", \"harness\", \"navigate the complexities\", \"look no further\", \"In conclusion\", \"Whether you're a … or a …\", \"It's important to note\", emoji. Vietnamese: \"Trong thời đại công nghệ số\", \"Trong thời đại 4.0\", \"Hãy cùng tìm hiểu\", \"Hãy cùng Microads\", \"không thể phủ nhận\", \"Tóm lại,\", \"Bài viết dưới đây sẽ\", \"vô cùng\", \"cực kỳ hữu ích\", \"một cách dễ dàng\" (repeated), emoji.";

function seo_dir(string $sub = ''): string { $d = MO_DATA . '/seo' . ($sub ? '/' . $sub : ''); if (!is_dir($d)) mkdir($d, 0750, true); return $d; }

function seo_log(string $msg): void {
    $line = date('Y-m-d H:i:s') . '  ' . $msg . "\n";
    file_put_contents(seo_dir('logs') . '/' . date('Y-m') . '.log', $line, FILE_APPEND | LOCK_EX);
    if (PHP_SAPI === 'cli') echo $line;
}

function seo_facts(): string { return (string)file_get_contents(__DIR__ . '/facts.md'); }

function seo_market(string $lang): string {
    return $lang === 'en'
        ? 'English-speaking readers worldwide (US, UK, India, Philippines, Australia, Canada, Africa…), mostly office workers, students, freelancers and small businesses on Windows who need Word/Excel/PDF without paying for a subscription.'
        : 'Người dùng Việt Nam: nhân viên văn phòng, kế toán, giáo viên, cán bộ hành chính, sinh viên, chủ cửa hàng nhỏ dùng Windows, cần soạn văn bản, bảng tính, ký số PDF mà không muốn trả phí bản quyền.';
}

/** Topics of the main microads.vn blog (company services) — the Office blog must not compete with them. */
function seo_main_site_topics(): array {
    $cache = seo_dir() . '/main-site-topics.json';
    $c = json_read($cache, []);
    if (($c['date'] ?? '') === date('Y-m-d')) return $c['topics'];
    $xml = seo_http_get('https://microads.vn/sitemap.xml')['body'] ?? '';
    preg_match_all('#<loc>https://microads\.vn/blog/([a-z0-9-]+)/?</loc>#', $xml, $m);
    $topics = array_values(array_unique(array_map(fn($s) => str_replace('-', ' ', $s), $m[1])));
    json_write($cache, ['date' => date('Y-m-d'), 'topics' => $topics]);
    return $topics;
}

function seo_topics(): array { return json_read(BLOG_DIR . '/topics.json', []) + ['vi' => [], 'en' => [], 'used' => []]; }
function seo_save_topics(array $t): void { if (!is_dir(BLOG_DIR)) mkdir(BLOG_DIR, 0750, true); json_write(BLOG_DIR . '/topics.json', $t); }

/** Titles + primary keywords already covered (published posts, backlog, used topics) for one language. */
function seo_covered(string $lang): array {
    $t = seo_topics();
    $out = [];
    foreach (blog_index_rebuild() as $p) if ($p['lang'] === $lang) $out[] = $p['title'] . ' [' . ($p['primaryKeyword'] ?? '') . ']';
    foreach ($t[$lang] as $x) $out[] = $x['title'] . ' [' . $x['primaryKeyword'] . '] (planned)';
    foreach ($t['used'] as $x) if (($x['lang'] ?? '') === $lang) $out[] = $x['title'] . ' [' . $x['primaryKeyword'] . '] (done/discarded)';
    return $out;
}

// ---------------------------------------------------------------------------------------------------------------- topics

function seo_plan_topics(string $lang, int $n = 12): int {
    $pillars = implode(', ', array_keys(BLOG_PILLARS));
    $covered = seo_covered($lang);
    $main = array_slice(seo_main_site_topics(), 0, 200);
    $language = $lang === 'en' ? 'English' : 'Vietnamese';
    $system = "You are an SEO strategist for Microads Office, a free Windows office suite (Word, Excel, PDF). You plan blog topics that a new, small site can realistically rank for: specific long-tail questions with real search demand, where we can add genuine first-hand value (we build a spreadsheet engine with 502 Excel functions, a .docx word processor and a PDF signer). Never plan thin, generic or near-duplicate topics.\n\nPRODUCT FACTS:\n" . seo_facts();
    $research = ai_call('topics-research', $system,
        "Target readers: " . seo_market($lang) . "\nWrite in $language.\nPillars: $pillars.\n\nUse web search (People Also Ask style questions, forums such as Reddit / Microsoft Q&A / Vietnamese Facebook groups and Webtretho for Vietnamese, top Google results) to find 25 concrete questions people really search for in these pillars right now. For each, note the exact query wording, the intent, and what current top results miss (information gain we could add). Avoid anything already covered:\n- " . implode("\n- ", $covered ?: ['(nothing yet)'])
        . "\n\nAlso avoid these topics of our company's main blog (custom software, security, ERP/CRM, hosting — not our focus):\n- " . implode("\n- ", $main ?: ['(none)']),
        null, 'low', 12000, 5);
    $schema = ['type' => 'object', 'additionalProperties' => false, 'required' => ['topics'], 'properties' => ['topics' => ['type' => 'array', 'items' => [
        'type' => 'object', 'additionalProperties' => false,
        'required' => ['title', 'primaryKeyword', 'secondaryKeywords', 'searchIntent', 'format', 'pillar', 'audience', 'uniqueAngle'],
        'properties' => [
            'title' => ['type' => 'string'], 'primaryKeyword' => ['type' => 'string'], 'secondaryKeywords' => ['type' => 'array', 'items' => ['type' => 'string']],
            'searchIntent' => ['type' => 'string', 'enum' => ['informational', 'commercial', 'transactional']],
            'format' => ['type' => 'string', 'enum' => array_keys(SEO_FORMATS)], 'pillar' => ['type' => 'string', 'enum' => array_keys(BLOG_PILLARS)],
            'audience' => ['type' => 'string'], 'uniqueAngle' => ['type' => 'string']]]]]];
    $res = ai_call('topics-plan', $system,
        "From the research below, choose the $n best topics for $language readers (" . seo_market($lang) . ").\nRules: every topic answers a different search query (no two topics share a primary keyword or intent); spread across ALL pillars and at least 6 different formats; titles are specific (numbers, exact function names, exact error messages, exact tasks), never clickbait; the primary keyword is the phrase people type (in $language); uniqueAngle says concretely what our article adds that current results lack.\nAlready covered (do not repeat or paraphrase):\n- " . implode("\n- ", $covered ?: ['(nothing yet)']) . "\n\nRESEARCH:\n" . $research,
        $schema, 'high', 16000);
    $t = seo_topics();
    $have = array_map(fn($x) => blog_ascii($x['primaryKeyword']), array_merge($t[$lang], array_filter($t['used'], fn($u) => ($u['lang'] ?? '') === $lang)));
    foreach (blog_index_rebuild() as $p) if ($p['lang'] === $lang) $have[] = blog_ascii($p['primaryKeyword'] ?? '');
    $added = 0;
    foreach ($res['topics'] as $topic) {
        $k = blog_ascii($topic['primaryKeyword']);
        if ($k === '' || in_array($k, $have, true)) continue;
        $have[] = $k;
        $topic['id'] = bin2hex(random_bytes(5));
        $topic['createdAt'] = date('c');
        $t[$lang][] = $topic;
        $added++;
    }
    seo_save_topics($t);
    seo_log("topics $lang: +$added (backlog " . count($t[$lang]) . ')');
    return $added;
}

/** Picks the next topic: least-covered pillar first, and a different format than the last post in that language. */
function seo_pick_topic(string $lang): ?array {
    $t = seo_topics();
    if (count($t[$lang]) < 4) { seo_plan_topics($lang); $t = seo_topics(); }
    if (!$t[$lang]) return null;
    $posts = array_values(array_filter(blog_index_rebuild(), fn($p) => $p['lang'] === $lang && ($p['status'] ?? '') === 'published'));
    $recent = array_slice($posts, 0, 12);
    $pillarCount = array_count_values(array_column($recent, 'pillar'));
    $lastFormats = array_slice(array_column($posts, 'format'), 0, 2);
    $cands = $t[$lang];
    usort($cands, function ($a, $b) use ($pillarCount, $lastFormats) {
        $sa = ($pillarCount[$a['pillar']] ?? 0) * 10 + (in_array($a['format'], $lastFormats, true) ? 25 : 0) + random_int(0, 6);
        $sb = ($pillarCount[$b['pillar']] ?? 0) * 10 + (in_array($b['format'], $lastFormats, true) ? 25 : 0) + random_int(0, 6);
        return $sa <=> $sb;
    });
    return $cands[0];
}

function seo_topic_done(string $lang, array $topic, string $result): void {
    $t = seo_topics();
    $t[$lang] = array_values(array_filter($t[$lang], fn($x) => $x['id'] !== $topic['id']));
    $t['used'][] = ['lang' => $lang, 'title' => $topic['title'], 'primaryKeyword' => $topic['primaryKeyword'], 'result' => $result, 'at' => date('c')];
    $t['used'] = array_slice($t['used'], -400);
    seo_save_topics($t);
}

// ---------------------------------------------------------------------------------------------------------------- writing

function seo_writer_system(string $lang): string {
    $formats = '';
    foreach (SEO_FORMATS as $k => $v) $formats .= "- $k: $v\n";
    $language = $lang === 'en' ? 'English (US spelling, international audience)' : 'Vietnamese (natural, full diacritics, Vietnamese office context: VND, Vietnamese names, Vietnamese regulations when relevant)';
    return <<<TXT
You are a senior writer on the Microads Office team. Microads builds Microads Office — a free Windows suite with a .docx word processor, an Excel-compatible spreadsheet engine with 502 functions and a PDF reader/signer. You write genuinely useful articles for people searching Google, in {$language}.

What makes an article publishable (Google's helpful, people-first content standard):
- Answer the searcher's question in the first 2–3 sentences, then go deeper. No throat-clearing intro.
- Add information gain: concrete examples with realistic sample data (tables), exact formulas with the expected result, exact error messages, exact keyboard shortcuts, edge cases and common mistakes. Every formula you show must be correct Excel syntax and every example result must be computed correctly.
- The article must stay useful to readers who use Microsoft Word/Excel or other tools; mention when steps are the same there.
- Mention Microads Office naturally only where it genuinely helps (typically 1–3 times), using ONLY the product facts below — never invent features, menus, platforms or prices; state limitations honestly when comparing. Download boxes are added by the page automatically: do not write download buttons or "download now" paragraphs.
- No invented statistics, studies, quotes, testimonials, prices of other products, or personal anecdotes. Facts that are not in the product facts or common knowledge must come from the research notes, with a link to that source.
- Avoid AI clichés and filler: {CLICHES}
- Use the structure that fits the format (guides below), not a fixed template. Vary headings; no heading called "Introduction" or "Conclusion".
- Length: what the intent needs — usually 900 to 1,800 words. Never pad.

Output fields:
- title: ≤ 60 characters when possible (max 70), contains the primary keyword naturally, specific, no clickbait, no ALL CAPS, a year only when content is time-sensitive.
- metaDescription: 140–158 characters, primary keyword + the concrete value of the article.
- slug: 3–7 lowercase ASCII words joined by hyphens ({$lang} slug without diacritics), no stop-word padding.
- excerpt: one sentence (≤ 160 characters) for article cards.
- bodyHtml: HTML using ONLY <h2> <h3> <p> <ul> <ol> <li> <strong> <em> <table> <thead> <tbody> <tr> <th> <td> <pre> <code> <kbd> <blockquote> <a href>. No <h1>, no images, no inline styles, no scripts.
  Internal links: 2–4 links to the related articles given (use their exact URLs) and, for formula topics, the free formula generator tool. External links: 0–3 authoritative sources from the research notes only (never invent URLs).
- faq: 3–5 real follow-up questions not already answered by a heading, each answered in 1–3 sentences.
- shot: a product screenshot that illustrates the article, or "none" when none fits.

Format guides:
{$formats}
PRODUCT FACTS (the only source for claims about Microads Office):
TXT . "\n" . seo_facts();
}

function seo_article_schema(): array {
    return ['type' => 'object', 'additionalProperties' => false,
        'required' => ['title', 'metaDescription', 'slug', 'excerpt', 'primaryKeyword', 'keywords', 'shot', 'bodyHtml', 'faq'],
        'properties' => [
            'title' => ['type' => 'string'], 'metaDescription' => ['type' => 'string'], 'slug' => ['type' => 'string'], 'excerpt' => ['type' => 'string'],
            'primaryKeyword' => ['type' => 'string'], 'keywords' => ['type' => 'array', 'items' => ['type' => 'string']],
            'shot' => ['type' => 'string', 'enum' => array_merge(['none'], array_keys(BLOG_SHOTS))], 'bodyHtml' => ['type' => 'string'],
            'faq' => ['type' => 'array', 'items' => ['type' => 'object', 'additionalProperties' => false, 'required' => ['q', 'a'], 'properties' => ['q' => ['type' => 'string'], 'a' => ['type' => 'string']]]]]];
}

/** Related articles the writer may link to (same language), plus fixed pages. */
function seo_link_candidates(string $lang, string $pillar): string {
    $posts = array_values(array_filter(blog_index_rebuild(), fn($p) => $p['lang'] === $lang && ($p['status'] ?? '') === 'published'));
    usort($posts, fn($a, $b) => [($b['pillar'] ?? '') === $pillar, $b['publishedAt']] <=> [($a['pillar'] ?? '') === $pillar, $a['publishedAt']]);
    $lines = [];
    foreach (array_slice($posts, 0, 25) as $p) $lines[] = '- ' . $p['title'] . ' → ' . blog_base($lang) . $p['slug'];
    $lines[] = '- ' . ($lang === 'en' ? 'Free AI Excel formula generator' : 'Công cụ tạo công thức Excel bằng AI miễn phí') . ' → ' . tool_path($lang);
    $lines[] = '- Microads Office (download page) → ' . home_path($lang);
    return implode("\n", $lines);
}

/** Checks done in code (no AI): lengths, structure, links, duplicates. Returns a list of problems. */
function seo_local_checks(array $a, string $lang, ?string $selfSlug = null): array {
    $problems = [];
    [$clean, $headings, $links] = blog_sanitize($a['bodyHtml']);
    $words = blog_word_count($clean);
    if ($words < 700) $problems[] = "Article body is too short ($words words); the intent needs more depth (examples, steps, edge cases).";
    if ($words > 3200) $problems[] = "Article body is too long ($words words); cut repetition.";
    if (count($headings) < 3) $problems[] = 'Use at least 3 <h2> sections.';
    if (mb_strlen($a['title']) > 70) $problems[] = 'Title longer than 70 characters.';
    $ml = mb_strlen($a['metaDescription']);
    if ($ml < 110 || $ml > 165) $problems[] = "Meta description is $ml characters (needs 140–158).";
    if (!blog_valid_slug($a['slug'])) $problems[] = 'Slug must be 2–13 lowercase ASCII words joined by hyphens.';
    $internal = array_filter($links, fn($l) => str_starts_with($l, '/office/') || str_starts_with($l, 'https://microads.vn/office/'));
    if (count($internal) < 1) $problems[] = 'Add 2–4 internal links from the provided list.';
    foreach ($internal as $l) {
        $path = preg_replace('#^https://microads\.vn#', '', $l);
        if (preg_match('#^/office/(en/)?blog/([a-z0-9-]+)$#', $path, $m) && !blog_get($m[1] ? 'en' : 'vi', $m[2])) $problems[] = "Internal link $l does not exist — use only the provided URLs.";
    }
    foreach (['vi', 'en'] as $l2) if ($l2 !== $lang) foreach ($internal as $l) if (str_contains($l, blog_base($l2))) $problems[] = "Link $l points to the other language blog.";
    $sh = blog_shingles($clean);
    foreach (blog_index_rebuild() as $p) {
        if ($p['lang'] !== $lang || $p['slug'] === $selfSlug) continue;
        $other = blog_get($lang, $p['slug']);
        if ($other && blog_similarity($sh, blog_shingles($other['bodyHtml'])) > 0.3) $problems[] = 'Too similar to the existing article "' . $p['title'] . '" — rewrite with a different angle and wording.';
    }
    if ($selfSlug === null && blog_get($lang, $a['slug'])) $problems[] = 'Slug already used — choose another slug.';
    if (preg_match('/<h1|<img|<script|style=/i', $a['bodyHtml'])) $problems[] = 'Remove <h1>, images, scripts and inline styles.';
    return $problems;
}

function seo_review(array $a, array $topic, string $lang): array {
    $schema = ['type' => 'object', 'additionalProperties' => false,
        'required' => ['score', 'factualErrors', 'qualityIssues', 'verdict', 'revisionInstructions'],
        'properties' => ['score' => ['type' => 'integer'], 'factualErrors' => ['type' => 'array', 'items' => ['type' => 'string']],
            'qualityIssues' => ['type' => 'array', 'items' => ['type' => 'string']], 'verdict' => ['type' => 'string', 'enum' => ['publish', 'revise', 'reject']],
            'revisionInstructions' => ['type' => 'string']]];
    $system = "You are a strict editor and fact-checker for the Microads Office blog. You protect the site from anything Google would consider low-quality, scaled or misleading content.\nScore 0–100 (publishable ≥ 80):\n- Correctness (35): every claim about Microads Office matches the facts below; every formula is valid Excel syntax; every worked example result is computed correctly; claims about other products are fair and verifiable.\n- Helpfulness and information gain (30): answers the query fast, concrete examples, would a reader bookmark it rather than return to Google.\n- Originality (15): not generic, not a paraphrase of common articles, not overlapping existing site articles.\n- Writing (20): natural " . ($lang === 'en' ? 'English' : 'Vietnamese with correct diacritics') . ", no AI clichés (" . SEO_CLICHES . "), no keyword stuffing, no hard selling, structure fits the format.\nfactualErrors: list every wrong statement (empty if none). verdict: publish (score ≥ 80 and no factual errors), revise (fixable), reject (wrong topic/thin/duplicate). revisionInstructions: precise edits for the writer.\n\nPRODUCT FACTS:\n" . seo_facts();
    return ai_call('review', $system, "Topic brief:\n" . json_encode($topic, JSON_UNESCAPED_UNICODE | JSON_PRETTY_PRINT)
        . "\n\nExisting articles on the site (must not overlap):\n" . implode("\n", array_slice(seo_covered($lang), 0, 80))
        . "\n\nArticle to review:\n" . json_encode($a, JSON_UNESCAPED_UNICODE | JSON_PRETTY_PRINT), $schema, 'medium', 16000);
}

/** Full pipeline for one new article. Returns the published post, or null when the topic was discarded. */
function seo_write_one(string $lang): ?array {
    $topic = seo_pick_topic($lang);
    if (!$topic) throw new RuntimeException("no topic available for $lang");
    seo_log("write $lang: \"{$topic['title']}\" ({$topic['pillar']}, {$topic['format']})");
    $spent0 = ai_spend(date('Y-m-d'));
    $language = $lang === 'en' ? 'English' : 'Vietnamese';
    $research = ai_call('research', "You are a meticulous research assistant for a $language article. Use web search to collect accurate, current facts. Prefer official sources (Microsoft Support / Learn for Excel and Word behaviour, Adobe or ETSI for PDF signatures, official Vietnamese government portals for Vietnamese regulations). Report what top-ranking pages cover and what they miss.",
        "Article topic: {$topic['title']}\nPrimary keyword: {$topic['primaryKeyword']}\nIntent: {$topic['searchIntent']}\nAudience: {$topic['audience']}\nAngle: {$topic['uniqueAngle']}\n\nReturn research notes: key facts with source URLs, exact syntax/menus/error messages where relevant, common questions people ask, pitfalls, and gaps in current top results. Plain text, concise bullet points.",
        null, 'low', 12000, 3);
    $system = seo_writer_system($lang);
    $brief = "TOPIC BRIEF\n" . json_encode($topic, JSON_UNESCAPED_UNICODE | JSON_PRETTY_PRINT)
        . "\n\nFORMAT: {$topic['format']} — " . SEO_FORMATS[$topic['format']]
        . "\n\nRELATED ARTICLES AND PAGES YOU MAY LINK TO (exact URLs):\n" . seo_link_candidates($lang, $topic['pillar'])
        . "\n\nRESEARCH NOTES:\n" . $research;
    $a = ai_call('write', $system, $brief . "\n\nWrite the article now.", seo_article_schema(), 'medium', 32000);
    $a['slug'] = blog_slugify($a['slug']) ?: blog_slugify($a['title']);
    $review = null;
    for ($round = 0; $round < 2; $round++) {
        $problems = seo_local_checks($a, $lang);
        $review = $problems ? null : seo_review($a, $topic, $lang);
        $ok = !$problems && $review['verdict'] === 'publish' && $review['score'] >= 80 && !$review['factualErrors'];
        if ($ok) break;
        if ($review && $review['verdict'] === 'reject') break;
        if ($round === 1) break;
        $feedback = $problems ? "Automatic checks failed:\n- " . implode("\n- ", $problems)
            : "Editor score {$review['score']}/100.\nFactual errors:\n- " . implode("\n- ", $review['factualErrors'] ?: ['none']) . "\nQuality issues:\n- " . implode("\n- ", $review['qualityIssues'] ?: ['none']) . "\nInstructions: " . $review['revisionInstructions'];
        seo_log("revise $lang: " . str_replace("\n", ' | ', mb_strimwidth($feedback, 0, 400, '…')));
        $a = ai_call('write', $system, $brief . "\n\nYOUR DRAFT:\n" . json_encode($a, JSON_UNESCAPED_UNICODE) . "\n\nFEEDBACK TO FIX:\n" . $feedback . "\n\nReturn the complete corrected article.", seo_article_schema(), 'medium', 32000);
        $a['slug'] = blog_slugify($a['slug']) ?: blog_slugify($a['title']);
    }
    $cost = round(ai_spend(date('Y-m-d')) - $spent0, 3);
    $problems = seo_local_checks($a, $lang);
    if ($problems || !$review || $review['verdict'] !== 'publish' || $review['score'] < 80 || $review['factualErrors']) {
        seo_topic_done($lang, $topic, 'discarded');
        $why = $problems ? implode('; ', $problems) : ($review ? "score {$review['score']}, verdict {$review['verdict']}, errors: " . implode('; ', $review['factualErrors']) : 'no review');
        seo_log("DISCARDED $lang \"{$topic['title']}\" — $why (cost \$$cost)");
        $drafts = seo_dir('drafts');
        json_write($drafts . '/' . date('Ymd-His') . "-$lang.json", ['topic' => $topic, 'article' => $a, 'review' => $review, 'problems' => $problems]);
        return null;
    }
    [$clean] = blog_sanitize($a['bodyHtml']);
    $now = date('c');
    $post = ['lang' => $lang, 'slug' => $a['slug'], 'title' => trim($a['title']), 'metaDescription' => trim($a['metaDescription']), 'excerpt' => trim($a['excerpt']),
        'primaryKeyword' => $a['primaryKeyword'], 'keywords' => array_slice($a['keywords'], 0, 8), 'pillar' => $topic['pillar'], 'format' => $topic['format'],
        'shot' => $a['shot'] === 'none' ? '' : $a['shot'], 'bodyHtml' => $clean, 'faq' => array_slice($a['faq'], 0, 6), 'wordCount' => blog_word_count($clean),
        'publishedAt' => $now, 'updatedAt' => $now, 'status' => 'published', 'review' => ['score' => $review['score'], 'issues' => $review['qualityIssues']],
        'costUsd' => $cost, 'topicId' => $topic['id'], 'source' => 'ai'];
    blog_save($post);
    seo_topic_done($lang, $topic, 'published');
    seo_log("PUBLISHED $lang /{$post['slug']} score {$review['score']} words {$post['wordCount']} cost \$$cost");
    $touched = [];
    try { $touched = seo_link_from_older($post); } catch (Throwable $e) { seo_log('internal links: ' . $e->getMessage()); }
    seo_indexnow(array_merge([blog_url($lang, $post['slug']), blog_url($lang), 'https://microads.vn/office/sitemap.xml'], $touched));
    return $post;
}

/**
 * Internal linking: adds one natural link to the new article from up to 2 older related articles (same language), so Google
 * discovers it quickly and understands the topic cluster. Changes one paragraph per older article; returns their URLs.
 */
function seo_link_from_older(array $new): array {
    $older = array_values(array_filter(blog_index_rebuild(), fn($p) => $p['lang'] === $new['lang'] && $p['slug'] !== $new['slug'] && ($p['status'] ?? '') === 'published'));
    usort($older, fn($a, $b) => [($b['pillar'] ?? '') === $new['pillar'], $b['publishedAt']] <=> [($a['pillar'] ?? '') === $new['pillar'], $a['publishedAt']]);
    $schema = ['type' => 'object', 'additionalProperties' => false, 'required' => ['index', 'html'], 'properties' => ['index' => ['type' => 'integer'], 'html' => ['type' => 'string']]];
    $target = blog_base($new['lang']) . $new['slug'];
    $urls = [];
    foreach (array_slice($older, 0, 2) as $o) {
        $post = blog_get($o['lang'], $o['slug']);
        if (!$post || str_contains($post['bodyHtml'], $target)) continue;
        preg_match_all('#<p>(.*?)</p>#s', $post['bodyHtml'], $m, PREG_OFFSET_CAPTURE);
        $paras = array_slice($m[0], 0, 40);
        if (count($paras) < 3) continue;
        $list = '';
        foreach ($paras as $i => [$html]) $list .= "[$i] $html\n";
        $r = ai_call('interlink', 'You add one internal link to an existing article. Pick the paragraph where a link to the new article helps the reader most, and return that paragraph rewritten with the smallest possible change: keep all existing text and markup, add one <a href="' . $target . '">descriptive anchor text</a> (anchor 2–6 words, natural, same language, not "click here"). If no paragraph fits naturally, return index -1.',
            "New article: {$new['title']} — {$new['excerpt']}\n\nParagraphs of \"{$post['title']}\":\n$list", $schema, 'low', 4000);
        $i = (int)$r['index'];
        if ($i < 0 || !isset($paras[$i]) || !str_contains($r['html'], $target) || substr_count($r['html'], '<a ') > substr_count($paras[$i][0], '<a ') + 1) continue;
        [$newPara] = blog_sanitize($r['html']);
        if (!str_starts_with($newPara, '<p>') || abs(mb_strlen(strip_tags($newPara)) - mb_strlen(strip_tags($paras[$i][0]))) > 160) continue;
        $post['bodyHtml'] = substr_replace($post['bodyHtml'], $newPara, $paras[$i][1], strlen($paras[$i][0]));
        $post['updatedAt'] = date('c');
        blog_save($post);
        $urls[] = blog_url($post['lang'], $post['slug']);
        seo_log("interlink {$post['lang']} /{$post['slug']} → /{$new['slug']}");
    }
    return $urls;
}

// ---------------------------------------------------------------------------------------------------------------- site health (beyond content)

/** Uptime probe (every cron tick): home page + update manifest. Daily stats in MO_DATA/seo/uptime/<date>.json. */
function seo_uptime_probe(): void {
    $f = seo_dir('uptime') . '/' . date('Y-m-d') . '.json';
    $d = json_read($f, ['checks' => 0, 'fails' => 0, 'slow' => 0, 'maxMs' => 0, 'sumMs' => 0, 'failures' => []]);
    foreach (['https://microads.vn/office/', 'https://microads.vn/office/api/manifest.json'] as $u) {
        $r = seo_http_get($u, [], 15);
        $d['checks']++;
        $d['sumMs'] += $r['ms'];
        $d['maxMs'] = max($d['maxMs'], $r['ms']);
        if (!in_array($r['status'], [200, 302], true)) { $d['fails']++; $d['failures'][] = date('H:i') . " $u HTTP {$r['status']} {$r['error']}"; seo_log("DOWN $u HTTP {$r['status']} {$r['error']}"); }
        elseif ($r['ms'] > 2000) $d['slow']++;
    }
    $d['failures'] = array_slice($d['failures'], -30);
    json_write($f, $d);
}

/** Lab + field performance from Google PageSpeed Insights (mobile), and SEO/accessibility/best-practice scores. */
function seo_pagespeed(string $url): array {
    $api = 'https://www.googleapis.com/pagespeedonline/v5/runPagespeed?strategy=mobile&url=' . rawurlencode($url)
        . '&category=performance&category=seo&category=accessibility&category=best-practices';
    $r = seo_http_get($api, [], 90);
    $j = json_decode($r['body'], true);
    if (!is_array($j) || empty($j['lighthouseResult'])) return ['url' => $url, 'error' => 'HTTP ' . $r['status'] . ' ' . mb_strimwidth($r['body'], 0, 160)];
    $lr = $j['lighthouseResult'];
    $cat = fn($k) => isset($lr['categories'][$k]['score']) ? (int)round($lr['categories'][$k]['score'] * 100) : null;
    $num = fn($k) => $lr['audits'][$k]['numericValue'] ?? null;
    $failed = [];
    foreach ($lr['audits'] as $id => $a) if (($a['score'] ?? 1) !== null && ($a['score'] ?? 1) < 0.5 && ($a['scoreDisplayMode'] ?? '') === 'binary') $failed[] = $id;
    $field = $j['loadingExperience']['metrics'] ?? [];
    return ['url' => $url, 'performance' => $cat('performance'), 'seo' => $cat('seo'), 'accessibility' => $cat('accessibility'), 'bestPractices' => $cat('best-practices'),
        'lcpMs' => $num('largest-contentful-paint'), 'cls' => $num('cumulative-layout-shift'), 'tbtMs' => $num('total-blocking-time'),
        'fieldLcp' => $field['LARGEST_CONTENTFUL_PAINT_MS']['category'] ?? null, 'fieldInp' => $field['INTERACTION_TO_NEXT_PAINT']['category'] ?? null,
        'fieldCls' => $field['CUMULATIVE_LAYOUT_SHIFT_SCORE']['category'] ?? null, 'failedAudits' => array_slice($failed, 0, 12)];
}

/** HTTPS certificate expiry and security headers of the site. */
function seo_security(): array {
    $out = ['certDaysLeft' => null, 'headers' => []];
    $ctx = stream_context_create(['ssl' => ['capture_peer_cert' => true, 'SNI_enabled' => true, 'peer_name' => 'microads.vn']]);
    $s = @stream_socket_client('ssl://microads.vn:443', $errno, $errstr, 15, STREAM_CLIENT_CONNECT, $ctx);
    if ($s) {
        $cert = openssl_x509_parse(stream_context_get_params($s)['options']['ssl']['peer_certificate'] ?? null);
        if ($cert) $out['certDaysLeft'] = (int)floor(($cert['validTo_time_t'] - time()) / 86400);
        fclose($s);
    }
    $h = seo_http_get('https://microads.vn/office/en/')['headers'];
    foreach (['strict-transport-security', 'x-content-type-options', 'referrer-policy', 'x-frame-options', 'content-security-policy', 'content-encoding', 'cache-control'] as $k) $out['headers'][$k] = $h[$k] ?? null;
    return $out;
}

/** Yesterday's visits per page (own counter), referrers and downloads (difference of the download counters). */
function seo_traffic(): array {
    $day = date('Y-m-d', strtotime('-1 day'));
    $views = json_read(seo_dir('views') . "/$day.json", []);
    $refs = [];
    foreach ($views as $k => $v) if (str_starts_with($k, '_ref:')) { $refs[substr($k, 5)] = $v; unset($views[$k]); }
    arsort($views); arsort($refs);
    $snapF = seo_dir() . '/downloads-snap.json';
    $snap = json_read($snapF, []);
    $now = json_read(MO_DATA . '/counters.json', []);
    $prev = $snap['counters'] ?? $now;
    $dl = [];
    foreach ($now as $k => $v) if ($v - ($prev[$k] ?? 0) > 0) $dl[$k] = $v - ($prev[$k] ?? 0);
    json_write($snapF, ['date' => date('Y-m-d'), 'counters' => $now]);
    $total = array_sum($views);
    return ['day' => $day, 'views' => $total, 'topPages' => array_slice($views, 0, 15, true), 'referrers' => array_slice($refs, 0, 15, true),
        'downloadsSinceLastAudit' => $dl, 'conversion' => $total ? round(array_sum($dl) / $total, 3) : null];
}

/** Improves an existing article (depth, accuracy, title/meta) — used by the daily audit for pages Google crawled but did not index. */
function seo_refresh(string $lang, string $slug, string $reason): bool {
    $post = blog_get($lang, $slug);
    if (!$post) return false;
    seo_log("refresh $lang /$slug: $reason");
    $topic = ['title' => $post['title'], 'primaryKeyword' => $post['primaryKeyword'], 'pillar' => $post['pillar'], 'format' => $post['format'], 'searchIntent' => 'informational', 'audience' => '', 'uniqueAngle' => 'Refresh: ' . $reason];
    $system = seo_writer_system($lang);
    $a = ai_call('refresh', $system, "Improve this published article so it deserves to be indexed and ranked. Reason: $reason\nKeep the same slug (\"{$post['slug']}\") and topic; add missing depth and examples, fix anything vague or outdated, keep internal links valid.\n\nRELATED ARTICLES AND PAGES YOU MAY LINK TO:\n" . seo_link_candidates($lang, $post['pillar'])
        . "\n\nCURRENT ARTICLE:\n" . json_encode(['title' => $post['title'], 'metaDescription' => $post['metaDescription'], 'slug' => $post['slug'], 'excerpt' => $post['excerpt'], 'primaryKeyword' => $post['primaryKeyword'], 'keywords' => $post['keywords'], 'shot' => $post['shot'] ?: 'none', 'bodyHtml' => $post['bodyHtml'], 'faq' => $post['faq']], JSON_UNESCAPED_UNICODE),
        seo_article_schema(), 'high', 32000);
    $a['slug'] = $post['slug'];
    $problems = seo_local_checks($a, $lang, $post['slug']);
    $review = $problems ? null : seo_review($a, $topic, $lang);
    if ($problems || $review['score'] < 80 || $review['factualErrors'] || $review['score'] < ($post['review']['score'] ?? 0)) {
        seo_log("refresh $lang /$slug kept old version (" . ($problems ? implode('; ', $problems) : "score {$review['score']}") . ')');
        return false;
    }
    [$clean] = blog_sanitize($a['bodyHtml']);
    $post['history'][] = ['at' => date('c'), 'title' => $post['title'], 'reason' => $reason];
    $post = array_merge($post, ['title' => trim($a['title']), 'metaDescription' => trim($a['metaDescription']), 'excerpt' => trim($a['excerpt']), 'keywords' => array_slice($a['keywords'], 0, 8),
        'shot' => $a['shot'] === 'none' ? '' : $a['shot'], 'bodyHtml' => $clean, 'faq' => array_slice($a['faq'], 0, 6), 'wordCount' => blog_word_count($clean), 'updatedAt' => date('c'),
        'review' => ['score' => $review['score'], 'issues' => $review['qualityIssues']], 'refreshedAt' => date('c')]);
    blog_save($post);
    seo_indexnow([blog_url($lang, $slug)]);
    seo_log("refreshed $lang /$slug score {$review['score']}");
    return true;
}

// ---------------------------------------------------------------------------------------------------------------- HTTP, IndexNow, GSC

function seo_http_get(string $url, array $headers = [], int $timeout = 25): array {
    $ch = curl_init($url);
    $respHeaders = [];
    curl_setopt_array($ch, [CURLOPT_RETURNTRANSFER => true, CURLOPT_FOLLOWLOCATION => false, CURLOPT_TIMEOUT => $timeout, CURLOPT_USERAGENT => 'MicroadsOfficeSeoCheck/1.0 (+https://microads.vn/office/)',
        CURLOPT_HTTPHEADER => $headers, CURLOPT_HEADERFUNCTION => function ($c, $h) use (&$respHeaders) { $p = explode(':', $h, 2); if (count($p) === 2) $respHeaders[strtolower(trim($p[0]))] = trim($p[1]); return strlen($h); }]);
    $t = microtime(true);
    $body = curl_exec($ch);
    $r = ['status' => (int)curl_getinfo($ch, CURLINFO_RESPONSE_CODE), 'body' => is_string($body) ? $body : '', 'headers' => $respHeaders, 'ms' => (int)((microtime(true) - $t) * 1000), 'error' => curl_error($ch)];
    curl_close($ch);
    return $r;
}

function seo_http_json(string $method, string $url, ?array $body, array $headers = []): array {
    $ch = curl_init($url);
    curl_setopt_array($ch, [CURLOPT_RETURNTRANSFER => true, CURLOPT_CUSTOMREQUEST => $method, CURLOPT_TIMEOUT => 40,
        CURLOPT_HTTPHEADER => array_merge(['Content-Type: application/json; charset=utf-8'], $headers)]);
    if ($body !== null) curl_setopt($ch, CURLOPT_POSTFIELDS, json_encode($body, JSON_UNESCAPED_SLASHES));
    $raw = curl_exec($ch);
    $code = (int)curl_getinfo($ch, CURLINFO_RESPONSE_CODE);
    curl_close($ch);
    return ['status' => $code, 'json' => json_decode((string)$raw, true), 'raw' => (string)$raw];
}

/** IndexNow: instant notification to Bing, Yandex, Seznam, Naver… (Google does not use IndexNow; it reads the sitemap). */
function seo_indexnow(array $urls): void {
    $key = blog_indexnow_key();
    $r = seo_http_json('POST', 'https://api.indexnow.org/indexnow', ['host' => 'microads.vn', 'key' => $key, 'keyLocation' => "https://microads.vn/office/$key.txt", 'urlList' => array_values(array_unique($urls))]);
    seo_log('indexnow ' . count($urls) . ' url(s): HTTP ' . $r['status']);
}

const GSC_SITE = 'https://microads.vn/office/';

function gsc_creds(): ?array { $c = json_read(MO_DATA . '/secrets/gsc.json', []); return !empty($c['client_email']) && !empty($c['private_key']) ? $c : null; }

function gsc_token(): ?string {
    static $tok = null;
    if ($tok) return $tok;
    $c = gsc_creds();
    if (!$c) return null;
    $b64 = fn($s) => rtrim(strtr(base64_encode($s), '+/', '-_'), '=');
    $now = time();
    $jwt = $b64(json_encode(['alg' => 'RS256', 'typ' => 'JWT'])) . '.' . $b64(json_encode(['iss' => $c['client_email'], 'scope' => 'https://www.googleapis.com/auth/webmasters',
        'aud' => 'https://oauth2.googleapis.com/token', 'iat' => $now, 'exp' => $now + 3600]));
    if (!openssl_sign($jwt, $sig, $c['private_key'], OPENSSL_ALGO_SHA256)) throw new RuntimeException('GSC: cannot sign with the service-account key');
    $ch = curl_init('https://oauth2.googleapis.com/token');
    curl_setopt_array($ch, [CURLOPT_RETURNTRANSFER => true, CURLOPT_POST => true, CURLOPT_TIMEOUT => 30,
        CURLOPT_POSTFIELDS => http_build_query(['grant_type' => 'urn:ietf:params:oauth:grant-type:jwt-bearer', 'assertion' => $jwt . '.' . $b64($sig)])]);
    $res = json_decode((string)curl_exec($ch), true);
    curl_close($ch);
    if (empty($res['access_token'])) throw new RuntimeException('GSC: token refused — ' . json_encode($res));
    return $tok = $res['access_token'];
}

function gsc_call(string $method, string $url, ?array $body = null): array {
    $tok = gsc_token();
    if (!$tok) return ['status' => 0, 'json' => null];
    return seo_http_json($method, $url, $body, ['Authorization: Bearer ' . $tok]);
}

function gsc_inspect(string $url): array {
    $r = gsc_call('POST', 'https://searchconsole.googleapis.com/v1/urlInspection/index:inspect', ['inspectionUrl' => $url, 'siteUrl' => GSC_SITE, 'languageCode' => 'en-US']);
    $s = $r['json']['inspectionResult']['indexStatusResult'] ?? [];
    return ['http' => $r['status'], 'verdict' => $s['verdict'] ?? '', 'coverage' => $s['coverageState'] ?? ($r['json']['error']['message'] ?? ''), 'lastCrawl' => $s['lastCrawlTime'] ?? '',
        'robots' => $s['robotsTxtState'] ?? '', 'indexing' => $s['indexingState'] ?? '', 'fetch' => $s['pageFetchState'] ?? '', 'googleCanonical' => $s['googleCanonical'] ?? ''];
}

function gsc_search(int $days, array $dims, int $limit = 100): array {
    $r = gsc_call('POST', 'https://www.googleapis.com/webmasters/v3/sites/' . rawurlencode(GSC_SITE) . '/searchAnalytics/query',
        ['startDate' => date('Y-m-d', strtotime("-$days days")), 'endDate' => date('Y-m-d', strtotime('-1 day')), 'dimensions' => $dims, 'rowLimit' => $limit]);
    return $r['json']['rows'] ?? [];
}

function gsc_submit_sitemap(): int {
    return gsc_call('PUT', 'https://www.googleapis.com/webmasters/v3/sites/' . rawurlencode(GSC_SITE) . '/sitemaps/' . rawurlencode('https://microads.vn/office/sitemap.xml'))['status'];
}

// ---------------------------------------------------------------------------------------------------------------- audit

/** Checks one page the way a crawler sees it. */
function seo_check_page(string $url, bool $expectIndexable = true): array {
    $r = seo_http_get($url);
    $issues = [];
    $info = ['url' => $url, 'status' => $r['status'], 'ms' => $r['ms'], 'kb' => (int)(strlen($r['body']) / 1024)];
    if ($r['status'] !== 200) { $issues[] = "HTTP {$r['status']}" . ($r['error'] ? " ({$r['error']})" : ''); return $info + ['issues' => $issues, 'links' => []]; }
    if ($r['ms'] > 1500) $issues[] = "slow response {$r['ms']} ms";
    if (stripos($r['headers']['x-robots-tag'] ?? '', 'noindex') !== false) $issues[] = 'X-Robots-Tag noindex';
    $doc = new DOMDocument();
    libxml_use_internal_errors(true);
    $doc->loadHTML('<?xml encoding="UTF-8">' . $r['body']);
    libxml_clear_errors();
    $xp = new DOMXPath($doc);
    $title = trim((string)$xp->evaluate('string(//title)'));
    $desc = trim((string)$xp->evaluate('string(//meta[@name="description"]/@content)'));
    $robots = (string)$xp->evaluate('string(//meta[@name="robots"]/@content)');
    $canon = (string)$xp->evaluate('string(//link[@rel="canonical"]/@href)');
    $h1 = $xp->query('//h1')->length;
    $info += ['title' => $title, 'titleLen' => mb_strlen($title), 'descLen' => mb_strlen($desc), 'canonical' => $canon, 'h1' => $h1];
    if ($expectIndexable && stripos($robots, 'noindex') !== false) $issues[] = 'meta robots noindex';
    if ($canon === '') $issues[] = 'missing canonical';
    elseif ($expectIndexable && rtrim(strtok($canon, '?'), '/') !== rtrim(strtok($url, '?'), '/')) $issues[] = "canonical points elsewhere ($canon)";
    if ($title === '' || mb_strlen($title) > 75) $issues[] = 'title missing or too long (' . mb_strlen($title) . ')';
    if (mb_strlen($desc) < 70 || mb_strlen($desc) > 170) $issues[] = 'meta description length ' . mb_strlen($desc);
    if ($h1 !== 1) $issues[] = "$h1 <h1> elements";
    foreach ($xp->query('//img[not(@alt)]') as $img) { $issues[] = 'image without alt'; break; }
    $links = [];
    foreach ($xp->query('//a[@href]') as $a) {
        $href = $a->getAttribute('href');
        if (str_starts_with($href, '/office/')) $href = 'https://microads.vn' . $href;
        if (str_starts_with($href, 'https://microads.vn/office/') && !str_contains($href, '/download/')) $links[] = strtok($href, '#');
    }
    return $info + ['issues' => $issues, 'links' => array_values(array_unique($links))];
}

function seo_audit(): array {
    $t0 = time();
    $report = ['date' => date('Y-m-d'), 'startedAt' => date('c'), 'pages' => [], 'site' => [], 'gsc' => null, 'actions' => []];
    $posts = blog_index_rebuild();
    $published = array_values(array_filter($posts, fn($p) => ($p['status'] ?? '') === 'published'));
    $urls = ['https://microads.vn/office/', 'https://microads.vn/office/en/', blog_url('vi'), blog_url('en'), 'https://microads.vn' . tool_path('vi'), 'https://microads.vn' . tool_path('en')];
    foreach ($published as $p) $urls[] = blog_url($p['lang'], $p['slug']);
    $allLinks = [];
    foreach ($urls as $u) { $c = seo_check_page($u); $report['pages'][] = $c; $allLinks = array_merge($allLinks, $c['links']); }
    // broken internal links
    $broken = [];
    foreach (array_diff(array_unique($allLinks), $urls) as $l) {
        if (str_contains($l, '?')) continue;
        $s = seo_http_get($l)['status'];
        if ($s >= 400) $broken[] = "$l → HTTP $s";
    }
    $report['site']['brokenLinks'] = $broken;
    // robots.txt + sitemap
    $robots = seo_http_get('https://microads.vn/robots.txt')['body'];
    $report['site']['robotsBlocksOffice'] = (bool)preg_match('#^\s*Disallow:\s*/office(/|\s|$)#mi', $robots);
    $report['site']['robotsListsSitemap'] = str_contains($robots, 'https://microads.vn/office/sitemap.xml');
    $sm = seo_http_get('https://microads.vn/office/sitemap.xml');
    preg_match_all('#<loc>([^<]+)</loc>#', $sm['body'], $m);
    $inSitemap = $m[1] ?? [];
    $report['site']['sitemapStatus'] = $sm['status'];
    $report['site']['sitemapUrls'] = count($inSitemap);
    $report['site']['missingFromSitemap'] = array_values(array_diff(array_map(fn($p) => blog_url($p['lang'], $p['slug']), $published), $inSitemap));
    // Google Search Console
    if (gsc_creds()) {
        try {
            $g = ['sitemapSubmit' => gsc_submit_sitemap(), 'inspections' => [], 'pages' => [], 'queries' => []];
            $toInspect = array_slice(array_values(array_filter($published, fn($p) => strtotime($p['publishedAt']) > time() - 60 * 86400)), 0, 60);
            foreach (array_merge(['https://microads.vn/office/', 'https://microads.vn/office/en/'], array_map(fn($p) => blog_url($p['lang'], $p['slug']), $toInspect)) as $u) $g['inspections'][$u] = gsc_inspect($u);
            foreach (gsc_search(28, ['page'], 100) as $row) $g['pages'][] = ['page' => $row['keys'][0], 'clicks' => $row['clicks'], 'impressions' => $row['impressions'], 'ctr' => round($row['ctr'], 4), 'position' => round($row['position'], 1)];
            foreach (gsc_search(28, ['query'], 50) as $row) $g['queries'][] = ['query' => $row['keys'][0], 'clicks' => $row['clicks'], 'impressions' => $row['impressions'], 'position' => round($row['position'], 1)];
            $report['gsc'] = $g;
        } catch (Throwable $e) { $report['gsc'] = ['error' => $e->getMessage()]; }
    }
    // beyond content: speed (PageSpeed Insights), uptime, HTTPS/security headers, visits and downloads
    $psiUrls = ['https://microads.vn/office/', 'https://microads.vn/office/en/'];
    if ($published) $psiUrls[] = blog_url($published[0]['lang'], $published[0]['slug']);
    $report['performance'] = array_map(fn($u) => seo_pagespeed($u), $psiUrls);
    $report['uptime'] = json_read(seo_dir('uptime') . '/' . date('Y-m-d', strtotime('-1 day')) . '.json', []) ?: json_read(seo_dir('uptime') . '/' . date('Y-m-d') . '.json', []);
    unset($report['uptime']['sumMs']);
    $report['security'] = seo_security();
    $report['traffic'] = seo_traffic();
    $report['cost'] = ['today' => round(ai_spend(date('Y-m-d')), 2), 'month' => round(ai_spend(date('Y-m')), 2), 'dailyBudget' => blog_settings()['dailyBudgetUsd'], 'avgArticle' => round(ai_article_cost(), 2)];
    $report['key'] = ['daysLeft' => ai_key_days_left()];
    $report['content'] = ['vi' => count(array_filter($published, fn($p) => $p['lang'] === 'vi')), 'en' => count(array_filter($published, fn($p) => $p['lang'] === 'en')),
        'backlog' => ['vi' => count(seo_topics()['vi']), 'en' => count(seo_topics()['en'])]];
    // Claude reads everything and decides on (limited) actions
    $summary = ['pagesWithIssues' => array_values(array_filter(array_map(fn($p) => $p['issues'] ? ['url' => $p['url'], 'issues' => $p['issues']] : null, $report['pages']))),
        'site' => $report['site'], 'gsc' => $report['gsc'], 'content' => $report['content'], 'cost' => $report['cost'], 'keyDaysLeft' => $report['key']['daysLeft'],
        'performance' => $report['performance'], 'uptime' => $report['uptime'], 'security' => $report['security'], 'traffic' => $report['traffic'],
        'recentReports' => array_map(fn($f) => ['date' => basename($f, '.json'), 'health' => json_read($f, [])['analysis']['health'] ?? '', 'summary' => mb_strimwidth(json_read($f, [])['analysis']['summaryVi'] ?? '', 0, 400, '…')], array_slice(array_reverse(glob(seo_dir('reports') . '/*.json') ?: []), 0, 5)),
        'posts' => array_map(fn($p) => ['lang' => $p['lang'], 'slug' => $p['slug'], 'title' => $p['title'], 'publishedAt' => substr($p['publishedAt'], 0, 10), 'refreshedAt' => substr($p['refreshedAt'] ?? '', 0, 10), 'score' => $p['review']['score'] ?? null], array_slice($published, 0, 120))];
    $schema = ['type' => 'object', 'additionalProperties' => false, 'required' => ['health', 'summaryVi', 'issues', 'actions'],
        'properties' => ['health' => ['type' => 'string', 'enum' => ['good', 'warning', 'bad']], 'summaryVi' => ['type' => 'string'],
            'issues' => ['type' => 'array', 'items' => ['type' => 'object', 'additionalProperties' => false, 'required' => ['severity', 'url', 'problemVi', 'fixVi'],
                'properties' => ['severity' => ['type' => 'string', 'enum' => ['high', 'medium', 'low']], 'url' => ['type' => 'string'], 'problemVi' => ['type' => 'string'], 'fixVi' => ['type' => 'string']]]],
            'actions' => ['type' => 'array', 'items' => ['type' => 'object', 'additionalProperties' => false, 'required' => ['type', 'lang', 'slug', 'reason'],
                'properties' => ['type' => ['type' => 'string', 'enum' => ['refresh_post', 'rewrite_meta']], 'lang' => ['type' => 'string', 'enum' => ['vi', 'en']], 'slug' => ['type' => 'string'], 'reason' => ['type' => 'string']]]]]];
    try {
        $analysis = ai_call('audit', "You are the technical SEO lead for https://microads.vn/office/ (a free office-suite download site with a bilingual blog, new domain section, launched November 2026). Each day you read the crawl report and Google Search Console data and decide what to fix. Be concrete and honest: explain WHY a page is not indexed (e.g. 'Discovered – currently not indexed' = Google has not crawled it yet, needs internal links/time; 'Crawled – currently not indexed' = Google judged the page not valuable enough → improve content; 'Duplicate, Google chose different canonical'; robots/noindex/404/5xx = technical). Look beyond indexing too: speed and Core Web Vitals (PageSpeed: LCP > 2.5 s, CLS > 0.1, TBT > 200 ms are problems), uptime failures, HTTPS certificate expiry (< 20 days), missing security/compression/cache headers, traffic and download conversion trends, referrers, and the trend versus the recent reports. Technical problems that need a code change go into issues with a precise fixVi for the developer. Write summaryVi (max 8 sentences: what changed since yesterday, the most important problem, what the system did) and issues in Vietnamese for the site owner.\nActions you may request (at most 2): refresh_post only for articles older than 10 days that Google crawled but did not index, or that get impressions but rank beyond position 15; rewrite_meta for articles with ≥ 100 impressions and CTR below 1%. Never refresh an article refreshed in the last 30 days. If nothing needs action, return an empty actions list.",
            json_encode($summary, JSON_UNESCAPED_UNICODE | JSON_UNESCAPED_SLASHES), $schema, 'medium', 16000);
    } catch (Throwable $e) {
        $analysis = ['health' => 'warning', 'summaryVi' => 'Không chạy được phần phân tích AI: ' . $e->getMessage(), 'issues' => [], 'actions' => []];
    }
    $report['analysis'] = $analysis;
    $s = blog_settings();
    $done = 0;
    foreach ($analysis['actions'] as $act) {
        if ($done >= (int)$s['refreshPerDay'] + 1) break;
        $post = blog_get($act['lang'], $act['slug']);
        if (!$post || strtotime($post['publishedAt']) > time() - 10 * 86400 || (!empty($post['refreshedAt']) && strtotime($post['refreshedAt']) > time() - 30 * 86400)) continue;
        if (ai_spend(date('Y-m-d')) + ai_article_cost() > (float)$s['dailyBudgetUsd'] + 0.3) break; // refreshes cost about one article
        try {
            $ok = $act['type'] === 'rewrite_meta' ? seo_rewrite_meta($post, $act['reason']) : seo_refresh($act['lang'], $act['slug'], $act['reason']);
            $report['actions'][] = $act + ['ok' => $ok];
            $done++;
        } catch (Throwable $e) { $report['actions'][] = $act + ['ok' => false, 'error' => $e->getMessage()]; }
    }
    $report['seconds'] = time() - $t0;
    json_write(seo_dir('reports') . '/' . date('Y-m-d') . '.json', $report);
    seo_log('audit done: health ' . $analysis['health'] . ', ' . count($analysis['issues']) . ' issue(s), ' . count($report['actions']) . ' action(s)');
    return $report;
}

function seo_rewrite_meta(array $post, string $reason): bool {
    $schema = ['type' => 'object', 'additionalProperties' => false, 'required' => ['title', 'metaDescription'], 'properties' => ['title' => ['type' => 'string'], 'metaDescription' => ['type' => 'string']]];
    $r = ai_call('meta', 'You improve search snippets. Title ≤ 60 characters with the primary keyword, specific and honest (no clickbait). Meta description 140–158 characters with the concrete value. Same language as the article.',
        json_encode(['reason' => $reason, 'primaryKeyword' => $post['primaryKeyword'], 'title' => $post['title'], 'metaDescription' => $post['metaDescription'], 'excerpt' => $post['excerpt']], JSON_UNESCAPED_UNICODE), $schema, 'medium', 4000);
    if (mb_strlen($r['title']) > 70 || mb_strlen($r['metaDescription']) < 110 || mb_strlen($r['metaDescription']) > 165) return false;
    $post['history'][] = ['at' => date('c'), 'title' => $post['title'], 'metaDescription' => $post['metaDescription'], 'reason' => 'meta: ' . $reason];
    $post['title'] = $r['title']; $post['metaDescription'] = $r['metaDescription']; $post['updatedAt'] = date('c'); $post['refreshedAt'] = date('c');
    blog_save($post);
    seo_log("meta rewritten {$post['lang']} /{$post['slug']}");
    return true;
}

// ---------------------------------------------------------------------------------------------------------------- daily schedule

/** Today's publishing plan: a random number of articles (min..max), split between the languages, at random times 07:30–21:30. */
function seo_schedule(): array {
    $path = BLOG_DIR . '/schedule.json';
    $s = json_read($path, []);
    if (($s['date'] ?? '') === date('Y-m-d')) return $s;
    $set = blog_settings();
    $n = random_int((int)$set['minPerDay'], max((int)$set['minPerDay'], (int)$set['maxPerDay']));
    $en = intdiv($n, 2) + ($n % 2 ? random_int(0, 1) : 0);
    $langs = array_merge(array_fill(0, $en, 'en'), array_fill(0, $n - $en, 'vi'));
    shuffle($langs);
    $start = strtotime(date('Y-m-d') . ' 07:30'); $end = strtotime(date('Y-m-d') . ' 21:30');
    $times = [];
    while (count($times) < $n) {
        $t = random_int($start, $end);
        $okGap = true;
        foreach ($times as $x) if (abs($x - $t) < 50 * 60) { $okGap = false; break; }
        if ($okGap) $times[] = $t;
    }
    sort($times);
    $slots = [];
    foreach ($langs as $i => $l) $slots[] = ['lang' => $l, 'at' => date('c', $times[$i]), 'done' => false, 'result' => ''];
    $s = ['date' => date('Y-m-d'), 'slots' => $slots];
    if (!is_dir(BLOG_DIR)) mkdir(BLOG_DIR, 0750, true);
    json_write($path, $s);
    seo_log("schedule " . date('Y-m-d') . ": $n article(s) — " . implode(', ', array_map(fn($x) => $x['lang'] . ' ' . substr($x['at'], 11, 5), $slots)));
    return $s;
}

function seo_tick(): void {
    $lock = fopen(seo_dir() . '/tick.lock', 'c');
    if (!$lock || !flock($lock, LOCK_EX | LOCK_NB)) { echo "busy\n"; return; }
    $set = blog_settings();
    $s = seo_schedule();
    // what articles may still spend today: daily budget − spent − reserve for tonight's audit − unused share of the formula tool
    $toolLeft = max(0.0, (float)$set['toolDailyBudgetUsd'] - ai_spend(date('Y-m-d'), 'tool'));
    $auditDone = is_file(seo_dir('reports') . '/' . date('Y-m-d') . '.json');
    $budgetLeft = (float)$set['dailyBudgetUsd'] - ai_spend(date('Y-m-d')) - ($auditDone ? 0 : (float)$set['auditReserveUsd']) - $toolLeft;
    $canWrite = $budgetLeft >= ai_article_cost() * 0.8;
    seo_uptime_probe();
    // "write now" requests from the admin page
    $now = json_read(seo_dir() . '/run-now.json', []);
    if ($now) {
        @unlink(seo_dir() . '/run-now.json');
        if (($now['action'] ?? '') === 'audit') seo_audit();
        elseif (in_array($now['lang'] ?? '', ['vi', 'en'], true) && $canWrite) { try { seo_write_one($now['lang']); } catch (Throwable $e) { seo_log('ERROR write-now: ' . $e->getMessage()); } }
    }
    if ($set['enabled'] && $canWrite) {
        foreach ($s['slots'] as $i => $slot) {
            if ($slot['done'] || strtotime($slot['at']) > time()) continue;
            $s['slots'][$i]['done'] = true;
            json_write(BLOG_DIR . '/schedule.json', $s);
            try { $p = seo_write_one($slot['lang']); $s['slots'][$i]['result'] = $p ? '/' . $p['slug'] : 'discarded'; }
            catch (Throwable $e) { $s['slots'][$i]['result'] = 'error: ' . $e->getMessage(); seo_log('ERROR write: ' . $e->getMessage()); }
            json_write(BLOG_DIR . '/schedule.json', $s);
            break; // one article per tick
        }
    } elseif ($set['enabled'] && !$canWrite) {
        foreach ($s['slots'] as $i => $slot) if (!$slot['done'] && strtotime($slot['at']) <= time()) {
            $s['slots'][$i]['done'] = true; $s['slots'][$i]['result'] = 'skipped: daily budget';
            json_write(BLOG_DIR . '/schedule.json', $s);
            seo_log('slot ' . $slot['lang'] . ' ' . substr($slot['at'], 11, 5) . ' skipped — daily budget $' . $set['dailyBudgetUsd'] . ' reached');
        }
    }
    // daily audit after the last slot (22:00 or later)
    if ((int)date('G') >= 22 && !is_file(seo_dir('reports') . '/' . date('Y-m-d') . '.json')) {
        try { seo_audit(); } catch (Throwable $e) { seo_log('ERROR audit: ' . $e->getMessage()); }
    }
    flock($lock, LOCK_UN);
}
