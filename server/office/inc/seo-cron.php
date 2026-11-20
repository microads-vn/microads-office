<?php
// Command line for the blog writer and daily SEO care (cron runs "tick" every 10 minutes as www-data).
//   php seo-cron.php tick            publish due articles of today's plan, run the nightly audit
//   php seo-cron.php write vi|en     write and publish one article now
//   php seo-cron.php topics vi|en    refill the topic backlog
//   php seo-cron.php audit           run the SEO audit now
//   php seo-cron.php status          plan, spend, backlog
declare(strict_types=1);
if (PHP_SAPI !== 'cli') { http_response_code(404); exit; }
set_time_limit(0);
require __DIR__ . '/bootstrap.php';
require __DIR__ . '/blog.php';
require __DIR__ . '/ai.php';
require __DIR__ . '/seo.php';
$cmd = $argv[1] ?? '';
$lang = in_array($argv[2] ?? '', ['vi', 'en'], true) ? $argv[2] : 'vi';
try {
    switch ($cmd) {
        case 'tick': seo_tick(); break;
        case 'write': $p = seo_write_one($lang); echo $p ? "published " . blog_url($lang, $p['slug']) . "\n" : "discarded\n"; break;
        case 'topics': seo_plan_topics($lang); break;
        case 'audit': $r = seo_audit(); echo $r['analysis']['summaryVi'] ?? '', "\n"; break;
        case 'status':
            echo json_encode(['schedule' => json_read(BLOG_DIR . '/schedule.json', []), 'spendToday' => ai_spend(date('Y-m-d')), 'spendMonth' => ai_spend(date('Y-m')),
                'backlog' => ['vi' => count(seo_topics()['vi']), 'en' => count(seo_topics()['en'])], 'posts' => count(blog_index_rebuild()), 'keyDaysLeft' => ai_key_days_left()],
                JSON_PRETTY_PRINT | JSON_UNESCAPED_UNICODE | JSON_UNESCAPED_SLASHES), "\n";
            break;
        default: fwrite(STDERR, "usage: tick | write vi|en | topics vi|en | audit | status\n"); exit(2);
    }
} catch (Throwable $e) {
    seo_log('ERROR ' . $cmd . ': ' . $e->getMessage());
    exit(1);
}
