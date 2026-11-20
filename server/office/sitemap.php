<?php
// /office/sitemap.xml — home pages (with hreflang pairs), blog lists, tool pages and every published article.
declare(strict_types=1);
require __DIR__ . '/inc/bootstrap.php';
require __DIR__ . '/inc/blog.php';
header('Content-Type: application/xml; charset=utf-8');
header('X-Content-Type-Options: nosniff');
$posts = blog_list();
$last = fn(?string $lang = null) => substr((string)(array_values(array_filter($posts, fn($p) => $lang === null || $p['lang'] === $lang))[0]['updatedAt'] ?? date('c')), 0, 10);
$pair = fn(string $vi, string $en) => '<xhtml:link rel="alternate" hreflang="vi" href="' . h($vi) . '"/><xhtml:link rel="alternate" hreflang="en" href="' . h($en) . '"/><xhtml:link rel="alternate" hreflang="x-default" href="' . h($en) . '"/>';
$B = 'https://microads.vn';
echo '<?xml version="1.0" encoding="UTF-8"?>' . "\n";
echo '<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9" xmlns:xhtml="http://www.w3.org/1999/xhtml" xmlns:image="http://www.google.com/schemas/sitemap-image/1.1">' . "\n";
foreach ([['/office/', '/office/en/'], [blog_base('vi'), blog_base('en')], [tool_path('vi'), tool_path('en')]] as [$vi, $en]) {
    foreach ([$vi, $en] as $u) echo '<url><loc>' . h($B . $u) . '</loc><lastmod>' . $last() . '</lastmod>' . $pair($B . $vi, $B . $en) . "</url>\n";
}
foreach ($posts as $p) {
    echo '<url><loc>' . h(blog_url($p['lang'], $p['slug'])) . '</loc><lastmod>' . h(substr($p['updatedAt'] ?? $p['publishedAt'], 0, 10)) . '</lastmod>'
        . '<image:image><image:loc>' . h($B . '/office/assets/blog/' . $p['pillar'] . '-' . $p['lang'] . '.jpg') . '</image:loc></image:image>'
        . (!empty($p['shot']) ? '<image:image><image:loc>' . h($B . '/office/assets/blog/shot-' . $p['shot'] . '.jpg') . '</image:loc></image:image>' : '')
        . "</url>\n";
}
echo '<url><loc>' . $B . '/office/privacy</loc></url>' . "\n";
echo "</urlset>\n";
