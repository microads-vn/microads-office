<?php
// Blog: /office/blog/ (Vietnamese) and /office/en/blog/ (English) — list, article, RSS feed.
// nginx: /office/(en/)blog/[slug|feed.xml] → blog.php?hl=vi|en[&slug=…|&feed=1]
declare(strict_types=1);
require __DIR__ . '/inc/bootstrap.php';
require __DIR__ . '/inc/blog.php';
require __DIR__ . '/inc/layout.php';
$LANG = ($_GET['hl'] ?? 'vi') === 'en' ? 'en' : 'vi';
$c = load_config();
$store = store_links($c);
$slug = (string)($_GET['slug'] ?? '');
header('X-Content-Type-Options: nosniff');
header('Referrer-Policy: strict-origin-when-cross-origin');
$T = fn($vi, $en) => $LANG === 'en' ? $en : $vi;
$pillarName = fn(string $k) => BLOG_PILLARS[$k][$LANG] ?? $k;

// ---------------------------------------------------------------- RSS feed
if (isset($_GET['feed'])) {
    header('Content-Type: application/rss+xml; charset=utf-8');
    $items = array_slice(blog_list($LANG), 0, 30);
    echo '<?xml version="1.0" encoding="UTF-8"?>' . "\n";
    echo '<rss version="2.0" xmlns:atom="http://www.w3.org/2005/Atom"><channel>';
    echo '<title>' . h($T('Blog Microads Office', 'Microads Office Blog')) . '</title><link>' . h(blog_url($LANG)) . '</link>';
    echo '<atom:link href="' . h(blog_url($LANG) . 'feed.xml') . '" rel="self" type="application/rss+xml"/>';
    echo '<description>' . h($T('Hướng dẫn Word, Excel, PDF và chữ ký số từ Microads Office.', 'Guides on documents, Excel formulas, PDF and digital signatures from Microads Office.')) . '</description>';
    echo '<language>' . ($LANG === 'en' ? 'en' : 'vi') . '</language>';
    foreach ($items as $p) {
        echo '<item><title>' . h($p['title']) . '</title><link>' . h(blog_url($LANG, $p['slug'])) . '</link><guid isPermaLink="true">' . h(blog_url($LANG, $p['slug'])) . '</guid>';
        echo '<pubDate>' . date(DATE_RSS, strtotime($p['publishedAt'])) . '</pubDate><description>' . h($p['metaDescription'] ?? '') . '</description></item>';
    }
    echo '</channel></rss>';
    exit;
}

// ---------------------------------------------------------------- article
$post = null;
if ($slug !== '') {
    $post = blog_get($LANG, $slug);
    if (!$post || ($post['status'] ?? '') !== 'published') {
        http_response_code(404);
        $post = null;
        $notFound = true;
    }
}

if (empty($notFound)) track_view(blog_base($LANG) . ($post ? $post['slug'] : ''));
$posts = $post ? [] : blog_list($LANG);
$page = max(1, (int)($_GET['page'] ?? 1));
$perPage = 12;
if (!$post && empty($notFound)) {
    $pages = max(1, (int)ceil(count($posts) / $perPage));
    if ($page > $pages) { http_response_code(404); $notFound = true; }
    $pagePosts = array_slice($posts, ($page - 1) * $perPage, $perPage);
}

if ($post) {
    $title = $post['title'];
    $desc = $post['metaDescription'];
    $canonical = blog_url($LANG, $post['slug']);
    $image = 'https://microads.vn/office/assets/blog/' . $post['pillar'] . '-' . $LANG . '.jpg';
    [$body, $headings] = blog_sanitize($post['bodyHtml']);
    $body = blog_body_with_cta($body, $LANG);
    $words = (int)($post['wordCount'] ?? blog_word_count($post['bodyHtml']));
    $related = blog_related($post);
    $alt = ['vi' => blog_base('vi'), 'en' => blog_base('en')];
    $ld = ['@context' => 'https://schema.org', '@graph' => [
        ['@type' => 'BlogPosting', '@id' => $canonical . '#article', 'headline' => $title, 'description' => $desc, 'image' => [$image],
            'datePublished' => $post['publishedAt'], 'dateModified' => $post['updatedAt'] ?? $post['publishedAt'], 'inLanguage' => $LANG,
            'wordCount' => $words, 'articleSection' => $pillarName($post['pillar']), 'keywords' => implode(', ', array_merge([$post['primaryKeyword']], $post['keywords'] ?? [])),
            'mainEntityOfPage' => $canonical,
            'author' => ['@type' => 'Organization', 'name' => $T('Đội ngũ Microads Office', 'Microads Office team'), 'url' => 'https://microads.vn/office/'],
            'publisher' => ['@type' => 'Organization', 'name' => 'Microads', 'url' => 'https://microads.vn/', 'logo' => ['@type' => 'ImageObject', 'url' => 'https://microads.vn/office/assets/office.png']],
            'about' => ['@type' => 'SoftwareApplication', 'name' => 'Microads Office', 'operatingSystem' => 'Windows 10, Windows 11', 'applicationCategory' => 'BusinessApplication',
                'offers' => ['@type' => 'Offer', 'price' => '0', 'priceCurrency' => 'USD'], 'downloadUrl' => 'https://microads.vn/office/download/MicroadsOffice_Setup.exe']],
        ['@type' => 'BreadcrumbList', 'itemListElement' => [
            ['@type' => 'ListItem', 'position' => 1, 'name' => 'Microads Office', 'item' => 'https://microads.vn' . home_path($LANG)],
            ['@type' => 'ListItem', 'position' => 2, 'name' => 'Blog', 'item' => blog_url($LANG)],
            ['@type' => 'ListItem', 'position' => 3, 'name' => $title, 'item' => $canonical]]],
    ]];
} else {
    $title = !empty($notFound) ? $T('Không tìm thấy bài viết', 'Article not found') : $T('Blog Microads Office — Hướng dẫn Word, Excel, PDF, chữ ký số', 'Microads Office Blog — Excel formulas, documents, PDF signing guides');
    $desc = $T('Hướng dẫn thực tế về hàm Excel, soạn thảo văn bản, ký số PDF và dùng AI trong công việc văn phòng — kèm phần mềm Microads Office miễn phí, không cần tài khoản.',
        'Practical guides to Excel formulas, documents, PDF digital signatures and AI at work — with the free, no-account Microads Office suite for Windows.');
    $canonical = blog_url($LANG) . ($page > 1 ? '?page=' . $page : '');
    $image = 'https://microads.vn/office/assets/blog/free-office-' . $LANG . '.jpg';
    $alt = ['vi' => blog_base('vi'), 'en' => blog_base('en')];
}
?><!doctype html>
<html lang="<?= $LANG ?>">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title><?= h($post ? $title . ' | Microads Office' : $title) ?></title>
<meta name="description" content="<?= h($desc) ?>">
<?php if (!empty($notFound)): ?><meta name="robots" content="noindex">
<?php else: ?><link rel="canonical" href="<?= h($canonical) ?>">
<?php if (!$post): ?><link rel="alternate" hreflang="vi" href="<?= h(blog_url('vi')) ?>"><link rel="alternate" hreflang="en" href="<?= h(blog_url('en')) ?>"><link rel="alternate" hreflang="x-default" href="<?= h(blog_url('en')) ?>">
<?php endif; ?>
<?php endif; ?>
<link rel="alternate" type="application/rss+xml" title="<?= h($T('Blog Microads Office', 'Microads Office Blog')) ?>" href="<?= h(blog_base($LANG)) ?>feed.xml">
<meta property="og:type" content="<?= $post ? 'article' : 'website' ?>">
<meta property="og:site_name" content="Microads Office">
<meta property="og:title" content="<?= h($title) ?>">
<meta property="og:description" content="<?= h($desc) ?>">
<meta property="og:url" content="<?= h($canonical) ?>">
<meta property="og:image" content="<?= h($image) ?>">
<meta property="og:image:width" content="1200"><meta property="og:image:height" content="630">
<meta property="og:locale" content="<?= $LANG === 'en' ? 'en_US' : 'vi_VN' ?>">
<?php if ($post): ?><meta property="article:published_time" content="<?= h($post['publishedAt']) ?>"><meta property="article:modified_time" content="<?= h($post['updatedAt'] ?? $post['publishedAt']) ?>">
<?php endif; ?>
<meta name="twitter:card" content="summary_large_image">
<?php if ($post): ?><script type="application/ld+json"><?= json_encode($ld, JSON_UNESCAPED_SLASHES | JSON_UNESCAPED_UNICODE | JSON_HEX_TAG) ?></script>
<?php endif; ?>
<link rel="icon" href="/office/assets/office.png">
<link rel="stylesheet" href="/office/assets/style.css?v=10">
</head>
<body class="blog">
<?php layout_header($LANG, $alt); ?>
<main class="wrap blog-main">
<?php if (!empty($notFound)): ?>
  <section class="card"><h1><?= h($title) ?></h1><p><?= $T('Bài viết không tồn tại hoặc đã được gỡ.', 'This article does not exist or was removed.') ?> <a href="<?= h(blog_base($LANG)) ?>"><?= $T('Xem các bài viết khác', 'See other articles') ?> →</a></p></section>
<?php elseif ($post): ?>
  <nav class="crumbs" aria-label="Breadcrumb"><a href="<?= h(home_path($LANG)) ?>">Microads Office</a> › <a href="<?= h(blog_base($LANG)) ?>">Blog</a> › <span><?= h($pillarName($post['pillar'])) ?></span></nav>
  <article class="post">
    <header class="post-head">
      <p class="kicker"><?= h($pillarName($post['pillar'])) ?></p>
      <h1><?= h($title) ?></h1>
      <p class="post-meta"><?= $T('Đội ngũ Microads Office', 'Microads Office team') ?> · <time datetime="<?= h($post['publishedAt']) ?>"><?= h(blog_date($post['publishedAt'], $LANG)) ?></time>
        <?php if (!empty($post['updatedAt']) && substr($post['updatedAt'], 0, 10) !== substr($post['publishedAt'], 0, 10)): ?> · <?= $T('Cập nhật', 'Updated') ?> <time datetime="<?= h($post['updatedAt']) ?>"><?= h(blog_date($post['updatedAt'], $LANG)) ?></time><?php endif; ?>
        · <?= blog_reading_minutes($words) ?> <?= $T('phút đọc', 'min read') ?></p>
      <img class="cover" src="/office/assets/blog/<?= h($post['pillar']) ?>-<?= $LANG ?>.jpg" alt="<?= h($title) ?>" width="1200" height="630" fetchpriority="high">
    </header>
    <?php if (count($headings) >= 3): ?>
    <details class="toc" open><summary><?= $T('Nội dung bài viết', 'In this article') ?></summary><ol>
      <?php foreach ($headings as [$id, $text]): ?><li><a href="#<?= h($id) ?>"><?= h($text) ?></a></li><?php endforeach; ?>
    </ol></details>
    <?php endif; ?>
    <div class="post-body">
      <?= $body ?>
      <?php if (!empty($post['shot']) && isset(BLOG_SHOTS[$post['shot']])): ?>
      <figure class="shot"><img src="/office/assets/blog/shot-<?= h($post['shot']) ?>.jpg" alt="<?= h(BLOG_SHOTS[$post['shot']][$LANG]) ?>" width="1200" height="675" loading="lazy"><figcaption><?= h(BLOG_SHOTS[$post['shot']][$LANG]) ?></figcaption></figure>
      <?php endif; ?>
      <?php if (!empty($post['faq'])): ?>
      <section class="faq" id="faq"><h2><?= $T('Câu hỏi thường gặp', 'Frequently asked questions') ?></h2>
        <?php foreach ($post['faq'] as $f): ?><details><summary><?= h($f['q']) ?></summary><p><?= h($f['a']) ?></p></details><?php endforeach; ?>
      </section>
      <?php endif; ?>
    </div>
    <?= blog_cta($LANG, 'end') ?>
    <section class="about-box">
      <img src="/office/assets/office.png" alt="" width="56" height="56">
      <div><b><?= $T('Về Microads Office', 'About Microads Office') ?></b>
        <p><?= $T('Microads Office là bộ ứng dụng Word, Excel và PDF miễn phí cho Windows do Công ty TNHH Công nghệ Microads (TP. Hồ Chí Minh) phát triển, mã nguồn mở GPL-3.0. Bài viết do đội ngũ Microads Office biên soạn với sự hỗ trợ của AI và được kiểm tra theo tính năng thật của phần mềm.',
            'Microads Office is a free Word, Excel and PDF suite for Windows built by Microads Technology Co., Ltd. (Ho Chi Minh City, Vietnam) and open source under GPL-3.0. Articles are written by the Microads Office team with AI assistance and checked against the software’s real features.') ?>
          <a href="https://github.com/microads-vn/microads-office"><?= $T('Mã nguồn', 'Source code') ?></a> · <a href="<?= h(home_path($LANG)) ?>"><?= $T('Tải phần mềm', 'Download') ?></a></p></div>
    </section>
  </article>
  <?php if ($related): ?>
  <section class="related"><h2><?= $T('Bài viết liên quan', 'Related articles') ?></h2><div class="post-grid">
    <?php foreach ($related as $r): ?><a class="post-card" href="<?= h(blog_base($LANG) . $r['slug']) ?>"><img src="/office/assets/blog/<?= h($r['pillar']) ?>-<?= $LANG ?>.jpg" alt="" width="600" height="315" loading="lazy"><b><?= h($r['title']) ?></b><span><?= h($r['excerpt'] ?? '') ?></span></a><?php endforeach; ?>
  </div></section>
  <?php endif; ?>
<?php else: ?>
  <header class="blog-head">
    <h1><?= $T('Blog Microads Office', 'Microads Office Blog') ?></h1>
    <p class="lead-dark"><?= $T('Hướng dẫn thực tế về Excel, văn bản, PDF và chữ ký số — dùng được ngay với phần mềm miễn phí, không cần tài khoản.', 'Practical guides to Excel formulas, documents, PDFs and digital signatures — usable right away with free software that needs no account.') ?></p>
  </header>
  <?php if (!$posts): ?>
    <section class="card"><p><?= $T('Bài viết đầu tiên sắp được đăng.', 'The first articles are coming soon.') ?></p></section>
  <?php else: ?>
  <div class="post-grid big">
    <?php foreach ($pagePosts as $r): ?><a class="post-card" href="<?= h(blog_base($LANG) . $r['slug']) ?>"><img src="/office/assets/blog/<?= h($r['pillar']) ?>-<?= $LANG ?>.jpg" alt="" width="600" height="315" loading="lazy"><small><?= h($pillarName($r['pillar'])) ?> · <?= h(blog_date($r['publishedAt'], $LANG)) ?></small><b><?= h($r['title']) ?></b><span><?= h($r['excerpt'] ?? '') ?></span></a><?php endforeach; ?>
  </div>
  <?php if ($pages > 1): ?><nav class="pager"><?php for ($i = 1; $i <= $pages; $i++): ?><a href="<?= h(blog_base($LANG) . ($i > 1 ? '?page=' . $i : '')) ?>" class="<?= $i === $page ? 'on' : '' ?>"><?= $i ?></a><?php endfor; ?></nav><?php endif; ?>
  <?php endif; ?>
  <?= blog_cta($LANG, 'end') ?>
<?php endif; ?>
</main>
<?php layout_footer($LANG, $c['promo']); ?>
<?php layout_download_notice($LANG, $store); ?>
</body>
</html>
