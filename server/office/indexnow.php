<?php
// /office/<key>.txt — IndexNow key file (proves to Bing, Yandex, Seznam, Naver… that microads.vn/office submits its own URLs).
declare(strict_types=1);
require __DIR__ . '/inc/bootstrap.php';
require __DIR__ . '/inc/blog.php';
$key = blog_settings()['indexnowKey'];
if (!preg_match('/^[0-9a-f]{32}$/', $key) || !hash_equals($key, (string)($_GET['k'] ?? ''))) { http_response_code(404); exit('Not found'); }
header('Content-Type: text/plain; charset=utf-8');
echo $key;
