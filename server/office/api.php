<?php
// GET /office/api/manifest.json  → update manifest for Microads Office clients.
declare(strict_types=1);
require __DIR__ . '/inc/bootstrap.php';
header('Content-Type: application/json; charset=utf-8');
header('Cache-Control: no-store, max-age=0');
header('X-Content-Type-Options: nosniff');
header('Access-Control-Allow-Origin: *');
$c = load_config();
record_check((string)($_GET['v'] ?? ''));
echo json_encode(build_manifest($c), JSON_UNESCAPED_UNICODE | JSON_UNESCAPED_SLASHES | JSON_PRETTY_PRINT);
