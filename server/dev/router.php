<?php
// Local test server for microads.vn/office (plays the nginx rules of server/deploy/nginx-office.conf):
//   php -S 127.0.0.1:8099 -t server server/dev/router.php
// Environment: MO_DATA_DIR=<test data folder>  MO_BASE_URL=http://127.0.0.1:8099/office  MO_DEV_SENDFILE=1  MO_DEV_INSECURE_COOKIES=1
// Used by server/dev/test_site.py. Never deployed (deploy.py only uploads server/office).
// The page scripts are required at GLOBAL scope (like php-fpm does): they use globals such as $LANG.
declare(strict_types=1);
$__uri = (string)parse_url($_SERVER['REQUEST_URI'], PHP_URL_PATH);
$__office = dirname(__DIR__) . '/office';
$__script = null;
if ($__uri === '/office') { header('Location: /office/', true, 301); return true; }
if (str_starts_with($__uri, '/office/inc/')) { http_response_code(404); echo 'Not found'; return true; }
if ($__uri === '/office/api/manifest.json') $__script = 'api.php';
elseif ($__uri === '/office/privacy') $__script = 'privacy.php';
elseif ($__uri === '/office/code-signing') $__script = 'code-signing.php';
elseif ($__uri === '/office/install.ps1') $__script = 'install.php';
elseif (preg_match('#^/office/download/([^/]+)$#', $__uri, $__m)) { $__script = 'download.php'; $_GET['f'] = rawurldecode($__m[1]); }
elseif ($__uri === '/office/') $__script = 'index.php';
elseif ($__uri === '/office/en/') { $__script = 'index.php'; $_GET['hl'] = 'en'; }
elseif (preg_match('#^/office/(en/)?blog/$#', $__uri, $__m)) { $__script = 'blog.php'; $_GET["hl"] = !empty($__m[1]) ? "en" : "vi"; }
elseif (preg_match('#^/office/(en/)?blog/feed\.xml$#', $__uri, $__m)) { $__script = 'blog.php'; $_GET["hl"] = !empty($__m[1]) ? "en" : "vi"; $_GET['feed'] = '1'; }
elseif (preg_match('#^/office/(en/)?blog/([a-z0-9-]+)/?$#', $__uri, $__m)) { $__script = 'blog.php'; $_GET["hl"] = !empty($__m[1]) ? "en" : "vi"; $_GET['slug'] = $__m[2]; }
elseif ($__uri === '/office/sitemap.xml') $__script = 'sitemap.php';
elseif ($__uri === '/office/en/excel-formula-generator') { $__script = 'tool.php'; $_GET['hl'] = 'en'; }
elseif ($__uri === '/office/cong-cu/tao-cong-thuc-excel') { $__script = 'tool.php'; $_GET['hl'] = 'vi'; }
elseif ($__uri === '/office/api/formula') $__script = 'tool-api.php';
elseif (preg_match('#^/office/([0-9a-f]{32})\.txt$#', $__uri, $__m)) { $__script = 'indexnow.php'; $_GET['k'] = $__m[1]; }
elseif ($__uri === '/office/admin/') $__script = 'admin/index.php';
elseif (preg_match('#^/office/([a-z/]+\.php)$#', $__uri, $__m) && is_file($__office . '/' . $__m[1])) $__script = $__m[1];
if ($__script === null) return false; // static files (assets) from the document root
$_SERVER['SCRIPT_FILENAME'] = $__office . '/' . $__script;
$_SERVER['SCRIPT_NAME'] = '/office/' . $__script;
chdir(dirname($_SERVER['SCRIPT_FILENAME']));
require $_SERVER['SCRIPT_FILENAME'];
return true;
