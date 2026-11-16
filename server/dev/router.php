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
elseif ($__uri === '/office/install.ps1') $__script = 'install.php';
elseif (preg_match('#^/office/download/([^/]+)$#', $__uri, $__m)) { $__script = 'download.php'; $_GET['f'] = rawurldecode($__m[1]); }
elseif ($__uri === '/office/') $__script = 'index.php';
elseif ($__uri === '/office/admin/') $__script = 'admin/index.php';
elseif (preg_match('#^/office/([a-z/]+\.php)$#', $__uri, $__m) && is_file($__office . '/' . $__m[1])) $__script = $__m[1];
if ($__script === null) return false; // static files (assets) from the document root
$_SERVER['SCRIPT_FILENAME'] = $__office . '/' . $__script;
$_SERVER['SCRIPT_NAME'] = '/office/' . $__script;
chdir(dirname($_SERVER['SCRIPT_FILENAME']));
require $_SERVER['SCRIPT_FILENAME'];
return true;
