<?php
// GET /office/install.ps1 → PowerShell install command:  irm https://microads.vn/office/install.ps1 | iex
// The script downloads the latest release, verifies its SHA-256 and the Microads ECDSA signature, then runs the setup.
declare(strict_types=1);
require __DIR__ . '/inc/bootstrap.php';
header('Content-Type: text/plain; charset=utf-8');
header('Cache-Control: no-cache, max-age=0');
header('X-Content-Type-Options: nosniff');
$s = (string)file_get_contents(__DIR__ . '/inc/install.ps1');
if (str_starts_with($s, "\xEF\xBB\xBF")) $s = substr($s, 3); // no BOM inside the string handed to Invoke-Expression
echo str_replace('__MANIFEST_URL__', MO_BASE_URL . '/api/manifest.json', $s);
