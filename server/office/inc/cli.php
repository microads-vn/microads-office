<?php
// Command line for the release script (run as www-data on the server):
//   php cli.php publish <exe> <sig.json> [--notes-vi=...] [--notes-en=...] [--force] [--no-latest]
//   php cli.php force on|off
//   php cli.php status
//   php cli.php websetup <MicroadsOffice_Setup.exe> [--replace]   publish the web installer (refuses to replace a different one)
declare(strict_types=1);
if (PHP_SAPI !== 'cli') { http_response_code(404); exit; }
require __DIR__ . '/bootstrap.php';
$args = array_slice($argv, 1);
$opt = function (string $name) use ($args): ?string { foreach ($args as $a) if (str_starts_with($a, "--$name=")) return substr($a, strlen($name) + 3); return null; };
$flag = fn(string $name) => in_array("--$name", $args, true);
switch ($args[0] ?? '') {
    case 'publish':
        $sig = json_decode((string)file_get_contents($args[2]), true);
        if (!is_array($sig)) { fwrite(STDERR, "bad signature file\n"); exit(2); }
        $notes = ['vi' => base64_decode($opt('notes-vi-b64') ?? '') ?: ($opt('notes-vi') ?? ''), 'en' => base64_decode($opt('notes-en-b64') ?? '') ?: ($opt('notes-en') ?? '')];
        [$ok, $msg] = publish_release($args[1], $sig['version'], $sig['signature'], $notes, !$flag('no-latest'), $flag('force'), true);
        echo $msg, "\n"; exit($ok ? 0 : 1);
    case 'force':
        $c = load_config();
        $c['forceUpdate'] = ($args[1] ?? '') === 'on' ? ['enabled' => true, 'commandId' => bin2hex(random_bytes(8)), 'issuedAt' => date('c')] : ['enabled' => false, 'commandId' => $c['forceUpdate']['commandId'] ?? '', 'issuedAt' => null];
        save_config($c); audit('force', $args[1] ?? ''); echo json_encode($c['forceUpdate']), "\n"; exit(0);
    case 'websetup':
        [$ok, $msg] = publish_websetup($args[1] ?? '', $flag('replace'));
        echo $msg, "\n"; exit($ok ? 0 : 1);
    case 'status':
        echo json_encode(build_manifest(load_config()), JSON_PRETTY_PRINT | JSON_UNESCAPED_UNICODE | JSON_UNESCAPED_SLASHES), "\n"; exit(0);
    default:
        fwrite(STDERR, "usage: publish|force|status|websetup\n"); exit(2);
}
