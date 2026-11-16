"""Local test of microads.vn/office (download page, web installer, PowerShell command, admin) with PHP's built-in server.

usage: python -I server/dev/test_site.py <php.exe> <work folder> <web installer exe> <full setup exe> <full setup .sig.json> [--keep-running]
Everything is written inside <work folder>. With --keep-running the server stays up on 127.0.0.1:8099 for the
installer / PowerShell integration tests (stop it by killing the printed PID).
"""
import hashlib, http.cookiejar, json, os, re, shutil, subprocess, sys, time, urllib.error, urllib.parse, urllib.request
sys.stdout.reconfigure(encoding='utf-8')

php, work, stub, full, sigf = sys.argv[1:6]
keep = '--keep-running' in sys.argv
ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
PORT = 8099
BASE = f'http://127.0.0.1:{PORT}/office'
data = os.path.join(work, 'data')
if os.path.isdir(work): shutil.rmtree(work)
for d in ('releases', 'sessions', 'logs', 'stats', 'tmp'): os.makedirs(os.path.join(data, d))
sig = json.load(open(sigf, encoding='utf-8'))
ver = sig['version']
full_name = f'MicroadsOffice_Setup_{ver}.exe'
try: os.link(full, os.path.join(data, 'releases', full_name))
except OSError: shutil.copy(full, os.path.join(data, 'releases', full_name))
json.dump({'latest': ver, 'releases': [{'version': ver, 'file': full_name, 'size': sig['size'], 'sha256': sig['sha256'], 'signature': sig['signature'],
           'date': '2026-11-16', 'publishedAt': '2026-11-16T18:30:00+07:00', 'notes': {'vi': 'Ghi chú thử nghiệm', 'en': 'Test notes'}, 'downloads': 0}]},
          open(os.path.join(data, 'config.json'), 'w', encoding='utf-8'))
ADMIN_PW = 'local-test-password-123'
env = dict(os.environ, MO_DATA_DIR=data, MO_BASE_URL=BASE, MO_DEV_SENDFILE='1', MO_DEV_INSECURE_COOKIES='1', TEMP=os.path.join(data, 'tmp'), TMP=os.path.join(data, 'tmp'))
phpini = ['-d', 'display_errors=1', '-d', 'error_reporting=E_ALL', '-d', f'session.save_path={os.path.join(data, "sessions")}',
          '-d', f'sys_temp_dir={os.path.join(data, "tmp")}', '-d', f'upload_tmp_dir={os.path.join(data, "tmp")}', '-d', 'opcache.enable_cli=0']
h = subprocess.run([php, *phpini, '-r', 'echo json_encode(["user"=>"admin","hash"=>password_hash($argv[1], PASSWORD_DEFAULT)]);', ADMIN_PW], capture_output=True, text=True, env=env)
open(os.path.join(data, 'admin.json'), 'w').write(h.stdout)

passed = failed = 0
def check(ok, what):
    global passed, failed
    passed += bool(ok); failed += (not ok)
    print(('PASS  ' if ok else 'FAIL  ') + what, flush=True)

def cli(*a):
    r = subprocess.run([php, *phpini, os.path.join(ROOT, 'server', 'office', 'inc', 'cli.php'), *a], capture_output=True, text=True, encoding='utf-8', env=env)
    return r.returncode, (r.stdout + r.stderr).strip()

sha = lambda b: hashlib.sha256(b).hexdigest()
stub_bytes = open(stub, 'rb').read()

# ---------------------------------------------------------------- CLI: publish the web installer
code, out = cli('websetup', os.path.join(data, 'missing.exe'))
check(code == 1, 'cli websetup: missing file refused')
code, out = cli('websetup', full)
check(code == 1 and '< 5 MB' in out, 'cli websetup: the 51 MB full setup is refused as web installer (' + out[:60] + ')')
code, out = cli('websetup', stub)
check(code == 0 and sha(stub_bytes) in out, 'cli websetup: web installer published')
other = os.path.join(data, 'tmp', 'other.exe'); open(other, 'wb').write(stub_bytes + b'\0changed')
code, out = cli('websetup', other)
check(code == 1 and 'Danh tiếng' in out,'cli websetup: a DIFFERENT web installer is refused (protects its reputation)')
code, out = cli('websetup', stub)
check(code == 0 and 'giữ nguyên' in out, 'cli websetup: same file again → kept unchanged')

# ---------------------------------------------------------------- server
srv = subprocess.Popen([php, *phpini, '-S', f'127.0.0.1:{PORT}', '-t', os.path.join(ROOT, 'server'), os.path.join(ROOT, 'server', 'dev', 'router.php')],
                       env=env, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
time.sleep(1.5)
def get(path, opener=None, data=None):
    req = urllib.request.Request(BASE + path, data=urllib.parse.urlencode(data).encode() if data else None)
    try:
        with (opener.open(req) if opener else urllib.request.urlopen(req)) as r: return r.status, dict(r.headers), r.read()
    except urllib.error.HTTPError as e: return e.code, dict(e.headers), e.read()
class NoRedirect(urllib.request.HTTPRedirectHandler):
    def redirect_request(self, *a, **k): return None
noredir = urllib.request.build_opener(NoRedirect)
def php_errors(html): return re.findall(r'(Warning|Notice|Deprecated|Fatal error|Parse error)</b>:|(Warning|Notice|Deprecated|Fatal error): ', html)

try:
    # ---- download page (Vietnamese)
    st, hd, body = get('/?lang=vi'); page = body.decode('utf-8')
    check(st == 200 and not php_errors(page), 'download page renders without PHP warnings')
    check('id="dl-main" href="download/MicroadsOffice_Setup.exe"' in page, 'main button → web installer MicroadsOffice_Setup.exe')
    check(f'id="dl-full" href="download/{full_name}"' in page and f'Bộ cài đầy đủ</a>' not in page and f'Bộ cài đầy đủ {ver}' in page.replace('</a>', ''), f'second button → full installer {ver}')
    check('luôn cài bản mới nhất' in page and f'<b>{ver}</b>' in page, 'meta line: web installer always installs the latest version')
    check(sha(stub_bytes) in page and sig['sha256'] in page, 'SHA-256 of both files shown')
    check('Vẫn chạy' in page and 'Tải tệp đáng ngờ xuống' in page and 'ít được tải xuống' in page and 'không phải phát hiện virus' in page,
          'safe-installation steps use the exact Chrome 157 wording (⋮ → Tải tệp đáng ngờ xuống) and Windows (Vẫn chạy)')
    check('google-site-verification' not in page, 'no Google verification tag until one is set')
    check('href="code-signing" id="lnk-code-signing">Code signing policy</a>' in page and 'github.com/microads-vn/microads-office' in page, 'home/download page links the Code signing policy and the source code')
    check('SignPath' not in page, 'home/download page makes no SignPath claim (application declined)')
    st, _, b = get('/code-signing'); cs = b.decode('utf-8')
    check(st == 200 and 'ECDSA P-256' in cs and 'SignPath' not in cs and not php_errors(cs),
          'code signing policy page describes the Microads release signature, no SignPath claim')
    check('Approvers' in cs or 'phê duyệt' in cs, 'code signing policy lists team roles')
    st, _, b = get('/privacy'); pv = b.decode('utf-8')
    check(st == 200 and 'Tắt kết nối tự động' in pv and '--no-auto-update' in pv, 'privacy policy explains how to switch automatic connections off')
    check(f'irm {BASE}/install.ps1 | iex' in page and 'data-copy="cmd-ps"' in page, 'PowerShell command with Copy button')
    check('MicroadsOffice_Setup.exe --silent' in page, 'silent install command for IT')
    check('cmd-winget' not in page, 'winget command hidden while no package ID is set')
    check('navigator.clipboard' in page and 'execCommand' in page, 'Copy buttons have a clipboard fallback')
    st, hd, body = get('/?lang=en'); en = body.decode('utf-8')
    check('Download Microads Office' in en and 'Full installer' in en and 'Install with a command' in en and 'Why the web installer?' in en and not php_errors(en), 'English page translated')
    st, hd, css = get('/assets/style.css')
    check(st == 200 and b'.cmd{' in css, 'stylesheet has the command-box styles')

    # ---- web installer download
    cfg_before = open(os.path.join(data, 'config.json'), 'rb').read()
    st, hd, body = get('/download/MicroadsOffice_Setup.exe')
    check(st == 200 and body == stub_bytes, 'web installer downloads byte-identical')
    check(hd.get('Content-Disposition') == 'attachment; filename="MicroadsOffice_Setup.exe"' and hd.get('Content-Type') == 'application/vnd.microsoft.portable-executable', 'download headers (attachment, PE type)')
    cfg = json.load(open(os.path.join(data, 'config.json'), encoding='utf-8'))
    cnt = json.load(open(os.path.join(data, 'counters.json'), encoding='utf-8'))
    check(cnt.get('MicroadsOffice_Setup.exe') == 1 and cfg['webSetup']['sha256'] == sha(stub_bytes), 'web installer download counted (counters.json)')
    check(open(os.path.join(data, 'config.json'), 'rb').read() == cfg_before, 'a download never rewrites config.json (no lost releases when downloads and publishing overlap)')
    st, hd, _ = get('/download/latest', noredir)
    check(st == 302 and hd.get('Location', '').endswith('/download/' + full_name), '/download/latest → full installer')
    st, hd, body = get('/download/' + full_name)
    check(st == 200 and sha(body) == sig['sha256'], 'full installer downloads byte-identical')
    check(json.load(open(os.path.join(data, 'counters.json'), encoding='utf-8')).get(full_name) == 1, 'full installer download counted')
    st, _, _ = get('/download/MicroadsOffice_Setup_2.0.1.exe%0A')
    check(st == 404, 'file name with a trailing newline refused')
    st, _, _ = get('/download/..%2Fconfig.json')
    check(st == 404, 'path tricks refused (404)')
    st, _, _ = get('/inc/install.ps1')
    check(st == 404, 'inc/ folder not reachable')

    # ---- manifest
    st, hd, body = get('/api/manifest.json?v=web-setup'); m = json.loads(body)
    check(m['latest']['version'] == ver and m['latest']['url'] == f'{BASE}/download/{full_name}', 'manifest gives the latest release to the web installer')
    check(m['promo']['showOnStart'] is True, 'manifest tells the apps to show the introduction at start')
    stats = json.load(open(os.path.join(data, 'stats', time.strftime('%Y-%m-%d') + '.json')))
    check(stats.get('web-setup') == 1, 'web installer runs counted separately ("web-setup")')

    # ---- PowerShell script
    st, hd, body = get('/install.ps1')
    check(st == 200 and hd.get('Content-Type') == 'text/plain; charset=utf-8', 'install.ps1 served as UTF-8 text')
    check(not body.startswith(b'\xef\xbb\xbf') and f"'{BASE}/api/manifest.json'".encode() in body and b'__MANIFEST_URL__' not in body, 'install.ps1: no BOM, manifest address filled in')

    # ---- admin
    cj = http.cookiejar.CookieJar(); op = urllib.request.build_opener(urllib.request.HTTPCookieProcessor(cj))
    csrf = lambda html: re.search(r'name="csrf" value="([0-9a-f]+)"', html).group(1)
    st, _, b = get('/admin/', op); hp = b.decode()
    st, _, b = get('/admin/', op, {'csrf': csrf(hp), 'action': 'login', 'user': 'admin', 'pass': ADMIN_PW}); hp = b.decode()
    check('Đăng xuất' in hp and not php_errors(hp), 'admin login')
    check('Trình cài đặt trực tuyến (nút tải chính)' in hp and sha(stub_bytes) in hp and '<b>1</b> lượt tải' in hp, 'admin overview: web installer, SHA-256 and download count')
    check('web-setup: 1' in hp, 'admin statistics list web installer runs')
    st, _, b = get('/admin/?tab=promo', op); hp = b.decode()
    check('id="winget"' in hp and 'name="wingetId"' in hp and not php_errors(hp), 'admin: winget package field')
    check('name="showOnStart"' in hp, 'admin: switch for the introduction at every start')
    st, _, b = get('/admin/', op, {'csrf': csrf(hp), 'action': 'winget', 'wingetId': 'bad id; rm -rf'}); hp = b.decode()
    check('Mã gói winget dạng' in hp, 'admin: invalid winget ID refused')
    st, _, b = get('/admin/', op, {'csrf': csrf(hp), 'action': 'winget', 'wingetId': 'Microads.Office'}); hp = b.decode()
    check('winget install --id Microads.Office -e' in hp, 'admin: winget ID saved')
    st, _, b = get('/?lang=vi'); page = b.decode()
    check('<code id="cmd-winget">winget install --id Microads.Office -e</code>' in page, 'download page shows the winget command once set')
    st, _, b = get('/admin/', op, {'csrf': csrf(hp), 'action': 'winget', 'wingetId': ''}); hp = b.decode()
    st, _, b = get('/?lang=vi')
    check('cmd-winget' not in b.decode(), 'clearing the winget ID hides the command again')
    st, _, b = get('/admin/?tab=promo', op); hp = b.decode()
    check('id="gsc"' in hp and 'name="gscVerification"' in hp, 'admin: Google Search Console verification field')
    st, _, b = get('/admin/', op, {'csrf': csrf(hp), 'action': 'gsc', 'gscVerification': 'bad code!'}); hp = b.decode()
    check('Mã xác minh không hợp lệ' in hp, 'admin: invalid verification code refused')
    tag = '<meta name="google-site-verification" content="AbC-123_xyzVerificationToken0987" />'
    st, _, b = get('/admin/', op, {'csrf': csrf(hp), 'action': 'gsc', 'gscVerification': tag}); hp = b.decode()
    st, _, b = get('/?lang=vi'); page = b.decode()
    check('<meta name="google-site-verification" content="AbC-123_xyzVerificationToken0987">' in page, 'pasting the whole <meta> tag puts the verification code on microads.vn/office/')
    st, _, b = get('/admin/', op, {'csrf': csrf(hp), 'action': 'gsc', 'gscVerification': ''}); hp = b.decode()
    st, _, b = get('/?lang=vi')
    check('google-site-verification' not in b.decode(), 'clearing the code removes the tag')
    st, _, b = get('/admin/', op, {'csrf': 'bad', 'action': 'winget', 'wingetId': 'X.Y'}); hp = b.decode()
    check('Phiên làm việc hết hạn' in hp, 'admin: CSRF still enforced')
    for tab in ('publish', 'releases', 'account'):
        st, _, b = get('/admin/?tab=' + tab, op)
        check(st == 200 and not php_errors(b.decode()), f'admin tab {tab} renders')
finally:
    if keep:
        print(f'SERVER_PID={srv.pid}')
    else:
        srv.kill()
print(f'\nTOTAL: {passed} passed, {failed} failed')
sys.exit(1 if failed else 0)
