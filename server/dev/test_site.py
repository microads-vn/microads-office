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
os.makedirs(os.path.join(data, 'blog', 'posts', 'vi')); os.makedirs(os.path.join(data, 'blog', 'posts', 'en'))
def post_json(lang, slug, title, body, status='published'):
    return {'lang': lang, 'slug': slug, 'title': title, 'metaDescription': 'Mô tả thử nghiệm đủ dài cho bài viết, có từ khóa chính và giá trị cụ thể của bài để người đọc bấm vào xem chi tiết.',
            'excerpt': 'Tóm tắt bài', 'primaryKeyword': 'ham xlookup', 'keywords': ['xlookup'], 'pillar': 'excel-formulas', 'format': 'how-to', 'shot': 'excel-explain',
            'bodyHtml': body, 'faq': [{'q': 'Câu hỏi 1?', 'a': 'Trả lời 1.'}], 'wordCount': 900, 'publishedAt': '2026-11-20T09:00:00+07:00', 'updatedAt': '2026-11-20T09:00:00+07:00',
            'status': status, 'review': {'score': 88, 'issues': []}, 'costUsd': 0.4}
evil = ('<h1>Tiêu đề lạ</h1><p>Đoạn mở <script>alert(1)</script><a href="javascript:alert(2)">xấu</a> <a href="https://support.microsoft.com/x">nguồn</a></p>'
        '<h2>Phần một</h2><p>Nội dung <img src=x onerror=alert(3)></p><h2>Phần hai</h2><p onclick="x()">B</p><h2>Phần ba</h2><p>C <a href="/office/blog/bai-khac-mau">liên kết</a></p>'
        '<table><tr><td style="color:red">1</td></tr></table>')
json.dump(post_json('vi', 'ham-xlookup-co-ban', 'Hàm XLOOKUP cơ bản', evil), open(os.path.join(data, 'blog', 'posts', 'vi', 'ham-xlookup-co-ban.json'), 'w', encoding='utf-8'), ensure_ascii=False)
json.dump(post_json('vi', 'bai-an-mau', 'Bài đã gỡ', '<p>x</p>', 'unpublished'), open(os.path.join(data, 'blog', 'posts', 'vi', 'bai-an-mau.json'), 'w', encoding='utf-8'), ensure_ascii=False)
json.dump(post_json('en', 'xlookup-basics-guide', 'XLOOKUP basics', '<h2>One</h2><p>Text</p>'), open(os.path.join(data, 'blog', 'posts', 'en', 'xlookup-basics-guide.json'), 'w', encoding='utf-8'), ensure_ascii=False)
idx = []
for lang, slug in (('vi', 'ham-xlookup-co-ban'), ('vi', 'bai-an-mau'), ('en', 'xlookup-basics-guide')):
    d = json.load(open(os.path.join(data, 'blog', 'posts', lang, slug + '.json'), encoding='utf-8')); d.pop('bodyHtml'); d.pop('faq'); idx.append(d)
json.dump(idx, open(os.path.join(data, 'blog', 'index.json'), 'w', encoding='utf-8'), ensure_ascii=False)
os.makedirs(os.path.join(data, 'seo'))
json.dump({'indexnowKey': '0123456789abcdef0123456789abcdef'}, open(os.path.join(data, 'seo', 'settings.json'), 'w'))
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
    check('id="dl-main" href="/office/download/MicroadsOffice_Setup.exe"' in page, 'main button → web installer MicroadsOffice_Setup.exe')
    check('id="dl-full"' not in page and 'Bộ cài đầy đủ' not in page, 'only the web installer is offered from microads.vn (no full-installer button)')
    check('luôn cài bản mới nhất' in page and f'<b>{ver}</b>' in page, 'meta line: web installer always installs the latest version')
    check(sha(stub_bytes) in page and sig['sha256'] not in page, 'SHA-256 of the web installer shown')
    check('Vẫn chạy' in page and 'Tải tệp đáng ngờ xuống' in page and 'ít được tải xuống' in page and 'không phải phát hiện virus' in page,
          'safe-installation steps use the exact Chrome 157 wording (⋮ → Tải tệp đáng ngờ xuống) and Windows (Vẫn chạy)')
    check('google-site-verification' not in page, 'no Google verification tag until one is set')
    check('href="/office/code-signing" id="lnk-code-signing">Code signing policy</a>' in page and 'github.com/microads-vn/microads-office' in page, 'home/download page links the Code signing policy and the source code')
    check('SignPath' not in page, 'home/download page makes no SignPath claim (application declined)')
    st, _, b = get('/code-signing'); cs = b.decode('utf-8')
    check(st == 200 and 'ECDSA P-256' in cs and 'SignPath' not in cs and not php_errors(cs),
          'code signing policy page describes the Microads release signature, no SignPath claim')
    check('Approvers' in cs or 'phê duyệt' in cs, 'code signing policy lists team roles')
    st, _, b = get('/privacy'); pv = b.decode('utf-8')
    check(st == 200 and 'Tắt kết nối tự động' in pv and '--no-auto-update' in pv, 'privacy policy explains how to switch automatic connections off')
    check('install.ps1' not in page and 'cai-bang-lenh' not in page and 'Cập nhật phần mềm' not in page, 'command-install and software-update boxes removed')
    check('cmd-winget' not in page, 'no winget command on the page')
    g = page.find('<details class="guide" id="cai-dat-an-toan">')
    check(page.find('id="dl-main"') < g and '<summary>Hướng dẫn cài đặt</summary>' in page and page.find('Vẫn chạy') > g and page.find('id="sha-web"') > g,
          'installation guide is a drop-down under the two download buttons (steps and SHA-256 inside)')
    check('id="dl-main" href="/office/download/MicroadsOffice_Setup.exe" data-warn>' in page
          and '<dialog class="warn" id="dl-warn"' in page and 'Chrome chặn tải về' in page and 'màn hình xanh' in page and 'không chứa virus' in page and 'id="dl-warn-go"' in page and 'Tiếp tục tải' in page,
          'both download buttons first show the short warning with a Continue download button')
    check('<button type="button" class="btn big store" id="dl-store" data-wait>' in page and 'id="store-wait"' in page and 'đang chờ Microsoft duyệt' in page
          and 'data-eta="2026-11-21T00:00:00+07:00"' in page and 'giờ Việt Nam (GMT+7)' in page and '21/11/2026' in page,
          'Microsoft Store button explains the review and counts down to 21/11/2026 (Vietnam time)')
    check('id="cac-phien-ban"' in page and '<td class="n"><p>Ghi chú thử nghiệm</p></td>' in page, 'versions table lists the release with its notes')
    st, hd, body = get('/?lang=en'); en = body.decode('utf-8')
    check('Download Microads Office' in en and 'Continue download' in en and 'Full installer' not in en and '<summary>Installation guide</summary>' in en and 'Which option?' in en and not php_errors(en), 'English page translated')
    st, hd, css = get('/assets/style.css')
    check(st == 200 and b'.guide-body{' in css and b'.versions thead{display:none}' in css, 'stylesheet has the guide drop-down and phone layout of the versions table')

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
    check('cmd-winget' not in page, 'download page keeps no command box even with a winget ID')
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
    st, _, b = get('/admin/?tab=promo', op); hp = b.decode()
    st, _, b = get('/admin/', op, {'csrf': csrf(hp), 'action': 'store', 'storeId': 'https://apps.microsoft.com/detail/9PHM0WX16WDR?hl=vi'}); hp = b.decode()
    st, _, b = get('/?lang=vi'); page = b.decode()
    sb = page.find('id="dl-store" href="https://apps.microsoft.com/detail/9PHM0WX16WDR"')
    check(sb > 0 and 'ms-store-badge-vi.svg' in page and sb < page.find('id="dl-main"') and 'Khuyên dùng' in page and 'Đã vượt qua quy trình kiểm duyệt của Microsoft' in page and 'id="store-wait"' not in page,
          'with a Store ID: the official Microsoft Store badge comes first, the .exe buttons below, no countdown')
    st, _, b = get('/en/'); en = b.decode()
    check('ms-store-badge-en.svg' in en and 'Recommended' in en, 'English home shows the English Store badge')
    st, _, b = get('/blog/ham-xlookup-co-ban'); art = b.decode()
    check(art.count('ms-store-badge-vi.svg') == 2 and 'hoặc tải trình cài đặt' in art and 'Đã qua kiểm duyệt của Microsoft' in art, 'article download boxes lead with the Microsoft Store badge')
    st, _, b = get('/admin/?tab=promo', op); hp = b.decode()
    get('/admin/', op, {'csrf': csrf(hp), 'action': 'store', 'storeId': ''})
    for tab in ('publish', 'releases', 'account', 'seo'):
        st, _, b = get('/admin/?tab=' + tab, op)
        check(st == 200 and not php_errors(b.decode()), f'admin tab {tab} renders')

    # ---- language URLs, blog, sitemap, feed, tool
    st, hd, _ = get('/?lang=en', noredir)
    check(st == 301 and hd.get('Location') == '/office/en/', '?lang=en → 301 /office/en/ (one URL per language)')
    st, _, b = get('/en/'); en = b.decode()
    check(st == 200 and '<html lang="en">' in en and '<link rel="canonical" href="https://microads.vn/office/en/">' in en and 'hreflang="vi"' in en and not php_errors(en), '/office/en/ is the English home with canonical + hreflang')
    check('href="/office/assets/style.css' in en and 'href="/office/download/MicroadsOffice_Setup.exe" data-warn' in en, 'English home uses absolute asset/download paths')
    check('"@type":"SoftwareApplication"' in en and '"price":"0"' in en, 'home has SoftwareApplication structured data (free)')
    st, _, b = get('/?lang=vi'); page = b.decode()
    check('id="bai-viet"' in page and '/office/blog/ham-xlookup-co-ban' in page and 'bai-an-mau' not in page, 'home lists the latest published articles only')
    st, _, b = get('/blog/'); bl = b.decode()
    check(st == 200 and 'Hàm XLOOKUP cơ bản' in bl and 'Bài đã gỡ' not in bl and not php_errors(bl), 'Vietnamese blog list')
    st, _, b = get('/blog/ham-xlookup-co-ban'); art = b.decode()
    check(st == 200 and not php_errors(art), 'article renders')
    check('<script>alert' not in art and 'javascript:' not in art and 'onerror' not in art and 'onclick' not in art and 'style="color' not in art and '<img src=x' not in art, 'article HTML is sanitised (no script, js links, handlers, styles, foreign images)')
    check(art.count('<h1') == 1 and 'target="_blank" rel="noopener"' in art and 'id="phan-mot"' in art, 'one <h1>; external links open safely; headings get anchors')
    check('"@type":"BlogPosting"' in art and '"@type":"BreadcrumbList"' in art and '<link rel="canonical" href="https://microads.vn/office/blog/ham-xlookup-co-ban">' in art, 'article: BlogPosting + breadcrumbs + canonical')
    check(art.count('class="cta cta-') == 2 and 'Tải Microads Office — miễn phí vĩnh viễn' in art and 'Không cần tài khoản' in art and 'id="dl-warn"' in art, 'article: two "free forever, no account" download boxes + download notice')
    check('class="toc"' in art and 'shot-excel-explain.jpg' in art and 'Câu hỏi thường gặp' in art, 'article: table of contents, product screenshot, FAQ')
    st, _, b = get('/blog/bai-an-mau'); x = b.decode()
    check(st == 404 and 'noindex' in x, 'unpublished article → 404 + noindex')
    st, _, _ = get('/blog/khong-co-bai-nay')
    check(st == 404, 'unknown article → 404')
    st, _, b = get('/en/blog/xlookup-basics-guide'); ea = b.decode()
    check(st == 200 and '<html lang="en">' in ea and 'Get Microads Office — free forever' in ea and 'No account' in ea, 'English article in English interface')
    st, hd, b = get('/sitemap.xml'); sm = b.decode()
    check(st == 200 and 'https://microads.vn/office/blog/ham-xlookup-co-ban' in sm and 'https://microads.vn/office/en/blog/xlookup-basics-guide' in sm and 'bai-an-mau' not in sm and 'hreflang="en"' in sm, 'sitemap: published articles + hreflang pairs')
    st, hd, b = get('/blog/feed.xml'); fd = b.decode()
    check(st == 200 and '<rss' in fd and 'ham-xlookup-co-ban' in fd and 'xlookup-basics-guide' not in fd, 'Vietnamese RSS feed')
    st, _, b = get('/0123456789abcdef0123456789abcdef.txt')
    check(st == 200 and b.decode() == '0123456789abcdef0123456789abcdef', 'IndexNow key file served')
    st, _, _ = get('/ffffffffffffffffffffffffffffffff.txt')
    check(st == 404, 'wrong IndexNow key → 404')
    st, _, b = get('/en/excel-formula-generator'); tl = b.decode()
    check(st == 200 and 'AI Excel Formula Generator' in tl and 'name="t" value="' in tl and not php_errors(tl), 'formula generator page (EN)')
    st, _, b = get('/cong-cu/tao-cong-thuc-excel'); tv = b.decode()
    check(st == 200 and 'Tạo công thức Excel bằng AI' in tv, 'formula generator page (VI)')
    req = urllib.request.Request(BASE + '/api/formula', data=json.dumps({'q': 'sum A1:A10', 't': 'bad'}).encode(), headers={'Content-Type': 'application/json'})
    try: urllib.request.urlopen(req); code = 200
    except urllib.error.HTTPError as e: code = e.code
    check(code == 403, 'formula API refuses calls without the page token')
    tok = re.search(r'name="t" value="([0-9a-f]+)"', tl).group(1)
    req = urllib.request.Request(BASE + '/api/formula', data=json.dumps({'q': 'sum A1:A10', 't': tok, 'website': 'spam'}).encode(), headers={'Content-Type': 'application/json'})
    try: urllib.request.urlopen(req); code = 200
    except urllib.error.HTTPError as e: code = e.code
    check(code == 403, 'formula API refuses bots that fill the hidden field')
finally:
    if keep:
        print(f'SERVER_PID={srv.pid}')
    else:
        srv.kill()
print(f'\nTOTAL: {passed} passed, {failed} failed')
sys.exit(1 if failed else 0)
