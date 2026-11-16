"""Publishes a signed Microads Office release to microads.vn/office.

usage: publish_release.py <MicroadsOffice_Setup_x.y.z.exe> <...exe.sig.json> [--notes-vi FILE] [--notes-en FILE] [--force] [--no-latest]
The server re-verifies the ECDSA signature before accepting the file.
"""
import argparse, base64, os, sys, time
sys.stdout.reconfigure(encoding='utf-8'); sys.stderr.reconfigure(encoding='utf-8')
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', '..', '.build', 'scripts'))
import vps

ap = argparse.ArgumentParser()
ap.add_argument('exe'); ap.add_argument('sig')
ap.add_argument('--notes-vi'); ap.add_argument('--notes-en')
ap.add_argument('--force', action='store_true'); ap.add_argument('--no-latest', action='store_true')
a = ap.parse_args()

def b64file(p):
    return base64.b64encode(open(p, encoding='utf-8').read().encode('utf-8')).decode() if p else ''

c = vps.client()
s = c.open_sftp()
stamp = time.strftime('%Y%m%d%H%M%S')
rexe = f'/var/www/microads-office-data/incoming/{stamp}.exe'
rsig = f'/var/www/microads-office-data/incoming/{stamp}.sig.json'
size = os.path.getsize(a.exe)
t0 = time.time()
last = [-10]
def _progress(done, total):
    pct = done * 100 // max(total, 1)
    if pct >= last[0] + 10:
        last[0] = pct
        print(f'  upload {pct}%', flush=True)
s.put(a.exe, rexe, callback=_progress)
print(f'  uploaded {size / 1048576:.1f} MB in {time.time() - t0:.0f}s')
s.put(a.sig, rsig)
s.close()
cmd = (f'chown www-data:www-data {rexe} {rsig} && runuser -u www-data -- php /var/www/microads.vn/office/inc/cli.php publish {rexe} {rsig}'
       f' --notes-vi-b64={b64file(a.notes_vi)} --notes-en-b64={b64file(a.notes_en)}'
       + (' --force' if a.force else '') + (' --no-latest' if a.no_latest else '') + f'; rc=$?; rm -f {rexe} {rsig}; exit $rc')
code, out, err = vps.run(cmd, c)
print(out.strip(), err.strip())
c.close()
sys.exit(code)
