"""Publishes the web installer (release/web-setup/MicroadsOffice_Setup.exe) to microads.vn/office/download/MicroadsOffice_Setup.exe.

usage: publish_websetup.py [MicroadsOffice_Setup.exe] [--replace]
The web installer never changes: the server keeps an identical file as it is and REFUSES a different one unless --replace,
because Chrome / SmartScreen reputation is tied to the file's SHA-256 (a new file starts again from zero).
"""
import argparse, hashlib, os, sys, time
sys.stdout.reconfigure(encoding='utf-8'); sys.stderr.reconfigure(encoding='utf-8')
HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, '..', '..', '.build', 'scripts'))
import vps

ap = argparse.ArgumentParser()
ap.add_argument('exe', nargs='?', default=os.path.join(HERE, '..', '..', 'release', 'web-setup', 'MicroadsOffice_Setup.exe'))
ap.add_argument('--replace', action='store_true')
a = ap.parse_args()
data = open(a.exe, 'rb').read()
sha = hashlib.sha256(data).hexdigest()
pinned = a.exe + '.sha256'
if os.path.exists(pinned) and open(pinned).read().strip() != sha:
    sys.exit(f'{a.exe} does not match {pinned} — refusing to publish a modified web installer')
c = vps.client()
s = c.open_sftp()
remote = f'/var/www/microads-office-data/incoming/websetup-{time.strftime("%Y%m%d%H%M%S")}.exe'
s.put(a.exe, remote)
s.close()
cmd = (f'chown www-data:www-data {remote} && runuser -u www-data -- php /var/www/microads.vn/office/inc/cli.php websetup {remote}'
       + (' --replace' if a.replace else '') + f'; rc=$?; rm -f {remote}; exit $rc')
code, out, err = vps.run(cmd, c)
print(out.strip(), err.strip())
code2, out2, _ = vps.run(f'sha256sum /var/www/microads-office-data/releases/MicroadsOffice_Setup.exe', c)
print('server copy:', out2.strip())
c.close()
sys.exit(code if code else (0 if sha in out2 else 1))
