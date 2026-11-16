"""Deploys microads.vn/office to the VPS (idempotent, with backups and automatic rollback)."""
import os, sys, io, time, posixpath, secrets, string, json
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', '..', '.build', 'scripts'))
import vps

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(HERE, '..', 'office')
WEB = '/var/www/microads.vn/office'
DATA = '/var/www/microads-office-data'
SITE = '/etc/nginx/sites-available/microads.vn'
POOL = '/etc/php/8.5/fpm/pool.d/microads-office.conf'
STAMP = time.strftime('%Y%m%d-%H%M%S')

def sh(c, cmd, check=True):
    code, out, err = vps.run(cmd, c)
    if check and code != 0:
        raise SystemExit(f'FAILED ({code}): {cmd}\n{out}\n{err}')
    return out

def main():
    c = vps.client()
    sftp = c.open_sftp()
    # 1. code
    sh(c, f'mkdir -p {WEB}.new && rm -rf {WEB}.new/*')
    for root, dirs, files in os.walk(SRC):
        rel = os.path.relpath(root, SRC).replace(os.sep, '/')
        rdir = WEB + '.new' + ('' if rel == '.' else '/' + rel)
        sh(c, f'mkdir -p "{rdir}"')
        for f in files:
            sftp.put(os.path.join(root, f), rdir + '/' + f)
    lint = sh(c, f'for f in $(find {WEB}.new -name "*.php"); do php -l "$f" || echo LINT_FAIL; done 2>&1', check=False)
    print(lint)
    if 'LINT_FAIL' in lint: raise SystemExit('PHP syntax error, nothing deployed')
    sh(c, f'if [ -d {WEB} ]; then mv {WEB} /var/www/office.bak-{STAMP}; fi; mv {WEB}.new {WEB}; chown -R root:root {WEB}; chmod -R a+rX,go-w {WEB}')
    # 2. data folder (kept across deploys)
    sh(c, f'mkdir -p {DATA}/releases {DATA}/sessions {DATA}/logs {DATA}/stats {DATA}/tmp {DATA}/incoming && chown -R www-data:www-data {DATA} && chmod 750 {DATA}')
    admin_pw = None
    code, _, _ = vps.run(f'test -f {DATA}/admin.json', c)
    if code != 0:
        alphabet = string.ascii_letters + string.digits
        admin_pw = 'Mo-' + ''.join(secrets.choice(alphabet) for _ in range(14))
        php = f"php -r 'echo json_encode([\"user\"=>\"admin\",\"hash\"=>password_hash($argv[1], PASSWORD_DEFAULT)]);' '{admin_pw}' > {DATA}/admin.json"
        sh(c, php + f' && chown www-data:www-data {DATA}/admin.json && chmod 640 {DATA}/admin.json')
    # 3. php-fpm pool
    sftp.put(os.path.join(HERE, 'php-fpm-office.conf'), POOL)
    sh(c, 'php-fpm8.5 -t 2>&1')
    # 4. nginx block (replace between markers, else insert before microads-help block)
    sh(c, f'cp {SITE} /root/microads.vn.nginx.bak-{STAMP}')
    conf = sh(c, f'cat {SITE}')
    block = open(os.path.join(HERE, 'nginx-office.conf'), encoding='utf-8').read()
    start, end = '    # >>> microads-office >>>', '    # <<< microads-office <<<\n'
    if start in conf:
        a = conf.index(start); b = conf.index(end, a) + len(end)
        conf = conf[:a] + block.rstrip('\n') + '\n' + conf[b:].lstrip('\n')
    else:
        anchor = '    # >>> microads-help >>>'
        if anchor not in conf: raise SystemExit('anchor not found in nginx config')
        conf = conf.replace(anchor, block + anchor, 1)
    with sftp.open(SITE + '.tmp', 'w') as fh: fh.write(conf.encode('utf-8'))
    sh(c, f'mv {SITE}.tmp {SITE}')
    code, out, err = vps.run('nginx -t 2>&1', c)
    if code != 0:
        sh(c, f'cp /root/microads.vn.nginx.bak-{STAMP} {SITE}')
        raise SystemExit('nginx -t failed, config restored:\n' + out + err)
    sh(c, 'systemctl reload php8.5-fpm && systemctl reload nginx')
    print('deployed', STAMP)
    if admin_pw:
        print('ADMIN_USER=admin')
        print('ADMIN_PASSWORD=' + admin_pw)
    sftp.close(); c.close()

if __name__ == '__main__':
    main()
