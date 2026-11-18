"""Build the 寶雅新品資訊 page from many 報價單 files.
Usage: python3 poya_build.py <manifest.json> <template.html> <out.html>
manifest: [{"file": path, "date": "YYYY.MM.DD"}, ...]   (one dropdown entry each)
Exports retail price (牌價) but never cost, wholesale price or margin. Set env POYA_PASSWORD to encrypt the data (required for the public site)."""
import sys, json, re
sys.path.insert(0, __import__('os').path.dirname(__import__('os').path.abspath(__file__)))
from poya_extract import extract, images

man, tpl, out = sys.argv[1:4]
BRANDS = [('AA', r'(^|[^A-Z])AA|香薰'), ('LUNASOL', r'LUNASOL|國-LU|(^|\s)LU[^A-Z]'), ('CHIC CHOC', r'CHIC ?CHOC'),
          ('DEW', r'DEW'), ('米蘭絕色', r'MILANO|米蘭'), ('suisai', r'(?i)suisai'), ('KANEBO', r'KANEBO|國-KN'), ('TFM', r'TFM')]

def brands(items):
    cnt = {}
    for it in items:
        b = next((b for b, p in BRANDS if re.search(p, it['name'])), '佳麗寶')
        cnt[b] = cnt.get(b, 0) + 1
    return '、'.join(sorted(cnt, key=lambda k: -cnt[k]))

sheets = []
for m in json.load(open(man, encoding='utf-8')):
    d = extract(m['file'])
    by_e, by_n = images(m['file'])
    for it in d['items']:
        it['img'] = by_e.get(it['ean']) or ('' if by_e else by_n.get(it['no'], ''))
    sheets.append(dict(date=m['date'], brands=brands(d['items']), period=d['period'], items=d['items']))
sheets.sort(key=lambda s: (s['date'], s['brands']), reverse=True)
payload = json.dumps(dict(sheets=sheets), ensure_ascii=False)
pw = __import__('os').environ.get('POYA_PASSWORD')
if pw:  # encrypt so the published file holds no readable data
    import os, base64
    from cryptography.hazmat.primitives import hashes
    from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
    from cryptography.hazmat.primitives.ciphers.aead import AESGCM
    salt, iv, n = os.urandom(16), os.urandom(12), 250000
    key = PBKDF2HMAC(algorithm=hashes.SHA256(), length=32, salt=salt, iterations=n).derive(pw.encode())
    ct = AESGCM(key).encrypt(iv, payload.encode(), None)
    e = lambda b: base64.b64encode(b).decode()
    payload = json.dumps(dict(s=e(salt), i=e(iv), n=n, c=e(ct)))
else:
    print('WARNING: POYA_PASSWORD not set - page is NOT encrypted')
html = open(tpl, encoding='utf-8').read().replace('/*__DATA__*/null', payload)
open(out, 'w', encoding='utf-8').write(html)
print(len(sheets), 'sheets,', sum(len(s['items']) for s in sheets), 'items')
for s in sheets:
    bad = [i['ean'] for i in s['items'] if not i['ok']]
    print(s['date'], s['brands'], len(s['items']), 'img', sum(1 for i in s['items'] if i['img']), ('BAD ' + ','.join(bad)) if bad else '')
