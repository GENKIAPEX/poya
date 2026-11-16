"""Decide which downloaded 報價單 files become dropdown entries.
Usage: python3 poya_manifest.py <files.json> <manifest_out.json>
files.json: [{"path": local .xlsx path (xls already converted), "title": Drive title, "parentId": Drive folder id}, ...]
Rules: date = yyyymmdd in the title; skip files with no items; skip exact duplicates (same barcode set);
skip a file whose barcode set is a strict subset of another file in the same Drive folder (an earlier draft)."""
import sys, json, re
sys.path.insert(0, __import__('os').path.dirname(__import__('os').path.abspath(__file__)))
from poya_extract import extract

files = json.load(open(sys.argv[1], encoding='utf-8'))
recs = []
for f in files:
    m = re.search(r'(20\d{2})(\d{2})(\d{2})', f['title'])
    if not m:
        continue
    d = extract(f['path'])
    if not d or not d['items']:
        continue
    recs.append(dict(f, date=f'{m.group(1)}.{m.group(2)}.{m.group(3)}', eans=frozenset(i['ean'] for i in d['items'])))
keep, seen = [], set()
for r in sorted(recs, key=lambda r: ('複本' in r['title'], r['title'])):
    if r['eans'] in seen:
        print('skip duplicate:', r['title']); continue
    if any(o is not r and o['parentId'] == r['parentId'] and r['eans'] < o['eans'] for o in recs):
        print('skip draft (subset):', r['title']); continue
    seen.add(r['eans']); keep.append(r)
json.dump([dict(file=r['path'], date=r['date'], src=r['title']) for r in keep], open(sys.argv[2], 'w', encoding='utf-8'),
          ensure_ascii=False, indent=1)
print(len(keep), 'entries')
