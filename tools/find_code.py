"""
Search the FactSet code catalog.

  python tools/find_code.py "net interest margin"
  python tools/find_code.py leverage --family FFI
  python tools/find_code.py NPL --verified          # only codes proven to return data
  python tools/find_code.py FFI_NPL_LOAN_RATIO       # exact code lookup

All words must match (code or description, case-insensitive).
"""
import argparse
import csv
from pathlib import Path

CATALOG = Path(__file__).resolve().parent.parent / 'reference' / 'FactSet_Code_Catalog.csv'

ap = argparse.ArgumentParser()
ap.add_argument('words', nargs='+')
ap.add_argument('--family', help='FF, FFI, FB, FE, P, FG, FREF, FMA')
ap.add_argument('--verified', action='store_true', help='only codes that returned numbers in a refreshed file')
ap.add_argument('--limit', type=int, default=40)
a = ap.parse_args()

words = [w.lower() for w in a.words]
rows = list(csv.DictReader(open(CATALOG, encoding='utf-8')))
hits = []
for r in rows:
    hay = (r['Code'] + ' ' + r['Description'] + ' ' + r['Section / Statement']).lower()
    if not all(w in hay for w in words):
        continue
    if a.family and not r['Family'].upper().startswith(a.family.upper()):
        continue
    ver = r['Verified in a refreshed client file?']
    if a.verified and not ver.startswith('VERIFIED'):
        continue
    hits.append(r)
# verified first, then exact-code matches, then shorter codes
hits.sort(key=lambda r: (not r['Verified in a refreshed client file?'].startswith('VERIFIED'),
                         r['Code'].lower() not in words, len(r['Code'])))
for r in hits[:a.limit]:
    v = r['Verified in a refreshed client file?'] or 'not yet tested'
    print(f"{r['Code']:<42} [{r['Family'].split(' ')[0]}] {r['Description'][:90]}")
    print(f"{'':42} {v}" + (f" | e.g. {r['Sample values returned'][:40]}" if r['Sample values returned'] else ''))
    if r['Working example formula']:
        print(f"{'':42} example: {r['Working example formula'][:150]}")
    elif r['Syntax pattern']:
        print(f"{'':42} syntax:  {r['Syntax pattern'][:150]}")
print(f'\n{len(hits)} match(es)' + (f', showing {a.limit}' if len(hits) > a.limit else ''))
