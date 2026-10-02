"""
Mark codes as VERIFIED in reference/FactSet_Code_Catalog.csv using a workbook the user
refreshed in Excel + FactSet.

  python tools/update_catalog.py REFRESHED.xlsx --source "HBAN code test (2026-10)"
  python tools/update_catalog.py REFRESHED.xlsx --dry-run

Only positive evidence is recorded: a code that returned numbers is set to
"VERIFIED - returned numbers" (with sample values and a working example formula if the
row had none). Codes missing from the catalog are appended. A code that returned no data
is NOT marked "avoid" automatically - an empty cell is often coverage (FB_ outside the US,
LCR for small banks) rather than a bad code; add those warnings by hand.
"""
import argparse
import csv
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from fds_kit import harvest_codes  # noqa: E402

CATALOG = Path(__file__).resolve().parents[1] / 'reference' / 'FactSet_Code_Catalog.csv'
VERIFIED = 'VERIFIED - returned numbers'
FAMILY = {
    'FFI': 'FFI (FactSet Fundamentals Industry - bank/SF/insurance detail, global)',
    'FF': 'FF (FactSet Fundamentals - standardized, global)',
    'FB': 'FB (FactSet regulatory - US bank holding cos. FR Y-9C / call reports)',
    'P': 'P (FactSet prices)', 'FG': 'FG (FactSet global reference)', 'FREF': 'FREF (FactSet reference)',
}

ap = argparse.ArgumentParser()
ap.add_argument('workbook')
ap.add_argument('--source', default=None, help='label for the "Source file" column of new rows')
ap.add_argument('--dry-run', action='store_true')
a = ap.parse_args()

with open(CATALOG, encoding='utf-8', newline='') as f:
    reader = csv.DictReader(f)
    fields = reader.fieldnames
    rows = list(reader)
by_code = {}
for r in rows:
    by_code.setdefault(r['Code'], r)

COL_V, COL_S, COL_X = 'Verified in a refreshed client file?', 'Sample values returned', 'Working example formula'
source = a.source or Path(a.workbook).name
updated, added = [], []
for key, d in sorted(harvest_codes(a.workbook).items()):
    if not d['numbers']:
        continue
    code = key.removeprefix('FE:')
    example = (d['example'] or '').split(': ', 1)[-1]
    samples = '; '.join(f'{x:.6g}' for x in d['samples'])
    r = by_code.get(code)
    if r is None:
        prefix = code.split('_')[0]
        r = {k: '' for k in fields}
        r.update({'Code': code, 'Family': FAMILY.get(prefix, 'FE (FactSet estimates item)' if key.startswith('FE:') else prefix),
                  'Description': '(added from a refreshed workbook - no FactSet description)', 'Source file': source})
        rows.append(r)
        by_code[code] = r
        added.append(code)
    if r[COL_V].startswith('VERIFIED'):
        continue
    if r[COL_V].startswith('TESTED'):
        print(f'  note: {code} was marked "returned no data" and now returned numbers - check its syntax/notes')
    r[COL_V] = VERIFIED
    r[COL_S] = r[COL_S] or samples
    r[COL_X] = r[COL_X] or example
    updated.append(code)

print(f'newly verified: {len(updated)}  (appended to catalog: {len(added)})')
for c in updated:
    print('  ', c, '(new row)' if c in added else '')
if not a.dry_run and updated:
    with open(CATALOG, 'w', encoding='utf-8', newline='') as f:
        w = csv.DictWriter(f, fieldnames=fields)
        w.writeheader()
        w.writerows(rows)
    print('catalog written:', CATALOG)
