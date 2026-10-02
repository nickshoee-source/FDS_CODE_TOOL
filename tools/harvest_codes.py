"""
After the user refreshes a workbook in Excel (FactSet add-in) and sends it back,
run this to see which codes actually returned data.

  python tools/harvest_codes.py REFRESHED.xlsx            # table to screen
  python tools/harvest_codes.py REFRESHED.xlsx --csv out.csv

A code with numbers > 0 is proven to work for that syntax; numbers == 0 means
check periodicity / currency / coverage, or the code is wrong.
"""
import argparse
import csv
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from fds_kit import harvest_codes  # noqa

ap = argparse.ArgumentParser()
ap.add_argument('workbook')
ap.add_argument('--csv')
a = ap.parse_args()
res = harvest_codes(a.workbook)
rows = sorted(res.items(), key=lambda kv: (kv[1]['numbers'] == 0, kv[0]))
for code, d in rows:
    flag = 'OK ' if d['numbers'] else 'NO DATA'
    print(f"{flag:8} {code:40} cells={d['cells']:5} numbers={d['numbers']:5} na/err={d['na_or_error']:5} "
          f"text={d['text']:5} samples={[round(x, 4) for x in d['samples']]}")
if a.csv:
    with open(a.csv, 'w', newline='') as fh:
        w = csv.writer(fh)
        w.writerow(['code', 'cells', 'numbers', 'na_or_error', 'text', 'samples', 'example'])
        for code, d in rows:
            w.writerow([code, d['cells'], d['numbers'], d['na_or_error'], d['text'], d['samples'], d['example']])
    print('wrote', a.csv)
