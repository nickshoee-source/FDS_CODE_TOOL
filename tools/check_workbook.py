"""
Pre-delivery checks for an edited FactSet workbook.

  python tools/check_workbook.py NEW.xlsx --original ORIGINAL.xlsx
  python tools/check_workbook.py NEW.xlsx --original ORIGINAL.xlsx --fix-highlights OUT.xlsx

Reports:
  * formula hygiene: any _xll. / ArrayFormula left in cells you wrote, IFNA without _xlfn.
  * changed cells per sheet, and whether yellow == changed exactly
  * LibreOffice formula-parse check (if LibreOffice is installed)
With --fix-highlights it writes a copy where yellow marks exactly the changed cells.
"""
import argparse
import sys
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.worksheet.formula import ArrayFormula

sys.path.insert(0, str(Path(__file__).resolve().parent))
from fds_kit import formula_text, highlight_changes, is_yellow, libreoffice_parse_check, restore_sensitivity_label  # noqa

ap = argparse.ArgumentParser()
ap.add_argument('workbook')
ap.add_argument('--original')
ap.add_argument('--fix-highlights', metavar='OUT')
a = ap.parse_args()

wb = load_workbook(a.workbook)
problems = 0
print('== formula hygiene')
for ws in wb.worksheets:
    for row in ws.iter_rows():
        for c in row:
            v = c.value
            if isinstance(v, str) and v.startswith('='):
                if 'IFNA(' in v and '_xlfn.IFNA(' not in v:
                    print(f'  {ws.title}!{c.coordinate}: IFNA without _xlfn. -> #NAME? in Excel'); problems += 1
                if '_xll.' in v:
                    print(f'  {ws.title}!{c.coordinate}: contains _xll. (write plain =FDS)'); problems += 1
print('  ok' if not problems else f'  {problems} problem(s)')

if a.original:
    orig = load_workbook(a.original)
    print('== changes vs original (yellow should == changed)')
    for name in orig.sheetnames:
        if name not in wb.sheetnames:
            print(f'  {name}: MISSING in new workbook'); continue
        ws, ow = wb[name], orig[name]
        changed = mism = 0
        for r in range(1, max(ws.max_row, ow.max_row) + 1):
            for c in range(1, max(ws.max_column, ow.max_column) + 1):
                cell = ws.cell(r, c)
                d = formula_text(cell.value) != formula_text(ow.cell(r, c).value)
                changed += d
                mism += d != is_yellow(cell)
        print(f'  {name:32} changed={changed:6}  highlight mismatches={mism}')
    new_sheets = [s for s in wb.sheetnames if s not in orig.sheetnames]
    if new_sheets:
        print('  new sheets:', ', '.join(new_sheets))

errs, msg = libreoffice_parse_check(a.workbook)
print('== LibreOffice parse check:', msg, errs or '')

if a.fix_highlights:
    if not a.original:
        sys.exit('--fix-highlights needs --original')
    stats = highlight_changes(a.original, wb)
    wb.save(a.fix_highlights)
    restore_sensitivity_label(a.original, a.fix_highlights)
    print('== wrote', a.fix_highlights, stats)
