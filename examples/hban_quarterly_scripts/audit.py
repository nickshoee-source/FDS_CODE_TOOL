"""
Full-workbook audit for the HBAN Quarterly template.

Usage:  python3 audit.py <file.xlsx>

Checks every sheet for:
  structure   metric blocks, block widths, period sequence, blank anchors
  identity    identifier list + company-name pull on every data row
  formulas    each data cell: is a formula (not a pasted value); identifier
              ref points at its OWN row; every period reference resolves to a
              populated period cell and to the RIGHT period for its column
  estimates   FE_TIMESERIES anchors only in the first period column, spill
              cells empty, date range = own sheet's 2027/1F..2029/4F cells
  cell-math   every reference is in-sheet, period-matched and populated
              (a spill target counts as populated)
  hygiene     array-formula ref matches its own cell, no cross-sheet refs,
              0.00 number format on data cells, title bar width, freeze pane,
              scratch cells outside the blocks, ghost formatting left behind
"""
import re
import sys
from collections import defaultdict
from openpyxl import load_workbook
from openpyxl.worksheet.formula import ArrayFormula
from openpyxl.utils import get_column_letter, column_index_from_string

sys.path.insert(0, '/home/claude')
from metric_formats import FMT, KIND, base_sheet
PATH = sys.argv[1]
wb = load_workbook(PATH, data_only=False)

HIST_P = [f'{y}/{q}F' for y in range(2020, 2027) for q in range(1, 5)]
EST_P = [f'{y}/{q}F' for y in (2027, 2028, 2029) for q in range(1, 5)]
IDS = ['FITB', 'ZION', 'RF', 'PNC', 'MTB', 'FHN', 'HBAN', 'USB', 'KEY', 'CFG', 'TFC']


def txt(c):
    v = c.value
    return v.text if isinstance(v, ArrayFormula) else v


def s_(v):
    return v if isinstance(v, str) else ''


def C(i):
    return get_column_letter(i)


def runs(ws, mr):
    out, cur, st = [], None, None
    for c in range(5, ws.max_column + 2):
        m = txt(ws.cell(row=mr, column=c))
        if not isinstance(m, str):
            m = None
        if m != cur:
            if cur is not None:
                out.append((cur, st, c - 1))
            cur, st = m, c
    return out


def next_period(p):
    y, q = int(p[:4]), int(p[5])
    return f'{y + (q == 4)}/{q % 4 + 1}F'


ISSUES = defaultdict(list)


def issue(sheet, kind, msg):
    ISSUES[(sheet, kind)].append(msg)


ID_RE = re.compile(r'(?<![A-Za-z_"])\$?D\$?(\d+)')
PREF_RE = re.compile(r'(?<![A-Z])(\$?)([A-Z]{1,3})\$(\d+)')          # period refs X$5
CELL_RE = re.compile(r'(?<![A-Z$!:])([A-Z]{1,3})(\d+)(?![0-9(])')      # plain refs AH8

for sheet in wb.sheetnames:
    if sheet in ('Scratch Notes', 'Data Check'):
        continue
    ws = wb[sheet]
    is_est = sheet.endswith('- Estimates')
    P = EST_P if is_est else HIST_P
    NP = len(P)
    W = NP + 1

    blocks = []
    for cat_row in (3, 22):
        if txt(ws.cell(row=cat_row + 1, column=5)) is not None:
            blocks.append(cat_row)

    spill_target = set()        # cells filled by a multi-cell array anchor
    spill_anchor = set()
    for row_ in ws.iter_rows():
        for cl in row_:
            if isinstance(cl.value, ArrayFormula) and ':' in cl.value.ref:
                a_, b_ = cl.value.ref.split(':')
                if a_ != b_:
                    spill_anchor.add(cl.coordinate)
                    ma = re.match(r'([A-Z]+)(\d+)', a_); mb = re.match(r'([A-Z]+)(\d+)', b_)
                    for rr_ in range(int(ma.group(2)), int(mb.group(2)) + 1):
                        for cc_ in range(column_index_from_string(ma.group(1)), column_index_from_string(mb.group(1)) + 1):
                            if f'{C(cc_)}{rr_}' != cl.coordinate:
                                spill_target.add(f'{C(cc_)}{rr_}')
    block_row_of = {}           # data row -> (cat_row, per_row)
    legit_cols = {}             # cat_row -> set of columns inside metric runs
    metric_start = {}           # (cat_row,label) -> start col
    anchor_first = {}           # (row, col) -> True if FE_TIMESERIES anchor lives at col

    # ------------------------------------------------ structure pass
    for cat_row in blocks:
        met, per, hdr, first = cat_row + 1, cat_row + 2, cat_row + 4, cat_row + 5
        last = first + 10
        for r in range(first, last + 1):
            block_row_of[r] = (cat_row, per)
        rs = runs(ws, met)
        legit_cols[cat_row] = set()
        for lab, s, e in rs:
            legit_cols[cat_row].update(range(s, e + 1))
            metric_start[(cat_row, lab)] = s
            if e - s + 1 != W:
                issue(sheet, 'structure', f'{lab}: block width {e - s + 1}, expected {W}')
            got = [txt(ws.cell(row=per, column=s + i)) for i in range(NP)]
            if got != P:
                issue(sheet, 'structure', f'{lab}: period labels wrong ({got[0]}..{got[-1]})')
            if txt(ws.cell(row=per, column=s + NP)) is not None:
                issue(sheet, 'structure', f'{lab}: anchor column has a period label')
            if len({txt(ws.cell(row=cat_row, column=s + i)) for i in range(W)}) != 1:
                issue(sheet, 'structure', f'{lab}: category label not uniform')
        # identity
        if [txt(ws.cell(row=r, column=4)) for r in range(first, last + 1)] != IDS:
            issue(sheet, 'identity', f'block row {cat_row}: identifier list differs')
        if txt(ws.cell(row=hdr, column=3)) != 'Company Name' or txt(ws.cell(row=hdr, column=4)) != 'Identifier':
            issue(sheet, 'identity', f'block row {cat_row}: header row {hdr} labels wrong')
        for r in range(first, last + 1):
            f = s_(txt(ws.cell(row=r, column=3)))
            if 'PROPER_NAME' not in f:
                issue(sheet, 'identity', f'C{r}: company-name pull missing ({f[:40]!r})')
            else:
                m = ID_RE.search(f)
                if not m or int(m.group(1)) != r:
                    issue(sheet, 'identity', f'C{r}: name pull points at wrong row')
        # duplicate labels
        labs = [l for l, _, _ in rs]
        for l in set(labs):
            if labs.count(l) > 1:
                issue(sheet, 'structure', f'duplicate metric label {l!r}')

    # ------------------------------------------------ formula pass
    for cat_row in blocks:
        met, per, first = cat_row + 1, cat_row + 2, cat_row + 5
        last = first + 10
        for lab, s, e in runs(ws, met):
            f_first = s_(txt(ws.cell(row=first, column=s)))
            is_ts = 'FE_TIMESERIES' in f_first
            for r in range(first, last + 1):
                for i in range(W):
                    col = s + i
                    cell = ws.cell(row=r, column=col)
                    v = txt(cell)
                    ref = f'{C(col)}{r}'
                    if i == NP:                                # anchor column
                        if v is not None:
                            issue(sheet, 'structure', f'{lab} anchor {ref} not blank')
                        continue
                    exp_fmt = FMT[KIND[(base_sheet(sheet), lab)]]
                    if cell.number_format != exp_fmt:
                        issue(sheet, 'format', f'{lab} {ref} number format {cell.number_format!r} expected {exp_fmt!r}')
                    if isinstance(cell.value, ArrayFormula) and ref not in spill_anchor and cell.value.ref not in (ref, f'{ref}:{ref}'):
                        issue(sheet, 'hygiene', f'{ref}: array formula ref {cell.value.ref} != own cell')
                    if is_ts:
                        if i == 0:
                            f = s_(v)
                            if 'FE_TIMESERIES' not in f:
                                issue(sheet, 'estimates', f'{lab} {ref}: anchor missing')
                                continue
                            if not f.startswith('=FDSRC('):
                                issue(sheet, 'estimates', f'{lab} {ref}: anchor not FDSRC')
                            idm = ID_RE.search(f)
                            if not idm or int(idm.group(1)) != r:
                                issue(sheet, 'formula', f'{lab} {ref}: identifier not own row')
                            pr = PREF_RE.findall(f)
                            if len(pr) != 2:
                                issue(sheet, 'estimates', f'{lab} {ref}: expected 2 date refs, got {len(pr)}')
                            else:
                                (_, c1, r1), (_, c2, r2) = pr
                                v1 = txt(ws.cell(row=int(r1), column=column_index_from_string(c1)))
                                v2 = txt(ws.cell(row=int(r2), column=column_index_from_string(c2)))
                                if (v1, v2) != (P[0], P[-1]):
                                    issue(sheet, 'estimates', f'{lab} {ref}: range {v1}..{v2}')
                                if int(r1) != per or int(r2) != per:
                                    issue(sheet, 'estimates', f'{lab} {ref}: date refs not on own period row {per}')
                            anchor_first[(r, s)] = True
                        elif v is not None:
                            issue(sheet, 'estimates', f'{lab} {ref}: spill cell not empty')
                        continue
                    # ---- regular per-cell data
                    if ref in spill_target:
                        continue
                    if v is None:
                        issue(sheet, 'formula', f'{lab} {ref}: empty data cell')
                        continue
                    if not s_(v).startswith('='):
                        issue(sheet, 'formula', f'{lab} {ref}: hardcoded value {v!r} (will not refresh)')
                        continue
                    f = s_(v)
                    if '!' in f:
                        issue(sheet, 'hygiene', f'{ref}: cross-sheet reference')
                    if 'FDS' in f.upper():
                        if is_est:
                            if not (f.startswith('=IF(DATE(') and '_xlfn.IFNA(' in f and '>TODAY(),"-",' in f):
                                issue(sheet, 'na-guard', f'{lab} {ref}: Estimates pull lacks the future-quarter IF / IFNA guard')
                        elif not (f.startswith('=_xlfn.IFNA(') or re.search(r'FDSR?C\("-",', f)):
                            issue(sheet, 'na-guard', f'{lab} {ref}: #N/A not converted to "-" (no _xlfn.IFNA)')
                        if re.search(r'(?<!_xlfn\.)\bIFNA\(', f):
                            issue(sheet, 'na-guard', f'{lab} {ref}: IFNA without _xlfn. prefix (Excel shows #NAME?)')
                        if lab == 'TSR' and '/100' in f:
                            issue(sheet, 'units', f'{ref}: TSR still divided by 100 (fraction, not whole %)')
                        if lab in ('NCO Ratio', 'Deposit Cost') and not f.endswith('*4,"-")') and not f.endswith('*4,"-"))'):
                            issue(sheet, 'units', f'{ref}: {lab} not annualized (x4)')
                        if lab == 'Efficiency Ratio' and not is_est and '*100' not in f:
                            issue(sheet, 'units', f'{ref}: Efficiency Ratio still a fraction')
                        idm = ID_RE.search(f)
                        if not idm or int(idm.group(1)) != r:
                            issue(sheet, 'formula', f'{lab} {ref}: identifier ref not own row')
                        prs = PREF_RE.findall(f)
                        own = P[i]
                        seen_own = False
                        for _, pc, prow in prs:
                            pv = txt(ws.cell(row=int(prow), column=column_index_from_string(pc)))
                            if pv is None:
                                issue(sheet, 'formula', f'{lab} {ref}: references BLANK cell {pc}${prow}')
                            elif pv == own:
                                seen_own = True
                            elif pv == next_period(own) and 'P_PRICE_RETURNS' in f:
                                pass
                            else:
                                issue(sheet, 'formula', f'{lab} {ref}: period ref {pc}${prow}={pv} but column is {own}')
                        if prs and not seen_own:
                            issue(sheet, 'formula', f'{lab} {ref}: never references its own period {own}')
                    else:
                        # cell-math
                        for rc, rr in CELL_RE.findall(f):
                            rci, rr = column_index_from_string(rc), int(rr)
                            tgt = txt(ws.cell(row=rr, column=rci))
                            if rr not in block_row_of:
                                issue(sheet, 'cellmath', f'{lab} {ref}: refers outside data rows ({rc}{rr})')
                                continue
                            tcat, tper = block_row_of[rr]
                            tp = txt(ws.cell(row=tper, column=rci))
                            if tp != P[i]:
                                issue(sheet, 'cellmath', f'{lab} {ref}: {rc}{rr} period {tp} != {P[i]}')
                            if ws.cell(row=rr, column=4).value != ws.cell(row=r, column=4).value:
                                issue(sheet, 'cellmath', f'{lab} {ref}: {rc}{rr} is a different bank')
                            if tgt is None:
                                # spill target is fine
                                spill_ok = any(anchor_first.get((rr, st)) or
                                               'FE_TIMESERIES' in s_(txt(ws.cell(row=rr, column=st)))
                                               for st in range(max(5, rci - NP + 1), rci + 1))
                                if not spill_ok:
                                    issue(sheet, 'cellmath', f'{lab} {ref}: {rc}{rr} is EMPTY')
                        if not f.upper().startswith('=IFERROR('):
                            ISSUES[(sheet, '_cm_unwrapped')].append(ref)
                        if lab in ('Buyback Yield', 'Capital Return Yield') and not f.endswith(')*100,"-")'):
                            issue(sheet, 'units', f'{ref}: {lab} not converted to whole-number %')

    # ------------------------------------------------ hygiene pass
    block_rows = set()
    for cat_row in blocks:
        block_rows |= {1, cat_row, cat_row + 1, cat_row + 2, cat_row + 4} | set(range(cat_row + 5, cat_row + 16))
    for row in ws.iter_rows():
        for cell in row:
            r, c = cell.row, cell.column
            v = txt(cell)
            if r == 39 and c == 3 and sheet.startswith('Valuation'):
                continue                                          # footnote
            in_block = False
            for cat_row in blocks:
                rows_b = {cat_row, cat_row + 1, cat_row + 2} | set(range(cat_row + 5, cat_row + 16))
                if r in rows_b and (c in legit_cols[cat_row] or c in (3, 4)):
                    in_block = True
                if r == cat_row + 4 and c in (3, 4):
                    in_block = True
            if r == 1 or (r, c) == (2, 1):          # title / units note
                in_block = True
            if in_block:
                continue
            if v is not None:
                issue(sheet, 'scratch', f'{C(c)}{r} = {str(v)[:70]}')
            elif cell.has_style and (cell.fill.fill_type or any(
                    getattr(cell.border, sd).style for sd in ('left', 'right', 'top', 'bottom'))):
                ISSUES[(sheet, '_ghost')].append(f'{C(c)}{r}')

    # title bar
    last_metric = max([max(legit_cols[b]) for b in blocks] or [4])
    merges = [m for m in ws.merged_cells.ranges if m.min_row == 1]
    if not merges:
        issue(sheet, 'hygiene', 'title not merged')
    elif merges[0].max_col != last_metric:
        issue(sheet, 'hygiene', f'title bar ends {C(merges[0].max_col)}, content ends {C(last_metric)}')
    if ws.freeze_panes != 'E8':
        issue(sheet, 'view', f'freeze pane at {ws.freeze_panes} (expected E8)')

# ------------------------------------------------ report
print(f'AUDIT: {PATH}\nsheets: {len(wb.sheetnames)}\n')
kinds = ['structure', 'identity', 'formula', 'estimates', 'cellmath', 'na-guard', 'units', 'format',
         'hygiene', 'view', 'scratch']
total = 0
for sheet in wb.sheetnames:
    lines = []
    for k in kinds:
        for m in ISSUES.get((sheet, k), []):
            lines.append(f'  [{k}] {m}')
    g = ISSUES.get((sheet, '_ghost'), [])
    if g:
        lines.append(f'  [ghost-format] {len(g)} empty cells still carry fills/borders '
                     f'(e.g. {", ".join(g[:4])})')
    u = ISSUES.get((sheet, '_cm_unwrapped'), [])
    if u:
        lines.append(f'  [cellmath] {len(u)} cell-math formulas not error-guarded')
    total += len(lines)
    print(f'{sheet}: {"clean" if not lines else f"{len(lines)} finding(s)"}')
    for l in lines[:25]:
        print(l)
    if len(lines) > 25:
        print(f'  ... +{len(lines) - 25} more')
print(f'\nTOTAL finding lines: {total}')
