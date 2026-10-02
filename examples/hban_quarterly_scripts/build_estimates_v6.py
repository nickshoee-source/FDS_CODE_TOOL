"""
V6: add a forward-looking "- Estimates" copy of every worksheet, covering
2027/1F .. 2029/4F (12 quarters), interleaved after its source sheet.

Each Estimates sheet is rebuilt rather than bulk-copied, because the period
count changes from 28 to 12, so every metric block shrinks from 29 columns
(28 periods + anchor) to 13 (12 periods + anchor) and every formula has to be
regenerated for its new position.

Preserved from the source sheet: metric order, categories, the exact FactSet
code of each metric, identifiers, company-name pulls, styling, number formats,
column grouping, and the Valuation share-price footnote.

NOT carried over: the loose experiment/scratch cells sitting outside the
metric blocks (Valuation CK19:CL20, Shareholder Returns BK20/BK21/DQ25,
Profitability E19:E20, Revenue & Fee Mix AD19/BG21, NIM & Margin ET19/CN20/
DQ20/ET20, Credit CN19, Capital E20).

Formulas are copied as-is apart from repositioning. Reported-fundamentals
codes (FF_/FB_/FA_) will not return data for 2027-2029 -- swapping those for
FE_ESTIMATE consensus equivalents is a separate decision for the user.
"""
import re
import copy
from openpyxl import load_workbook
from openpyxl.worksheet.formula import ArrayFormula
from openpyxl.utils import get_column_letter, column_index_from_string

SRC = '/home/claude/Quarterly_Template_V5.xlsx'
OUT = '/home/claude/Quarterly_Template_V6.xlsx'

SRC_PERIODS = 28
EST_PERIODS = [f'{y}/{q}F' for y in (2027, 2028, 2029) for q in range(1, 5)]
N_EST = len(EST_PERIODS)                 # 12
EST_BLOCK_W = N_EST + 1                  # 13
N_BANKS = 11

# Excel caps sheet names at 31 chars
NAME_OVERRIDE = {'Balance Sheet Evolution': 'Balance Sheet Evol'}

wb = load_workbook(SRC, data_only=False)


def txt(c):
    v = c.value
    return v.text if isinstance(v, ArrayFormula) else v


def s_(v):
    return v if isinstance(v, str) else ''


def C(i):
    return get_column_letter(i)


def block_rows(cat_row):
    first = cat_row + 5
    return cat_row, cat_row + 1, cat_row + 2, cat_row + 4, first, first + N_BANKS - 1


def metric_runs(ws, met_row):
    out, cur, st = [], None, None
    for c in range(5, ws.max_column + 2):
        m = txt(ws.cell(row=met_row, column=c))
        if not isinstance(m, str):
            m = None
        if m != cur:
            if cur is not None:
                out.append((cur, st, c - 1))
            cur, st = m, c
    return out


ID_RE = re.compile(r'(?<![A-Za-z_])\$?D\$?\d+')
PER_RE = re.compile(r'(?<![A-Z$])(\$?)([A-Z]{1,3})(\$?)(\d+)')


def parameterize(formula, own_start):
    t = s_(formula).replace('_xll.', '')
    t = ID_RE.sub('$D{ROW}', t)

    def repl(m):
        d1, col, d2, row = m.groups()
        k = column_index_from_string(col) - own_start
        if 0 <= k <= SRC_PERIODS:
            return f'{d1}{{P+{k}}}{d2}{row}'
        return m.group(0)
    return PER_RE.sub(repl, t)


def render(template, row, start_col, i):
    out = template.replace('{ROW}', str(row))
    for k in range(SRC_PERIODS + 1):
        tok = f'{{P+{k}}}'
        if tok in out:
            out = out.replace(tok, C(start_col + i + k))
    return out


# ------------------------------------------------ cell-math rebuild recipes
def cm_buyback(P, r, i):
    return f"={C(P[(3,'Share Repurchases')]+i)}{r}/{C(P[(3,'Market Cap (QTR_R)')]+i)}{r}"

def cm_capreturn(P, r, i):
    return (f"=(({C(P[(3,'Dividend Per Share')]+i)}{r}"
            f"*{C(P[(3,'Common Shares Outstanding')]+i)}{r})"
            f"+{C(P[(3,'Stock Repurchase CF')]+i)}{r})"
            f"/{C(P[(3,'Market Value (Company)')]+i)}{r}")

def cm_totalrev(P, r, i):
    return f"={C(P[(3,'NII')]+i)}{r}+{C(P[(3,'Non-Interest Income')]+i)}{r}"

def cm_cof(P, r, i):
    return (f"=(({C(P[(3,'Interest Expense Total')]+i)}{r}"
            f"/{C(P[(3,'Average Interest Bearing Liabilities')]+i)}{r})*100)*4")

def cm_tbv(P, r, i):
    rr = r - 19
    return (f"={C(P[(3,'Total Assets EOP')]+i)}{rr}"
            f"-{C(P[(3,'Intangible Assets')]+i)}{rr}"
            f"-{C(P[(3,'Total Liabilities')]+i)}{rr}")

CELLMATH = {
    ('Shareholder Returns', 3, 'Buyback Yield'): cm_buyback,
    ('Shareholder Returns', 3, 'Capital Return Yield'): cm_capreturn,
    ('Revenue & Fee Mix', 3, 'Total Revenue'): cm_totalrev,
    ('NIM & Margin', 3, 'Cost of Funds'): cm_cof,
    ('Balance Sheet Evolution', 22, 'TBV'): cm_tbv,
}

source_names = list(wb.sheetnames)
report = []

for src_name in source_names:
    src = wb[src_name]
    base = NAME_OVERRIDE.get(src_name, src_name)
    est_name = f'{base} - Estimates'
    assert len(est_name) <= 31, est_name

    dst = wb.create_sheet(title=est_name)
    dst.sheet_view.showGridLines = src.sheet_view.showGridLines

    # title
    t = s_(txt(src.cell(row=1, column=1)))
    new_title = re.sub(r'Quarterly \(.*?\)', f'Estimates ({EST_PERIODS[0]} – {EST_PERIODS[-1]})', t)
    if new_title == t:
        new_title = f'{src_name} — Estimates ({EST_PERIODS[0]} – {EST_PERIODS[-1]})'
    dst.cell(row=1, column=1).value = new_title
    dst.cell(row=1, column=1)._style = copy.copy(src.cell(row=1, column=1)._style)
    dst.row_dimensions[1].height = src.row_dimensions[1].height

    for col in ('A', 'B', 'C', 'D'):
        if src.column_dimensions[col].width:
            dst.column_dimensions[col].width = src.column_dimensions[col].width

    positions, metrics_written = {}, []

    for cat_row in (3, 22):
        _, met_row, per_row, id_hdr, first, last = block_rows(cat_row)
        if txt(src.cell(row=met_row, column=5)) is None:
            continue

        # ---- header + identifier columns
        for col in (3, 4):
            sc, dc = src.cell(row=id_hdr, column=col), dst.cell(row=id_hdr, column=col)
            dc.value = txt(sc)
            dc._style = copy.copy(sc._style)
        for r in range(first, last + 1):
            sc, dc = src.cell(row=r, column=4), dst.cell(row=r, column=4)
            dc.value = txt(sc)
            dc._style = copy.copy(sc._style)
            sn, dn = src.cell(row=r, column=3), dst.cell(row=r, column=3)
            dn.value = f'=FDS(D{r},"PROPER_NAME(,,,""COMPANY_SHORT"")")'
            dn._style = copy.copy(sn._style)

        runs = metric_runs(src, met_row)

        # ---- snapshot templates/styles from source
        snap = {}
        for lab, s, e in runs:
            is_cm = (src_name, cat_row, lab) in CELLMATH
            snap[lab] = {
                'category': txt(src.cell(row=cat_row, column=s)),
                'is_cm': is_cm,
                'template': None if is_cm else parameterize(
                    txt(src.cell(row=first, column=s)), s),
                # style/width per slot: periods 0..11 from source 0..11, anchor from source anchor
                'slot_src': [s + i for i in range(N_EST)] + [s + SRC_PERIODS],
            }

        # ---- assign positions, write labels/periods/styles/FDS formulas
        for idx, (lab, s, e) in enumerate(runs):
            start = 5 + idx * EST_BLOCK_W
            positions[(cat_row, lab)] = start
            m = snap[lab]
            for i in range(EST_BLOCK_W):
                col = start + i
                ssrc = m['slot_src'][i]
                for r in [cat_row, met_row, per_row] + list(range(first, last + 1)):
                    dst.cell(row=r, column=col)._style = copy.copy(
                        src.cell(row=r, column=ssrc)._style)
                dst.cell(row=cat_row, column=col).value = m['category']
                dst.cell(row=met_row, column=col).value = lab
                dst.cell(row=per_row, column=col).value = (
                    EST_PERIODS[i] if i < N_EST else None)
                w = src.column_dimensions[C(ssrc)].width
                if w:
                    dst.column_dimensions[C(col)].width = w
                dst.column_dimensions[C(col)].outlineLevel = 1 if i < N_EST else 0
                if i < N_EST and not m['is_cm']:
                    for r in range(first, last + 1):
                        cell = dst.cell(row=r, column=col)
                        cell.value = render(m['template'], r, start, i)
                        cell.number_format = '0.00'
                elif i < N_EST:
                    for r in range(first, last + 1):
                        dst.cell(row=r, column=col).number_format = '0.00'
            metrics_written.append((cat_row, lab, C(start)))

        # ---- cell-math once all positions known
        for lab, s, e in runs:
            if not snap[lab]['is_cm']:
                continue
            fn = CELLMATH[(src_name, cat_row, lab)]
            start = positions[(cat_row, lab)]
            for i in range(N_EST):
                for r in range(first, last + 1):
                    cell = dst.cell(row=r, column=start + i)
                    cell.value = fn(positions, r, i)
                    cell.number_format = '0.00'

    # ---- footnote (Valuation only)
    fn_cell = src.cell(row=39, column=3)
    if txt(fn_cell) is not None:
        d = dst.cell(row=39, column=3)
        d.value = txt(fn_cell)
        d._style = copy.copy(fn_cell._style)

    # ---- title merge + freeze
    last_col = max(c for _, _, c in
                   [(0, 0, column_index_from_string(x[2])) for x in metrics_written]) \
        if metrics_written else 4
    last_col = 4 + EST_BLOCK_W * max(
        len([1 for cr, l, c in metrics_written if cr == cat]) for cat in
        {cr for cr, l, c in metrics_written}) if metrics_written else 4
    dst.merge_cells(start_row=1, start_column=1, end_row=1, end_column=last_col)
    for c in range(1, last_col + 1):
        dst.cell(row=1, column=c)._style = copy.copy(src.cell(row=1, column=1)._style)
    dst.freeze_panes = f'E{block_rows(3)[4]}'
    if dst.sheet_properties.outlinePr is not None:
        dst.sheet_properties.outlinePr.summaryRight = True

    report.append((src_name, est_name, last_col, metrics_written))

# ------------------------------------------------------ interleave the order
order = []
for n in source_names:
    order.append(n)
    order.append(f'{NAME_OVERRIDE.get(n, n)} - Estimates')
wb._sheets = [wb[n] for n in order]

wb.save(OUT)
print('saved', OUT, '\n')
for src_name, est_name, last_col, mw in report:
    print(f'{est_name:32} ends {C(last_col):>4}   {len(mw)} metrics')
print(f'\nPeriods on every Estimates sheet: {EST_PERIODS[0]} .. {EST_PERIODS[-1]} ({N_EST} quarters)')
print('\nSheet order:')
for i, n in enumerate(wb.sheetnames, 1):
    print(f'  {i:>2}. {n}')
