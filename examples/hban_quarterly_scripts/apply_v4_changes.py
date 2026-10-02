"""
Apply the 8/31/26 round of edits to the user's live Quarterly template.

Ground rules:
  * Work IN PLACE on the user's file. They hand-edited many formulas
    (FB_ROTCE, FB_INT_MGN, FF_CHARGE_OFFS_LOANS_PCT, P_PRICE_RETURNS,
    EPS_NONGAAP, ...). Never regenerate a sheet from scratch.
  * Touch ONLY the cells the instructions call for. Every other formula
    keeps its exact existing text (including its `_xll.` prefix and array
    wrapper, which is how Excel saved it after a successful refresh).
  * New / rewritten formulas are written as plain `=FDS(...)` strings with
    no `_xll.` prefix and no array wrapper -- established earlier in this
    project as the form that avoids Excel's `{}` / `@` artifacts.
  * Every metric block is exactly 29 columns: 28 quarterly period columns
    (2020/1F .. 2026/4F) + 1 blank anchor column.
"""
import re
import copy
from openpyxl import load_workbook
from openpyxl.worksheet.formula import ArrayFormula
from openpyxl.utils import get_column_letter, column_index_from_string

SRC = '/home/claude/Quarterly_Template_V3_8-31-26_Edits.xlsx'
OUT = '/home/claude/Quarterly_Template_V4.xlsx'

N_PERIODS = 28
BLOCK_W = 29          # 28 period columns + 1 anchor column
N_BANKS = 11

wb = load_workbook(SRC, data_only=False)


# ---------------------------------------------------------------- helpers
def txt(cell):
    """Formula text regardless of array-formula wrapping."""
    v = cell.value
    return v.text if isinstance(v, ArrayFormula) else v


def metric_runs(ws, met_row):
    """[(label, start_col, end_col)] for each contiguous metric-label run."""
    runs, cur, start = [], None, None
    for c in range(5, ws.max_column + 2):
        m = txt(ws.cell(row=met_row, column=c))
        if not isinstance(m, str):
            m = None
        if m != cur:
            if cur is not None:
                runs.append((cur, start, c - 1))
            cur, start = m, c
    return runs


def find_metric(ws, met_row, label):
    for lab, s, e in metric_runs(ws, met_row):
        if lab == label:
            return s, e
    raise KeyError(f"metric {label!r} not found on {ws.title!r} row {met_row}")


def last_used_col(ws, met_row):
    runs = metric_runs(ws, met_row)
    return max(e for _, s, e in runs)


def block_rows(cat_row):
    """(cat_row, met_row, per_row, first_data_row, last_data_row)"""
    first = cat_row + 5
    return cat_row, cat_row + 1, cat_row + 2, first, first + N_BANKS - 1


def clone_style(ws, src_col, dst_col, cat_row):
    """Copy the per-cell styling of one whole metric block column-for-column."""
    _, met_row, per_row, first, last = block_rows(cat_row)
    for r in [cat_row, met_row, per_row] + list(range(first, last + 1)):
        s = ws.cell(row=r, column=src_col)
        d = ws.cell(row=r, column=dst_col)
        d._style = copy.copy(s._style)


def parameterize(formula, own_start_letter):
    """
    Turn a concrete FDS formula into a template with {ROW} / {PCOL} tokens.
    Strips the `_xll.` prefix so the result is a plain =FDS(...) formula.
    """
    t = formula.replace('_xll.', '')
    t = re.sub(r'(\$?D)\d+', r'\1{ROW}', t)
    t = re.sub(r'(?<![A-Z])' + own_start_letter + r'\$5(?![0-9])', '{PCOL}$5', t)
    return t


# ---------------------------------------------------------------- writers
def write_metric_block(ws, cat_row, start_col, category, label,
                       formula_fn, style_src_col, group=True):
    """
    Lay down a full 29-column metric block at start_col.
    formula_fn(data_row, period_col_letter, period_index) -> formula str | None
    """
    _, met_row, per_row, first, last = block_rows(cat_row)
    src_periods = [txt(ws.cell(row=per_row, column=style_src_col + i))
                   for i in range(N_PERIODS)]

    for i in range(BLOCK_W):
        col = start_col + i
        letter = get_column_letter(col)
        clone_style(ws, style_src_col + i, col, cat_row)

        ws.cell(row=cat_row, column=col).value = category
        ws.cell(row=met_row, column=col).value = label
        ws.cell(row=per_row, column=col).value = (
            src_periods[i] if i < N_PERIODS else None)

        for r in range(first, last + 1):
            cell = ws.cell(row=r, column=col)
            cell.value = (formula_fn(r, letter, i) if i < N_PERIODS else None)
            cell.number_format = '0.00'

        ws.column_dimensions[letter].width = (
            ws.column_dimensions[get_column_letter(style_src_col + i)].width or 13)
        if group and i < N_PERIODS:
            ws.column_dimensions[letter].outlineLevel = 1
    return start_col + BLOCK_W - 1


def rewrite_metric_formulas(ws, cat_row, start_col, formula_fn):
    """Replace the formulas of an existing metric block, leaving labels/styling."""
    _, _, _, first, last = block_rows(cat_row)
    for i in range(N_PERIODS):
        col = start_col + i
        letter = get_column_letter(col)
        for r in range(first, last + 1):
            c = ws.cell(row=r, column=col)
            c.value = formula_fn(r, letter, i)
            c.number_format = '0.00'


def clear_metric_block(ws, cat_row, start_col, end_col):
    _, met_row, per_row, first, last = block_rows(cat_row)
    for col in range(start_col, end_col + 1):
        for r in [cat_row, met_row, per_row] + list(range(first, last + 1)):
            ws.cell(row=r, column=col).value = None


def move_metric_block(ws, cat_row, src_start, dst_start):
    """
    Relocate a metric block, regenerating its formulas for the new columns.
    Only safe for blocks whose formulas reference their own period cells.
    """
    _, met_row, per_row, first, last = block_rows(cat_row)
    label = txt(ws.cell(row=met_row, column=src_start))
    category = txt(ws.cell(row=cat_row, column=src_start))
    src_letter = get_column_letter(src_start)
    template = parameterize(txt(ws.cell(row=first, column=src_start)), src_letter)

    def fn(r, letter, i):
        return template.format(ROW=r, PCOL=letter)

    write_metric_block(ws, cat_row, dst_start, category, label, fn,
                       style_src_col=src_start)
    return label, category, template


def fds(code_and_args):
    """=FDS($D{ROW},"<code>") builder factory for a period-driven code."""
    def fn(r, letter, i):
        return f'=FDS($D{r},"{code_and_args.format(PCOL=letter)}")'
    return fn


CHANGELOG = []


def log(sheet, what):
    CHANGELOG.append((sheet, what))


# ================================================================ 1. VALUATION
ws = wb['Valuation']
# Footnote goes under the Market Value block (rows 27-37) because that is
# where Share Price actually lives.
FOOT_ROW = 39
ws.cell(row=FOOT_ROW, column=3).value = (
    'Note: Share Price is as of the last trading day of each fiscal quarter shown.')
from openpyxl.styles import Font
ws.cell(row=FOOT_ROW, column=3).font = Font(
    name='Arial', italic=True, size=9, color='595959')
log('Valuation', f'Share Price footnote added at C{FOOT_ROW}')


# ====================================================== 2. SHAREHOLDER RETURNS
ws = wb['Shareholder Returns']
CAT = 3
_, MET, PER, FIRST, LAST = block_rows(CAT)
cursor = last_used_col(ws, MET)

sr_cols = {}
new_metrics = [
    ('Shareholder Returns', 'Share Repurchases',
     'FF_SHS_REPURCH(TOTAL_VAL_REPURCH,QTR,"&{PCOL}$5&")'),
    ('Shareholder Returns', 'Market Cap (QTR_R)',
     'FF_MKT_VAL(QTR_R,"&{PCOL}$5&")'),
    ('Shareholder Returns', 'Common Shares Outstanding',
     'FF_COM_SHS_OUT(QTR_R,"&{PCOL}$5&")'),
    ('Shareholder Returns', 'Stock Repurchase CF',
     'FF_STK_PURCH_CF(QTR_R,"&{PCOL}$5&")'),
    ('Shareholder Returns', 'Market Value (Company)',
     'FREF_MARKET_VALUE_COMPANY("&{PCOL}$5&",,,,,0,,""LEGACY"")'),
    ('Capital Return', 'Dividends Paid',
     'FF_DIV_CF(QTR,"&{PCOL}$5&")'),
]
for category, label, code in new_metrics:
    start = cursor + 1
    cursor = write_metric_block(ws, CAT, start, category, label,
                                fds(code), style_src_col=5)
    sr_cols[label] = start
    log('Shareholder Returns',
        f'added "{label}" at {get_column_letter(start)}:{get_column_letter(cursor)}')

# Buyback Yield -> traceable cell math: Share Repurchases / Market Cap (QTR_R)
bb_start, _ = find_metric(ws, MET, 'Buyback Yield')
rep, mkt = sr_cols['Share Repurchases'], sr_cols['Market Cap (QTR_R)']
rewrite_metric_formulas(
    ws, CAT, bb_start,
    lambda r, letter, i: f'={get_column_letter(rep + i)}{r}'
                         f'/{get_column_letter(mkt + i)}{r}')
log('Shareholder Returns',
    'Buyback Yield now = Share Repurchases cell / Market Cap (QTR_R) cell')

# Capital Return Yield -> traceable cell math from the four component columns
cry_start, _ = find_metric(ws, MET, 'Capital Return Yield')
dps_start, _ = find_metric(ws, MET, 'Dividend Per Share')
shs, scf, mvc = (sr_cols['Common Shares Outstanding'],
                 sr_cols['Stock Repurchase CF'],
                 sr_cols['Market Value (Company)'])
rewrite_metric_formulas(
    ws, CAT, cry_start,
    lambda r, letter, i:
        f'=(({get_column_letter(dps_start + i)}{r}*{get_column_letter(shs + i)}{r})'
        f'+{get_column_letter(scf + i)}{r})/{get_column_letter(mvc + i)}{r}')
log('Shareholder Returns',
    'Capital Return Yield now = ((DPS x Shares Out) + Stock Repurchase CF) '
    '/ Market Value (Company), all cell refs')

# Dividend Payout Ratio: LTM -> QTR
dpr_start, _ = find_metric(ws, MET, 'Dividend Payout Ratio')
rewrite_metric_formulas(ws, CAT, dpr_start,
                        fds('FF_PAY_OUT_RATIO(QTR,"&{PCOL}$5&")'))
log('Shareholder Returns', 'Dividend Payout Ratio basis LTM -> QTR')


# ======================================================= 3. REVENUE & FEE MIX
ws = wb['Revenue & Fee Mix']
CAT = 3
_, MET, PER, FIRST, LAST = block_rows(CAT)
tr_start, _ = find_metric(ws, MET, 'Total Revenue')
nii_start, _ = find_metric(ws, MET, 'NII')
nonii_start, _ = find_metric(ws, MET, 'Non-Interest Income')
rewrite_metric_formulas(
    ws, CAT, tr_start,
    lambda r, letter, i: f'={get_column_letter(nii_start + i)}{r}'
                         f'+{get_column_letter(nonii_start + i)}{r}')
log('Revenue & Fee Mix',
    'Total Revenue now = NII cell + Non-Interest Income cell')


# ============================================================ 4. NIM & MARGIN
ws = wb['NIM & Margin']
CAT = 3
_, MET, PER, FIRST, LAST = block_rows(CAT)
cursor = last_used_col(ws, MET)

ie_start = cursor + 1
cursor = write_metric_block(ws, CAT, ie_start, 'Margin', 'Interest Expense Total',
                            fds('FF_INT_EXP_TOT(QTR,"&{PCOL}$5&")'),
                            style_src_col=5)
log('NIM & Margin',
    f'added "Interest Expense Total" at {get_column_letter(ie_start)}:{get_column_letter(cursor)}')

ail_start = cursor + 1
cursor = write_metric_block(ws, CAT, ail_start, 'Margin',
                            'Average Interest Bearing Liabilities',
                            fds('FF_BK_AVG_LIABS_INTB(QTR_TOT,"&{PCOL}$5&")'),
                            style_src_col=5)
log('NIM & Margin',
    f'added "Average Interest Bearing Liabilities" at '
    f'{get_column_letter(ail_start)}:{get_column_letter(cursor)}')

cof_start, _ = find_metric(ws, MET, 'Cost of Funds')
rewrite_metric_formulas(
    ws, CAT, cof_start,
    lambda r, letter, i:
        f'=(({get_column_letter(ie_start + i)}{r}'
        f'/{get_column_letter(ail_start + i)}{r})*100)*4')
log('NIM & Margin',
    'Cost of Funds now = ((Interest Expense Total / Avg IB Liabilities) x 100) x 4, cell refs')


# ================================================== 5. BALANCE SHEET EVOLUTION
ws = wb['Balance Sheet Evolution']

# -- block 1 (Assets / Loans, rows 8-18): add Intangible Assets + Total Liabilities
CAT1 = 3
_, MET1, PER1, FIRST1, LAST1 = block_rows(CAT1)
cursor = last_used_col(ws, MET1)

intang_start = cursor + 1
cursor = write_metric_block(ws, CAT1, intang_start, 'Assets', 'Intangible Assets',
                            fds('FF_INTANG(QTR_R,"&{PCOL}$5&")'),
                            style_src_col=5)
log('Balance Sheet Evolution',
    f'added "Intangible Assets" at {get_column_letter(intang_start)}:{get_column_letter(cursor)}')

liabs_start = cursor + 1
cursor = write_metric_block(ws, CAT1, liabs_start, 'Liabilities', 'Total Liabilities',
                            fds('FF_LIABS(QTR_R,"&{PCOL}$5&")'),
                            style_src_col=5)
log('Balance Sheet Evolution',
    f'added "Total Liabilities" at {get_column_letter(liabs_start)}:{get_column_letter(cursor)}')

assets_start, _ = find_metric(ws, MET1, 'Total Assets EOP')

# -- block 2 (Deposits / Capital, rows 27-37): drop Average Deposits, compact left
CAT2 = 22
_, MET2, PER2, FIRST2, LAST2 = block_rows(CAT2)
runs2 = metric_runs(ws, MET2)
avg_dep_start, avg_dep_end = find_metric(ws, MET2, 'Average Deposits')

after = [(lab, s, e) for lab, s, e in runs2 if s > avg_dep_end]
old_last_end = max(e for _, _, e in runs2)

# shift each following block left by one block width, regenerating formulas
for lab, s, e in after:
    move_metric_block(ws, CAT2, s, s - BLOCK_W)
clear_metric_block(ws, CAT2, old_last_end - BLOCK_W + 1, old_last_end)
log('Balance Sheet Evolution',
    'removed "Average Deposits"; Tangible Common Equity / Avg TCE / TBV / TBVPS '
    'shifted left one block')

# -- TBV -> Total Assets - Intangible Assets - Total Liabilities (cell refs)
tbv_start, _ = find_metric(ws, MET2, 'TBV')
ROW_OFFSET = FIRST2 - FIRST1          # block2 row -> matching block1 row
rewrite_metric_formulas(
    ws, CAT2, tbv_start,
    lambda r, letter, i:
        f'={get_column_letter(assets_start + i)}{r - ROW_OFFSET}'
        f'-{get_column_letter(intang_start + i)}{r - ROW_OFFSET}'
        f'-{get_column_letter(liabs_start + i)}{r - ROW_OFFSET}')
log('Balance Sheet Evolution',
    f'TBV now = Total Assets - Intangible Assets - Total Liabilities '
    f'(cell refs into rows {FIRST1}-{LAST1})')


# ============================================================== 6. CAPITAL
ws = wb['Capital']
CAT = 3
_, MET, PER, FIRST, LAST = block_rows(CAT)
dp_start, dp_end = find_metric(ws, MET, 'Dividends Paid')
clear_metric_block(ws, CAT, dp_start, dp_end)
log('Capital', 'removed "Dividends Paid" (moved to Shareholder Returns)')


# ================================================================== save
for name in wb.sheetnames:
    sh = wb[name]
    if sh.sheet_properties.outlinePr is not None:
        sh.sheet_properties.outlinePr.summaryRight = True

wb.save(OUT)
print('saved', OUT)
print()
for sheet, what in CHANGELOG:
    print(f'  [{sheet}] {what}')
