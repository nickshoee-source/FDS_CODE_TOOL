"""
Scan every cached FactSet value from the user's refreshed 8/31 workbook for
values that are probably wrong. Values are first converted into V10 display
units (e.g. NCO x4, Efficiency Ratio x100) and every flag is mapped to the
cell where that value now lives in V10.

Tests
  A  constant series      one bank shows the same value every quarter
  B  duplicate series     two metrics return identical numbers everywhere
  C  out of range         outside a plausible band for a US regional bank
  D  outlier              far outside that bank's own history
  E  jump                 balance-sheet item moves >30% in one quarter
  F  cross-checks         NIBD + IB deposits vs total; dividend yield vs DPS/price
  G  data gaps            #N/A share per metric / runs of #N/A
"""
import json
import statistics as st
from collections import defaultdict
from openpyxl import load_workbook
from openpyxl.worksheet.formula import ArrayFormula
from openpyxl.utils import get_column_letter

SRC = '/home/claude/Quarterly_Template_V3_8-31-26_Edits.xlsx'
NEW = '/home/claude/Quarterly_Template_V10.xlsx'
F = load_workbook(SRC)
V = load_workbook(SRC, data_only=True)
N = load_workbook(NEW)


def t(c):
    v = c.value
    return v.text if isinstance(v, ArrayFormula) else v


def runs(ws, mr):
    out, cur, s0 = [], None, None
    for c in range(5, ws.max_column + 2):
        m = t(ws.cell(row=mr, column=c))
        if not isinstance(m, str):
            m = None
        if m != cur:
            if cur is not None:
                out.append((cur, s0, c - 1))
            cur, s0 = m, c
    return out


PER = [f'{y}/{q}F' for y in range(2020, 2027) for q in range(1, 5)]
RENAME = {'Market Cap (QTR_R)': 'Market Cap'}
SCALE = {'Efficiency Ratio': 100, 'TSR': 100, 'Buyback Yield': 100,
         'Capital Return Yield': 100, 'NCO Ratio': 4, 'Deposit Cost': 4}
SKIP = {'ACL Ratio', 'Average Deposits', 'Dividends Paid'}         # removed / was on wrong code
FIXED_BEFORE = {'Effective Tax Rate', 'TBVPS'}                       # 8/31 values were broken, fixed since

# plausible bands for US regional banks, in V10 display units
BAND = {
    'P/E': (3, 40), 'P/B': (0.3, 3.0), 'P/TBV': (0.4, 4.5), 'Share Price': (3, 400),
    'Market Cap': (500, 200000), 'Enterprise Value': (0, 400000),
    'TSR': (-60, 60), 'Dividend Yield': (0.5, 10), 'Buyback Yield': (0, 5),
    'Dividend Per Share': (0, 3), 'Capital Return Yield': (0, 8), 'Dividend Payout Ratio': (0, 150),
    'ROTCE': (-15, 35), 'ROE': (-15, 25), 'ROA': (-1.5, 2.5), 'EPS': (-5, 12),
    'Efficiency Ratio': (35, 95), 'NIM': (1.5, 5.0), 'EA Yield': (1.5, 8.0), 'Cost of Funds': (0, 6),
    'Deposit Cost': (0, 5), 'Interest Bearing Deposit Cost': (0, 6), 'Total Funding Cost': (0, 6),
    'Net Interest Spread': (0.5, 5), 'Taxable Equivalent NIM': (1.5, 5.0),
    'NCO Ratio': (-0.3, 3), 'NPL Ratio': (0, 5), 'NPA Ratio': (0, 5), 'CET1 Ratio': (6, 16),
    'Consensus ROTCE': (-15, 35), 'Consensus EPS': (-5, 12), 'Branch Count': (50, 5000),
}
POSITIVE = {'Net Income', 'PPNR', 'Total Revenue', 'NII', 'Non-Interest Income', 'NIE',
            'Personnel Expense', 'Consensus Revenue', 'Consensus NII', 'Consensus Fee Income',
            'Tangible Common Equity', 'Average Tangible Common Equity', 'TBV', 'Market Cap',
            'Enterprise Value'}
BALANCE = {'Total Assets EOP', 'Average Assets', 'Earning Assets EOP', 'Total Loans EOP',
           'Total Deposits EOP', 'Tangible Common Equity', 'Risk Weighted Assets',
           'NIBD Balance', 'Interest Bearing Deposits', 'Total Deposits', 'Branch Count'}

# ---- collect data: series[(sheet, metric)][bank] = [(period, value_or_NA)]
series = defaultdict(lambda: defaultdict(list))
banks = {}
for sh in F.sheetnames:
    wf, wv = F[sh], V[sh]
    for cr in (3, 22):
        if t(wf.cell(row=cr + 1, column=5)) is None:
            continue
        for lab, s0, e in runs(wf, cr + 1):
            lab2 = RENAME.get(lab, lab)
            for r in range(cr + 5, cr + 16):
                bank = wv.cell(row=r, column=4).value
                banks[(sh, r)] = bank
                for i in range(28):
                    v = wv.cell(row=r, column=s0 + i).value
                    if isinstance(v, (int, float)):
                        v = v * SCALE.get(lab2, 1)
                    series[(sh, lab2)][(bank, r)].append((PER[i], v))

# ---- V10 location for (sheet, metric, row, period)
v10pos = {}
for sh in N.sheetnames:
    if sh.endswith('- Estimates') or sh in ('Scratch Notes',):
        continue
    ws = N[sh]
    for cr in (3, 22):
        if t(ws.cell(row=cr + 1, column=5)) is None:
            continue
        for lab, s0, e in runs(ws, cr + 1):
            v10pos[(sh, lab)] = (s0, cr)


def cell_of(sh, lab, row, period):
    s0, cr = v10pos[(sh, lab)]
    return f'{get_column_letter(s0 + PER.index(period))}{row}'


flags = []


def flag(kind, sev, sh, lab, bank=None, row=None, period=None, value=None, why='', cause=''):
    cell = cell_of(sh, lab, row, period) if (row and period and (sh, lab) in v10pos) else ''
    flags.append(dict(kind=kind, sev=sev, sheet=sh, metric=lab, bank=bank or '', period=period or '',
                      cell=cell, value=value, why=why, cause=cause))


nums = lambda pts: [v for p, v in pts if isinstance(v, (int, float))]

for (sh, lab), bybank in series.items():
    if lab in SKIP:
        continue
    allv = [v for pts in bybank.values() for v in nums(pts)]
    tot = sum(len(p) for p in bybank.values())
    na = sum(1 for pts in bybank.values() for p, v in pts if v == '#N/A')

    # A: constant series
    const = [b for (b, r), pts in bybank.items() if len(nums(pts)) >= 8 and len(set(round(x, 6) for x in nums(pts))) == 1]
    if len(const) >= 3:
        flag('A', 'fixed' if lab in FIXED_BEFORE else 'likely error', sh, lab,
             why=f'same value in every quarter for {len(const)} of 11 banks',
             cause='period reference pointed at a blank cell (fixed in an earlier round)' if lab in FIXED_BEFORE
             else 'period is not reaching the FactSet code')
        if lab in FIXED_BEFORE:
            continue

    if lab in FIXED_BEFORE:
        continue

    # G: data gaps
    if tot and na / tot >= 0.10:
        worst = sorted(((sum(1 for p, v in pts if v == '#N/A'), b) for (b, r), pts in bybank.items()), reverse=True)[:3]
        flag('G', 'check', sh, lab, why=f'{100 * na / tot:.0f}% of cells are #N/A (now shows "-"); most gaps: '
             + ', '.join(f'{b} {n}' for n, b in worst if n), cause='FactSet has no data for these quarters/banks')

    for (bank, row), pts in bybank.items():
        vals = [(p, v) for p, v in pts if isinstance(v, (int, float))]
        if not vals:
            continue
        # C: out of plausible band
        if lab in BAND:
            lo, hi = BAND[lab]
            for p, v in vals:
                if not (lo <= v <= hi):
                    flag('C', 'check', sh, lab, bank, row, p, v, f'outside plausible range {lo} to {hi}')
        if lab in POSITIVE:
            for p, v in vals:
                if v < 0:
                    flag('C', 'check', sh, lab, bank, row, p, v, 'negative value for a normally-positive item')
        # D: outlier vs own history (robust)
        xs = [v for p, v in vals]
        if len(xs) >= 8:
            med = st.median(xs)
            mad = st.median([abs(x - med) for x in xs]) or 1e-9
            for p, v in vals:
                z = abs(v - med) / (1.4826 * mad)
                if z > 10 and abs(v - med) > 0.5 * abs(med) + 1e-9:
                    flag('D', 'check', sh, lab, bank, row, p, v,
                         f'{z:.0f} robust std-devs from this bank\'s own median ({med:,.2f})')
        # E: balance-sheet jumps
        if lab in BALANCE:
            for (p0, v0), (p1, v1) in zip(vals, vals[1:]):
                if v0 and PER.index(p1) - PER.index(p0) == 1:
                    ch = v1 / v0 - 1
                    if ch > 0.30 or ch < -0.25:
                        flag('E', 'check', sh, lab, bank, row, p1, v1,
                             f'{ch:+.0%} vs prior quarter ({v0:,.0f})')

# B: duplicate series across metrics
keys = [k for k in series if k[1] not in SKIP]
for i, a in enumerate(keys):
    for b in keys[i + 1:]:
        if a[1] == b[1]:
            continue
        va = [v for pts in series[a].values() for p, v in pts]
        vb = [v for pts in series[b].values() for p, v in pts]
        pairs = [(x, y) for x, y in zip(va, vb) if isinstance(x, (int, float)) and isinstance(y, (int, float))]
        if len(pairs) > 100 and all(abs(x - y) < 1e-9 for x, y in pairs):
            flag('B', 'info' if {a[1], b[1]} == {'Total Deposits EOP', 'Total Deposits'} else 'likely error',
                 a[0], a[1], why=f'identical to {b[0]} / {b[1]} in all {len(pairs)} cells')

# F: cross-checks
dc = series[('Deposit Competition', 'NIBD Balance')]
ib = series[('Deposit Competition', 'Interest Bearing Deposits')]
td = series[('Deposit Competition', 'Total Deposits')]
for key in td:
    for (p, t_), (_, n_), (_, i_) in zip(td[key], dc[key], ib[key]):
        if all(isinstance(x, (int, float)) for x in (t_, n_, i_)) and t_:
            gap = (n_ + i_) / t_ - 1
            if abs(gap) > 0.05:
                flag('F', 'check', 'Deposit Competition', 'Total Deposits', key[0], key[1], p, t_,
                     f'NIBD + interest-bearing = {n_ + i_:,.0f}, {gap:+.0%} vs total deposits')
dy, dps, px = (series[('Shareholder Returns', 'Dividend Yield')], series[('Shareholder Returns', 'Dividend Per Share')],
               series[('Valuation', 'Share Price')])
for key in dy:
    k2 = (key[0], key[1] + 19)                                      # Valuation share price is in the lower block
    for (p, y), (_, d), (_, pr) in zip(dy[key], dps[key], px.get(k2, [])):
        if all(isinstance(x, (int, float)) for x in (y, d, pr)) and pr and y:
            implied = d * 4 / pr * 100
            if abs(implied / y - 1) > 0.25:
                flag('F', 'check', 'Shareholder Returns', 'Dividend Yield', key[0], key[1], p, y,
                     f'DPS x4 / price implies {implied:.2f}% vs {y:.2f}% shown')

json.dump(flags, open('/home/claude/value_flags.json', 'w'), indent=1, default=str)
from collections import Counter
print('flags:', len(flags))
print(Counter((f['kind'], f['sev']) for f in flags))
print(Counter((f['sheet'], f['metric']) for f in flags).most_common(40))
