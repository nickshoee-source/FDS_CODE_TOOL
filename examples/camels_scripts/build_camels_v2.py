"""
CAMELS workbook (new client) -- fill blacked-out cells, replace wrong codes,
yellow-highlight every changed cell, Quarterly-Template-style formatting.

Rules followed
  * Only codes whose item was seen returning real numbers in earlier refreshed
    FactSet files are used as fixes/fallbacks ("verified").  Anything else is
    written in but listed as PLEASE CONFIRM.
  * Fallbacks use the client's own FQL '@' syntax (X@Y = use Y when X is NA),
    so banks that already have data keep exactly the same value.
  * New / changed formulas are written as plain =FDSC(...) (no _xll., no array
    wrapper) so nothing extra shows in the formula bar.
"""
import re
from copy import copy
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.formatting.rule import FormulaRule
from openpyxl.utils import get_column_letter as L
from openpyxl.worksheet.formula import ArrayFormula

SRC = '/home/claude/camels/original.xlsx'
OUT = '/home/claude/camels/CAMELS_FDS_Code_Analysis_V2.xlsx'
wb = load_workbook(SRC)
FM, H = wb['Formatted Annual Camels'], wb['CAMELS HISTORICAL ANNUAL']
NTM, EST, LTM = wb['CAMELS NTM'], wb['CAMELS ESTIMATE ANNUAL'], wb['CAMELS LTM']

YELLOW = PatternFill('solid', fgColor='FFFFFF00')
FIRST, LAST = 3, 671            # CAMELS HISTORICAL ANNUAL company columns C..YU (669 banks)
log = []                        # (sheet, row, metric, cells, before, after, why, status)


def txt(v):
    return v.text if isinstance(v, ArrayFormula) else v


def fdsc(idref, code):
    return f'=FDSC("-",{idref},"{code}")'


# ---------------------------------------------------------------- historical
def ffi(item, ccy=False, y='{Y}'):
    return f'{item}(ANN_L,{y}' + (',,,,USD)' if ccy else ')')


def ff(item, ccy=False, y='{Y}'):
    return f'{item}(ANN_R,{y}' + (',,,,USD)' if ccy else ')')


HIST = {
    # row: (new FQL with {Y} placeholder, why, status)
    # One code family per cell. FF_ codes are the ones that returned real data in
    # the earlier HBAN client files; rows with no stored FF_ code keep the client's FFI_ code.
    7:  (ff('FF_BK_COM_EQ_TIER1_RATIO'), 'Replaced FFI_COM_EQ_TIER1_RATIO with FF_ code used in earlier client work', 'Verified code'),
    11: (ff('FF_BK_LEV_RATIO'), 'Replaced FFI_LEV_RATIO_ADVT (data for only 51 of 669 banks) with FF_ code used in earlier client work', 'Verified code - confirm definition'),
    12: (ff('FF_ASSETS_RISK_WGHT', True), 'Replaced FFI_RWA with FF_ code used in earlier client work (USD)', 'Verified code'),
    15: (ff('FF_LOAN_LOSS_PROV', True), 'Replaced FB_PROV_CR_LOSS_OTH (provision on OTHER assets - blank for 659 of 669 banks) with loan loss provision', 'Verified code'),
    19: (ff('FF_BK_EFF_RATIO'), 'Replaced FFI_EFF_RATIO with FF_ code used in earlier client work (whole-number %)', 'Verified code'),
    22: (ff('FF_ROTCE'), 'Replaced FFI_ROTCE with FF_ code used in earlier client work', 'Verified code'),
    23: (ff('FF_ROA'), 'Replaced FFI_ROA with FF_ code used in earlier client work', 'Verified code'),
    24: ('FB_INT_MGN(ANN,{Y})', 'Replaced FFI_NIM_NIS_DIFF (= NIM minus net interest SPREAD, median 0.21) with the NIM code used in earlier client work (no FF_ NIM code returned data there)', 'Verified code - confirm global coverage'),
    25: (ff('FF_INT_INC_NET', True), 'Replaced FFI_INT_INC_NET with FF_ code used in earlier client work (USD)', 'Verified code'),
    26: (ff('FF_NON_INT_INC', True), 'Replaced FFI_NON_INT_INC with FF_ code used in earlier client work (USD)', 'Verified code'),
    28: (ff('FF_BK_LOAN_TOT', True), 'Replaced FFI_LOAN_ADV_TOT with FF_ code used in earlier client work (USD)', 'Verified code'),
    29: (ff('FF_DEPS', True), 'Replaced FFI_DEPS_TOT with FF_ code used in earlier client work (USD)', 'Verified code'),
    30: (ff('FF_BK_LOAN_TOT', True) + '/' + ff('FF_DEPS', True), 'Same FF_ codes as the Loans and Deposits rows', 'Verified code'),
    31: (ff('FF_BK_LIQ_COVG_RATIO'), 'Replaced FFI_BK_LIQ_COVG_RATIO with FF_ code used in earlier client work', 'Verified code'),
    34: (ffi('FFI_LOAN_AMORT_CUST', True), 'No FF_ code stored - kept FFI_, added USD (it was pulling LOCAL currency: yen, rupiah, dong...)', 'Verified code'),
    35: (ffi('FFI_LOAN_AMORT_BK', True), 'No FF_ code stored - kept FFI_, added USD (it was pulling LOCAL currency)', 'Verified code'),
    36: (ff('FF_INVEST_TOT', True), 'Row had no code (blacked out). FF_INVEST_TOT = "Total Investments - Banks"', 'PLEASE CONFIRM'),
}
for _code, *_ in HIST.values():
    assert not ('FFI_' in _code and ('FF_' in _code or 'FB_' in _code)) and '@' not in _code, _code

for r, (code, why, status) in HIST.items():
    before = txt(H.cell(r, FIRST).value)
    for c in range(FIRST, LAST + 1):
        col = L(c)
        H.cell(r, c).value = fdsc(f'{col}2', code.replace('{Y}', f'"&{col}4&"'))
        H.cell(r, c).fill = YELLOW
    assert '{Y}' in code and 'Y}' not in H.cell(r, FIRST).value
    log.append(('CAMELS HISTORICAL ANNUAL', r, H.cell(r, 2).value, f'{L(FIRST)}{r}:{L(LAST)}{r}',
                (before or '(blank)').replace('_xll.', ''), txt(H.cell(r, FIRST).value), why, status))

# Row 38: Financial institutions deposits share = 1 - customer share; #VALUE! when row 37 is "-"
before = txt(H.cell(38, FIRST).value)
for c in range(FIRST, LAST + 1):
    col = L(c)
    H.cell(38, c).value = f'=IFERROR(1-{col}37,"-")'
    H.cell(38, c).fill = YELLOW
log.append(('CAMELS HISTORICAL ANNUAL', 38, H.cell(38, 2).value, f'C38:{L(LAST)}38', before,
            H.cell(38, FIRST).value, 'Showed #VALUE! (blacked out) whenever row 37 had no data; now shows "-"', 'Formula fix'))

# ------------------------------------------------------------ formatted sheet
FC0, FC1 = 4, 672               # D..YV  (Formatted col k  <->  Historical col k-1)
DIV_ROWS = {19: 21, 20: 22, 21: 23, 22: 24, 29: 31, 30: 32}
for fr, hr in DIV_ROWS.items():
    before = FM.cell(fr, FC0).value
    for c in range(FC0, FC1 + 1):
        FM.cell(fr, c).value = f"=IFERROR('CAMELS HISTORICAL ANNUAL'!{L(c - 1)}{hr}/100,\"-\")"
        FM.cell(fr, c).fill = YELLOW
    log.append(('Formatted Annual Camels', fr, FM.cell(fr, 3).value, f'D{fr}:YV{fr}', before, FM.cell(fr, FC0).value,
                'Dividing "-" by 100 gave #VALUE! (blacked out); now shows "-" when there is no data', 'Formula fix'))
for c in range(FC0, FC1 + 1):
    FM.cell(34, c).value = f"='CAMELS HISTORICAL ANNUAL'!{L(c - 1)}36"
    FM.cell(34, c).fill = YELLOW
log.append(('Formatted Annual Camels', 34, FM.cell(34, 3).value, 'D34:YV34', '(blank)', FM.cell(34, FC0).value,
            'Row was empty (blacked out); now linked to the new code on the historical sheet', 'Formula fix'))

# metric labels on the formatted sheet whose underlying code changed
F2H = {5: 7, 9: 11, 10: 12, 13: 15, 17: 19, 20: 22, 21: 23, 22: 24, 23: 25, 24: 26, 26: 28, 27: 29, 28: 30,
       29: 31, 32: 34, 33: 35, 34: 36, 36: 38}
LABEL_ROWS = set(F2H) | set(DIV_ROWS) | {34}

# ------------------------------------------------------- single-company sheets
def put(ws, r, value, why, status):
    before = txt(ws.cell(r, 3).value)
    ws.cell(r, 3).value = value
    ws.cell(r, 3).fill = YELLOW
    log.append((ws.title, r, ws.cell(r, 2).value, f'C{r}', (before or '(blank)').replace('_xll.', '') if isinstance(before, str) else before,
                value, why, status))


for ws, per in ((NTM, 'NTMA,'), (EST, 'ANN_ROLL,"&C4&"')):
    fe = lambda item: f"FE_ESTIMATE({item},MEAN,{per},NOW,,,'CURRENCY=USD')"
    put(ws, 7, fdsc('B1', fe('COMCAP_RATIO_TIER1')),
        'Was "estimates not available" - COMCAP_RATIO_TIER1 is the consensus CET1 ratio item', 'Verified code')
    put(ws, 16, fdsc('B1', fe('LOAN_NONPERF') + '/' + fe('LOAN_LOSS_RSRV')),
        'Was plain NPL (not a ratio) - now NPL / loan loss reserve', 'Verified code')
    put(ws, 17, fdsc('B1', fe('LOAN_NONPERF') + '/' + fe('LOAN_GROSS')),
        'Was plain NPL (not a ratio) - now NPL / gross loans', 'Verified code')
    put(ws, 31, 'estimates not available',
        'Old formula was (current assets - inventory) / current liabilities = QUICK ratio, not LCR; no LCR estimate item exists', 'Removed wrong formula')

put(LTM, 17, fdsc('B1', 'FF_LOAN_NONPERF(LTMSG,0,,,,USD)/FF_BK_LOAN_TOT(LTMSG,0,,,,USD)'),
    'Was plain NPL estimate (not a ratio) - now NPL / total loans, same style as row 16', 'Verified code')
put(LTM, 24, fdsc('B1', "FE_ESTIMATE(INT_INC_MARGIN,MEAN,LTMA,,NOW,,,'CURRENCY=USD')"),
    'Replaced FFI_NIM_NIS_DIFF (NIM minus spread) with NIM, LTM actual - same style as row 15', 'PLEASE CONFIRM')
put(LTM, 34, fdsc('B1', 'FFI_LOAN_AMORT_CUST(LTM_L,0,,,,USD)@FFI_LOAN_AMORT_CUST(LTM_SEMI_L,0,,,,USD)'),
    'Added USD (was local currency)', 'Verified code')
put(LTM, 35, fdsc('B1', 'FFI_LOAN_AMORT_BK(LTM_L,0,,,,USD)@FFI_LOAN_AMORT_BK(LTM_SEMI_L,0,,,,USD)'),
    'Added USD (was local currency)', 'Verified code')
put(LTM, 36, fdsc('B1', 'FF_INVEST_TOT(LTMSG,0,,,,USD)'),
    'Row had no code. FF_INVEST_TOT = "Total Investments - Banks"', 'PLEASE CONFIRM')
put(LTM, 38, '=IFERROR(1-C37,"-")', 'Row had no formula - 1 minus customer-deposit share, same as historical sheet', 'Formula fix')

# ------------------------------------------------------------ formatting
PCT_WHOLE = '#,##0.00"%"_);(#,##0.00"%")'   # value already x100 (14.44 -> 14.44%)
PCT = '0.00%_);(0.00%)'                     # fraction (0.1444 -> 14.44%)
DOLLAR = '$#,##0.00_);($#,##0.00)'
ROWFMT = {4: DOLLAR, 5: PCT_WHOLE, 6: DOLLAR, 7: PCT_WHOLE, 8: PCT_WHOLE, 9: PCT_WHOLE, 10: DOLLAR,
          12: DOLLAR, 13: DOLLAR, 14: PCT, 15: PCT, 17: PCT_WHOLE,
          19: PCT, 20: PCT, 21: PCT, 22: PCT, 23: DOLLAR, 24: DOLLAR,
          26: DOLLAR, 27: DOLLAR, 28: PCT, 29: PCT, 30: PCT,
          32: DOLLAR, 33: DOLLAR, 34: DOLLAR, 35: PCT, 36: PCT}
SECTIONS = (3, 11, 16, 18, 25, 31)
HDR_FILL = PatternFill('solid', fgColor='FFD9E7F5')
HDR_FONT = Font(name='Arial', size=9, bold=True, color='FF1F4E78')
SEC_FONT = Font(name='Arial', size=10, bold=True, color='FF1F4E78')
DATA_FONT = Font(name='Arial', size=10)
CENTER = Alignment(horizontal='center', vertical='center')

for r in range(2, 37):
    for c in range(1, FC1 + 1):
        cell = FM.cell(r, c)
        if r == 2:
            if c >= 3:
                cell.font = HDR_FONT
                cell.fill = HDR_FILL
                cell.alignment = Alignment(horizontal='center' if c > 3 else 'left', vertical='center', wrap_text=True)
            continue
        if r in SECTIONS:
            if c >= 2:
                cell.fill = HDR_FILL
            cell.font = SEC_FONT if c == 3 else Font(name='Arial', size=9, bold=True, color='FF1F4E78')
            if c == 2:
                cell.alignment = CENTER
            continue
        if c == 3:
            cell.font = Font(name='Arial', size=10)
            if r in LABEL_ROWS:
                cell.fill = YELLOW
            continue
        if c >= 2:
            cell.font = Font(name='Arial', size=10, bold=(c == 2))
            cell.alignment = CENTER
            if r in ROWFMT:
                cell.number_format = ROWFMT[r]
FM['B3'].value = 'AVERAGE'
FM['C2'].font = Font(name='Arial', size=11, bold=True, color='FF1F4E78')
FM['A1'].value = '$ in millions (USD). Yellow = cell filled in or code changed.'
FM['A1'].font = Font(name='Arial', italic=True, size=9, color='FF595959')
FM.row_dimensions[2].height = 40
for c in range(FC0, FC1 + 1):
    FM.column_dimensions[L(c)].width = 15
FM.column_dimensions['A'].width = 11.14
FM.column_dimensions['B'].width = 14
FM.column_dimensions['C'].width = 47
FM.freeze_panes = 'D3'
FM.sheet_view.topLeftCell = 'A1'
FM.sheet_view.selection[0].activeCell = 'D4'
FM.sheet_view.selection[0].sqref = 'D4'
# grey "-" like the Quarterly Template (client's black fill for 0 / errors is kept)
grey = FormulaRule(formula=['IFERROR(B4="-",FALSE)'], font=Font(color='FF9E9E9E'))
FM.conditional_formatting.add('B4:YV36', grey)
grey.priority = 10          # unique; client's black-out rules use 1-3 and 7-9

# historical sheet: labels + IDs frozen, even column widths so 669 banks scroll cleanly
H.freeze_panes = 'C5'
H.sheet_view.topLeftCell = 'A1'
for c in range(FIRST, LAST + 1):
    H.column_dimensions[L(c)].width = 14
for c in range(FIRST, LAST + 1):
    H.cell(3, c).alignment = Alignment(wrap_text=True, vertical='top')
H.row_dimensions[3].height = 45

# Red theme on the historical sheet (red #B00B1C, white, light red tint)
from openpyxl.styles import Border, Side
THEME_RED = 'FFB00B1C'
RED_FILL = PatternFill('solid', fgColor=THEME_RED)
TINT_FILL = PatternFill('solid', fgColor='FFF7E1E4')
RED_LINE = Border(bottom=Side(style='medium', color=THEME_RED))
H_SECTIONS = (5, 13, 18, 20, 27, 33)
for c in range(1, LAST + 1):
    for r in (2, 3):
        cell = H.cell(r, c)
        cell.fill = RED_FILL
        cell.font = Font(name='Arial', size=9, bold=True, color='FFFFFFFF')
        cell.alignment = Alignment(horizontal='left' if c < FIRST else 'center', vertical='center', wrap_text=(r == 3))
    y = H.cell(4, c)
    y.fill = TINT_FILL
    y.font = Font(name='Arial', size=9, bold=True, color=THEME_RED)
    y.alignment = Alignment(horizontal='center')
    y.border = RED_LINE
    for r in H_SECTIONS:
        s = H.cell(r, c)
        s.fill = TINT_FILL
        s.font = Font(name='Arial', size=10, bold=True, color=THEME_RED)
    for r in range(6, 39):
        if r in H_SECTIONS:
            continue
        d = H.cell(r, c)
        d.font = Font(name='Arial', size=10, bold=False, color='FF262626')
        if c >= FIRST:
            d.alignment = Alignment(horizontal='center')
H['A1'].font = Font(name='Arial', size=11, bold=True, color=THEME_RED)
H.column_dimensions['B'].width = 44

# ------------------------------------------------------------ change log sheet
CL = wb.create_sheet('Change Log')
CL.append(['Sheet', 'Row', 'Metric', 'Cells', 'Before (first cell)', 'After (first cell)', 'Why', 'Status'])
for row in log:
    CL.append(list(row))
for c in CL[1]:
    c.font = HDR_FONT
    c.fill = HDR_FILL
for r in CL.iter_rows(min_row=2):
    for c in r:
        c.font = Font(name='Arial', size=9)
        c.alignment = Alignment(wrap_text=True, vertical='top')
    if 'CONFIRM' in str(r[7].value):
        r[7].fill = YELLOW
for col, w in zip('ABCDEFGH', (26, 6, 34, 14, 60, 60, 60, 24)):
    CL.column_dimensions[col].width = w
CL.freeze_panes = 'A2'

wb.save(OUT)

# put back the file's Microsoft sensitivity label part (openpyxl drops it)
import zipfile, shutil
ORIG = zipfile.ZipFile(SRC)
tmp = OUT + '.tmp'
with zipfile.ZipFile(OUT) as zin, zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
    for item in zin.infolist():
        data = zin.read(item.filename)
        if item.filename == '[Content_Types].xml':
            data = data.replace(b'</Types>', b'<Override PartName="/docMetadata/LabelInfo.xml" '
                                b'ContentType="application/vnd.ms-office.classificationlabels+xml"/></Types>')
        if item.filename == '_rels/.rels':
            data = data.replace(b'</Relationships>', b'<Relationship Type="http://schemas.microsoft.com/office/2020/02/'
                                b'relationships/classificationlabels" Target="docMetadata/LabelInfo.xml" Id="rIdLbl1"/>'
                                b'</Relationships>')
        zout.writestr(item, data)
    zout.writestr('docMetadata/LabelInfo.xml', ORIG.read('docMetadata/LabelInfo.xml'))
shutil.move(tmp, OUT)
print('saved', OUT, 'changes logged:', len(log))
