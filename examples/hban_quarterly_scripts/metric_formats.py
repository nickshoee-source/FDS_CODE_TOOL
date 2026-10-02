"""
Single source of truth for how every metric is displayed.

Unit classification is evidence-based -- taken from the values FactSet actually
returned in the user's refreshed 8/31 workbook and the HBAN estimates template:
  * FactSet percentages come back as WHOLE numbers (ROE 10.7 = 10.7%).
  * Four metrics came back as FRACTIONS (0.605 = 60.5%) and are converted to
    whole numbers (x100) so every % in the workbook uses one convention:
      historical Efficiency Ratio (FF_EFF_RATIO), TSR (formula divided by 100),
      Buyback Yield and Capital Return Yield (cell math).
  * $ values are in MILLIONS (HBAN total assets 210,228 = $210B; market cap
    25,204 = $25B); per-share values are plain dollars; shares in millions.
"""

FMT = {
    'PCT':    '#,##0.00"%"_);(#,##0.00"%")',      # whole-number percent
    'DOLLAR': '$#,##0.00_);($#,##0.00)',            # $M or $/share
    'DPS':    '$#,##0.00#_);($#,##0.00#)',          # dividends: 3rd decimal only if needed ($0.155)
    'MULT':   '[<0]"NM";#,##0.00"x"',              # valuation multiple
    'PE':     '[>100]"NM";[<0]"NM";#,##0.00"x"',   # P/E: NM when >100x or negative
    'NUM':    '#,##0.00_);(#,##0.00)',              # share counts (millions)
    'COUNT':  '#,##0_);(#,##0)',                    # whole counts
}

KIND = {
    # Valuation
    ('Valuation', 'P/E'): 'PE', ('Valuation', 'P/B'): 'MULT', ('Valuation', 'P/TBV'): 'MULT',
    ('Valuation', 'Market Cap'): 'DOLLAR', ('Valuation', 'Enterprise Value'): 'DOLLAR',
    ('Valuation', 'Share Price'): 'DOLLAR',
    # Shareholder Returns
    ('Shareholder Returns', 'TSR'): 'PCT', ('Shareholder Returns', 'Dividend Yield'): 'PCT',
    ('Shareholder Returns', 'Dividend Payout Ratio'): 'PCT',
    ('Shareholder Returns', 'Dividends Paid'): 'DOLLAR',
    ('Shareholder Returns', 'Share Repurchases'): 'DOLLAR',
    ('Shareholder Returns', 'Market Cap'): 'DOLLAR',
    ('Shareholder Returns', 'Buyback Yield'): 'PCT',
    ('Shareholder Returns', 'Dividend Per Share'): 'DPS',
    ('Shareholder Returns', 'Common Shares Outstanding'): 'NUM',
    ('Shareholder Returns', 'Stock Repurchase CF'): 'DOLLAR',
    ('Shareholder Returns', 'Market Value (Company)'): 'DOLLAR',
    ('Shareholder Returns', 'Capital Return Yield'): 'PCT',
    # Profitability
    ('Profitability', 'ROTCE'): 'PCT', ('Profitability', 'ROE'): 'PCT', ('Profitability', 'ROA'): 'PCT',
    ('Profitability', 'EPS'): 'DOLLAR', ('Profitability', 'Net Income'): 'DOLLAR',
    ('Profitability', 'PPNR'): 'DOLLAR', ('Profitability', 'Effective Tax Rate'): 'PCT',
    # Revenue & Fee Mix
    ('Revenue & Fee Mix', 'NII'): 'DOLLAR', ('Revenue & Fee Mix', 'Non-Interest Income'): 'DOLLAR',
    ('Revenue & Fee Mix', 'Total Revenue'): 'DOLLAR',
    # Efficiency
    ('Efficiency', 'Efficiency Ratio'): 'PCT', ('Efficiency', 'NIE'): 'DOLLAR',
    ('Efficiency', 'Personnel Expense'): 'DOLLAR',
    # NIM & Margin
    ('NIM & Margin', 'NIM'): 'PCT', ('NIM & Margin', 'EA Yield'): 'PCT',
    ('NIM & Margin', 'Interest Expense Total'): 'DOLLAR',
    ('NIM & Margin', 'Average Interest Bearing Liabilities'): 'DOLLAR',
    ('NIM & Margin', 'Cost of Funds'): 'PCT', ('NIM & Margin', 'Deposit Cost'): 'PCT',
    ('NIM & Margin', 'Interest Bearing Deposit Cost'): 'PCT',
    ('NIM & Margin', 'Total Funding Cost'): 'PCT', ('NIM & Margin', 'Net Interest Spread'): 'PCT',
    ('NIM & Margin', 'Taxable Equivalent NIM'): 'PCT',
    # Balance Sheet Evolution
    ('Balance Sheet Evolution', 'Total Assets EOP'): 'DOLLAR',
    ('Balance Sheet Evolution', 'Intangible Assets'): 'DOLLAR',
    ('Balance Sheet Evolution', 'Total Liabilities'): 'DOLLAR',
    ('Balance Sheet Evolution', 'Average Assets'): 'DOLLAR',
    ('Balance Sheet Evolution', 'Earning Assets EOP'): 'DOLLAR',
    ('Balance Sheet Evolution', 'Total Loans EOP'): 'DOLLAR',
    ('Balance Sheet Evolution', 'Total Deposits EOP'): 'DOLLAR',
    ('Balance Sheet Evolution', 'Tangible Common Equity'): 'DOLLAR',
    ('Balance Sheet Evolution', 'Average Tangible Common Equity'): 'DOLLAR',
    ('Balance Sheet Evolution', 'TBV'): 'DOLLAR', ('Balance Sheet Evolution', 'TBVPS'): 'DOLLAR',
    # Deposit Competition
    ('Deposit Competition', 'NIBD Balance'): 'DOLLAR',
    ('Deposit Competition', 'Interest Bearing Deposits'): 'DOLLAR',
    ('Deposit Competition', 'Total Deposits'): 'DOLLAR',
    # Credit
    ('Credit', 'NCO Ratio'): 'PCT', ('Credit', 'NPL Ratio'): 'PCT', ('Credit', 'NPA Ratio'): 'PCT',
    # Capital
    ('Capital', 'CET1 Ratio'): 'PCT', ('Capital', 'Risk Weighted Assets'): 'DOLLAR',
    # Estimates (consensus)
    ('Estimates', 'Consensus EPS'): 'DOLLAR', ('Estimates', 'Consensus Revenue'): 'DOLLAR',
    ('Estimates', 'Consensus NII'): 'DOLLAR', ('Estimates', 'Consensus Fee Income'): 'DOLLAR',
    ('Estimates', 'Consensus ROTCE'): 'PCT',
    ('Estimates', 'Consensus CET1'): 'DOLLAR',      # COM_EQUITY_TIER1 = CET1 capital, $M
    # Franchise
    ('Franchise', 'Branch Count'): 'COUNT',
}

# per-share $ (not millions) -- used only for the units note wording
PER_SHARE = {('Valuation', 'Share Price'), ('Shareholder Returns', 'Dividend Per Share'),
             ('Profitability', 'EPS'), ('Balance Sheet Evolution', 'TBVPS'),
             ('Estimates', 'Consensus EPS')}

EST_NAME_BASE = {'Balance Sheet Evol': 'Balance Sheet Evolution'}


def base_sheet(sheet):
    b = sheet.replace(' - Estimates', '')
    return EST_NAME_BASE.get(b, b)
