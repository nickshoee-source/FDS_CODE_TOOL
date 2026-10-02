# Project history (for context)

## 1. HBAN peer quarterly template
- **Goal:** FactSet dashboard for Huntington (HBAN) vs peers FITB, ZION, RF, PNC, MTB, FHN, HBAN,
  USB, KEY, CFG, TFC. Started from `FactSet_MASTER_FILE_-HBAN_V1` (annual + quarterly codes) and
  `Key_Metric_Mapping_V2` (metric list by dashboard tab).
- **Final:** HBAN_Quarterly_Template_V10.xlsx (kept privately).
  - 12 category sheets (Valuation, Shareholder Returns, Profitability, Revenue & Fee Mix, Efficiency,
    NIM & Margin, Balance Sheet Evolution, Deposit Competition, Credit, Capital, Estimates, Franchise),
    each with stacked blocks (rows 4/23 metric headers, rows 5/24 periods, banks in rows 8–18/27–37,
    tickers in column D), 28 quarters 2020/1F–2026/4F per metric.
  - Matching "- Estimates" sheets 2027/1F–2029/4F using FE_TIMESERIES / FE_TIMESERIES_VALUATION spills
    (codes taken from the HBAN `All_Estimates_Formulas` template); metrics with no FE item use a
    date-gated FDS that fills itself in once the quarter is reported.
  - Formatting: $ (millions), whole-number %, multiples with x, P/E "NM" > 100x, `_xlfn.IFNA` → "-",
    grey "-" conditional format, unit note in A2, freeze panes E8; Data Check sheet listing suspect values.
- **Key decisions / lessons**
  - `{}` and `@` appeared because cells were written as `_xll.` array formulas → write plain `=FDS(`.
  - Efficiency (FF_EFF_RATIO) is a fraction → ×100; TSR: removed `/100`; NCO and deposit cost ×4.
  - FB_INT_MGN for NIM and FB_ROTCE for ROTCE (FF versions returned NA quarterly).
  - Stacked blocks must reference their own period row (a wrong reference produced a constant series).
  - Leave the user's FDSRC spill outputs alone (they look like hardcodes).
- **Open items flagged to the user:** Total Funding Cost and IB Deposit Cost codes returned the same
  values as Cost of Funds; "Earning Assets EOP" is actually an average (QTR_INTE); TSR window is
  labelled one quarter early (start = column's quarter, end = next column).

## 2. CAMELS global bank screen (global-bank client)
- **Original:** the client's original CAMELS workbook (refreshed; kept privately) — 669 global
  banks (SEDOL IDs), FFI_ codes, fiscal year 2025, USD; `Formatted Annual Camels` links to
  `CAMELS HISTORICAL ANNUAL`; NTM / Estimate / LTM single-company sheets.
- **Final:** CAMELS_V5_Changes_Highlighted.xlsx (kept privately).
  - Every metric on the historical sheet uses an FFI → FF → FB fallback chain (`@`).
  - Fixed: provision (`FB_PROV_CR_LOSS_OTH` was "other assets"), NIM (`FFI_NIM_NIS_DIFF` was
    NIM−spread), leverage (`_ADVT` advanced approach → `FFI_LEV_RATIO_RPT`), local-currency
    amortized-cost loans (+USD), empty "Financial Investments, Derivatives, P/L assets" row,
    `/100` and `1-x` on "-" text (#VALUE!), NTM/Estimate ratio rows that pulled plain NPL,
    an "LCR" that was really a quick ratio.
  - Red theme colours (red B00B1C / tint F7E1E4) on all sheets; freeze panes; Change Log sheet;
    yellow = exactly the cells changed vs the original.
- **Still to confirm with the user after refresh:** cross-family `@` chains populate (an earlier
  test of mixed FFI/FF returned nothing); `FF_BK_COM_EQ_TIER1_TOT`, `FF_DEPS_CUST`, `FF_INVEST_TOT`
  are unverified; the investments row is a sum of three FFI items (NA if any piece is missing).

## 3. Reference files supplied by the user
`reference/factset_code_files/`: FF standard template codes, FFI industry templates (Banks — full,
with regulatory capital / asset quality / supplemental sheets; Specialty Finance; Insurance), and the
Regulatory Bank Codes glossary (FB_ items for FR Y-9C holding companies and call-report banks).
All are merged into `reference/FactSet_Code_Catalog.csv`.
