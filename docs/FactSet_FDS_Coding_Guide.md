# FactSet `=FDS` Coding Guide — Knowledge Base for Claude

> **How to use this file**
> - **Claude Code:** put this file (rename to `CLAUDE.md` if you like) and `FactSet_Code_Catalog.csv` in the project folder. Tell Claude: *"Read FactSet_FDS_Coding_Guide.md before writing any FactSet formula, and look codes up in FactSet_Code_Catalog.csv."*
> - **New Claude conversation:** attach both files and say the same thing.
> - The catalog has ~9,700 codes (FFI, FF, FB, FE, P, FG, FREF, FMA) with descriptions, source file, and — for ~210 of them — proof that they returned real data in a refreshed client workbook.
>
> Built from: HBAN quarterly template project (V1 master → V10), the HBAN estimates template (All_Estimates_Formulas), the CAMELS global-bank screen (V1 → V5), and FactSet's code reference files (FF SF template, FFI Banks / Specialty Finance / Insurance templates, Regulatory Bank Codes "FFB").

---

## 0. Rules of engagement (how this user wants work done)

1. **Formulas must be clean.** Only `=FDS(...)`, `=FDSC(...)`, `=FDSR(...)`, `=FDSRC(...)` plus ordinary Excel. **Nothing extra** — no `_xll.` prefix, no `{ }` array braces, no `@` in the Excel formula bar. (When writing with Python/openpyxl, write the plain string `=FDS(...)`, never an `ArrayFormula` with `_xll.`; that is what produced `{}` and `@` in downloads.)
2. **Triple-check.** Verify every code against this guide / the catalog, and check units with real returned values when available.
3. **Only change what you are 100% sure of.** Anything less → put it on a **"please confirm"** list for the user, with the worksheet name, row/metric, the code used, and why it is uncertain.
4. **Highlight every changed cell in yellow** (`FFFF00`). Yellow must mean *only* "changed from the original" — remove yellow from cells you did not change, and don't use yellow for anything else.
5. **Summarize what you did NOT change** (and why), always with worksheet names.
6. **Sample first on big rollouts.** For a sheet with hundreds of companies, build the first 5 companies, let the user refresh in FactSet, then roll out to all.
7. Keep a **Change Log** sheet (sheet, row, metric, cells, before, after, why, status) when editing a client file.
8. Never overwrite the user's own FactSet spill outputs (cells holding values right of an `FDSRC` anchor) — they look like hardcodes but are live spill results.

---

## 1. The four Excel functions

| Function | Returns | When no data | Use for |
|---|---|---|---|
| `=FDS(id,"CODE(...)")` | one value | `#N/A` (true Excel error) | standard single-cell pulls |
| `=FDSC("-",id,"CODE(...)")` | one value | the default text you give (`"-"`) | when you want a clean "-" instead of #N/A |
| `=FDSR(id,"CODE(...)")` | a row of values (spills right) | #N/A | multi-period pulls |
| `=FDSRC("-",id,"CODE(...)")` | a row (spills right) | default text | **FE_TIMESERIES** estimate rows |

- `id` = a cell reference holding the identifier: US ticker (`HBAN`, `FITB`), ticker-exchange (`JPM-US`, `SAN-ES`, `UBSG-CH`, `NAB-AU`), SEDOL (`762084`, `BRJDJ6`), or `"DUMMY"` for non-company items like `P_DATE`.
- The code string is FactSet **FQL**. Build periods into it with `"&cell&"`:
  `=FDS($D8,"FF_ASSETS(QTR,"&E$5&")")` where `E$5` holds `2025/1F`.
- Text inside the FQL string that itself needs quotes uses doubled quotes: `FREF_MARKET_VALUE_COMPANY("&E$5&",,,,,0,,""LEGACY"")`.
- For **spills** (`FDSR`/`FDSRC`): write the formula only in the **first cell** and leave the cells to the right empty for the spill.
- Company name: `=FDS(C2,"FG_COMPANY_NAME")`.

### Errors you will see
| Result | Meaning | Fix |
|---|---|---|
| `#N/A` / `"-"` | no data for that company/period | normal gap, or wrong periodicity |
| `ERROR: Company or industry not found (FG-23)` | the **FFI** industry database does not cover this company (≈20% of global banks) | add an FF_ / FB_ fallback (see §5) |
| `ERROR: Item not available (FGF-L21)` | code not valid for this company's template (e.g. bank item on an insurer) | use a different code |
| `#VALUE!` after cell math | dividing/subtracting the text `"-"` | wrap in `IFERROR(...,"-")` |
| `#NAME?` | Excel function missing its prefix (see §7) or add-in not loaded | `_xlfn.IFNA`, open in Excel with FactSet |

---

## 2. Code families

| Prefix | Database | Coverage | Typical syntax |
|---|---|---|---|
| `FF_` | FactSet Fundamentals (standardized) | global, all companies | `FF_X(QTR,"2025/1F")`, `FF_X(ANN_R,2025,,,,USD)` |
| `FFI_` | FactSet Fundamentals **Industry** (bank / specialty-finance / insurance detail, IFRS line items, regulatory capital) | global banks, but **not all** (FG-23) | `FFI_X(ANN_L,2025,,,,USD)` |
| `FB_` | FactSet **Regulatory** bank data (FR Y-9C, call reports) — the user's "FFB" file | **US** bank holding cos. & banks only | `FB_X(QTR,"2025/1F")`, `FB_X(ANN,2025,,,RF)`, `FB_X(ANN,2025,,,RF,USD)` |
| `FE_` | FactSet **Estimates** (consensus) — functions wrapping an *item* | covered companies | `FE_ESTIMATE(ITEM,MEAN,...)`, `FE_TIMESERIES(ITEM,...)` |
| `P_` | Prices | all | `P_PRICE(date)`, `P_PRICE_RETURNS(2,start,end)` |
| `FG_` | Global reference / prices | all | `FG_COMPANY_NAME`, `FG_PRICE(0)` |
| `FREF_` | Reference / market value | all | `FREF_MARKET_VALUE_COMPANY(date,,,,,0,,""LEGACY"")` |
| `FMA_` | Market aggregates | portfolios | `FMA_PE(0,0,,,""PORTAGG"",""MEANR"")` |

**Choosing a family for banks**
- **US regional bank, quarterly history:** `FF_` / `FF_BK_` first; `FB_` for regulatory ratios (NIM, ROTCE, branches).
- **Global banks, annual:** `FFI_` first (richest bank detail, IFRS lines like "loans at amortized cost"), then `FF_`, then `FB_` (US only).
- **Forecasts:** `FE_` items.

---

## 3. Arguments: periodicity, period, currency

### 3.1 Periodicity (first argument)
| Code | Meaning | Seen with |
|---|---|---|
| `QTR` | quarterly as originally reported | FF_, FB_ |
| `QTR_R` | quarterly, restated | FF_ |
| `ANN` | annual as reported | FF_, FB_ |
| `ANN_R` | annual restated | FF_ |
| `ANN_L` | annual, **latest** filing (FFI convention) | FFI_ |
| `LTM` | last twelve months | FF_PAY_OUT_RATIO(LTM,…) |
| `LTM_L`, `LTM_SEMI_L` | FFI LTM / LTM built from semi-annual reporters | FFI_ |
| `LTMSG` | LTM used by the client for FF_ balance items | FF_ |
| `QTR_INTE`, `QTR_TCE`, `QTR_TOT`, `QTR_NINTB_LIABS`, `QTR_INTB_DEPS` | **item-specific** variants (see §6) | FF_BK_ items |

### 3.2 Period (second argument)
- `0` = latest, `-1` = one period back, `+1` = next (estimates).
- Fiscal quarter: `2025/1F` … `2025/4F` (string built from a header cell).
- Fiscal year: `2025`.
- Prices: a date or fiscal period string; `NOW`.

### 3.3 Currency & options
- **FF_ / FFI_ currency = 6th argument:** `FFI_RWA(ANN_L,2025,,,,USD)`. **Without it, FFI returns local currency** (yen, rupiah…) — never compare banks without `USD`. Ratios don't need it.
- **FB_:** `FB_X(ANN,2025,,,RF)`; monetary `FB_X(ANN,2025,,,RF,USD)` (the `RF` form was verified with `FB_COM_EQ_TIER1_RATIO(ANN,2024,,,RF)`).
- **FF_ with options string (used & working for HBAN):** `FF_ROA(QTR_R,"&E$5&",,,,,,""D=1"")`, also ROE and ROTCE.
- **FE_:** currency in the options string: `'CURRENCY=USD'` or `'CURRENCY=RPT'`.

---

## 4. Estimates (FE_) syntax

### 4.1 Single values
```
FE_ESTIMATE(ITEM, MEAN, QTR,      "2026/1F", 0,,,'')          quarter
FE_ESTIMATE(ITEM, MEAN, QTR_ROLL, +1,        0,,,'')          next quarter (rolling)
FE_ESTIMATE(ITEM, MEAN, ANN_ROLL, 2026, NOW,,,'CURRENCY=USD') fiscal year
FE_ESTIMATE(ITEM, MEAN, NTMA,     ,     NOW,,,'CURRENCY=USD') next twelve months
FE_ESTIMATE(ITEM, MEAN, LTMA,     ,     NOW,,,'CURRENCY=USD') last twelve months ACTUAL
FE_VALUATION(PE, MEAN, ANN_ROLL, +1, 0,,,'')                   valuation ratios (P/E, ROA %)
FE_GROWTH(LOAN_NET, MEAN, ANN_ROLL, +1, 0CY, 0,,,'')
FE_ESTIMATE_DATE(RPT_DATE,,QTR_ROLL,+1,'MM/DD/YYYY',0,,,'')    next report date (use FDSC)
```

### 4.2 Time series (spill across quarters) — the HBAN standard
```
=FDSRC("-",$D8,"FE_TIMESERIES(ITEM,MEAN,"&E$5&","&P$5&",FQ,'BKRACTMED=1,WIN=0,CURRENCY=RPT,UNITS=AUTO,DATE=NOW')")
=FDSRC("-",$D8,"FE_TIMESERIES_VALUATION(ITEM,MEAN,"&E$5&","&P$5&",FQ,'BKRACTMED=1,WIN=0,UNITS=AUTO,DATE=NOW,CALC=LTMA')")
=FDSRC("-",$B$12,"FE_TIMESERIES_GUIDANCE(COST_INCOME,LOW,2026/4F,2029/2F,FQ,'BKRACTMED=1,WIN=0,CURRENCY=RPT,UNITS=AUTO,DATE=NOW')")
```
- Start/end periods from header cells (`E$5` first quarter, `P$5` 12th quarter). `FQ` = fiscal quarters, `FY` = years.
- Use `FE_TIMESERIES_VALUATION` (with `CALC=LTMA`) for ratio items: PE, PBPS, P_TBPS, DIV_YLD, ROA, ROE, NPL, NETCHARGE_LOANNET, NETINC_ROA_PCT, PEG…; `FE_TIMESERIES` for everything else.

### 4.3 FE item dictionary (all verified in the HBAN estimates template unless marked "-")
| Item | Meaning | Item | Meaning |
|---|---|---|---|
| EPS / EAG / CUSTOM_EPS / EPS_NONGAAP | EPS / GAAP / non-GAAP | NETINTERESTINC, INT_INC_NET | Net interest income |
| SALES | Total revenue | NON_INT_INC, INCFEESCOM | Non-interest / fee income |
| INT_INC_MARGIN | **NIM (%)** | INT_INCOME | Interest income |
| COST_INCOME | **Efficiency ratio (%)** | OPEREXPEN | Operating expense (NIE) |
| SAL_BENEFITS | Compensation | COMP_RATIO | Comp ratio (%) |
| EBIT_ADJ | **PPNR** (pre-tax pre-provision) | PTP / PTPA / PTPBG | Pretax income |
| NETPROFIT / NETBG / BFNG / NET_INC | Net income variants | TAX_EXPENSE | Tax |
| ROTE | RoTE (%) — used for "Consensus ROTCE" | ROA, ROE (valuation) | % |
| COM_EQUITY_TIER1 | CET1 **capital** ($) | COMCAP_RATIO_TIER1 | **CET1 ratio (%)** |
| CAP_RATIO_TIER1 / CAP_RATIO_TOT | Tier 1 / total capital ratio | LEV_RATIO | Leverage ratio |
| ASSETS_RISK_WGHT | RWA | TCE_TA | TCE/TA (%) |
| TOTASSET | Total assets | AVG_EARN_ASSETS | Avg earning assets |
| LOAN_GROSS / LOAN_NET / LOAN_NET_AVG | Loans | LOAN_LOSS_RSRV | Allowance |
| DEPS / DEPS_AVG | Deposits EOP / avg | AVG_NINTB_DEPS / AVG_INTB_DEPS | Avg NIB / IB deposits (often "-") |
| LOAN_NONPERF, ASSETS_NONPERF | NPLs / NPAs | NPL_COV_RTO | NPL coverage (%) |
| LOANOREO_NONPERF | NPAs / loans+OREO (%) | NETCHARGE_LOANNET | NCO / avg loans (%), **already annualized** |
| NET_CHARGE_OFFS, PROVLOANS, LOAN_PROV, RSRV_BUILD | credit $ | LOANLOSSPROV_PCT, LOANLOSSRSV_PCT | % |
| NETDIV | DPS | SHS_REPURCH | Buybacks |
| BPS_TANG, BVPS / BPS | TBVPS, BVPS | PE, PBPS, P_TBPS, DIV_YLD, PEG | valuation |
| INTANG, TOTGW, PFD_STK, SHEQUITY | balance sheet | MV | market value (returned NA) |
| COMLOAD_RATIO_T1, LEV_RATIO_TIER1, SERVICE_CHRG, CARD_INCOME, SGA, CFF, EBITR, CURRENTASSETS | returned "-" for HBAN | NETINC_ROA_PCT | ROA (%) |
- Fractions not %: `EBIT_ROA`, `EBIT_RET_SHEQUITY`, `NETINC_SHEQUITY` (0.018 = 1.8%).

---

## 5. Fallback chains with `@` (inside the FQL string)

`"X@Y@Z"` = use X; if X is NA use Y; then Z. Arithmetic (`+ - * /`, parentheses) also works inside the string.

- **Proven:** same-family chains, e.g. the client's LTM sheet
  `FFI_COM_EQ_TIER1_RATIO(LTM_L,0)@FFI_COM_EQ_TIER1_RATIO(LTM_SEMI_L,0)` returned data.
- **Unproven / caution:** **cross-family** chains (`FFI_…@FF_…@FB_…`). In one test the user reported mixed FFI+FF cells returned nothing. The current CAMELS V5 file uses cross-family chains on the user's request — **always test on 5 companies before rolling out**, and if a row goes blank, remove one code at a time (an invalid code anywhere in the chain can blank the whole cell).
- **Ratios:** fall back item-by-item: `(LOANS_FFI@LOANS_FF@LOANS_FB)/(DEPS_FFI@DEPS_FF@DEPS_FB)`.
- To convert a whole-number % to a fraction: `"(A@B@C)/100"` (parenthesize the chain).
- `@` does **not** trigger on a returned `0`, and probably not on the FG-23 "company not found" error.

---

## 6. Verified bank code library (returned real data in refreshed client files)

### 6.1 US bank — quarterly history (HBAN template, final V10 codes)
Pattern: `=_xlfn.IFNA(FDS($D8,"CODE(PERIODICITY,"&E$5&")"),"-")`

| Metric | Code (periodicity) | Units / notes |
|---|---|---|
| P/E, P/B, P/TBV | `FF_PE(QTR)`, `FF_PBK(QTR)`, `FF_PBK_TANG(QTR)` | x |
| Market cap | `FF_MKT_VAL(QTR)` | $M |
| Enterprise value | `FF_ENTRPR_VAL(QTR)` | $M |
| Share price | `P_PRICE("&E$5&")` | $ |
| TSR | `P_PRICE_RETURNS(2,"&E$5&","&F$5&")` | whole % (do **not** /100); start = this quarter, end = next quarter's header |
| Dividend yield | `FF_DIV_YLD(QTR)` | % |
| Payout ratio | `FF_PAY_OUT_RATIO(QTR)` (or LTM) | % |
| Dividends paid | `FF_DIV_CF(QTR)` | $M |
| Share repurchases | `FF_SHS_REPURCH(TOTAL_VAL_REPURCH,QTR,"&E$5&")` | $M |
| DPS | `FF_DPS(QTR)` | $ |
| Shares outstanding | `FF_COM_SHS_OUT(QTR_R)` | millions |
| Stock repurchase CF | `FF_STK_PURCH_CF(QTR_R)` | $M |
| Company market value | `FREF_MARKET_VALUE_COMPANY("&E$5&",,,,,0,,""LEGACY"")` | $M |
| ROTCE | `FB_ROTCE(QTR)` (`FF_ROTCE` quarterly returned NA) | % |
| ROE / ROA | `FF_ROE(QTR,…,,,,,,""D=1"")` / `FF_ROA(QTR_R,…,,,,,,""D=1"")` | % |
| EPS / Net income / PPNR | `FF_EPS(QTR_R)` / `FF_NET_INC(QTR)` / `FF_PPNR(QTR)` | |
| Effective tax rate | `FF_TAX_RATE(QTR_R)` | % |
| NII / Non-interest income | `FF_INT_INC_NET(QTR_R)` / `FF_NON_INT_INC(QTR)` | $M; revenue = sum |
| Efficiency ratio | `FF_EFF_RATIO(QTR_R)` **×100** (returns fraction) — or `FF_BK_EFF_RATIO` (whole %) | % |
| NIE / Personnel | `FF_NON_INT_EXP(QTR)` / `FF_LABOR_EXP(QTR)` | $M |
| NIM | `FB_INT_MGN(QTR)` | % |
| Earning-asset yield | `FF_YLD_INT_EARN_ASSETS(QTR_R)` | % |
| Total interest expense | `FF_INT_EXP_TOT(QTR)` | $M |
| Avg interest-bearing liabilities | `FF_BK_AVG_LIABS_INTB(QTR_TOT)` | $M |
| Cost of funds | cell math `(IntExp / AvgIBL)*100*4` | % annualized |
| Deposit cost | `FF_COST_DEPS(QTR)` **×4** (quarterly rate) | % |
| IB liability cost | `FF_BK_INT_COST_INTB_AVG(QTR)` | % (duplicated cost of funds in test) |
| Funding cost | `FF_BK_COST(QTR_NINTB_LIABS)` | % (duplicated cost of funds in test) |
| Net interest spread / TE NIM | `FF_BK_INT_SPREAD(QTR)` / `FF_BK_INT_MGN_TAX_EQV(QTR)` | % |
| Total assets / liabilities / intangibles | `FF_ASSETS(QTR)` / `FF_LIABS(QTR_R)` / `FF_INTANG(QTR_R)` | $M |
| Average assets | `FF_BK_AVG_ASSETS(QTR)` | $M |
| Avg earning assets | `FF_BK_AVG_ASSETS(QTR_INTE)` | $M (it is an **average**, not EOP) |
| Total loans | `FF_BK_LOAN_TOT(QTR)` | $M |
| Deposits | `FF_DEPS(QTR)`, `FF_DEPS_INTB(QTR)`, `FF_DEPS_NINTB(QTR)` | $M |
| TCE / avg TCE | `FF_COM_EQ_TANG(QTR)` / `FF_BK_SUPPL_AVG(QTR_TCE)` | $M |
| TBVPS | `FF_BPS_TANG(QTR)` | $ |
| NCO ratio | `FF_CHARGE_OFFS_LOANS_PCT(QTR_R)` **×4** | % annualized |
| NPL ratio / NPA ratio | `FF_NONPERF_LOAN_PCT(QTR_R)` / `FF_NPA_ASSETS_PCT(QTR)` | % |
| CET1 ratio | `FF_BK_COM_EQ_TIER1_RATIO(QTR)` | % |
| RWA | `FF_ASSETS_RISK_WGHT(QTR)` | $M |
| Branch count | `FB_DOM_OFFCE_NUM(QTR)` | count |
| Consensus EPS (history) | `FE_ESTIMATE(EPS_NONGAAP,MEAN,QTR,"&E$5&",0,,,'')` | |
| Consensus revenue / fees / ROTCE | `FE_ESTIMATE(SALES|INCFEESCOM|ROTE,MEAN,QTR_ROLL,"&E$5&",0,,,'')` | |
| Consensus NII / CET1 $ | `FE_ESTIMATE(INT_INC_NET,MEAN,QTR,…)` / `FE_ESTIMATE(COM_EQUITY_TIER1,MEAN,QTR,…,,,,'')` | |

Annual (HBAN master, `ANN`/`ANN_R`, period 0): `FF_BK_EFF_RATIO`, `FF_BK_LEV_RATIO(ANN,0)`, `FF_BK_LIQ_COVG_RATIO(ANN_R,0)`, `FF_LOAN_LOSS_PROV(ANN_R,0)`, `FF_LOAN_GROSS`, `FF_LOAN_NONPERF`, `FF_ROTCE(ANN_R)`, `FF_INT_EXP_NET`, `FF_NET_INCOME(ANN)`, `FFI_CHRG_OFF_NET(ANN_L,2024)` (NCO $, matched JPM 2024 = $8.6B), `FB_COM_EQ_TIER1_RATIO(ANN,2024,,,RF)`.

### 6.2 Global banks — CAMELS annual (FFI first, then FF, then FB)
`=FDSC("-",C2,"CHAIN")` with `C2` = ID, `C4` = fiscal year; `{Y}` below = `"&C4&"`.

| Metric | FFI (primary) | FF fallback | FB fallback (US only) |
|---|---|---|---|
| CET1 capital | `FFI_COM_EQ_TIER1_TOT(ANN_L,{Y},,,,USD)` ✔ | `FF_BK_COM_EQ_TIER1_TOT` (unconfirmed) | `FB_COM_EQ_TIER1(ANN,{Y},,,RF,USD)` |
| CET1 ratio | `FFI_COM_EQ_TIER1_RATIO` ✔ | `FF_BK_COM_EQ_TIER1_RATIO` ✔ | `FB_COM_EQ_TIER1_RATIO` ✔ |
| Tier 1 capital | `FFI_TIER1_CAP …USD` ✔ | `FF_TIER1_CAP` | `FB_TIER1_CAP` |
| Tier 1 ratio | `FFI_CAP_RATIO_TIER1` ✔ | `FF_CAP_RATIO_TIER1` | `FB_CAP_RATIO_TIER1` |
| Total capital ratio | `FFI_CAP_RATIO_TOT` ✔ | `FF_CAP_RATIO_TOT` | `FB_CAP_RATIO_TOT` |
| Leverage ratio | `FFI_LEV_RATIO_RPT` (reported) — *not* `_ADVT` (advanced approach, 51/669 banks) | `FF_BK_LEV_RATIO` ✔ | `FB_LEV_RATIO` |
| RWA | `FFI_RWA …USD` ✔ | `FF_ASSETS_RISK_WGHT` ✔ | `FB_ASSETS_RISK_WGHT` |
| NPLs | `FFI_NPL_LOAN_ADV` | `FF_LOAN_NONPERF` ✔ | `FB_NONPERF_LOAN` |
| Provision | `FFI_LOAN_LOSS_PROV` | `FF_LOAN_LOSS_PROV` ✔ | `FB_LOAN_LOSS_PROV` |
| NPL / allowance | `FFI_NPL_LOAN_LOSS_RSRV_RATIO` (÷100 for fraction) | `FF_NONPERF_LOAN_LOSS_RSRV` | — |
| NPL / loans | `FFI_NPL_LOAN_RATIO` (user-supplied, e.g. `FFI_NPL_LOAN_RATIO(ANN_L,0)`) | `FF_NONPERF_LOAN_PCT` ✔ | `FB_NONPERF_LOAN_PCT` |
| Efficiency | `FFI_EFF_RATIO` ✔ (whole %) | `FF_BK_EFF_RATIO` ✔ | `FB_EFF_RATIO` |
| ROTE | `FFI_ROTE` ✔ | — | `FB_ROTE` |
| ROTCE | `FFI_ROTCE` ✔ | `FF_ROTCE` ✔ | `FB_ROTCE` ✔ |
| ROA | `FFI_ROA` ✔ | `FF_ROA` ✔ | `FB_ROA` |
| NIM | `FFI_AVG_BAL_INT_RATE_NET_MGN` — **not** `FFI_NIM_NIS_DIFF` (= NIM minus spread) | `FF_INT_MGN` | `FB_INT_MGN` ✔ |
| NII / Non-II | `FFI_INT_INC_NET` ✔ / `FFI_NON_INT_INC` ✔ | `FF_INT_INC_NET` ✔ / `FF_NON_INT_INC` ✔ | `FB_INT_INC_NET` / `FB_NON_INT_INC` |
| Loans | `FFI_LOAN_ADV_TOT` ✔ | `FF_BK_LOAN_TOT` ✔ | `FB_TOT_HFI_HFS_UNEARN_INC` |
| Deposits | `FFI_DEPS_TOT` ✔ | `FF_DEPS` ✔ | `FB_DEPS` |
| LCR | `FFI_BK_LIQ_COVG_RATIO` ✔ | `FF_BK_LIQ_COVG_RATIO` ✔ | — |
| NSFR | `FFI_BK_NSFR` ✔ | — | — |
| Loans to customers @ amortized cost | `FFI_LOAN_AMORT_CUST …USD` ✔ | `FF_LOAN_NET` (proxy) | `FB_TOT_HFI_HFS_UNEARN_INC` (proxy) |
| Loans to FIs @ amortized cost | `FFI_LOAN_AMORT_BK …USD` ✔ | — | — |
| Investments + derivatives + FVTPL | `FFI_SECS_INVEST + FFI_TRADE_ACCT + FFI_DERIV_HEDGE` | `FF_INVEST_TOT` (unconfirmed) | — |
| Customer deposits share | `FFI_DEPS / FFI_DEPS_TOT` ✔ | `FF_DEPS_CUST` (unconfirmed) / `FF_DEPS` | `FB_DEPS` |

✔ = returned real numbers in a refreshed client file.

### 6.3 Global banks — LTM and estimates (CAMELS NTM/Estimate/LTM sheets)
- LTM (FFI): `FFI_X(LTM_L,0)@FFI_X(LTM_SEMI_L,0)`; FF balance items: `FF_LOAN_NONPERF(LTMSG,0,,,,USD)`.
- LTM actuals from estimates DB: `FE_ESTIMATE(LOAN_PROV,MEAN,LTMA,,NOW,,,'CURRENCY=USD')`.
- NTM: `FE_ESTIMATE(ITEM,MEAN,NTMA,,NOW,,,'CURRENCY=USD')`; fiscal year: `FE_ESTIMATE(ITEM,MEAN,ANN_ROLL,"&C4&",NOW,,,'CURRENCY=USD')`.
- Ratios from estimates: NPL/allowance = `LOAN_NONPERF/LOAN_LOSS_RSRV`; NPL/loans = `LOAN_NONPERF/LOAN_GROSS`; CET1 ratio = `COMCAP_RATIO_TIER1`.
- There is **no LCR / NSFR / ROTCE estimate item** → write "estimates not available".

---

## 7. Units & conventions (checked against real returned values)

- **Percentages come back as whole numbers** (ROE 10.7 = 10.7%) for FF_, FFI_, FB_ and most FE_ items. Keep one convention per workbook (HBAN workbook: whole numbers with a `"%"` format).
- **Exceptions returning fractions:** `FF_EFF_RATIO` (0.605), FE `EBIT_ROA`, `EBIT_RET_SHEQUITY`, `NETINC_SHEQUITY`; any `X/Y` you compute inside FQL (e.g. NPL/loans, loans/deposits).
- **Quarterly rates that need ×4:** `FF_CHARGE_OFFS_LOANS_PCT`, `FF_COST_DEPS`, cost of funds built from quarterly interest expense. (FE `NETCHARGE_LOANNET` is already annualized.)
- **Money:** $ millions; per-share in dollars; shares in millions. FFI/FF without `USD` = local currency.
- **Common mistakes found in client drafts:** wrong item with a similar name (`FB_PROV_CR_LOSS_OTH` = provision on *other* assets; `FFI_NIM_NIS_DIFF` = NIM–NIS difference; `FFI_LEV_RATIO_ADVT` = advanced-approach), a ratio row pulling a plain amount, an "LCR" built as (current assets–inventory)/current liabilities (= quick ratio), `/100` applied to text, `1-x` on text.
- **Known invalid / empty:** `FF_PROV_LOAN_LOSS` (use `FF_LOAN_LOSS_PROV`), `FF_BK_COM_EQ_TIER1` (negative/NA — not CET1 capital), `FF_LOSS_RATIO` (insurance), `FF_ROTCE(QTR)` for HBAN (use `FB_ROTCE`), `FF_INT_MGN(QTR_R)` for HBAN (use `FB_INT_MGN`).
- A series with the **same value every quarter** means the period reference points at an empty cell (check that each block references its own period row).

---

## 8. Excel wrappers & layout patterns

- **NA → "-" for historical pulls:** `=_xlfn.IFNA(FDS(...),"-")`. When writing the file with openpyxl you **must** include `_xlfn.` (Excel 2013 function) or Excel shows `#NAME?`. `IFERROR` needs no prefix.
- **Cell math:** `=IFERROR((DQ8/ET8)*100,"-")`, `=IFERROR(E8+AH8,"-")`, `=IFERROR(1-C37,"-")`.
- **Estimate sheets — skip FactSet for future quarters (no FE item):**
  `=IF(DATE(VALUE(LEFT(E$5,4)),VALUE(MID(E$5,6,1))*3+1,0)>TODAY(),"-",_xlfn.IFNA(FDS($D8,"FF_X(QTR,"&E$5&")"),"-"))`
- **Grey "-" conditional format:** formula `IFERROR(E8="-",ISNA(E8))`, font `9E9E9E`.
- **Number formats used:**
  - Whole-number %: `#,##0.00"%"_);(#,##0.00"%")`
  - Fraction %: `0.00%_);(0.00%)`
  - $ millions / per share: `$#,##0.00_);($#,##0.00)`; DPS `$#,##0.00#_);($#,##0.00#)`
  - Multiple: `[<0]"NM";#,##0.00"x"`; P/E: `[>100]"NM";[<0]"NM";#,##0.00"x"`
  - Shares: `#,##0.00_);(#,##0.00)`; counts `#,##0_);(#,##0)`
- **Freeze panes** so labels + IDs stay visible while scrolling hundreds of companies (HBAN: `E8`; CAMELS formatted: `D3`; CAMELS source: `C5`).
- **Unit note** in the top-left cell: "$ in millions (USD)…".
- **Styles used:** Quarterly Template — Arial, header fill `D9E7F5` with `1F4E78` bold text, period row `F2F2F2`. Red client theme — red `B00B1C` header with white bold text, light tint `F7E1E4` section rows with red text, red rule under the header; change highlights `FFFF00`.

### Template structures
- **HBAN quarterly template:** one sheet per category (Valuation, Shareholder Returns, Profitability, Revenue & Fee Mix, Efficiency, NIM & Margin, Balance Sheet Evolution, Deposit Competition, Credit, Capital, Estimates, Franchise). Two stacked blocks per sheet (metric header row 4 / 23, period row 5 / 24, 11 banks in rows 8–18 / 27–37, tickers in column D). Each metric = 28 quarter columns (2020/1F–2026/4F). "- Estimates" copies cover 2027/1F–2029/4F (12 quarters) using FE_TIMESERIES spills. Peers: FITB, ZION, RF, PNC, MTB, FHN, HBAN, USB, KEY, CFG, TFC.
- **CAMELS screen:** `CAMELS HISTORICAL ANNUAL` holds the codes (IDs row 2, names row 3 `FG_COMPANY_NAME`, fiscal year row 4, metrics rows 6–38, ~669 banks across columns); `Formatted Annual Camels` links to it (some rows `/100`), average column `=_xlfn.AGGREGATE(1,6,range)`; black-fill conditional format for 0/errors marks missing data. Single-company NTM / Estimate / LTM sheets use the ID in `B1`.

---

## 9. Working with the files in Python (Claude Code)

- Use **openpyxl**. Read formulas with `load_workbook(p)`, cached values with `load_workbook(p, data_only=True)`. Cached values in a refreshed client file are the best evidence of what a code returns — check them before trusting a code.
- Existing FactSet cells load as `ArrayFormula` with `_xll.` text; get the text with `cell.value.text`. When rewriting, write a plain string `'=FDS(...)'`.
- openpyxl does **not** shift formulas when inserting/deleting rows — regenerate formulas instead.
- Preserve multi-cell array spills (an `ArrayFormula` whose `ref` spans several cells).
- openpyxl drops some package parts (e.g. Microsoft sensitivity label `docMetadata/LabelInfo.xml`); copy it back from the original zip if present.
- **Verify before delivering:** LibreOffice headless convert to .ods and search for `Err:5xx` (formula parse errors); diff every cell against the original and confirm *changed ⇔ yellow*; render a small mock to PNG to eyeball formatting.
- Look up codes fast: `grep -i "net interest margin" FactSet_Code_Catalog.csv`, filter the "Verified" column first.

---

## 10. Checklist for a new FDS task

1. Identify company universe & ID type (ticker, ticker-exchange, SEDOL) and the period grid (quarters `YYYY/QF`, fiscal years, LTM/NTM).
2. For each metric: pick the family (§2), find the code (§6 → catalog, prefer **Verified**), set periodicity + period + USD (§3).
3. Decide units and annualization (§7); add ×100 / ×4 / ÷100 only where the evidence says so.
4. Wrap: `FDSC("-",…)` or `_xlfn.IFNA(FDS(…),"-")`; IFERROR for cell math; date-gate future quarters.
5. Format: number formats, freeze panes, unit note, grey "-", client colour scheme.
6. Highlight changes yellow; write the Change Log; list "please confirm" items and "not changed" items with worksheet names.
7. Sample 5 companies → user refreshes → roll out.
