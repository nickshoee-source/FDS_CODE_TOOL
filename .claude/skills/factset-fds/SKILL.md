---
name: factset-fds
description: Build, fix or extend Excel workbooks that use FactSet =FDS / =FDSC / =FDSRC formulas (bank metrics, CAMELS screens, peer templates, consensus estimates). Use whenever the user asks for FactSet codes, FDS formulas, or edits to a FactSet template.
---

# FactSet FDS workflow

Read `docs/FactSet_FDS_Coding_Guide.md` first (syntax, families, verified codes, units, traps).

## 1. Understand the workbook
- Open with openpyxl twice: formulas (`load_workbook(p)`) and cached values (`data_only=True`).
- Map: ID cells (ticker / SEDOL), period cells (`YYYY/QF`, fiscal year), metric rows/blocks,
  which sheet holds codes vs which sheet only links/formats.
- If the file was refreshed, run `python tools/harvest_codes.py FILE.xlsx` to see which codes return
  data, and summarise cached values per metric (counts of numbers / "-" / errors, medians) to spot
  wrong items (e.g. a "NIM" with median 0.2, amounts in local currency, ratios as plain amounts).
- Find "blacked out"/missing cells via conditional formatting rules (0 / ISERROR fills).

## 2. Choose codes
- `python tools/find_code.py <words> --verified` first, then without `--verified`.
- Family order: US quarterly → FF_/FF_BK_ then FB_; global annual → FFI_ then FF_ then FB_ (US only);
  forecasts → FE_ items (FE_TIMESERIES spill with FDSRC for multi-quarter estimates).
- Add `,,,,USD` to FF_/FFI_ money items; FB money items `(ANN,yr,,,RF,USD)`.
- Units: whole-number % except the fractions listed in the guide; ×4 for quarterly rates
  (FF_CHARGE_OFFS_LOANS_PCT, FF_COST_DEPS).
- Fallbacks: `chain(a, b, c)` → `a@b@c`. Cross-family chains are unproven — sample first.
- Anything not 100% certain → "please confirm" list.

## 3. Build
- Use `tools/fds_kit.py` builders; write plain strings (no `_xll.`, no ArrayFormula).
- Historical pulls: `ifna(fds(...))` or `fdsc(...)`; cell math in `IFERROR(...,"-")`;
  estimate sheets without an FE item: `future_gate(period_cell, formula)`.
- Formatting: number formats from `FMT`, `grey_dash_rule`, `freeze(ws, cell)`, even column widths,
  theme `quarterly` (blue) or `red` or the client's own, unit note in the top-left.
- Big universes: write only the first 5 companies, save as `..._SAMPLE_5.xlsx`, stop and ask the
  user to refresh. Roll out only after they confirm.

## 4. Track changes
- `highlight_changes(original, wb)` so yellow == changed exactly; `write_change_log(wb, rows)`.
- `restore_sensitivity_label(original, saved)` after `wb.save`.

## 5. Verify before delivering
- `python tools/check_workbook.py NEW.xlsx --original ORIGINAL.xlsx` → hygiene ok, highlight
  mismatches 0, no parse errors.
- Re-read 2–3 formulas per changed row; confirm references point at the right ID/period cells.

## 6. Report to the user (short)
- What changed (by worksheet), codes used, the "please confirm" list, what was NOT changed and why.
