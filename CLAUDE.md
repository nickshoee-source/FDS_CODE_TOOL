# FactSet =FDS Excel Project

This project builds and fixes Excel workbooks that pull bank data with the FactSet Excel add-in
(`=FDS`, `=FDSC`, `=FDSR`, `=FDSRC`). Claude writes the formulas with Python/openpyxl; the user
opens the file in Excel (FactSet add-in), refreshes, and sends back the refreshed file as evidence.

## Read first
- `docs/FactSet_FDS_Coding_Guide.md` — syntax, code families, verified code library, units, traps.
  **Read it before writing any FactSet formula.**
- `docs/PROJECT_HISTORY.md` — what was built for each client and the decisions made.
- Use the `factset-fds` skill (`.claude/skills/factset-fds/`) for the step-by-step workflow.

## Rules (non-negotiable — the user's standing preferences)
1. Formulas are clean: only `=FDS(...)`/`=FDSC(...)`/`=FDSRC(...)` + normal Excel. Never write
   `_xll.` or ArrayFormula objects (they show up as `{}` and `@`). Write plain strings.
2. Excel 2013+ functions written by openpyxl need `_xlfn.` (`_xlfn.IFNA`) or Excel shows #NAME?.
3. Triple-check codes: prefer codes marked VERIFIED in the catalog; check units against real values.
4. Change only what you are 100% sure of. Everything else goes on a "please confirm" list with
   worksheet name, row/metric, code, and reason.
5. Yellow (`FFFF00`) = changed from the original, and nothing else. Verify with
   `python tools/check_workbook.py NEW.xlsx --original ORIGINAL.xlsx` (mismatches must be 0).
6. Always report what you did NOT change, with worksheet names.
7. Large rollouts (hundreds of companies): build the first 5 companies, ask the user to refresh,
   then roll out.
8. Keep a "Change Log" sheet in edited client files. Never overwrite the user's FDSRC spill results.
9. Save new versions as new files (`..._V2.xlsx`); never edit the user's original in place.

## Frequency codes (first argument of every FF_/FFI_/FB_ code)
`QTR` = quarter, `QTR_R` = quarter rolling (user's definition; older notes said "restated" — see guide
§3.1), `ANN` = annual, `LTM` = last twelve months. FFI_ uses `ANN_L` / `LTM_L`. Estimates use their own
rolling forms `QTR_ROLL` / `ANN_ROLL` / `NTMA`. Example: `FF_ASSETS(QTR,"&E$5&")`, `FF_EPS(QTR_R,"&E$5&")`.

## Tools (run from the project root)
| Command | Purpose |
|---|---|
| `python tools/find_code.py <words> [--family FFI] [--verified]` | search 9,684 FactSet codes |
| `python tools/harvest_codes.py REFRESHED.xlsx` | which codes returned data after the user refreshed |
| `python tools/update_catalog.py REFRESHED.xlsx --source "label"` | mark codes that returned numbers as VERIFIED in the catalog |
| `python tools/check_workbook.py NEW.xlsx --original OLD.xlsx [--fix-highlights OUT.xlsx]` | hygiene, yellow==changed, parse check |
| `from tools.fds_kit import *` | formula builders (`fds`, `fdsc`, `fdsrc`, `ifna`, `future_gate`, `chain`, `ff`, `ffi`, `fb`, `fe_timeseries`, `fe_estimate`), styles (`THEMES['quarterly'|'red']`, `FMT`), `highlight_changes`, `write_change_log`, `freeze`, `restore_sensitivity_label` |

## Folders
- `reference/FactSet_Code_Catalog.csv` — every code + description + VERIFIED flag + working example.
- `reference/factset_code_files/` — FactSet's own code lists (FF, FFI banks/SF/insurance, FB regulatory).
- `work/` (git-ignored, local only) — client workbooks, including refreshed files (ground truth for
  what codes return). Never commit client files; this repository is public.
- `examples/` — past build scripts (patterns only; paths inside point to the old workspace). Deliverables are kept privately, not in the repo.

## Environment
- Python 3.10+, `pip install -r requirements.txt` (openpyxl). LibreOffice is optional (parse check).
- FactSet only calculates inside Excel with the add-in; Claude cannot refresh data. Outside Excel,
  FactSet cells show #NAME? — that is expected.
