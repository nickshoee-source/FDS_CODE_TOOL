# FactSet FDS kit for Claude Code

Everything Claude Code needs to build and fix FactSet `=FDS` Excel workbooks the way we did in the
HBAN and CAMELS projects.

## Setup (one time, ~10 minutes)
1. Install **Python 3.10+** (python.org; on Windows tick "Add Python to PATH").
2. In this folder: `pip install -r requirements.txt`
3. Optional: install **LibreOffice** (lets the checker catch broken formulas).
4. Install **Claude Code** (desktop app's Code tab, or the CLI — see
   https://code.claude.com/docs/en/overview) and sign in with your Claude account.
5. Open this folder in Claude Code (desktop: choose the folder; CLI: `cd` here and run `claude`).
   Claude reads `CLAUDE.md` automatically and has the `factset-fds` skill.

Tip: make this folder a git repository (`git init && git add . && git commit -m "kit"`) so every
change Claude makes can be reviewed and undone.

## Daily use
- Drop the client's workbook into a `work/` folder and ask, for example:
  - *"Read work/Client.xlsx. Fill the blacked-out cells with the right FactSet codes, highlight
    changes yellow, and give me a list of anything you're not 100% sure about."*
  - *"Add NIM, CET1 ratio and NCO ratio for these 15 tickers, quarterly 2021/1F–2026/4F, Quarterly
    Template formatting."*
  - *"Here's the refreshed file — which codes returned no data?"*
- Open the result in Excel, refresh FactSet, send the refreshed file back to Claude to check.

## What's inside
| Path | What |
|---|---|
| `CLAUDE.md` | always-loaded rules + map of the kit |
| `.claude/skills/factset-fds/SKILL.md` | step-by-step workflow skill |
| `docs/FactSet_FDS_Coding_Guide.md` | the full coding guide |
| `docs/PROJECT_HISTORY.md` | what was built for each client and why |
| `reference/FactSet_Code_Catalog.csv` | 9,684 codes, VERIFIED flags, working examples, warnings |
| `reference/factset_code_files/` | FactSet's FF / FFI / FB code lists |
| `work/` (git-ignored) | your private client workbooks; never committed |
| `tools/fds_kit.py` | formula builders, styles, change highlighting, verification |
| `tools/find_code.py` | search codes: `python tools/find_code.py net interest margin --verified` |
| `tools/harvest_codes.py` | after a refresh: which codes returned data |
| `tools/update_catalog.py` | after a refresh: mark working codes VERIFIED in the catalog |
| `tools/check_workbook.py` | pre-delivery checks (hygiene, yellow == changed, parse) |
| `examples/` | past build scripts (patterns only) |
| `samples/HBAN_FDS_Code_Test_SAMPLE.xlsx` | test workbook (~1,700 formulas) for HBAN, peers and 3 global banks; rebuild with `python samples/build_code_test.py` |

## Confidentiality
This repository is public. Client workbooks (refreshed files, deliverables, instructions) are **not**
included. Keep them in the git-ignored `work/` folder on your own machine. The FactSet codes, catalog
and tools are not confidential.
