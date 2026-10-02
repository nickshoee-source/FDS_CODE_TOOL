# Examples

These are the actual scripts used during the HBAN and CAMELS projects. They are **patterns to copy**,
not ready-to-run tools: file paths inside point to the old workspace (`/home/claude/...`), and several
HBAN scripts were one-shot steps in a chain (V4 → V10) and are **not idempotent** (e.g. `fix_v9.py`
would double-apply ×100). For new work use `tools/fds_kit.py`.

- `hban_quarterly_scripts/` — template restructure (V4/V5), estimates sheets with FE_TIMESERIES (V6/V7),
  formatting + IFNA + date-gated estimates (V9), value checks and the Data Check sheet (V10),
  `metric_formats.py` (metric → number format map), `audit.py` (structure/formula/period checks).
- `camels_scripts/` — CAMELS fixes (V2), FFI@FF@FB chains on a 5-bank sample then all banks (V4/V5),
  `changes_only.py` (yellow == changed), `harvest.py` / `catalog.py` (how the code catalog was built).
- The final deliverables (HBAN Quarterly Template V10, CAMELS V5) are client files and are kept privately,
  not in this repository.
