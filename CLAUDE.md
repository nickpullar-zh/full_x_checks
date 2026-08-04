# X-Checks Full Application — Infrastructure (`main`)

## Repository

- **Repo:** https://github.com/nickpullar-zh/full_x_checks
- **Infrastructure branch:** `main` — shared plumbing only, no strategy code
- **Strategy branches** (each carries its own strategy module + tests + UAT):
  - `obsolete_to_v0.4-X-Checks` — X-Checks comparison (superseded by v0.4; kept for reference)
  - `v0.2-Grouping_By` — Grouping By comparison (old pre-split codebase; superseded by v0.3-Grouping_By)
  - `v0.3-Grouping_By` — Grouping By comparison (ported to modern branch structure)
  - `v0.5-Accounting-Principles` — Accounting Principles (shipped)
  - `v0.6-Conditions` — Conditions (in progress)

## Branch architecture

`main` carries **only infrastructure** that every strategy depends on:

- App skeleton: `main.py`, `progress_dialog.py`, `file_upload_ui.py`, `file_upload_config.py`, `task_configs.py`, `task_registry.py`
- Shared base class: `strategies/base_strategy.py` (file loading, Excel writing, formatting, logging)
- Build/packaging: `build.py` with PyInstaller spec generation, splash, fonts
- Branding: Zurich brand fonts under `templates/fonts/`, splash template, output template
- Configuration scaffolding: `config.py`, `exceptions.py`, `version.py`

`task_registry.TASK_REGISTRY` and `task_configs` start **empty** on `main`. Each strategy branch:

1. Adds its `UploadTaskConfig` to `task_configs.py`
2. Registers it in `TASK_REGISTRY` and adds the matching `if False:` import for PyInstaller
3. Adds a `BUILDS` debug entry in `build.py` if it wants a debug EXE
4. Adds its own `_DEBUG_FILES_*` dict in `main.py` and an entry in `_DEBUG_FILES_MAP`
5. Adds its own change-log section to that branch's `CLAUDE.md`

Strategy branches must NEVER depend on code from another strategy branch. The infrastructure on `main` is the only shared surface.

---

## Code Reuse Rule

**Before writing any new function, block, or constant, ask: "Does this already exist in `BaseStrategy` or elsewhere in the codebase?" If yes, reuse it. Never duplicate.**

- Shared colours, fills, and fonts → `BaseStrategy` class constants (`FILL_GREEN`, `FONT_RED`, etc.)
- Shared loading or annotation logic → `BaseStrategy` method
- Strategy-specific formatting logic → the strategy's own `apply_output_formatting()`; Full Run delegates via `_PrefixedWorkbook` shim, never reimplements
- If you find yourself writing the same pattern in two places, stop and centralise it first

---

## Change Log Policy

**Every change, no matter how small, must be logged here. This is not optional and must never be skipped.**

- When a change is **proposed**, add it to the change log immediately with status `PROPOSED`.
- When a change is **implemented and confirmed**, update the entry to `DONE`.
- No change may be silently applied. If a change is rejected or abandoned, mark it `REJECTED` with a reason.
- **Claude must update the change log as part of implementing the change — not as an afterthought, and not only when reminded.** Updating the log is the final step before marking any task complete.
- **Every change must also bump the version in `version.py`.** The version is displayed in the app title bar, UI label, and Excel processing log so users can confirm which version they are running when reporting issues.

---

## Change Log

### v1.0.93 — Lower minimum screen resolution to 1280×720 (2026-08-04)

| # | Change | Status | Notes |
|---|--------|--------|-------|
| 1 | `main.py`: `MIN_WIDTH` 1920→1280, `MIN_HEIGHT` 1080→720. | DONE | |
| 2 | `version.py`: bump to `1.0.93`. | DONE | |

### v1.0.92 — Scrollable upload form; minimum 1920×1080 screen size check (2026-08-04)

| # | Change | Status | Notes |
|---|--------|--------|-------|
| 1 | `file_upload_ui.py`: replace two-column overflow approach with a scrollable canvas. File fields and output directory sit inside a canvas that is capped to fit the usable screen height; controls (checkboxes, Proceed, version) are in a fixed frame below the canvas. Mousewheel scrolling enabled. | DONE | Works correctly on all strategies regardless of field count. |
| 2 | `main.py`: screen size check at startup — if width < 1920 or height < 1080, shows an error dialog explaining the requirement and exits without launching the app. | DONE | |
| 3 | `version.py`: bump to `1.0.92`. | DONE | |

### v1.0.91 — Full Run dialog: 20px screen margin; two-column layout when too tall (2026-08-04)

| # | Change | Status | Notes |
|---|--------|--------|-------|
| 1 | `file_upload_ui.py`: add `_SCREEN_MARGIN = 20` px; `_set_position` clamps all four edges with this margin. | DONE | |
| 2 | `file_upload_ui.py`: `_build_ui` measures dialog height after initial render; if it exceeds usable screen height minus margins, destroys the single-column fields panel and rebuilds it as two side-by-side columns (split at the first SectionConfig after the midpoint). | DONE | |
| 3 | `file_upload_ui.py`: extract field-building into `_build_fields_panel(parent, fields, hint_wrap, two_col)` reused by both layouts. Add `_usable_screen_height()` helper. | DONE | |
| 4 | `version.py`: bump to `1.0.91`. | DONE | |

### v1.0.90 — AP pastel tab colour darkened to distinguish from Conditions (2026-08-04)

| # | Change | Status | Notes |
|---|--------|--------|-------|
| 1 | `accounting_principles.py` + `full_run.py`: AP pastel changed from `B4C6E7` to `8BAFC7` (darker mid-blue) to distinguish it from Conditions pastel `BDD7EE`. | DONE | |
| 2 | `version.py`: bump to `1.0.90`. | DONE | |

### v1.0.89 — Tab colour coding for all individual strategy outputs (2026-08-04)

| # | Change | Status | Notes |
|---|--------|--------|-------|
| 1 | `base_strategy.py`: add `TAB_COLOUR` / `TAB_COLOUR_PASTEL` class constants (default `""`); add `_apply_tab_colours(workbook)` — Comparison tab gets main colour, other data sheets get pastel, Processing Log gets grey. | DONE | |
| 2 | `x_checks.py`: `TAB_COLOUR="70AD47"` (green), `TAB_COLOUR_PASTEL="D9EAD3"`; call `_apply_tab_colours` at end of `apply_output_formatting`. | DONE | |
| 3 | `grouping_by.py`: `TAB_COLOUR="ED7D31"` (orange), `TAB_COLOUR_PASTEL="FCE4D6"`; call `_apply_tab_colours`. | DONE | |
| 4 | `accounting_principles.py`: `TAB_COLOUR="23366F"` (dark blue), `TAB_COLOUR_PASTEL="B4C6E7"`; call `_apply_tab_colours`. | DONE | |
| 5 | `conditions.py`: `TAB_COLOUR="2167AE"` (Zurich blue), `TAB_COLOUR_PASTEL="BDD7EE"`; call `_apply_tab_colours`. | DONE | |
| 6 | `full_run.py`: add `STRATEGY_COLOURS_PASTEL` dict; reorder `apply_output_formatting` to delegate cell formatting first, then overwrite tab colours; supporting sheets get pastel, Comparison gets main colour, Processing Log grey. | DONE | |
| 7 | `version.py`: bump to `1.0.89`. | DONE | |

### v1.0.88 — Fix Full Run SectionConfig crash in _load_files, full_run.py, main.py (2026-08-04)

| # | Change | Status | Notes |
|---|--------|--------|-------|
| 1 | `strategies/base_strategy.py`: `_load_files` column_map and signals_map comprehensions now filter to `isinstance(f, FileFieldConfig)` — SectionConfig has no `.label`, `.required_columns`, or `.header_signals`. | DONE | |
| 2 | `strategies/full_run/full_run.py`: `strategy_labels` set comprehension guarded with `isinstance(f, FileFieldConfig)`. | DONE | |
| 3 | `main.py`: `required_labels` and `extra` set comprehensions in debug path guarded with `isinstance(f, FileFieldConfig)`. | DONE | |
| 4 | `version.py`: bump to `1.0.88`. | DONE | |

### v1.0.87 — Fix Full Run Proceed button never enabling (SectionConfig in _check_ready) (2026-08-04)

| # | Change | Status | Notes |
|---|--------|--------|-------|
| 1 | `file_upload_ui.py`: `_check_ready` now skips non-`FileFieldConfig` entries (i.e. `SectionConfig` dividers) before checking `field.required`. Previously the loop hit a `SectionConfig`, threw `AttributeError` on `.required`, and left the button permanently disabled for Full Run. | DONE | Individual strategies have no `SectionConfig` in their field lists so were unaffected. |
| 2 | `version.py`: bump to `1.0.87`. | DONE | |

### v1.0.86 — KEL sheet name: editable on individual strategies, locked on Full Run (2026-08-04)

| # | Change | Status | Notes |
|---|--------|--------|-------|
| 1 | `file_upload_config.py`: add `sheet_editable: bool = False` to `FileFieldConfig`. When True the combobox becomes free-text after file selection (`state="normal"`); when False it stays read-only. | DONE | |
| 2 | `task_configs.py`: all four strategy KEL fields gain `sheet_editable=True` so users can type a custom sheet name if needed. Full Run KEL field keeps `sheet_editable=False` (locked, with the `sheet_note`). | DONE | |
| 3 | `file_upload_ui.py`: `_browse_file` and `_apply_prefill` use `state="normal"` vs `"readonly"` based on `field.sheet_editable`. | DONE | |
| 4 | `version.py`: bump to `1.0.86`. | DONE | |

### v1.0.85 — Full Run KEL: grey out sheet name with explanatory note (2026-08-04)

| # | Change | Status | Notes |
|---|--------|--------|-------|
| 1 | `file_upload_config.py`: add `sheet_note: str = ""` field to `FileFieldConfig`. When non-empty, the sheet combobox stays permanently disabled and the note is shown below it. | DONE | |
| 2 | `file_upload_ui.py`: render the sheet note as a grey label below the combobox when `field.sheet_note` is set. `_browse_file` skips combobox population and remains disabled for locked fields. | DONE | |
| 3 | `task_configs.py`: Full Run config now supplies its own `Known Exception List` field with `default_sheet="(per strategy)"` and `sheet_note="Sheet name is set automatically per strategy (X-Checks / Grouping By / Accounting Principles / Conditions)."` | DONE | |
| 4 | `version.py`: bump to `1.0.85`. | DONE | |

### v1.0.84 — KEL file picker: default sheet matches the strategy sheet name (2026-08-04)

| # | Change | Status | Notes |
|---|--------|--------|-------|
| 1 | `task_configs.py`: Known Exception List `default_sheet` updated for all four strategies — X-Checks: `"X-Checks"`, Grouping By: `"Grouping By"`, Accounting Principles: `"Accounting Principles"`, Conditions: `"Conditions"`. Previously all defaulted to `"Known Exceptions"` which doesn't exist in the shared workbook. | DONE | The file picker auto-selects this sheet when the user browses to `known_exception_list.xlsx`. The Full Run is unaffected — each strategy already hardcodes the correct sheet name when calling `_annotate_known_exceptions`. |
| 2 | `version.py`: bump to `1.0.84`. | DONE | |

### v1.0.83 — FX-22: add missing REF_BASE|REF_BASE row to Conditions expected output (2026-08-04)

| # | Change | Status | Notes |
|---|--------|--------|-------|
| 1 | `docs/generate_fixture_uat.py`: FX-22 row list extended with `REF_BASE\|REF_BASE - Not Matched`. This row is produced because `GB_REF_XC_KEY` has `Reference X-Check (Condition)=REF_BASE`; Conditions picks up that cell value and generates an output row with no matching FIP entry. | DONE | |
| 2 | `docs/20260804 Fixture_UAT_v1.0.83 Test Plan.xlsx`: regenerated. | DONE | |
| 3 | `version.py`: bump to `1.0.83`. | DONE | |

### v1.0.82 — Comprehensive fixture UAT accuracy pass (2026-08-03)

| # | Change | Status | Notes |
|---|--------|--------|-------|
| 1 | FX-05 `XC_QU_YTD` reverted to Match — FX-05 supplies no GCoA, so EBX produces VAL_YTD matching FIP. MisMatch only occurs in FX-06e where GCoA is supplied. | DONE | |
| 2 | FX-06 items a/b reordered to Formula Match → Variables Match sequence. Item b corrected: Formula Match=MisMatch for XC_FORMULA_MISMATCH. | DONE | |
| 3 | FX-06e expected result expanded to show all 4 match columns in sequence. | DONE | |
| 4 | FX-07 Comparison row count corrected: 30 → 31. | DONE | |
| 5 | FX-07b counts corrected: Matched=17 (was 16), total=31 (was 30). | DONE | |
| 6 | FX-09 "all 30 rows" → "all 31 rows". | DONE | |
| 7 | FX-22 Conditions row count corrected: 15 → 16. | DONE | |
| 8 | FX-12/13/14 GB file paths updated to flat layout (`{F}\\gb_*`; `xc_pub.xlsx`; `known_exception_list.xlsx`). | DONE | |
| 9 | FX-17/19 AP file paths updated to flat layout (`{F}\\ap_*`; `xc_pub.xlsx`; `known_exception_list.xlsx`). | DONE | |
| 10 | FX-22/24 Conditions file paths updated to flat layout (`{F}\\cond_*`; `xc_pub.xlsx`; `known_exception_list.xlsx`). | DONE | |
| 11 | `docs/20260803 Fixture_UAT_v1.0.82 Test Plan.xlsx`: regenerated. | DONE | |
| 12 | `version.py`: bump to `1.0.82`. | DONE | |

### v1.0.81 — Fix FX-05 and FX-06e: XC_QU_YTD result is MisMatch, not Match (2026-08-03)

| # | Change | Status | Notes |
|---|--------|--------|-------|
| 1 | `docs/generate_fixture_uat.py`: FX-05 `XC_QU_YTD` corrected to `MisMatch` — EBX produces `QU_YTD(ACC_QU)<=0` but FIP is parsed as `VAL_YTD(ACC_QU)<=0` (no QU substitution in FIP). | DONE | |
| 2 | `docs/generate_fixture_uat.py`: FX-06e expected result corrected to `Formula Match = MisMatch`; FIP formula description corrected to `VAL_YTD(ACC_QU)<=0`. | DONE | |
| 3 | `docs/20260803 Fixture_UAT_v1.0.81 Test Plan.xlsx`: regenerated. | DONE | |
| 4 | `version.py`: bump to `1.0.81`. | DONE | |

### v1.0.80 — Fix FX-05 row order: XC_QU_YTD moved to correct alphabetical position (2026-08-03)

| # | Change | Status | Notes |
|---|--------|--------|-------|
| 1 | `docs/generate_fixture_uat.py`: FX-05 row list reordered so `XC_QU_YTD` appears between `XC_PCT_FORMAT` and `XC_REORDER_MATCH` (correct alphabetical order matching actual output). | DONE | |
| 2 | `docs/20260803 Fixture_UAT_v1.0.80 Test Plan.xlsx`: regenerated. | DONE | |
| 3 | `version.py`: bump to `1.0.80`. | DONE | |

### v1.0.79 — Fix FX-05 test plan: XC_EXCL_MISMATCH row result is MisMatch (2026-08-03)

| # | Change | Status | Notes |
|---|--------|--------|-------|
| 1 | `docs/generate_fixture_uat.py`: FX-05 `XC_EXCL_MISMATCH` entry corrected from `Match` to `MisMatch` — Formula Match=Match but Formula Match (Excl)=MisMatch, so the row-level result is MisMatch. | DONE | |
| 2 | `docs/20260803 Fixture_UAT_v1.0.79 Test Plan.xlsx`: regenerated. | DONE | |
| 3 | `version.py`: bump to `1.0.79`. | DONE | |

### v1.0.78 — Merge all four KEL files into one known_exception_list.xlsx (2026-07-31)

| # | Change | Status | Notes |
|---|--------|--------|-------|
| 1 | `test_data/generate_all_fixtures.py`: replace four separate `_make_xc/gb/ap/cond_kel()` functions with one `_make_known_exception_list()` that writes a single `known_exception_list.xlsx` with one sheet per strategy (X-Checks, Grouping By, Accounting Principles, Conditions, Instructions). | DONE | |
| 2 | All four logic test scripts updated to reference `known_exception_list.xlsx`. | DONE | |
| 3 | `docs/generate_fixture_uat.py`: KEL file references updated; test plan regenerated. | DONE | |
| 4 | Removed from git: `xc_kel.xlsx`, `gb_kel.xlsx`, `ap_kel.xlsx`, `cond_kel.xlsx`. | DONE | |
| 5 | `version.py`: bump to `1.0.78`. | DONE | |

### v1.0.77 — Remove stale fixture files and old sub-folders from git (2026-07-31)

| # | Change | Status | Notes |
|---|--------|--------|-------|
| 1 | Removed from git: `fixtures/fip_xc.txt`, `fixtures/fip_ZQ9_*.xlsx`, `fixtures/mapping.txt`, `fixtures/ap/`, `fixtures/cond/`, `fixtures/gb/` (all superseded by flat layout). | DONE | Empty sub-folders on disk locked by OneDrive — delete manually in Explorer. |
| 2 | `version.py`: bump to `1.0.77`. | DONE | |

### v1.0.76 — Update fixture UAT plan with new flat file names, GCoA test, corrected row counts (2026-07-31)

| # | Change | Status | Notes |
|---|--------|--------|-------|
| 1 | `docs/generate_fixture_uat.py`: all file references updated to flat layout (xc_fip.txt, xc_gcoa.xlsx, gb_fip_ZQ9_VALFLDGR.xlsx, gb_mapping.txt, ap_fip_ZQ9_VALMSG.xlsx, cond_fip_ZQ9_VALMETH.xlsx, xc_kel.xlsx). Add FX-06e GCoA QU_YTD test case. FX-05 row count 30→31. Cond row count 15→16. Full Run file list updated with all 9 fixture files. | DONE | |
| 2 | `docs/20260731 Fixture_UAT_v1.0.75 Test Plan.xlsx`: generated. | DONE | |
| 3 | `version.py`: bump to `1.0.76`. | DONE | |

### v1.0.75 — Consolidated flat fixture layout; GCoA test; all strategy files in fixtures/ (2026-07-31)

| # | Change | Status | Notes |
|---|--------|--------|-------|
| 1 | `test_data/generate_all_fixtures.py`: new single generator producing all fixture files into `test_data/fixtures/` flat folder. Naming: `xc_` prefix for XC-only, `gb_`/`ap_`/`cond_` for strategy-specific, plain names for shared. Includes `xc_gcoa.xlsx` (GCoA file for QU_YTD test). Merges all strategy pub rows into one `xc_pub.xlsx`. | DONE | |
| 2 | All four logic test scripts updated to use flat paths, explicit label strings, and corrected row counts (XC=75, GB=22, AP=26, Cond=26 — total 149 assertions, all pass). | DONE | |
| 3 | `version.py`: bump to `1.0.75`. | DONE | |

### v1.0.74 — Update both UAT plans for new X-Checks output structure (2026-07-31)

| # | Change | Status | Notes |
|---|--------|--------|-------|
| 1 | `docs/generate_fixture_uat.py`: FX-07 updated to list all 7 sheets; add FX-07b (Logic) verifying filtered sheet counts (16 Matched, 7 MisMatched, 7 Not Found); FX-08 updated for colour coding on filtered sheets; Full Run sheet list updated with 6 XC sheets. | DONE | |
| 2 | `docs/generate_full_app_uat.py`: FA-07 and FA-08 updated to match new 7-sheet output structure and filtered counts. | DONE | |
| 3 | `docs/20260731 Fixture_UAT_v1.0.73 Test Plan.xlsx` + `docs/20260731 Full_Application_v1.0.73 Test Plan.xlsx`: generated. | DONE | |
| 4 | `version.py`: bump to `1.0.74`. | DONE | |

### v1.0.73 — X-Checks output: add source sheets and three filtered Comparison sheets (2026-07-31)

| # | Change | Status | Notes |
|---|--------|--------|-------|
| 1 | `strategies/x_checks/x_checks.py`: output now contains 6 sheets: EBX Data (raw cross checks all), FIP Data (parsed FIP results), Comparison (full, unchanged), Matched Data (all 4 match cols = Match or MisMatch(Excepted)), MisMatched Data (any col = MisMatch or MisMatch(Excepted)), Not Found Data (any col = Not Found). | DONE | |
| 2 | `version.py`: bump to `1.0.73`. | DONE | |

### v1.0.72 — Update FIP file descriptions to match actual raw export column headers (2026-07-31)

| # | Change | Status | Notes |
|---|--------|--------|-------|
| 1 | `task_configs.py`: VALFLDGR description updated to list exact columns (ValidRule, Long Text, Field name). VALMSG description updated to list exact raw ZQ9_VALMSG columns; removes misleading Key column mention. VALMETH description updated to list exact positional columns. | DONE | |
| 2 | `version.py`: bump to `1.0.72`. | DONE | |

### v1.0.71 — Fix CF fill: set bgColor = fgColor for solid fill to render in Excel (2026-07-31)

| # | Change | Status | Notes |
|---|--------|--------|-------|
| 1 | `docs/generate_fixture_uat.py`: CF `PatternFill` now sets both `fgColor` and `bgColor` to `FFD6E4F7`. Excel requires both for a conditional format solid fill to display. | DONE | |
| 2 | `docs/20260731 Fixture_UAT_v1.0.70 Test Plan.xlsx`: regenerated with correct CF. | DONE | |
| 3 | `version.py`: bump to `1.0.71`. | DONE | |

### v1.0.70 — Regenerate fixture UAT plan at v1.0.69 (2026-07-31)

| # | Change | Status | Notes |
|---|--------|--------|-------|
| 1 | `docs/20260731 Fixture_UAT_v1.0.69 Test Plan.xlsx`: regenerated. CF rule confirmed present (A4:J42, formula $C4="Logic", fill FFD6E4F7). | DONE | |
| 2 | `version.py`: bump to `1.0.70`. | DONE | |

### v1.0.69 — Pub first, FIP second in all strategies; Full Run with section dividers and titles (2026-07-31)

| # | Change | Status | Notes |
|---|--------|--------|-------|
| 1 | `file_upload_config.py`: add `SectionConfig(title="")` dataclass for dividers and section titles. | DONE | |
| 2 | `file_upload_ui.py`: import `SectionConfig`; update field loop to render a `ttk.Separator` + optional bold label when it encounters a `SectionConfig`. | DONE | |
| 3 | `task_configs.py`: all four strategies reordered to Pub first, FIP second, other files, KEL last. `_build_full_run_config` replaced with explicit ordered layout using `SectionConfig` dividers per strategy. | DONE | |
| 4 | `strategies/grouping_by/grouping_by.py`: replace index-based `file_fields[0/1/2].label` with literal label strings (safe after reorder). | DONE | |
| 5 | `version.py`: bump to `1.0.69`. | DONE | |

### v1.0.68 — Standardise file field order: FIP first, Pub second, KEL last (2026-07-31)

| # | Change | Status | Notes |
|---|--------|--------|-------|
| 1 | `task_configs.py`: AP reordered to FIP File (VALMSG), X-Checks Publication File, Validation Methods File, Known Exception List. Conditions reordered to FIP File (ZQ9_VALMETH), X-Checks Publication File, Known Exception List. X-Checks and Grouping By were already correct. | DONE | |
| 2 | `version.py`: bump to `1.0.68`. | DONE | |

### v1.0.67 — Fixture UAT plan: Logic row colour via conditional formatting on column C (2026-07-31)

| # | Change | Status | Notes |
|---|--------|--------|-------|
| 1 | `docs/generate_fixture_uat.py`: replace static per-cell `PatternFill` on Logic rows with a `FormulaRule` conditional format on `A4:J{last_row}` — fills light blue when `$C4="Logic"`. This means a tester can change "Whole App" to "Logic" in column C and the row colour updates automatically. | DONE | |
| 2 | `docs/20260731 Fixture_UAT_v1.0.66 Test Plan.xlsx`: regenerated. | DONE | |
| 3 | `version.py`: bump to `1.0.67`. | DONE | |

### v1.0.66 — Workbooks open on Comparison sheet by default (2026-07-30)

| # | Change | Status | Notes |
|---|--------|--------|-------|
| 1 | `strategies/base_strategy.py`: `write_excel_output` now sets `workbook.active` to the first sheet whose name contains "Comparison" before saving. Falls back to first sheet if none found. Applies to all strategies. | DONE | |
| 2 | `version.py`: bump to `1.0.66`. | DONE | |

### v1.0.65 — Clarify KEL test cases: explicitly state Process only differences must be unchecked (2026-07-30)

| # | Change | Status | Notes |
|---|--------|--------|-------|
| 1 | `docs/generate_fixture_uat.py`: FX-09, FX-13, FX-19 KEL test cases now explicitly state "'Process only differences' must be UNCHECKED" and step updated to "Uncheck... then click Proceed". Without this, the default checkbox causes the comparison to be filtered to 2 rows and the KEL rows never appear. | DONE | |
| 2 | `docs/20260730 Fixture_UAT_v1.0.64 Test Plan.xlsx`: regenerated. | DONE | |
| 3 | `version.py`: bump to `1.0.65`. | DONE | |

### v1.0.64 — Field label: (optional) on new line; wraplength 180px (2026-07-30)

| # | Change | Status | Notes |
|---|--------|--------|-------|
| 1 | `file_upload_ui.py`: optional field label now uses `\n(optional)` so the field name and "(optional)" are on separate lines. Required fields keep `*` on the same line. wraplength reduced to 180px. | DONE | |
| 2 | `version.py`: bump to `1.0.64`. | DONE | |

### v1.0.63 — Shorten button label; fix field label wrapping (2026-07-30)

| # | Change | Status | Notes |
|---|--------|--------|-------|
| 1 | `main.py`: "Return to Strategy Selection" → "Return to Selection" (fits the button). | DONE | |
| 2 | `file_upload_ui.py`: field label wraplength 150 → 200 so "X-Checks Publication File *" fits on one line. | DONE | |
| 3 | `version.py`: bump to `1.0.63`. | DONE | |

### v1.0.62 — Fix "Start" to "Proceed" in both test plan generators (2026-07-30)

| # | Change | Status | Notes |
|---|--------|--------|-------|
| 1 | `docs/generate_fixture_uat.py` + `docs/generate_full_app_uat.py`: replace all 18 occurrences of "click/press Start" with "click/press Proceed" to match the actual button label. | DONE | |
| 2 | `docs/20260730 Fixture_UAT_v1.0.61 Test Plan.xlsx` + `docs/20260730 Full_Application_v1.0.61 Test Plan.xlsx`: regenerated. | DONE | |
| 3 | `version.py`: bump to `1.0.62`. | DONE | |

### v1.0.61 — Generate test plans at v1.0.60 (2026-07-30)

| # | Change | Status | Notes |
|---|--------|--------|-------|
| 1 | `docs/20260730 Fixture_UAT_v1.0.60 Test Plan.xlsx`: generated with Internal_Use_Only label. | DONE | |
| 2 | `docs/20260730 Full_Application_v1.0.60 Test Plan.xlsx`: generated with Internal_Use_Only label. | DONE | |
| 3 | `version.py`: bump to `1.0.61`. | DONE | |

### v1.0.60 — Apply Internal_Use_Only label to generated test plans (2026-07-30)

| # | Change | Status | Notes |
|---|--------|--------|-------|
| 1 | `docs/generate_fixture_uat.py` + `docs/generate_full_app_uat.py`: call `ExcelLabeler` after saving the workbook to apply the Internal_Use_Only MIP label. Failure is swallowed silently so the script still completes. | DONE | |
| 2 | `docs/20260730 Fixture_UAT_v1.0.59 Test Plan.xlsx` + `docs/20260730 Full_Application_v1.0.59 Test Plan.xlsx`: regenerated with label applied. | DONE | |
| 3 | `version.py`: bump to `1.0.60`. | DONE | |

### v1.0.59 — Output both UAT test plans at v1.0.58 (2026-07-30)

| # | Change | Status | Notes |
|---|--------|--------|-------|
| 1 | `docs/20260730 Fixture_UAT_v1.0.58 Test Plan.xlsx`: generated. | DONE | |
| 2 | `docs/20260730 Full_Application_v1.0.58 Test Plan.xlsx`: generated. | DONE | |
| 3 | `version.py`: bump to `1.0.59`. | DONE | |

### v1.0.58 — Fixture UAT plan expanded: GB/AP/Conditions each list every row with expected output (2026-07-30)

| # | Change | Status | Notes |
|---|--------|--------|-------|
| 1 | `docs/generate_fixture_uat.py`: GB, AP, and Conditions sections rewritten to the same level of rigour as X-Checks. Each now has dedicated test cases covering: full comparison output (every row named with exact EBX Key/Result or X-Check/Event/FIP/Actual/Match), differences mode, KEL annotation, output structure, colour coding. Full Run row counts updated to reflect actual fixture outputs (GB=14, AP=11, Cond=15). Total 44 test cases (17 Logic, 27 Whole App). | DONE | |
| 2 | `docs/20260730 Fixture_UAT_v1.0.57 Test Plan.xlsx`: generated output. | DONE | |
| 3 | `version.py`: bump to `1.0.58`. | DONE | |

### v1.0.57 — "Close" after success returns to Strategy Selection, not exit (2026-07-30)

| # | Change | Status | Notes |
|---|--------|--------|-------|
| 1 | `progress_dialog.py`: add `on_success` callback parameter; on successful completion the action button calls `on_success()` instead of exiting. `on_dismiss` (cancel/error) still returns to the upload form. Debug mode (no callbacks) still exits. | DONE | |
| 2 | `main.py`: pass `on_success=_return_to_selector` which calls `self.root.deiconify()` to show the task selector. Rename button text from `"Close"` to `"Return to Strategy Selection"`. | DONE | |
| 3 | `version.py`: bump to `1.0.57`. | DONE | |

### v1.0.56 — Progress dialog: copy button right-aligned with tooltip; green separator; fix [[Complete]] (2026-07-30)

| # | Change | Status | Notes |
|---|--------|--------|-------|
| 1 | `progress_dialog.py`: copy button moved to its own row right-aligned with the text area (same pattern as settings gear). Tooltip added: "Copy the whole contents of the Processing Log to the clipboard". Add `_Tooltip` class inline. | DONE | |
| 2 | `progress_dialog.py`: separator — remove blank lines before dashes; colour dashes green (reuse `"matched"` tag). | DONE | |
| 3 | `main.py`: fix `[[Complete]]` double brackets — was passing `"[Complete]"` as the `file` arg; `append_entry` adds its own brackets. Changed to `"Complete"`. | DONE | |
| 4 | `version.py`: bump to `1.0.56`. | DONE | |

### v1.0.55 — File dialogs remember last file location; folder picks don't update it (2026-07-30)

| # | Change | Status | Notes |
|---|--------|--------|-------|
| 1 | `file_upload_ui.py`: add `_last_dir`, `_get_initial_dir()`, `_set_last_dir()`. File pickers use `_get_initial_dir()` and call `_set_last_dir()` on success. Directory picker uses `_get_initial_dir()` but does NOT update `_last_dir`. | DONE | |
| 2 | `known_exception_builder.py`: remove local `_exe_dir()`; import `_get_initial_dir` and `_set_last_dir` from `file_upload_ui`. Both file pickers call `_set_last_dir()` on success. | DONE | |
| 3 | `version.py`: bump to `1.0.55`. | DONE | |

### v1.0.54 — Update fixture UAT test plan for v1.0.52/53 changes (2026-07-30)

| # | Change | Status | Notes |
|---|--------|--------|-------|
| 1 | `docs/generate_fixture_uat.py`: FX-05 row count 26 → 30 (4 new diff rows); add XC_DIFF_YELLOW, XC_DIFF_GREEN, XC_DIFF_ORANGE, XC_DIFF_NO_TOC to row table. FX-10 rewritten to describe both .txt output (3 entries) and Comparison filtering; document colour semantics (yellow=Changed in scope, green=New in scope, orange=Removed excluded). FX-25 XC row count 26 → 30. | DONE | |
| 2 | `docs/20260730 Fixture_UAT_v1.0.53 Test Plan.xlsx`: generated output. | DONE | |
| 3 | `version.py`: bump to `1.0.54`. | DONE | |

### v1.0.53 — File dialogs open in EXE folder by default (2026-07-30)

| # | Change | Status | Notes |
|---|--------|--------|-------|
| 1 | `file_upload_ui.py`: add `_exe_dir()` helper; pass `initialdir=_exe_dir()` to `askopenfilename` and `askdirectory`. | DONE | |
| 2 | `known_exception_builder.py`: same — add `_exe_dir()` and pass to `asksaveasfilename` and `askopenfilename`. | DONE | |
| 3 | `version.py`: bump to `1.0.53`. | DONE | |

### v1.0.52 — Fix "Process only differences" filtering across all four strategies (2026-07-30)

| # | Change | Status | Notes |
|---|--------|--------|-------|
| 1 | `strategies/x_checks/x_check_no_selection.py`: step 2 now excludes rows where Type of change = "Removed" (orange) in addition to blank. | DONE | |
| 2 | `strategies/x_checks/x_checks.py`: `_write_x_check_no_list` now returns the in-scope list; `process()` filters `df_comparison` to that list when diff=True. | DONE | |
| 3 | `strategies/grouping_by/grouping_by.py`: add `_diff_in_scope_xchecks()` which reads Grouping By cell fills via openpyxl; filter Comparison when diff=True to rows whose X-Check part is in scope. | DONE | |
| 4 | `strategies/accounting_principles/accounting_principles.py`: extend `_select_in_scope_x_checks()` with an event-column colour pass; only X-Checks where at least one event column is yellow or green are in scope for diff=True. | DONE | Conditions was already correct — no change needed. |
| 5 | `test_data/generate_test_fixtures.py`: add XC_DIFF_YELLOW (Changed/yellow), XC_DIFF_GREEN (New/green), XC_DIFF_ORANGE (Removed/orange — excluded), XC_DIFF_NO_TOC (blank — excluded) rows with correct fills. | DONE | |
| 6 | `test_data/generate_gb_fixtures.py` + `generate_ap_fixtures.py`: add colour rows for diff=True testing. | DONE | |
| 7 | All three logic test suites updated and passing: XC 71/71, GB 22/22, AP 26/26. | DONE | |
| 8 | `version.py`: bump to `1.0.52`. | DONE | |

### v1.0.51 — Fix datetime import; add Copy Log button with green tick (2026-07-30)

| # | Change | Status | Notes |
|---|--------|--------|-------|
| 1 | `main.py`: move `from datetime import datetime` to module-level import (was only inside `_run_debug`), causing `NameError: name 'datetime' is not defined` on the regular run path. Remove duplicate local import from `_run_debug`. | DONE | |
| 2 | `progress_dialog.py`: add Copy Log button (📋) to button row. On click, copies full log text to clipboard and shows ✅ for 3 seconds before reverting to 📋. | DONE | |
| 3 | `version.py`: bump to `1.0.51`. | DONE | |

### v1.0.50 — Logic test fixtures and test suites for GB, AP, and Conditions (2026-07-30)

| # | Change | Status | Notes |
|---|--------|--------|-------|
| 1 | `test_data/generate_gb_fixtures.py`: generates `test_data/fixtures/gb/` — 11-row EBX pub, FIP VALFLDGR, mapping file, KEL. Covers: standard match/not-in-FIP, Reference X-Check override, multi-value Grouping By, deduplication, mapping-to-ignore, unmapped field, blank ValidRule, KEL annotation, KEL wrong fingerprint. | DONE | |
| 2 | `test_data/generate_ap_fixtures.py`: generates `test_data/fixtures/ap/` — EBX pub, FIP VALMSG, KEL. Covers: Warning match/mismatch, Both severity (w and e), grey binding fallback, unknown method skipped, blank event col skipped, raw ZQ9_VALMSG Key construction, all 4 selection filters, KEL annotation. Uses real validation_methods.xlsx. | DONE | |
| 3 | `test_data/generate_cond_fixtures.py`: generates `test_data/fixtures/cond/` — EBX pub, FIP VALMETH, KEL. Covers: all 5 CONDITION_COLS, Reference X-Check override, Not Matched, multi-col, differences mode (yellow+green+white), FIP column renaming, KEL annotation. Documents that Not Matched rows cannot be KEL-annotated (blank FIP Data). | DONE | |
| 4 | `test_data/run_gb_logic_tests.py`: 15 assertions — all pass. | DONE | |
| 5 | `test_data/run_ap_logic_tests.py`: 22 assertions — all pass. | DONE | |
| 6 | `test_data/run_cond_logic_tests.py`: 26 assertions — all pass. | DONE | |
| 7 | `version.py`: bump to `1.0.50`. | DONE | |

### v1.0.49 — Move version label down by its own height (2026-07-30)

| # | Change | Status | Notes |
|---|--------|--------|-------|
| 1 | `main.py` + `file_upload_ui.py`: version label top padding increased to 14px (approx one label height) to move it down. | DONE | |
| 2 | `version.py`: bump to `1.0.49`. | DONE | |

### v1.0.48 — Version label in grey at bottom-left of every screen (2026-07-30)

| # | Change | Status | Notes |
|---|--------|--------|-------|
| 1 | `main.py`: add grey `v{VERSION}` label at bottom-left of task selector frame. | DONE | |
| 2 | `file_upload_ui.py`: add grey `v{VERSION}` label at bottom-left of every file upload form, below the Proceed button. | DONE | |
| 3 | `version.py`: bump to `1.0.48`. | DONE | |

### v1.0.47 — Regenerate both UAT test plans at v1.0.46 (2026-07-30)

| # | Change | Status | Notes |
|---|--------|--------|-------|
| 1 | `docs/20260730 Fixture_UAT_v1.0.46 Test Plan.xlsx`: regenerated. | DONE | |
| 2 | `docs/20260730 Full_Application_v1.0.46 Test Plan.xlsx`: regenerated. | DONE | |
| 3 | `version.py`: bump to `1.0.47`. | DONE | |

### v1.0.46 — Fix missing return True in X-Checks and Collect Live X-Checks (2026-07-30)

| # | Change | Status | Notes |
|---|--------|--------|-------|
| 1 | `strategies/x_checks/x_checks.py`: add `return True` at end of `process()`. Missing return caused `execute()` to receive `None` → treated as failure → completion line never emitted. | DONE | |
| 2 | `strategies/collect_live_x_checks/collect_live_x_checks.py`: same fix — add `return True` at end of `process()`. | DONE | |
| 3 | `version.py`: bump to `1.0.46`. | DONE | |

### v1.0.45 — Enclose completion timestamp in square brackets (2026-07-30)

| # | Change | Status | Notes |
|---|--------|--------|-------|
| 1 | `main.py`: completion message timestamp format changed from `%Y%m%d %H%M%S` to `[%Y%m%d %H%M%S]`. | DONE | |
| 2 | `version.py`: bump to `1.0.45`. | DONE | |

### v1.0.44 — Green completion message with timestamp on all strategies (2026-07-30)

| # | Change | Status | Notes |
|---|--------|--------|-------|
| 1 | `main.py`: both `run_processing` completion lines now emit `[Complete]  {yyyymmdd hhmmss}  Processing has completed successfully` instead of the plain "Processing complete" text. | DONE | |
| 2 | `progress_dialog.py`: `_MATCHED_KEYWORDS` updated from `"processing complete"` to `"completed successfully"` to match the new wording and colour the line green. | DONE | |
| 3 | `version.py`: bump to `1.0.44`. | DONE | |

### v1.0.43 — Update fixture UAT test plan for 26-row X-Checks (2026-07-30)

| # | Change | Status | Notes |
|---|--------|--------|-------|
| 1 | `docs/generate_fixture_uat.py`: FX-05 updated to list all 26 X-Check fixture rows with descriptions; FX-06 expanded to four column-level checks (Variables Match, operator-only mismatch, excl match, excl mismatch); FX-08 updated for new per-row colour coding; FX-09 updated for 2-entry KEL (XC_KEL_MISMATCH annotated, XC_KEL_NO_MATCH not annotated); FX-10 updated for INACTIVE and yellow-category exclusions; FX-25 row count updated to 26. | DONE | |
| 2 | `docs/20260730 Fixture_UAT_v1.0.42 Test Plan.xlsx`: generated output. | DONE | |
| 3 | `version.py`: bump to `1.0.43`. | DONE | |

### v1.0.42 — X-Checks output: per-row colour coding + MisMatch (Excepted) (2026-07-30)

| # | Change | Status | Notes |
|---|--------|--------|-------|
| 1 | `strategies/x_checks/x_checks.py`: `apply_output_formatting` rewritten as a single per-row pass. Rules: all Match → X-Check No. green; any MisMatch → red (beats Not Found); any Not Found (no MisMatch) → orange; any MisMatch with valid KEL annotation → rewrite cell to "MisMatch (Excepted)", colour blue; X-Check No. blue when all bad columns are excepted. Removes the old conditional-formatting approach for comparison columns. | DONE | |
| 2 | `version.py`: bump to `1.0.42`. | DONE | |

### v1.0.41 — Comprehensive X-Checks fixtures + KEL fingerprint no-match test (2026-07-30)

| # | Change | Status | Notes |
|---|--------|--------|-------|
| 1 | `test_data/generate_test_fixtures.py`: expanded X-Checks fixtures from 11 to 26 rows covering every transformation path: ABS wrapping, LC_YTD/CONST_LC, percentage format, ff suffix, subtraction, non-zero limit, >= operator, excl.acc.type (match and mismatch), REX→ToM correction, INACTIVE filter, yellow Category filter, XC_ALL_MISMATCH (both formula and vars wrong), XC_FORMULA_MISMATCH (operator differs, vars same), XC_KEL_NO_MATCH (KEL entry with wrong fingerprint → no annotation). Also adds @2A@ excl account type lines to FIP block helper. | DONE | |
| 2 | `test_data/run_logic_tests.py`: updated to 58 assertions covering all 26 X-Check rows plus column-level checks (Variables Match, Formula Match (Excl), KEL no-match). All 58 pass. | DONE | |
| 3 | `version.py`: bump to `1.0.41`. | DONE | |

### v1.0.40 — X-Check No. cell turns blue when row has a Known Exception (2026-07-30)

| # | Change | Status | Notes |
|---|--------|--------|-------|
| 1 | `strategies/x_checks/x_checks.py`: `apply_output_formatting` now also colours the `X-Check No.` cell blue on any row where the `Known Exception` column is populated. Both cells are coloured in the same pass. | DONE | |
| 2 | `version.py`: bump to `1.0.40`. | DONE | |

### v1.0.39 — Add XC_KEL_MISMATCH fixture and KEL logic test (2026-07-30)

| # | Change | Status | Notes |
|---|--------|--------|-------|
| 1 | `test_data/generate_test_fixtures.py`: add `XC_KEL_MISMATCH` EBX row and FIP block (formula mismatch using ACC999); add `_make_known_exception_list()` which runs the comparison after other fixtures are written to extract exact fingerprint values, then builds `known_exception_list.xlsx` with one X-Checks entry keyed to XC_KEL_MISMATCH. | DONE | |
| 2 | `test_data/run_logic_tests.py`: add FX-09a/b/c — verify XC_KEL_MISMATCH stays MisMatch, gets reason populated, and non-mismatch rows have blank Known Exception column. All 38 assertions pass. | DONE | |
| 3 | `docs/generate_fixture_uat.py`: update FX-09 to reference `known_exception_list.xlsx` fixture; update FX-05 and FX-25 row counts to 11. | DONE | |
| 4 | `docs/20260730 Fixture_UAT_v1.0.38 Test Plan.xlsx`: generated output. | DONE | |
| 5 | `version.py`: bump to `1.0.39`. | DONE | |

### v1.0.38 — Logic tests 34/34 pass; fix X-Checks row count to 10 (2026-07-28)

| # | Change | Status | Notes |
|---|--------|--------|-------|
| 1 | `test_data/run_logic_tests.py`: new script running all 12 Logic test cases (FX-05/06/10/12/16/20/21/25) against the fixture files. All 34 assertions pass. | DONE | |
| 2 | `test_data/run_logic_tests.py` + `docs/generate_fixture_uat.py`: correct X-Checks expected row count from 9 → 10. XC_DIFF_EXCLUDED has an Account No so extract_ebx produces a formula; with no matching FIP block it generates a Not Found row in the Comparison. The X-Check No Selection filter (differences mode) correctly excludes it from the .txt output — these are independent. | DONE | |
| 3 | `docs/20260728 Fixture_UAT_v1.0.37 Test Plan.xlsx`: regenerated with corrected row counts. | DONE | |
| 4 | `version.py`: bump to `1.0.38`. | DONE | |

### v1.0.37 — Add green-cell fixture row for Conditions differences mode (2026-07-28)

| # | Change | Status | Notes |
|---|--------|--------|-------|
| 1 | `test_data/generate_test_fixtures.py`: add `COND_DIFF_GREEN` row with green fill on the Applicable Quarters condition cell; add matching FIP entry. Fix: green rule checks the condition cell itself, not column A. | DONE | |
| 2 | `docs/generate_fixture_uat.py`: update FX-21 to expect both yellow and green rows (2 rows) in differences mode. | DONE | |
| 3 | `docs/20260728 Fixture_UAT_v1.0.36 Test Plan.xlsx`: generated output. | DONE | |
| 4 | `version.py`: bump to `1.0.37`. | DONE | |

### v1.0.36 — Add differences-mode test coverage to fixtures and UAT plan (2026-07-28)

| # | Change | Status | Notes |
|---|--------|--------|-------|
| 1 | `test_data/generate_test_fixtures.py`: add four new rows to `xc_pub.xlsx`: `XC_DIFF_IN_SCOPE` (Status=ACTIVE, Type of change=Changed → survives X-Check No Selection), `XC_DIFF_EXCLUDED` (same but Exclude Z-Core=X → filtered out), `COND_DIFF_YELLOW` (Applicable Quarters cell has yellow openpyxl fill → collected in differences mode), `COND_DIFF_WHITE` (plain white → not collected). Add matching FIP entries. | DONE | |
| 2 | `docs/generate_fixture_uat.py`: add FX-10 (X-Check No Selection differences mode) and FX-21 (Conditions differences mode with yellow cell); renumber subsequent cases; update row counts in FX-05, FX-20, FX-25. | DONE | |
| 3 | `docs/20260728 Fixture_UAT_v1.0.35 Test Plan.xlsx`: generated output (37 cases). | DONE | |
| 4 | `version.py`: bump to `1.0.36`. | DONE | |

### v1.0.35 — Fixture UAT plan expanded to standalone (2026-07-28)

| # | Change | Status | Notes |
|---|--------|--------|-------|
| 1 | `docs/generate_fixture_uat.py`: rewritten as a 36-case standalone plan covering all strategies, UI, output structure, colour coding, sensitivity labelling, and error handling. Each case carries a 'Test Type' column: Logic (10 cases, light blue rows — run after every code change) or Whole App (26 cases — run for release sign-off). | DONE | |
| 2 | `docs/20260728 Fixture_UAT_v1.0.34 Test Plan.xlsx`: generated output. | DONE | |
| 3 | `version.py`: bump to `1.0.35`. | DONE | |

### v1.0.34 — Rename production EXE to X-Checks_FullRun (2026-07-28)

| # | Change | Status | Notes |
|---|--------|--------|-------|
| 1 | `build.py`: production build `name` changed from `X-Checks_{VERSION}` to `X-Checks_FullRun_{VERSION}`. Output: `dist\X-Checks_FullRun_v{VERSION}.exe`. | DONE | |
| 2 | `version.py`: bump to `1.0.34`. | DONE | |

### v1.0.33 — Fixture-based UAT test plan (2026-07-28)

| # | Change | Status | Notes |
|---|--------|--------|-------|
| 1 | `docs/generate_fixture_uat.py`: new generator producing an 8-case UAT plan against the minimal fixture files. Each test case lists the exact fixture files, expected Comparison sheet rows with X-Check ID → result mapping, and colour-coding verification. | DONE | |
| 2 | `docs/20260728 Fixture_UAT_v1.0.32 Test Plan.xlsx`: generated output. | DONE | |
| 3 | `version.py`: bump to `1.0.33`. | DONE | |

### v1.0.32 — Update UAT test plan to v1.0.31 (2026-07-28)

| # | Change | Status | Notes |
|---|--------|--------|-------|
| 1 | `docs/generate_full_app_uat.py`: updated AP FIP file reference to `20260728_VALMSG_File_direct_from_FIP.XLSX` (raw ZQ9_VALMSG, 17,184 rows); updated FA-15 to note app builds Key column from MK + ValidRule; updated FA-16 output sheet list (EBX, FIP, Comparison, Processing Log); added FA-26–28 covering Settings gear menu, Known Exception Builder dialog, and KEL build-and-open. Former FA-26–32 renumbered to FA-29–35. | DONE | |
| 2 | `docs/20260728 Full_Application_v1.0.31 Test Plan.xlsx`: generated output (35 test cases). | DONE | |
| 3 | `version.py`: bump to `1.0.32`. | DONE | |

### v1.0.31 — Minimal test fixtures for all four strategies (2026-07-28)

| # | Change | Status | Notes |
|---|--------|--------|-------|
| 1 | `test_data/generate_test_fixtures.py`: new generator script producing 7 minimal fixture files under `test_data/fixtures/`. Each X-Check ID encodes its expected output, covering every comparison outcome per strategy (Match/MisMatch/Not Found for X-Checks; Matched/Not in FIP for GB; Match/MisMatch for AP; Matched/Not Matched for Conditions). Also covers TOM→ToM correction, period-thousands correction, and the Reference X-Check (Condition) key-prefix override. | DONE | XC_REORDER_MATCH produces Formula Match=MisMatch due to a known edge case in the reorder logic (formula produced by string replacement is invalid for simple two-variable addition); documented in the generator. |
| 2 | `test_data/fixtures/`: 7 generated fixture files committed. | DONE | |
| 3 | `version.py`: bump to `1.0.31`. | DONE | |

### v1.0.30 — AP strategy: build Key column from raw ZQ9_VALMSG FIP export (2026-07-28)

| # | Change | Status | Notes |
|---|--------|--------|-------|
| 1 | `strategies/accounting_principles/accounting_principles.py`: after loading `fip_df`, if no `Key` column is present, build it from `MK + '|' + ValidRule`. Raw ZQ9_VALMSG exports have `MK` (validation method code, e.g. `V900A`) and `ValidRule` (X-Check No., e.g. `A001_09`) but no pre-built `Key`. Files that already have a `Key` column pass through unchanged. If neither column is present, log an error and abort. | DONE | |
| 2 | `version.py`: bump to `1.0.30`. | DONE | |

### v1.0.29 — Apply Internal Use Only label to built Known Exception List (2026-07-21)

| # | Change | Status | Notes |
|---|--------|--------|-------|
| 1 | `known_exception_builder.py`: after file is verified on disk, apply the `Internal_Use_Only` MIP sensitivity label via `ExcelLabeler`. Best-effort — failure is silently swallowed so the dialog still closes and the file is still opened. | DONE | |
| 2 | `version.py`: bump to `1.0.29`. | DONE | |

### v1.0.28 — Centre main window on screen at launch (2026-07-21)

| # | Change | Status | Notes |
|---|--------|--------|-------|
| 1 | `main.py`: add `_centre_main_window()` — centres the task selector on screen at launch, clamped 60px from every edge. Called after `update_idletasks()` so the window size is known. | DONE | |
| 2 | `version.py`: bump to `1.0.28`. | DONE | |

### v1.0.27 — Builder dialog: clamp position to screen (2026-07-21)

| # | Change | Status | Notes |
|---|--------|--------|-------|
| 1 | `known_exception_builder.py`: `_centre()` now clamps the dialog position to a 40px margin from every screen edge, preventing it from appearing off-screen when the parent window is near the top-left corner. | DONE | |
| 2 | `version.py`: bump to `1.0.27`. | DONE | |

### v1.0.26 — Builder "Save as" hint text (2026-07-21)

| # | Change | Status | Notes |
|---|--------|--------|-------|
| 1 | `known_exception_builder.py`: replace "Not set" placeholder with "Click Browse and select a folder, then type the filename". | DONE | |
| 2 | `version.py`: bump to `1.0.26`. | DONE | |

### v1.0.25 — Settings button opens popup menu (2026-07-21)

| # | Change | Status | Notes |
|---|--------|--------|-------|
| 1 | `main.py`: `_on_settings()` now opens a `tk.Menu` popup below the ⚙ button rather than launching the builder directly. "Build Known Exception List…" is the first entry; future settings items can be added as `menu.add_command()` calls. Builder moved to `_open_known_exception_builder()`. | DONE | |
| 2 | `version.py`: bump to `1.0.25`. | DONE | |

### v1.0.24 — Known Exception List Builder refinements (2026-07-21)

| # | Change | Status | Notes |
|---|--------|--------|-------|
| 1 | `main.py`: Start button restored to `columnspan=2` (centred); ⚙ moved to row 3 `column=1 sticky=e` (below and right). | DONE | |
| 2 | `known_exception_builder.py`: add "Open file after building" checkbox (on by default); verify file exists on disk after `wb.save()` before closing dialog; open file via `start` shell command if checkbox checked. | DONE | |
| 3 | `version.py`: bump to `1.0.24`. | DONE | |

### v1.0.23 — Known Exception List Builder (2026-07-21)

| # | Change | Status | Notes |
|---|--------|--------|-------|
| 1 | `known_exception_builder.py`: new module with `KnownExceptionBuilderDialog` — modal tkinter dialog with output file picker, optional comparison output import, strategy detection, and openpyxl Excel builder. Creates one sheet per strategy with fingerprint-column headers derived from `STRATEGY_DEFINITIONS`, a guidance row (row 2, always skipped by `_load_known_exceptions`), and optionally pre-fills mismatch rows from a comparison output. | DONE | |
| 2 | `main.py`: restructure button row — Start at `column=0 sticky=ew`, ⚙ at `column=1 sticky=e`. Add `_on_settings()` which opens `KnownExceptionBuilderDialog`. | DONE | |
| 3 | `build.py`: add `"known_exception_builder"` to `hidden_imports`. | DONE | |
| 4 | `version.py`: bump to `1.0.23`. | DONE | |

### v1.0.22 — Full Application UAT test plan — complete with all test files (2026-07-17)

| # | Change | Status | Notes |
|---|--------|--------|-------|
| 1 | `docs/generate_full_app_uat.py`: updated with all 11 test files now confirmed present in `test_data\` (added `validation methods.xlsx`, `20260602 VALMSG (Accounting Principle)_Original.XLSX`, `20260602 VALMETH (Conditions).xlsx`). Correct sheet names throughout. AP FIP 17 072 rows, Conditions FIP 4 816 rows, differences 92 pairs, full 130 pairs. Conditions test cases expanded to FA-18/FA-19 (diff + full runs). All case IDs renumbered sequentially to FA-32. | DONE | |
| 2 | `docs/20260717 Full_Application_v1.0.21 Test Plan.xlsx`: generated output. | DONE | |
| 3 | `version.py`: bump to `1.0.22`. | DONE | |

### v1.0.21 — Full Application UAT test plan (2026-07-17)

| # | Change | Status | Notes |
|---|--------|--------|-------|
| 1 | `docs/generate_full_app_uat.py`: new generator producing a 30-case UAT test plan covering all 6 tasks, Processing Log entries (output path + expected sensitivity label), colour coding, Full Run combined output, and error/stop handling. All file references and counts derived from `test_data\` files: XC pub 664 extracted / 653 comparison rows (525 formula match, 125 mismatch, 3 not found); GB FIP 12 348 rows / 6 776 processed / 19 mapping entries; AP pub 3 345 rows; Known Exceptions 2 rows. | DONE | |
| 2 | `docs/20260717 Full_Application_v1.0.20 Test Plan.xlsx`: generated output. | DONE | |
| 3 | `version.py`: bump to `1.0.21`. | DONE | |

### v1.0.20 — Port Conditions fixes from v0.6.11 (2026-07-17)

| # | Change | Status | Notes |
|---|--------|--------|-------|
| 1 | `strategies/conditions/fip.py`: rename `CONCAT_COL` `"Concatenated"` → `"Key (Concatenated)"`. | DONE | |
| 2 | `strategies/conditions/compare.py`: import `CONCAT_COL` from `fip.py`; replace hardcoded `"Concatenated"` string with the constant. | DONE | Hardcoded string would have caused a `KeyError` after the rename. |
| 3 | `strategies/conditions/extract.py`: when building concat keys, resolve `effective_xc` — use the row's `"Reference  X-Check (Condition)"` value if non-blank, otherwise fall back to `"X-Check No."`. | DONE | |
| 4 | `strategies/base_strategy.py`: move `df_log` snapshot and `Processing Log` sheet write to end of `with` block, after all `log_step` calls; add "Output written to" inside the block; add expected sensitivity label directly via `log.append()` (Excel log only, not dialog). | DONE | |
| 5 | `version.py`: bump to `1.0.20`. | DONE | |

### v1.0.19 — Progress dialog: only final completion message is green (2026-07-02)

| # | Change | Status | Notes |
|---|--------|--------|-------|
| 1 | `progress_dialog.py`: `_MATCHED_KEYWORDS` reduced to `("processing complete",)` — intermediate lines containing "complete", "successfully", "matched" etc. were incorrectly turning green. Only the final completion message is green. | DONE | |
| 2 | `version.py`: bump to `1.0.19`. | DONE | |

### v1.0.18 — Architecture refactor: centralise colours, exception annotation, Full Run formatting (2026-07-02)

| # | Change | Status | Notes |
|---|--------|--------|-------|
| 1 | `base_strategy.py`: add shared colour constants (`FILL_GREEN`, `FONT_GREEN`, `FILL_RED`, `FONT_RED`, `FILL_ORANGE`, `FONT_ORANGE`, `FILL_BLUE`, `FONT_BLUE`) as class-level attributes. All strategies inherit. | DONE | |
| 2 | `base_strategy.py`: add `_annotate_known_exceptions(df, exc_path, sheet_name, fingerprint_columns)` — encapsulates the full load→key-build→annotate→log pattern. Returns annotated df or `False` on load error. | DONE | |
| 3 | `x_checks.py`, `grouping_by.py`, `accounting_principles.py`, `conditions.py`: replace ~15-line Known Exception boilerplate blocks with single `_annotate_known_exceptions()` call. Remove all local `PatternFill`/`Font` definitions; reference base constants. | DONE | |
| 4 | `grouping_by.py`: add missing `apply_output_formatting()` (Result column: green/orange). | DONE | |
| 5 | `full_run.py`: add `_PrefixedWorkbook` shim that maps unprefixed sheet names to prefixed ones. `apply_output_formatting` now delegates to each strategy instance via the shim — no per-strategy column knowledge in Full Run. Remove all `PatternFill`/`Font` definitions from Full Run. | DONE | |
| 6 | `CLAUDE.md`: add Code Reuse Rule section above Change Log Policy. | DONE | |
| 7 | `version.py`: bump to `1.0.18`. | DONE | |

### v1.0.17 — Conditions Comparison column: string values + colour coding (2026-07-02)

| # | Change | Status | Notes |
|---|--------|--------|-------|
| 1 | `conditions/compare.py`: change `Comparison` column from boolean `True`/`False` to strings `"Matched"`/`"Not Matched"`. Update summary keys accordingly. | DONE | Booleans can't be matched by CellIsRule string formulas. |
| 2 | `conditions.py`: update log step to use new summary keys; add `apply_output_formatting` with green/red conditional formatting on the `Comparison` column. | DONE | |
| 3 | `full_run.py`: update Cond formatter to use `"Matched"`/`"Not Matched"` strings. | DONE | |
| 4 | `version.py`: bump to `1.0.17`. | DONE | |

### v1.0.16 — Fix comparison column colour coding in Full Run output (2026-07-02)

| # | Change | Status | Notes |
|---|--------|--------|-------|
| 1 | `full_run.py`: `apply_output_formatting` now applies match/mismatch colour coding to all `"— Comparison"` sheets in the combined workbook. Sheet names are prefixed so exact-name checks in individual strategies don't fire. | DONE | |
| 2 | `version.py`: bump to `1.0.16`. | DONE | |

### v1.0.15 — Comparison sheet named "Comparison" and always last in all strategies (2026-07-02)

| # | Change | Status | Notes |
|---|--------|--------|-------|
| 1 | `conditions.py`: rename `"Conditions"` sheet → `"Comparison"`, move it last (after Working Sheet, FIP Data). Update summaries key. | DONE | |
| 2 | `x_checks.py`: rename `"X-Checks Comparison"` → `"Comparison"` (sheets dict + apply_output_formatting). | DONE | |
| 3 | Grouping By and Accounting Principles already use `"Comparison"` as the last sheet — no change needed. | DONE | |
| 4 | `version.py`: bump to `1.0.15`. | DONE | |

### v1.0.14 — Full Run debug EXE bundles only Q2 X-Checks Data folder (2026-07-02)

| # | Change | Status | Notes |
|---|--------|--------|-------|
| 1 | `build.py`: add `test_data_subdir` option — when set, only that subfolder of `test_data/` is bundled instead of the whole folder. `full_run_debug` uses `"Q2 X-Checks Data"`. | DONE | |
| 2 | `main.py`: all `_DEBUG_FILES_*` dicts updated to use files from `test_data/Q2 X-Checks Data/` with correct sheet names. | DONE | |
| 3 | `version.py`: bump to `1.0.14`. | DONE | |

### v1.0.13 — Remove FIP Key column guard from Conditions (2026-07-02)

| # | Change | Status | Notes |
|---|--------|--------|-------|
| 1 | `main.py`: all `_DEBUG_FILES_*` dicts updated to use files from `test_data/Q2 X-Checks Data/`. Correct sheet names set for each file. | DONE | |
| 2 | `version.py`: bump to `1.0.14`. | DONE | |

### v1.0.13 — Remove FIP Key column guard from Conditions (2026-07-02)

| # | Change | Status | Notes |
|---|--------|--------|-------|
| 1 | `conditions.py`: remove the `"Key"` first-column guard — it was causing Conditions to abort even with correct ZQ9_VALMETH files. `process_fip()` already raises a clear ValueError if the file has the wrong structure. | DONE | |
| 2 | `version.py`: bump to `1.0.13`. | DONE | |

### v1.0.12 — Progress dialog: colour coding, timestamp format, completion separator (2026-07-01)

| # | Change | Status | Notes |
|---|--------|--------|-------|
| 1 | `progress_dialog.py`: restore colour coding — green for matched/complete/success lines, orange for not-found/not-matched/mismatch lines, red (bold) for errors. | DONE | |
| 2 | `progress_dialog.py` + `base_strategy.py`: reorder line format to `[File]  [yyyymmdd hhmmss]  step  (count)`. | DONE | |
| 3 | `progress_dialog.py` + `main.py`: add `append_separator()` — writes two blank lines then a `----` separator line before "Processing complete". | DONE | |
| 4 | `version.py`: bump to `1.0.12`. | DONE | |

### v1.0.11 — Full Run: abort on strategy failure; exclude Collect Live X-Checks (2026-07-01)

| # | Change | Status | Notes |
|---|--------|--------|-------|
| 1 | `full_run.py`: check return value of `strategy.process()` — if `False`, log an abort message and return `False` immediately. Previously ignored. | DONE | |
| 2 | `full_run.py`: exceptions inside a sub-strategy now also abort Full Run rather than continuing. | DONE | |
| 3 | `full_run.py`: exclude `"Collect Live X-Checks"` from Full Run iteration (alongside `"Full Run"` itself) — it produces no output sheets and should not run as part of a Full Run. | DONE | |
| 4 | `version.py`: bump to `1.0.11`. | DONE | |

### v1.0.10 — Add Collect Live X-Checks strategy to dropdown (2026-07-01)

| # | Change | Status | Notes |
|---|--------|--------|-------|
| 1 | `strategies/collect_live_x_checks/` — ported from `v0.4-X-Check-No-Selection`. Runs the X-Check selection pipeline (Status / Type of Change / Exclude Z-Core / yellow Category), writes a `.txt` file of in-scope X-Check Nos, and copies the list to the Windows clipboard. | DONE | |
| 2 | `task_configs.py`: add `COLLECT_LIVE_X_CHECKS_UPLOAD_CONFIG` (X-Checks Publication File only). | DONE | |
| 3 | `task_registry.py`: register `"Collect Live X-Checks"` as the first entry — appears at the top of the dropdown above X-Checks. | DONE | |
| 4 | `main.py`: add `_DEBUG_FILES_COLLECT_LIVE_X_CHECKS` debug dict. | DONE | |
| 5 | `build.py`: add `collect_debug` build entry + hidden imports. | DONE | |
| 6 | `version.py`: bump to `1.0.10`. | DONE | |

### v1.0.9 — Reject FIP files with Key column; fix typo; remove Layout B (2026-07-01)

| # | Change | Status | Notes |
|---|--------|--------|-------|
| 1 | `conditions.py`: add guard — if the FIP file's first column is `"Key"`, log an error and abort. Users must upload the raw ZQ9_VALMETH extract, not a pre-processed file. | DONE | |
| 2 | `conditions/fip.py`: remove Layout B (pre-processed Key column path) entirely — the raw 8-column layout is the only supported input. | DONE | |
| 3 | `conditions.py`: fix typo `process_only_differenceserences` → `process_only_differences`. | DONE | |
| 4 | `version.py`: bump to `1.0.9`. | DONE | |

### v1.0.8 — Progress dialog: timestamp on each line; fix false-positive red colouring (2026-07-01)

| # | Change | Status | Notes |
|---|--------|--------|-------|
| 1 | `progress_dialog.py`: removed `"exception"` from `_ERROR_KEYWORDS` — it matched `"Known Exception List"` causing informational lines to appear red. Lines that reference the exception list are not errors. | DONE | |
| 2 | `progress_dialog.py` + `base_strategy.py`: timestamp `[yyyymmdd hhmmss]` appended after the step description on each progress dialog line. | DONE | |
| 3 | `version.py`: bump to `1.0.8`. | DONE | |

### v1.0.7 — Sheet name field: combobox populated from workbook (2026-07-01)

| # | Change | Status | Notes |
|---|--------|--------|-------|
| 1 | `file_upload_ui.py`: replace sheet name `ttk.Entry` with `ttk.Combobox`. On file select, reads all sheet names from the workbook, selects the configured default if present otherwise the first sheet, and populates the dropdown with all available sheets. Prefill path also populates the combobox. | DONE | |
| 2 | `version.py`: bump to `1.0.7`. | DONE | |

### v1.0.6 — Fix process_only_differences coupling in strategies (2026-07-01)

| # | Change | Status | Notes |
|---|--------|--------|-------|
| 1 | `x_checks.py`, `accounting_principles.py`: replace `self.process_only_differences` with `files.get("process_only_differences", False)`. Strategies now read directly from the `files` dict and are no longer coupled to `execute()` having been called first. | DONE | |
| 2 | `full_run.py`: remove workaround that manually set `strategy.process_only_differences` before calling `process()` — no longer needed. | DONE | |
| 3 | `version.py`: bump to `1.0.6`. | DONE | |

### v1.0.5 — Full Run bug fixes (2026-07-01)

| # | Change | Status | Notes |
|---|--------|--------|-------|
| 1 | `full_run.py`: set `strategy.process_only_differences` before calling `strategy.process()`. Full Run bypasses `execute()` which normally sets this attribute, causing `AttributeError` on X-Checks and Accounting Principles. | DONE | |
| 2 | `grouping_by.py`: cast `"Reference  X-Check (Condition)"` column to string before `.str.strip()` to handle NaN/float values that caused `AttributeError: Can only use .str accessor with string values`. | DONE | |
| 3 | `version.py`: bump to `1.0.5`. | DONE | |

### v1.0.4 — Mapping File accepts .txt extension (2026-07-01)

| # | Change | Status | Notes |
|---|--------|--------|-------|
| 1 | `task_configs.py`: Mapping File `file_types` updated to `("CSV / Text Files", "*.csv *.txt")` so both extensions are selectable in the file picker. | DONE | |
| 2 | `version.py`: bump to `1.0.4`. | DONE | |

### v1.0.3 — Fix missing FIP File (ZQ9_VALMETH) in Full Run dialog (2026-07-01)

| # | Change | Status | Notes |
|---|--------|--------|-------|
| 1 | `task_configs.py`: rename Conditions FIP field label `"FIP File"` → `"FIP File (ZQ9_VALMETH)"` so it is distinct from X-Checks `"FIP File"` and not deduplicated out of the Full Run dialog. | DONE | Caused by v1.0.2 rename of X-Checks "FIP file" → "FIP File" creating a collision. |
| 2 | `conditions.py`, `main.py`, `docs/generate_uat.py`: update all references to `"FIP File"` (Conditions context) → `"FIP File (ZQ9_VALMETH)"`. | DONE | |
| 3 | `version.py`: bump to `1.0.3`. | DONE | |

### v1.0.2 — Naming consistency pass (2026-07-01)

| # | Change | Status | Notes |
|---|--------|--------|-------|
| 1 | `x_checks/compare.py`, `ebx_extraction.py`, `fip_extraction.py`, `x_checks.py`: rename column `X-Check Number` → `X-Check No.` throughout to match all other strategies. | DONE | |
| 2 | `task_configs.py`, `x_checks.py`, `main.py`, `test_data/run_new.py`, `docs/generate_uat.py`: rename field label `"FIP file"` → `"FIP File"` (capitalised, matching all other FIP file labels). | DONE | |
| 3 | `x_checks.py`, `grouping_by.py`: rename log step label `"Compare"` → `"Comparison"` to match Accounting Principles and Conditions. | DONE | |
| 4 | `grouping_by.py`: rename output sheet tab `"Compare"` → `"Comparison"` to match the consistent `"[Strategy] Comparison"` naming pattern. | DONE | |
| 5 | `conditions.py`: rename local variable `process_only_diff` → `process_only_differences` to match standard naming throughout the codebase. | DONE | |
| 6 | `version.py`: bump to `1.0.2`. | DONE | |

### v1.0.1 — Shared Known Exception List across all strategies (2026-07-01)

| # | Change | Status | Notes |
|---|--------|--------|-------|
| 1 | `base_strategy.py`: add shared `_load_known_exceptions(path, sheet_name, fingerprint_columns)`. One workbook, one sheet per strategy; exception fires only when all fingerprint column values match exactly; annotation-only (result columns unchanged). Falls back to `"Known Exceptions"` sheet name for X-Checks backwards compatibility. | DONE | |
| 2 | `x_checks.py`: replace private `_load_known_exceptions()` with call to base; change to full-fingerprint tuple key (8 columns); remove result-column rewriting — now annotation-only. | DONE | |
| 3 | `task_configs.py`: add optional `Known Exception List` field to `GROUPING_BY_UPLOAD_CONFIG`, `ACCOUNTING_PRINCIPLES_UPLOAD_CONFIG`, `CONDITIONS_UPLOAD_CONFIG`. Full Run deduplicates to one field automatically. | DONE | |
| 4 | `grouping_by.py`: add exception annotation after compare; fingerprint = `["EBX Key"]`. | DONE | |
| 5 | `accounting_principles.py`: add exception annotation after compare; fingerprint = `["X-Check No.", "Event", "Expected", "FIP", "Actual", "Method"]`. | DONE | |
| 6 | `conditions.py`: add exception annotation after compare; fingerprint = `["EBX Data", "FIP Data"]`; `Comparison` boolean unchanged. | DONE | |
| 7 | `version.py`: bump to `1.0.1`. | DONE | |

### v1.0.0 — Full application: all strategies combined (in progress 2026-06-30)

| # | Change | Status | Notes |
|---|--------|--------|-------|
| 1 | `strategies/grouping_by/` — ported from `v0.3.1-Grouping_By`. | DONE | |
| 2 | `strategies/accounting_principles/` — ported from `v0.5-Accounting-Principles`. | DONE | |
| 3 | `strategies/conditions/` — ported from `v0.6-Conditions`. | DONE | |
| 4 | `strategies/full_run/` — ported from `v0.7-Full-Run`. Dynamically runs all registered strategies and combines output into one workbook with colour-coded tabs. | DONE | |
| 5 | `strategies/base_strategy.py` — updated to v0.5 version which adds `_detect_header_row()` for dynamic Excel header detection. Used by Accounting Principles. | DONE | |
| 6 | `task_configs.py` — all four configs combined; `_build_full_run_config()` added. | DONE | |
| 7 | `task_registry.py` — all four strategies registered; Full Run last. | DONE | |
| 8 | `main.py` — all four `_DEBUG_FILES_*` dicts added to `_DEBUG_FILES_MAP`. | DONE | |
| 9 | `build.py` — BUILDS list updated with one prod entry + four debug entries; hidden_imports extended with all strategy submodules. | DONE | |
| 10 | `file_upload_config.py` — added `header_signals` field to `FileFieldConfig` (ported from v0.5); required by Accounting Principles config. | DONE | |
| 11 | `version.py` — bumped to `1.0.0`. | DONE | |
| 12 | `docs/generate_uat.py` + `docs/20260630 X-Checks_v1.0.0 Test Plan.xlsx` — functional test plan covering launch, all four strategies, error handling, and version checks. | DONE | |
| 13 | `strategies/x_checks/` — ported from `v0.4-X-Check-No-Selection`; registered in `task_configs.py`, `task_registry.py`, `main.py`, `build.py`. X-Checks now included in Full Run (colour: green). | DONE | |
| 14 | `docs/generate_uat.py` + `docs/20260630 X-Checks_v1.0.0 Test Plan.xlsx` — updated test plan to include X-Checks strategy (5 workflow cases, Files Required rows, dropdown count, Full Run sheet counts). | DONE | |
| 15 | `file_upload_ui.py`: hint labels (field descriptions) now all wrap at the same width — Pass 2 applies `max_hint_width` as `wraplength` to every hint label so all descriptions are consistent and none overflow past the Browse button. | DONE | |
| 16 | `task_configs.py`: GCoA Publication File description — removed `\n` line break, added full stop after "sheet" so it reads as a single sentence. | DONE | |
| 17 | `file_upload_ui.py`: reduced `HINT_WRAP_LENGTH` from 800 → 533 (and fallback 400 → 267) to make the dialog ~1/3 narrower. | DONE | |

### v0.4.7 — Sensitivity-label hook in BaseStrategy (completed 2026-06-19)

Ported from `v0.5-Accounting-Principles` (commit `e17698d`) so every strategy on every branch automatically applies a Microsoft Information Protection sensitivity label to the Excel workbooks it produces.

| # | Change | Status | Notes |
|---|--------|--------|-------|
| 1 | New `strategies/sensitivity.py` with `ExcelLabeler`. Wraps `Workbook.SensitivityLabel.CreateLabelInfo` + `AssignmentMethod=PRIVILEGED` + `LabelName`/`LabelId`/`SiteId` + `SetLabel` via pywin32. Caches one `Excel.Application` COM session per run. | DONE | Tenant `SITE_ID` and 7 label GUIDs imported verbatim from the user's VBA `SetLabelInfo` module. |
| 2 | `BaseStrategy.write_excel_output()` calls `_apply_sensitivity_label(path)` after each save. `DEFAULT_SENSITIVITY_LEVEL = "Internal_Use_Only"`. Failure logs `[Sensitivity] Could not apply label: <reason>` and the run continues. | DONE | Strategies can override `DEFAULT_SENSITIVITY_LEVEL` if they need a different default. |
| 3 | `BaseStrategy.execute()` `finally` block closes the cached labeler so Excel exits cleanly on success or error. | DONE | |
| 4 | `build.py`: `hidden_imports` extended with `strategies.sensitivity`, `win32com.client`, `win32com`, `pythoncom`, `pywintypes`. | DONE | PyInstaller can't statically resolve `win32com.client.DispatchEx`. |
| 5 | `tests/test_sensitivity.py`: 6 unit tests covering the level → (LabelId, LabelName) mapping, unknown-level error, and the missing-file failure path. | DONE | Excel COM is mocked; round-trip verified live on the v0.5 branch. |
| 6 | `version.py`: bump to `0.4.7`. | DONE | |

### v0.4.6 — Infrastructure refresh on main (completed 2026-06-15)

This commit ports every infrastructure improvement made on the v0.3 / v0.4 branches between v0.3.5 and v0.4.6 onto `main`, while stripping all strategy code (X-Checks, Grouping By, accounting/conditions stubs, their tests, their UAT artefacts). Strategy branches stay where they are; future strategies branch from this clean infra `main`.

| # | Change | Status | Notes |
|---|--------|--------|-------|
| 1 | App: `main.py` adopts task registry + lazy strategy loading + Zurich-font registration; debug-mode test-file dicts cleared (each strategy branch supplies its own). | DONE | Form pre-fill on cancel/error preserved. |
| 2 | Progress dialog: error-keyword detection styles error lines bold red; plays Windows critical-stop chime; permanent "Exit Application" button alongside Stop/Close/Return-to-Form; "Return to Form" callback for cancel/error paths. | DONE | v0.3.13 + v0.4.3 + v0.4.4 + v0.4.5 combined. |
| 3 | File upload: configurable checkboxes; "Process only differences" default ON; tooltips on experimental checkboxes; suppress console window. | DONE | Strategy branches consume the checkbox values via `files["..."]`. |
| 4 | `BaseStrategy`: `execute()` returns True on success; `build_output_path()` accepts optional `extension` kwarg; load-time error → `Return to Form`; `apply_conditional_formatting` scans all rows. | DONE | |
| 5 | Build: `build.py` generates a PyInstaller spec with `Splash(always_on_top=False)`; explicit `hiddenimports` mechanism (each strategy branch fills the list); per-strategy debug builds via `BUILDS` config; CLI keys `prod`/`debug`/...; splash generated from `splash_template.png` with current version stamped at build time. | DONE | |
| 6 | Branding: Zurich brand fonts (Light/Regular/Medium/Semibold/Bold + Ogg-Regular) bundled under `templates/fonts/`; brand colours applied across UI surfaces (forms, progress dialog, splash). | DONE | |
| 7 | `task_configs.py` + `task_registry.py` reduced to empty scaffolds. | DONE | |
| 8 | `version.py` set to `0.4.6` to match the last released version of the combined codebase before the split. | DONE | |
| 9 | All strategy modules removed: `strategies/x_checks/`, `strategies/accounting_principles.py`, `strategies/conditions.py`, `strategies/grouping_by.py`. All strategy tests removed: `tests/test_compare.py`, `tests/test_ebx_extraction.py`, `tests/test_fip_extraction.py`, `tests/test_integration.py`, `tests/test_load_known_exceptions.py`, `tests/test_variable_builder.py`, `tests/fixtures/`, `tests/generate_golden_fixtures.py`. | DONE | These remain on `obsolete_to_v0.4-X-Checks` / `v0.4-X-Check-No-Selection` where they belong. |

---

## Strategy-branch contract

When a new strategy branches from `main`:

1. **Pick a branch name** following the `vMAJOR.MINOR-Display-Name` pattern (e.g. `v0.5-Accounting-Principles`).
2. **Bump version** appropriately — minor for a new strategy, patch for changes within an existing one.
3. **Add the strategy module** under `strategies/<strategy_name>/` (or as a single `.py` for very small strategies). It must subclass `BaseStrategy` and implement `process(loaded_files, files)`.
4. **Add the upload config** to `task_configs.py` and **register** it in `task_registry.py` (entry + `if False:` import + `_lazy()` factory).
5. **Add a debug entry** to `BUILDS` in `build.py` if you want a `_Debug_<Name>` EXE; add a matching `_DEBUG_FILES_<NAME>` dict in `main.py` and register it in `_DEBUG_FILES_MAP`.
6. **Add hidden-imports** to `build.py`'s `hidden_imports` list for each submodule of the strategy (PyInstaller cannot see lazy imports).
7. **Tests** under `tests/test_<strategy>.py`. Run `pytest` after every change; do NOT mark a change DONE while tests fail.
8. **Change log on this strategy branch's `CLAUDE.md`** — under no circumstances should strategy-specific entries leak back to `main`'s change log.

When merging an infra fix back to `main`, do it as a separate commit on `main` (or a PR `main ← <branch>`) that touches **only** infrastructure files. Strategy commits and infra commits stay strictly separate.

---

## Development Notes

- **Run `pytest` after every code change.** All tests must pass before a change is marked DONE.
- Use **`build.py prod`** for the production EXE. Strategy branches add their own debug keys.
- Generated outputs in `test_data/X-Checks Output/` and ad-hoc diagnostic scripts at the repo root are not part of the app and are not committed.
