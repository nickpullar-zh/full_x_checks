# X-Checks Full Application — Infrastructure (`main`)

## Repository

- **Repo:** https://github.com/nickpullar-zh/full_x_checks
- **Infrastructure branch:** `main` — shared plumbing only, no strategy code
- **Strategy branches** (each carries its own strategy module + tests + UAT):
  - `v0.3-X-Checks` — X-Checks comparison + Collect Live X-Checks (X-Checks family)
  - `v0.2-Grouping_By` — Grouping By comparison
  - `v0.5-Accounting-Principles` — Accounting Principles (shipped)
  - `v0.6-Conditions` — Conditions (in progress)
  - `v0.7-Full-Run` — Full Run (in progress)

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

## Change Log Policy

**Every change, no matter how small, must be logged here. This is not optional and must never be skipped.**

- When a change is **proposed**, add it to the change log immediately with status `PROPOSED`.
- When a change is **implemented and confirmed**, update the entry to `DONE`.
- No change may be silently applied. If a change is rejected or abandoned, mark it `REJECTED` with a reason.
- **Claude must update the change log as part of implementing the change — not as an afterthought, and not only when reminded.** Updating the log is the final step before marking any task complete.
- **Every change must also bump the version in `version.py`.** The version is displayed in the app title bar, UI label, and Excel processing log so users can confirm which version they are running when reporting issues.

---

## Change Log

### v0.6.0 — Conditions strategy (completed 2026-06-22)

Full implementation of the Conditions strategy on `v0.6-Conditions`. Compares X-Check condition data from the publication file (yellow/green cells) against the FIP ZQ9_VALMETH extract, producing a 4-sheet output workbook.

| # | Change | Status | Notes |
|---|--------|--------|-------|
| 1 | `task_configs.py`: add `CONDITIONS_UPLOAD_CONFIG` (2 file fields: X-Checks Publication File, FIP File). | DONE | |
| 2 | `task_registry.py`: register `"Conditions"` entry + `if False:` import. | DONE | |
| 3 | `strategies/conditions/extract.py`: `_is_yellow`, `_is_green` helpers + `extract_conditions(pub_path, sheet_name)` → working DataFrame with X-Check No., 5 condition value columns, 5 concat columns. | DONE | Handles openpyxl rgb/theme/indexed colour variants. |
| 4 | `strategies/conditions/fip.py`: `process_fip(df)` — rename 8 columns per spec, add `Concatenated` key column (`Normal X-Check No \| Condition No`). | DONE | |
| 5 | `strategies/conditions/compare.py`: `compare(working_df, fip_df)` → results DataFrame (True/False/blank per X-Check × condition) + summary dict. | DONE | |
| 6 | `strategies/conditions/conditions.py`: `Conditions(BaseStrategy)` wiring extract → fip → compare → `write_excel_output`. `apply_output_formatting` applies green/red fills to True/False cells. | DONE | |
| 7 | `strategies/conditions/__init__.py`: re-exports `Conditions` for lazy import. | DONE | |
| 8 | `main.py`: add `_DEBUG_FILES_CONDITIONS` dict + entry in `_DEBUG_FILES_MAP`. | DONE | Uses existing test_data files. |
| 9 | `build.py`: add `conditions_debug` BUILDS entry + hidden_imports for all 5 conditions submodules. | DONE | |
| 10 | `tests/test_conditions.py`: 20 unit tests for extract, fip, compare + integration test for process(). All pass. | DONE | |
| 11 | `version.py`: bump to `0.6.0`. | DONE | |

### v0.6.1 — Conditions output format aligned to reference workbook (completed 2026-06-22)

Restructure the Conditions comparison output to match the 3-column format used in the Q2 2026 Final Cross Checks Summary workbook (`EBX Data | FIP Data | Comparison`), with one row per `XCheck|ConditionValue` pair instead of one row per X-Check with 5 match columns.

| # | Change | Status | Notes |
|---|--------|--------|-------|
| 1 | `strategies/conditions/compare.py`: rewrite `compare()` to emit one row per pair with columns `EBX Data`, `FIP Data`, `Comparison` (True/False). Summary dict updated accordingly. | DONE | |
| 2 | `strategies/conditions/conditions.py`: rename output sheet from `"Comparison Results"` to `"Conditions"`; remove wide-format green/red cell formatting (reference has no fills). Update `apply_output_formatting` accordingly. | DONE | |
| 3 | `tests/test_conditions.py`: update `TestCompare` tests to match new 3-column row-per-pair output shape. | DONE | |
| 4 | `version.py`: bump to `0.6.1`. | DONE | |

### v0.6.2 — Fix extraction rule: only collect condition cells that are themselves yellow or green (completed 2026-06-22)

The previous logic also collected rows where the X-Check No. cell was green and the condition cell had any value, regardless of the condition cell's own colour. The correct rule (matching the manual workbook process) is: a condition cell is included only if it is itself yellow or green **and** has a non-blank value.

| # | Change | Status | Notes |
|---|--------|--------|-------|
| 1 | `strategies/conditions/extract.py`: remove the `xc_is_green` path that collects condition cells based on the X-Check No. cell colour. Keep only: yellow condition cell → collect; green condition cell with value → collect. | DONE | |
| 2 | `tests/test_conditions.py`: added `TestExtractionRule` class with 4 tests confirming the corrected rule. | DONE | |
| 3 | `version.py`: bump to `0.6.2`. | DONE | |

### v0.6.3 — Honour "process only differences" checkbox in Conditions extraction (completed 2026-06-22)

When the checkbox is **unchecked**: collect every non-blank condition cell regardless of colour (full file).
When the checkbox is **checked**: collect only condition cells that are yellow or green (changed/new rows — ~20 rows).

| # | Change | Status | Notes |
|---|--------|--------|-------|
| 1 | `strategies/conditions/extract.py`: add `process_only_differences` parameter to `extract_conditions()`. When `False`, skip the colour check and collect all non-blank condition cells. When `True` (default behaviour), keep the yellow/green-only rule. | DONE | |
| 2 | `strategies/conditions/conditions.py`: pass `files["process_only_differences"]` through to `extract_conditions()`. Log message updated to show mode. | DONE | |
| 3 | `tests/test_conditions.py`: added 2 tests for both modes in `TestExtractionRule`. | DONE | |
| 4 | `version.py`: bump to `0.6.3`. | DONE | |

### v0.6.3 (UAT) — Conditions UAT Test Plan workbook (completed 2026-06-23)

| # | Change | Status | Notes |
|---|--------|--------|-------|
| 1 | `docs/generate_uat.py`: self-contained openpyxl script generating the Conditions UAT workbook in the Zurich reference format (Overview + Test Cases + Sign-off). 22 test cases covering launch, file selection, differences-only mode, full-file mode, output structure, data accuracy, sensitivity label, stop/return-to-form, and error handling. | DONE | |
| 2 | `docs/20260623 Conditions_v0.6.3 Test Plan.xlsx`: generated output committed to repo. | DONE | |

### v0.7.0 — Full Run strategy (completed 2026-06-23)

Adds a new "Full Run" task that runs every registered strategy sequentially, combining all output sheets into a single colour-coded workbook with one shared Processing Log.

| # | Change | Status | Notes |
|---|--------|--------|-------|
| 1 | `strategies/full_run/full_run.py`: `FullRun(BaseStrategy)` — iterates `TASK_REGISTRY`, partitions inputs per strategy, monkey-patches `write_excel_output` to capture sheets, prefixes tab names, applies per-strategy tab colours in `apply_output_formatting`. | DONE | |
| 2 | `strategies/full_run/__init__.py`: re-exports `FullRun`. | DONE | |
| 3 | `task_configs.py`: add `_build_full_run_config(registry)` — builds merged `UploadTaskConfig` from all registered strategies, deduplicating file fields by label. | DONE | |
| 4 | `task_registry.py`: register `"Full Run"` entry (added last so `_build_full_run_config` sees all other entries) + `if False:` import. | DONE | |
| 5 | `main.py`: add `_DEBUG_FILES_FULL_RUN` dict (reuses Conditions test files) + entry in `_DEBUG_FILES_MAP`. | DONE | |
| 6 | `build.py`: add `full_run_debug` BUILDS entry + hidden imports for `strategies.full_run` and `strategies.full_run.full_run`. | DONE | |
| 7 | `tests/test_full_run.py`: 17 unit tests covering `_unique_name`, `_build_full_run_config`, `FullRun.process` (sheet capture, prefixing, skip self, exception resilience, deduplication), and `apply_output_formatting` (tab colours). All pass. | DONE | |
| 8 | `version.py`: bump to `0.7.0`. | DONE | |

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
| 9 | All strategy modules removed: `strategies/x_checks/`, `strategies/accounting_principles.py`, `strategies/conditions.py`, `strategies/grouping_by.py`. All strategy tests removed: `tests/test_compare.py`, `tests/test_ebx_extraction.py`, `tests/test_fip_extraction.py`, `tests/test_integration.py`, `tests/test_load_known_exceptions.py`, `tests/test_variable_builder.py`, `tests/fixtures/`, `tests/generate_golden_fixtures.py`. | DONE | These remain on `v0.3-X-Checks` (along with `v0.4`'s extensions) where they belong. |

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
