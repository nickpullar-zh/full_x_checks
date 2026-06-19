# X-Checks Full Application — Infrastructure (`main`)

## Repository

- **Repo:** https://github.com/nickpullar-zh/full_x_checks
- **Infrastructure branch:** `main` — shared plumbing only, no strategy code
- **Strategy branches** (each carries its own strategy module + tests + UAT):
  - `v0.3-X-Checks` — X-Checks comparison + Collect Live X-Checks (X-Checks family)
  - `v0.2-Grouping_By` — Grouping By comparison
  - `v0.5-Accounting-Principles` — Accounting Principles (in progress)

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

### v0.5.12 — MIP sensitivity label applied to all generated Excel outputs (completed 2026-06-19)

| # | Change | Status | Notes |
|---|--------|--------|-------|
| 1 | New module `strategies/sensitivity.py` with `ExcelLabeler`. Mirrors the VBA `SetLabelInfo` procedure (`Workbook.SensitivityLabel.CreateLabelInfo` + `AssignmentMethod=PRIVILEGED` + `LabelName`/`LabelId`/`SiteId` + `SetLabel`) via pywin32. Caches a single `Excel.Application` COM instance per strategy run so multiple writes share startup cost. | DONE | `_LABELS` table imports the seven Label IDs and the Tenant `SITE_ID` from your VBA module verbatim. |
| 2 | `BaseStrategy.write_excel_output()`: after the workbook saves, call `_apply_sensitivity_label(path)` which lazily creates the `ExcelLabeler` and applies `DEFAULT_SENSITIVITY_LEVEL = "Internal_Use_Only"`. Failure logs `[Sensitivity] Could not apply label: <reason>` and the run continues. | DONE | Wired in `BaseStrategy` so every current and future strategy on every branch picks it up automatically. |
| 3 | `BaseStrategy.execute()`: wrapped the main try block with a `finally` that closes the cached labeler so Excel exits cleanly when the strategy finishes (success or error). | DONE | |
| 4 | `build.py`: added `strategies.sensitivity` plus `win32com.client`, `win32com`, `pythoncom`, `pywintypes` to `hidden_imports`. | DONE | PyInstaller can't statically resolve `win32com.client.DispatchEx`. |
| 5 | `tests/test_sensitivity.py`: 6 new unit tests covering the level → (LabelId, LabelName) mapping, unknown-level error, and missing-file/unknown-level failure paths on the labeler. | DONE | 26 tests passing. Smoke run on the live AP fixtures applied + verified `Internal_Use_Only` label. |
| 6 | `version.py`: bump to `0.5.12`. | DONE | |

### v0.5.11 — UAT plan for Accounting Principles (completed 2026-06-19)

| # | Change | Status | Notes |
|---|--------|--------|-------|
| 1 | `docs/generate_uat.py` authored on this branch in the standard reference format (Sheet1 narrative walkthrough in Aptos Narrow + Sheet2 Zurich-branded sectioned tables). | DONE | Re-runnable on every version bump; filename auto-stamps date + version. |
| 2 | Sections covered: Files Required (3 fields); General UI + Dialog (15 cases including v0.5.8 label wrap, v0.4.3 bold-red errors, v0.4.4 chime, v0.4.5 Exit Application, v0.5.9 Close→selector + Return-to-Form pre-fill); Detailed Field Interaction (9 cases); Workflow-Specific (19 cases including smoke run, output structure, all match rules, v0.5.10 black/grey font priority + dedup, v0.5.4 punctuation-insensitive matching, v0.5.5 process-only-differences filter ON/OFF, error visibility paths). | DONE | |
| 3 | Generated `docs/20260619 X-Checks_AccountingPrinciples_v0.5.11 Test Plan.xlsx`. | DONE | |

### v0.5.11 — Output workbook adds EBX & FIP sheets; rename to Comparison (completed 2026-06-19)

| # | Change | Status | Notes |
|---|--------|--------|-------|
| 1 | `accounting_principles.process()`: rename main sheet `Accounting Principles` → `Comparison`. Add `EBX` sheet (cc_df filtered to in-scope X-Check Nos). Add `FIP` sheet (FIP rows whose V-code is in the validation-methods subset). Sheet order written: EBX → FIP → Comparison → Processing Log. | DONE | Lets the user audit what the strategy was reading without opening the source files. |
| 2 | `apply_output_formatting`: target sheet renamed to `Comparison`. | DONE | Conditional formatting (green Match / red MisMatch) preserved. |
| 3 | `version.py`: bump to `0.5.11`. | DONE | 20 tests pass. Smoke test produces EBX 1499 rows × 95 cols, FIP 5066 rows × 11 cols, Comparison 318 rows × 7 cols. |

### v0.5.10 — Black/grey font priority; one row per (X-Check, V-code) (completed 2026-06-18)

The validation methods file uses grey font (theme=1, tint > 0) to mark "reference copies" of a method that already appears in another column. The previous comparator emitted one row per matching event definition, producing duplicates whenever a V-code appeared (in any colour) under multiple events.

| # | Change | Status | Notes |
|---|--------|--------|-------|
| 1 | `validation_methods.py`: new `MethodBinding` dataclass and `parse_method_bindings()` that returns one record per (V-code, event, severity, font, column) cell occurrence. `_font_kind(cell)` returns `"grey"` for `theme=1, tint>0` cells, else `"black"`. The legacy `parse_validation_methods()` + `EventDefinition` helpers are kept so the existing test suite keeps passing. | DONE | |
| 2 | `compare.py`: new `compare_with_bindings()` consumed by the strategy. For each (V-code, X-Check) FIP hit it walks bindings sorted by `(font priority [black<grey], leftmost column first)` and emits ONE row attributed to the first binding whose cross-checks-all column has a non-empty actual letter. Old `compare()` retained as a backwards-compat wrapper for tests. | DONE | |
| 3 | `accounting_principles.process()`: switched from `parse_validation_methods` + `compare` to `parse_method_bindings` + `compare_with_bindings`. | DONE | |
| 4 | `version.py`: bump to `0.5.10`. | DONE | 20 tests pass. Live data: 318 rows / 318 unique (X-Check, Method) — duplicates eliminated. |

### v0.5.9 — Close → strategy selector; Exit Application → quit (completed 2026-06-16)

| # | Change | Status | Notes |
|---|--------|--------|-------|
| 1 | `progress_dialog.py`: `_on_stop_or_close` no longer calls `sys.exit()` on the success branch. Both success (Close) and failure (Return to Form) now invoke `on_dismiss(success)` with a boolean flag, leaving the exit decision to the caller. Hard exit stays on the dedicated **Exit Application** button. | DONE | The dialog itself no longer makes the policy choice between "exit" and "return to start" — main.py decides per task. |
| 2 | `main.py` `_run_task`: `on_dismiss` callback now branches on the success flag. Success → `root.deiconify()` to return to the strategy selector for a fresh pick. Cancel/error → existing prefill+rerun behaviour. | DONE | Debug mode (`_run_debug`) keeps `on_dismiss=None` so it still hard-exits on any close — debug builds are intended to run a single task and stop. |
| 3 | `version.py`: bump to `0.5.9`. | DONE | 20 tests passing. |

### v0.5.8 — Field-label wrap fix (X-Checks Publication File "*" on same row) (completed 2026-06-16)

| # | Change | Status | Notes |
|---|--------|--------|-------|
| 1 | `file_upload_ui.py`: bump `wraplength` on the field-label widgets from `150` to `220` (file-field labels and the Output Directory label). The Accounting Principles dialog's "X-Checks Publication File *" was wrapping; the `*` ended up on row 2. | DONE | The dialog auto-sizes around the wider label so it grows in width by ~70 px overall. |
| 2 | `version.py`: bump to `0.5.8`. | DONE | |

### v0.5.7 — Centralise EXE strategy label (completed 2026-06-16)

| # | Change | Status | Notes |
|---|--------|--------|-------|
| 1 | `build.py`: introduce `STRATEGY_LABEL = "AccountingPrinciples"` constant. Both BUILDS entries derive their `name` from it (debug: `X-Checks_Debug_<LABEL>_<VERSION>`, prod: `X-Checks_<LABEL>_<VERSION>`). Renaming the strategy is now a one-line change instead of two. | DONE | Replaces v0.5.6's spot fix where I edited only the prod entry. Future strategy branches set `STRATEGY_LABEL` once and both EXE filenames stay consistent. |
| 2 | `version.py`: bump to `0.5.7`. | DONE | |

### v0.5.6 — Rename prod EXE to include strategy name (completed 2026-06-16)

| # | Change | Status | Notes |
|---|--------|--------|-------|
| 1 | `build.py` BUILDS prod entry: `name` changed from `f"X-Checks_{VERSION}"` to `f"X-Checks_AccountingPrinciples_{VERSION}"` so the production EXE's filename identifies the strategy carried by this branch. | DONE | Was a leftover from when `main` was refreshed before strategies were split into per-branch packaging. |
| 2 | `version.py`: bump to `0.5.6`. | DONE | |

### v0.5.5 — Wire 'Process only differences' into Accounting Principles (completed 2026-06-16)

| # | Change | Status | Notes |
|---|--------|--------|-------|
| 1 | `accounting_principles.process()`: when `self.process_only_differences` is True, run an in-strategy port of v0.4.1's `select_x_check_nos` pipeline (drop INACTIVE rows, keep non-blank Type of change, drop X-Checks with Exclude Z-Core = X, drop X-Checks whose Category cell is yellow #FFFF00) and use those as the in-scope X-Check Nos. When False, fall back to every unique non-blank X-Check No. (existing behaviour). | DONE | The pipeline is reimplemented locally rather than imported from `strategies.x_checks` to honour the architectural rule: strategy branches don't depend on each other's modules. |
| 2 | `_select_in_scope_x_checks` helper: case-insensitive column resolution; reads the EBX file with openpyxl to detect yellow Category cells (auto-detects header row 1–6 by signal `X-Check No. + Status + Type of change`). | DONE | |
| 3 | `version.py`: bump to `0.5.5`. | DONE | 20 tests passing. Filter narrows the live `20260313 Cross Checks All.xlsx` from 189 unique X-Checks to 152. |

### v0.5.4 — Punctuation-insensitive event-column matching (completed 2026-06-16)

Reconciled the strategy output (322 rows in v0.5.3) against the spreadsheet template's manual extract (435 rows). Root cause: cross-checks-all uses `DE GAAP RFD` (space) while validation methods uses `DE-GAAP RFD` (hyphen) — the strategy's exact-string column lookup missed the entire DE-GAAP family.

| # | Change | Status | Notes |
|---|--------|--------|-------|
| 1 | `compare._norm_event_name`: lowercase + strip spaces and hyphens. Used as the canonical key for matching validation-methods event names against cross-checks-all column headers. | DONE | |
| 2 | `compare._build_event_to_column`: pre-builds {event_name → actual cross-checks-all column header} once. Handles pandas's `.1`/`.2` suffix on duplicate column names by stripping a trailing `.<digits>` before normalising, and keeps the first occurrence so the bogus duplicate is ignored. | DONE | |
| 3 | `compare.compare`: looks up the cross-checks-all column via the new map instead of `d.event in cross_checks_df.columns`. | DONE | |
| 4 | `tests/test_accounting_principles.py`: 2 new tests (`test_compare_event_name_punctuation_insensitive`, `test_compare_pandas_dot_n_dedup_column_ignored`). | DONE | 20 tests passing. |
| 5 | `version.py`: bump to `0.5.4`. | DONE | |
| - | Output now produces 407 rows (covers all 402 unique (X-Check, Method) pairs from the spreadsheet, plus 1 corrected typo: spreadsheet template has `V815S` and `V110` while the validation methods file has `V851S` and `V1100` — the strategy uses the validation methods values, so it correctly emits `V851S\|A336_00` which the spreadsheet's `V815S` block missed). | INFO | |
| - | Why 407 vs 435: the spreadsheet has separate method-blocks per (Event, Method) so a method like `V900W` that serves both `IFRS New RFD Warning` AND `IFRS New SFD Warning` produces two blocks → two rows per X-Check. The strategy collapses both into the same FIP join. Both representations are correct; the strategy's grain is (X-Check × Event × Method) which matches what the user asked for in the spec. | INFO | |

### v0.5.3 — Header-row auto-detection in BaseStrategy._load_files (completed 2026-06-16)

| # | Change | Status | Notes |
|---|--------|--------|-------|
| 1 | `FileFieldConfig`: add optional `header_signals: list[str]` field — when set, _load_files scans the first 6 rows of the sheet for a row whose cells include ALL of these names (case-insensitive, stripped) and uses that as the header row. | DONE | Defaults to None = treat row 1 as header (existing behaviour). |
| 2 | `BaseStrategy._detect_header_row()`: walks rows 1-6 with openpyxl read-only, returns the matching 0-indexed row or 0 as fallback so the caller still gets a DataFrame and the strategy's own 'column not found' check can surface the right error. | DONE | |
| 3 | `_load_files`: pass `header=` from the detector into pd.read_excel. | DONE | |
| 4 | `task_configs.py`: EBX field for Accounting Principles gets `header_signals=["X-Check No.", "Status", "Type of change"]`. | DONE | These three column names coexist on the header row of every Cross Checks All sheet seen so far. |
| 5 | `accounting_principles.process()`: drop the explicit `pd.read_excel(... header=1)` and `pd.read_excel(fip_path, ...)` reads — both files are already loaded by BaseStrategy via the upload form's path/sheet inputs. | DONE | Strategy now just consumes `loaded_files`. |
| 6 | `version.py`: bump to `0.5.3`. | DONE | 18 tests passing; same 322-row output on the test fixtures. |

### v0.5.2 — Abort lines styled as errors; abort no longer reads as success (completed 2026-06-16)

| # | Change | Status | Notes |
|---|--------|--------|-------|
| 1 | `progress_dialog._ERROR_KEYWORDS`: extend with `aborting`, `aborted`, `cannot`, `invalid`, `missing`, `not found`. Any line containing these now styles bold-red and plays the error chime. | DONE | Closes the case where `Required column 'X-Check No.' not found — aborting` was rendered in default dark blue. |
| 2 | `BaseStrategy.execute()`: return `bool(self.process(...))` instead of unconditional `True`. Strategies that bail early (return `None`) now signal failure to `run_processing`, which switches the dialog button to **Return to Form** instead of **Close** + "Processing complete". | DONE | |
| 3 | `accounting_principles.process()`: return `True` at the bottom of the success path. Existing early-return-on-error branches stay as-is and now correctly propagate `False`. | DONE | |
| 4 | `version.py`: bump to `0.5.2` | DONE | 18 tests passing. |

### v0.5.1 — Separate FIP File upload field (completed 2026-06-15)

| # | Change | Status | Notes |
|---|--------|--------|-------|
| 1 | `task_configs.py`: add 3rd `FileFieldConfig` "FIP File (VALMSG)" with `default_sheet="FIP Methods Rules and Condition"`, required. | DONE | EBX file no longer needs to also contain the FIP sheet. |
| 2 | `accounting_principles.py`: read FIP from its own path; honour user-supplied sheet names for both the EBX 'cross checks all' sheet and the FIP sheet. | DONE | |
| 3 | `main.py`: add the FIP entry to `DEBUG_FILES_ACCOUNTING_PRINCIPLES`. | DONE | Debug build still uses the same workbook for both EBX and FIP — the file containing both sheets — but the strategy treats them as separate paths. |
| 4 | `version.py`: bump to `0.5.1`. | DONE | 18 tests passing. |

### v0.5.0 — Accounting Principles strategy (completed 2026-06-15, branch v0.5-Accounting-Principles)

First release of the Accounting Principles task. Compares the severity letter recorded for each X-Check on the EBX `cross checks all` sheet against the W/E recorded in `FIP Methods Rules and Condition` (the VALMSG dump), guided by a `Validation Methods` workbook that defines which methods correspond to which Validation Events.

| # | Change | Status | Notes |
|---|--------|--------|-------|
| 1 | New strategy `strategies/accounting_principles/` with three modules: `validation_methods.py` (parser for the validation methods xlsx), `compare.py` (FIP-gated comparator + match table), `accounting_principles.py` (`AccountingPrinciples(BaseStrategy)`). | DONE | Pipeline: parse Validation Methods → load `cross checks all` (header row 2) → load `FIP Methods Rules and Condition` from same workbook → walk FIP, emit one row per (X-Check, Event, Method) where the X-Check letter is non-empty and a FIP row exists. Match if FIP letter (W/E) == cross-checks-all letter (w/e). |
| 2 | Match rules: `Warning↔w`, `Error↔e`, `Both↔(w or e)` — Both events use whichever methods (warning-row or error-row) match the actual letter. Empty actual → no row. | DONE | Both-events with the same method in both row groups don't double-count. |
| 3 | `task_configs.py`: `ACCOUNTING_PRINCIPLES_UPLOAD_CONFIG` with two file fields (Validation Methods File, X-Checks Publication File). FIP Methods Rules and Condition is read from the SAME workbook as the X-Checks Publication File so the user only picks the file once. | DONE | Strategy reads cross-checks-all with `header=1` directly (BaseStrategy's `_load_files` defaults to `header=0`, which is wrong for this sheet). |
| 4 | `task_registry.py`: register `Accounting Principles` task; add lazy-import factory and PyInstaller hint. | DONE | |
| 5 | `build.py`: new `BUILDS` entry `ap` for the debug build (`X-Checks_Debug_AccountingPrinciples_v<ver>`); add accounting_principles submodules to `hidden_imports`. | DONE | |
| 6 | `main.py`: `DEBUG_FILES_ACCOUNTING_PRINCIPLES` dict; `_DEBUG_FILES_MAP` registers it under `Accounting Principles`. `DEBUG_TASK` defaults to `Accounting Principles` on this branch. | DONE | |
| 7 | `test_data/`: copy reference fixtures `validation methods.xlsx` and `20260602 VALMSG (Accounting Principle).xlsx` for debug builds. | DONE | |
| 8 | `tests/test_accounting_principles.py`: 18 unit tests covering `_extract_method_codes`, `parse_validation_methods` (Warning-only, Error-only, independent W+E, Both-via-merged-cell, dash-as-blank, subset filter), and `compare` (Match, MisMatch, FIP-missing-no-row, actual-empty-no-row, Both-w, Both-e, no-double-count, out-of-scope-xcheck filtered). | DONE | All 18 passing. |
| 9 | `version.py`: bump to `0.5.0` (first release of v0.5 line). | DONE | |
| - | UI multi-select + persistence to `%APPDATA%/X-Checks/accounting_principles.json` for the Validation Events subset. | DEFERRED | v0.5.0 ships with the 27-event default hardcoded; v0.5.1 adds the form-driven multi-select. |

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
