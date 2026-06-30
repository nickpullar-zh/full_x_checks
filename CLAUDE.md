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

## Change Log Policy

**Every change, no matter how small, must be logged here. This is not optional and must never be skipped.**

- When a change is **proposed**, add it to the change log immediately with status `PROPOSED`.
- When a change is **implemented and confirmed**, update the entry to `DONE`.
- No change may be silently applied. If a change is rejected or abandoned, mark it `REJECTED` with a reason.
- **Claude must update the change log as part of implementing the change — not as an afterthought, and not only when reminded.** Updating the log is the final step before marking any task complete.
- **Every change must also bump the version in `version.py`.** The version is displayed in the app title bar, UI label, and Excel processing log so users can confirm which version they are running when reporting issues.

---

## Change Log

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
