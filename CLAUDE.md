# X-Checks Full Application — Development Guide

## Repository

- **Repo:** https://github.com/nickpullar-zh/full_x_checks
- **Working branch:** `v0.3-X-Checks`
- **Archive branch (read-only):** `X-Checks_v0.3_Parity_with_Original_X-Checks_Archive`

---

## Change Log Policy

**Every change, no matter how small, must be logged here.**

- When a change is **proposed**, add it to the change log immediately with status `PROPOSED`.
- When a change is **implemented and confirmed**, update the entry to `DONE`.
- No change may be silently applied. If a change is rejected or abandoned, mark it `REJECTED` with a reason.

---

## Change Log

### v0.3.1 — Post-Parity Improvements (in progress)

| # | Change | Status | Notes |
|---|--------|--------|-------|
| 1 | `compare.py`: guard against empty `bracket_vars` before `[0]` access in `_compare_formulas` | DONE |
| 2 | `fip_extraction.py`: add `_safe_split` helper and replace all unsafe `.split()[N]` calls | DONE |
| 3 | `ebx_extraction.py`: robust null handling — `fillna('')` before `astype(str)`, replace all `'nan'` string comparisons with `''` | DONE |
| 4 | Re-introduce `variable_builder` for EBX variables; un-comment `Variables Match (Builder)` and `FIP Variable (Builder)` columns | DONE |
| 5 | `fip_extraction.py`: extract FIP block-delimiter strings as module-level constants | DONE |
| 6 | `ebx_extraction.py`: remove always-true dead condition `if len(df) - 1 >= index:` | DONE |
| 7 | `fip_extraction.py`: replace boolean flags in `_get_x_check_information` with a `_ParseState` enum | DONE |

### v0.3.2 — Investigation Fixes (in progress)

| # | Change | Status | Notes |
|---|--------|--------|-------|
| 1 | `compare.py`: `_compare_formulas` — use `str.replace(..., count=1)` to avoid replacing all VAL_YTD occurrences | DONE | Added `count=1` to both reorder paths. Parity confirmed on both test pairs. |
| 2 | `x_checks.py`: guard against empty DataFrame before `sort_values` | DONE | Early return with log message if compare() returns no rows. |
| 3 | `variable_builder.py`: guard against empty `fs_accounts` in `_build_variable_name` | DONE | Filters empty strings; returns `'<blank>'` if no valid accounts remain. |
| 4 | `fip_extraction.py`: replace `remaining_checks` list with set to fix O(n²) `.remove()` | DONE | Changed to `set` with `.discard()` (O(1) per call). |
| 5 | `fip_extraction.py`: fix outer loop re-iteration over consumed lines | REJECTED | The re-iteration is load-bearing — 9 X-Checks (A453_00, A454_00, A647_00, A833_00, AL111_70, L124_00, LA001_09, LS602_00, S124_70) are found INSIDE other X-Check blocks. A while-loop fix caused these 9 to show "Not Found". The for-loop re-iteration must be preserved. |
| 6 | `compare.py`: `_compare_variables` — guard against empty string producing false match | DONE | Filters empty strings after split before comparing. |
| 7 | `variable_builder.py`: tighten `.replace('.0', '')` to only strip trailing `.0` | DONE | Uses `endswith('.0')` slice. |
| 8 | `fip_extraction.py`: save incomplete variable if state machine exits without `_BLOCK_END` | DONE | Saves in-progress variable at end of loop if FS_ACCOUNT or MOV_GENERAL state with accumulated accounts. |
| 9 | `x_checks.py`: guard against missing `X-Check No.` column with clear error | DONE | Logs clear message and returns early. | `_ParseState` enum with 5 states (SEARCHING, FORMULA, VARIABLE, FS_ACCOUNT, MOV_GENERAL). Key finding during implementation: ALC/MAT and the 'Movement Type' header line all read values from the same line that triggers the transition. Parity confirmed on both test pairs. | Outer condition removed, inner block de-indented one level. Parity confirmed on both test pairs. | 7 constants defined (`_SEGMENT_END`, `_BLOCK_END`, `_SEPARATOR`, `_BLANK_LINE`, `_FORMULA_HEADER`, `_VAR_HEADER`, `_FS_ACCT_BREAK`). All literal usages replaced. Parity confirmed on both test pairs. | EBX Variables now built via `build_variables_string`. Builder comparison columns live in output (columns H–I). Formatting applied to all 3 match columns. Test pair 2: 0 differences. Test pair 1: same 2 LR048_17 improvements from change 3. | 11 sites updated. Side-effect: LR048_17 (all-null operator/limit record) no longer produces `nan` in formula string or `Movement Types:nan` — these were bugs in the original. Test pair 2: 0 differences. Test pair 1: 2 intentional improvements on LR048_17. | Replaced 11 bare `.split()[N]` accesses (lines 196, 228–236, 247, 250, 276, 288–289) with `_safe_split(line, N)`. Parity confirmed on both test pairs. | Both reorder paths (addition-only and single-minus) can crash with IndexError if neither regex matches. Added `if bracket_vars:` and `if bracket_vars and len(ebx_vars) >= 2:` guards. Parity confirmed on both test pairs. |

---

### v0.3 — Parity with Original X-Checks (completed 2026-05-12)

| # | Change | Status | Notes |
|---|--------|--------|-------|
| 1 | Implement `XChecks.process()` — full pipeline replacing stub | DONE | Calls extract_ebx, extract_fip, compare, write_excel_output |
| 2 | Add `DEBUG_FILES_X_CHECKS` config block to `main.py` | DONE | Allows headless testing without UI |
| 3 | Add `variable_builder.py` | DONE | Standardised variable string builder; currently commented out in EBX, active in FIP as backup |
| 4 | Integrate `variable_builder` into `fip_extraction.py` | DONE | Stored as `FIP Variable (Builder)` alongside primary `FIP Variables` |
| 5 | `compare.py`: `_compare_formulas()` returns normalised formula | DONE | Stores reordered formula in output, matching old Compare_Files.py behaviour |
| 6 | `compare.py`: `ebx_by_xcheck` uses first-occurrence logic | DONE | Matches old code's `EBXFile.index[...][0]`; fixes 6–7 X-Checks with duplicate entries |
| 7 | `ebx_extraction.py`: build variables string from `_create_variable()` output | DONE | Reverted from `build_variables_string`; required for parity |
| 8 | `x_checks.py`: `x_check_list` read from raw DataFrame, not extraction results | DONE | Ensures X-Checks with no Account rows (e.g. CON_UK_CH) are still searched in FIP |
| 9 | `compare.py`: simplify "Not Found" labels | DONE | Changed `Not Found (EBX)` / `Not Found (FIP)` → `Not Found` to match old output |
| 10 | `x_checks.py`: sort output by X-Check Number | DONE | Matches old Format_Excel_File alphabetical sort |
| 11 | `x_checks.py`: remove summary block, write headers at row 1 | DONE | Required for sheet-level identity with old "All Data" sheet |
| 12 | `base_strategy.py`: fix `apply_conditional_formatting` header search | DONE | Now scans all rows rather than row 1 only; supports sheets with summary blocks |
| 13 | `compare.py`: temporarily comment out builder columns | DONE | `Variables Match (Builder)` and `FIP Variable (Builder)` disabled pending re-introduction |

---

## Current State (as of 2026-05-12, commit 830d683)

The new X-Checks strategy produces output **identical** to the original application, verified across two test file pairs (0 differences in 653 rows; 0 differences in 189 rows). See `test_data/X-Checks_Parity_Proof.md`.

### Key files

| File | Purpose |
|------|---------|
| `strategies/x_checks/x_checks.py` | Main strategy — orchestrates extraction, comparison, output |
| `strategies/x_checks/ebx_extraction.py` | EBX Excel extraction |
| `strategies/x_checks/fip_extraction.py` | FIP text file parsing |
| `strategies/x_checks/compare.py` | EBX vs FIP comparison |
| `strategies/x_checks/variable_builder.py` | Standardised variable builder (ready for re-introduction) |
| `strategies/base_strategy.py` | Shared infrastructure — file loading, Excel output, formatting |
| `main.py` | App entry point; DEBUG_FILES_X_CHECKS block added |

### Known items pending re-introduction (commented out, not deleted)

- `EBX Variable (Builder)` — EBX variables to be rebuilt using `variable_builder` once the identical-output baseline is confirmed and improvements are in scope
- `Variables Match (Builder)` / `FIP Variable (Builder)` — builder-based comparison columns, disabled for parity phase
- `.variable_builder import build_variables_string` in `ebx_extraction.py` — commented out, ready to re-enable

---

## Development Notes

- Always run both `run_original.py` and `run_new.py` against the same test files before confirming any change
- Use `run_compare_outputs.py` to verify zero differences after each change
- Test scripts live in `test_data/`; update file paths before each run
- Diagnostic scripts (`diagnose_ebx*.py`, `diagnose_con_uk_ch.py`) are untracked investigation tools, not part of the app
