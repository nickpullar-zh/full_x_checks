# X-Checks Application — Output Parity Proof

## Summary

The new X-Checks application (`strategies/x_checks/`) has been verified to produce output **identical** to the original X-Checks application (`X-Checks/`) across two independent test file pairs.

---

## Test Results

| Test | Files | Old Rows | New Rows | Differences |
|------|-------|----------|----------|-------------|
| Test 1 | 20251205 EPM X-Checks - Original.xlsx / 20251205 FIP X-Checks - Original.txt | 653 | 653 | **0** |
| Test 2 | 20260313 Cross Checks All.xlsx / 20260318 FIP X-Checks.txt | 189 | 189 | **0** |

Both tests produced zero field-level differences across all rows and columns.

---

## Test Method

For each test pair, both applications were run against the same input files and their outputs compared row by row using `run_compare_outputs.py`.

**Original app pipeline** (`run_original.py`):
1. `FIPExtraction()` — parses FIP text, writes `FIPExtraction {TIMESTAMP}.xlsx`
2. `EBXExtraction1()` — processes EBX Excel file, writes `EBXExtraction {TIMESTAMP}.xlsx`
3. `Compare_Files()` — compares both extractions, writes `ComparedFiles {TIMESTAMP}.xlsx`
4. `Format_Excel_File()` — formats and splits into sheets (All Data / Matched / MisMatched / Not Found / Summary)

**New app pipeline** (`run_new.py`):
1. `extract_ebx()` — processes EBX DataFrame in memory
2. `extract_fip()` — parses FIP text in memory
3. `compare()` — compares both in memory
4. `write_excel_output()` — writes `{TIMESTAMP}_X-Checks Comparison.xlsx`

The comparison (`run_compare_outputs.py`) reads the **All Data** sheet from the original output and the **X-Checks Comparison** sheet from the new output, and compares every field of every row after aligning on X-Check Number.

---

## Issues Found and Resolved During Parity Work

| # | Issue | Root Cause | Fix Applied |
|---|-------|-----------|-------------|
| 1 | FIP Formula ordering difference (47 cases) | Old `Compare_Files.py` stored the normalised (variable-reordered) formula; new code stored the original | `_compare_formulas()` now returns the modified formula alongside the match result |
| 2 | EBX Variables/Formula wrong for 6–7 X-Checks | `compare.py` dict comprehension kept the **last** duplicate X-Check entry; old code used `[0]` (first) | `ebx_by_xcheck` built with first-occurrence logic |
| 3 | EBX Variables built from wrong stage | New code called `build_variables_string` on `group_accounts()` output; old code used `create_variable()` output directly | Reverted to `'|'.join(v['Variable'] for v in dict_variables_output.values())` |
| 4 | Missing X-Check with no Account rows | `x_check_list` was derived from `extract_ebx()` results, skipping X-Checks where all rows have `Account No. = nan` | `x_check_list` now read from all unique `X-Check No.` values in the raw DataFrame |
| 5 | Not Found label mismatch | New code wrote `Not Found (EBX)` / `Not Found (FIP)`; old wrote `Not Found` | Simplified to `Not Found` |
| 6 | Row ordering mismatch | Old `Format_Excel_File` sorted All Data alphabetically; new code wrote in FIP-first order | `df_comparison` sorted by X-Check Number before writing |

---

## Output Files in This Archive

### Test Pair 1 — 20251205 Original Files

| File | Description |
|------|-------------|
| `20251205 EPM X-Checks - Original.xlsx` | EBX input file |
| `20251205 FIP X-Checks - Original.txt` | FIP input file |
| `X-Checks Output/ComparedFiles 20260512 113929.xlsx` | Original app output |
| `X-Checks Output/20260512_113951_X-Checks Comparison.xlsx` | New app output |

### Test Pair 2 — 20260313/20260318 Files

| File | Description |
|------|-------------|
| `20260313 Cross Checks All.xlsx` | EBX input file |
| `20260318 FIP X-Checks.txt` | FIP input file |
| `X-Checks Output/ComparedFiles 20260512 112837.xlsx` | Original app output |
| `X-Checks Output/20260512_113118_X-Checks Comparison.xlsx` | New app output |

---

## New App — Key Source Files

| File | Purpose |
|------|---------|
| `strategies/x_checks/x_checks.py` | Main strategy — orchestrates extraction, comparison, output |
| `strategies/x_checks/ebx_extraction.py` | Extracts X-Check formulas and variables from EBX Excel file |
| `strategies/x_checks/fip_extraction.py` | Parses FIP validation rule text output |
| `strategies/x_checks/compare.py` | Compares EBX and FIP results, produces comparison rows |
| `strategies/x_checks/variable_builder.py` | Standardised variable string builder (ready for future use) |

---

## Test Scripts

| Script | Purpose |
|--------|---------|
| `test_data/run_original.py` | Runs the original X-Checks app pipeline headlessly |
| `test_data/run_new.py` | Runs the new X-Checks app pipeline headlessly |
| `test_data/run_compare_outputs.py` | Compares All Data vs X-Checks Comparison sheet row by row |

---

*Parity verified: 2026-05-12*
