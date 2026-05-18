---
name: parity-check
description: Run run_new.py then run_compare_outputs.py against the current test pair and report any differences. Use after every change to strategies/x_checks/ files before confirming the change as DONE.
disable-model-invocation: true
---

Run the X-Checks parity verification workflow in two steps:

**Step 1 — generate fresh output:**
Run `test_data/run_new.py` using the Python interpreter. This writes a new timestamped file into `test_data/X-Checks Output/`.

**Step 2 — compare against reference:**
Before running `test_data/run_compare_outputs.py`, check that `NEW_FILE` in that script points to the file just generated in Step 1. Update the path if it doesn't. Then run the script.

**Report the result clearly:**
- Total X-Checks in OLD vs NEW
- Any rows only in OLD or only in NEW
- Total field-level differences, broken down by column
- Verdict: **PARITY INTACT** (0 differences) or **REGRESSION FOUND** (list differences)

If differences exist, note whether each looks like an intentional improvement (e.g. a known bug fix) or an unexpected regression.
