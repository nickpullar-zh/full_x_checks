"""
Compare X-Check|Condition pairs from the publication working sheet
against the FIP Concatenated key set.

Output format matches the reference workbook (Q2 2026 Final Cross Checks Summary):
  EBX Data    — XCheck|ConditionValue key from the publication
  FIP Data    — matching FIP key (same value when found, blank when not found)
  Comparison  — "Matched" / "Not Matched"

One row per pair (not one row per X-Check).
"""

import pandas as pd
from strategies.conditions.extract import CONDITION_COLS
from strategies.conditions.fip import CONCAT_COL


def compare(working_df: pd.DataFrame, fip_df: pd.DataFrame) -> tuple[pd.DataFrame, dict]:
    """
    Args:
        working_df  — output of extract_conditions(); contains X-Check No. and
                      5 concat columns named "<condition_col> (Concat)"
        fip_df      — output of process_fip(); contains a "Concatenated" column

    Returns:
        results_df  — columns: EBX Data, FIP Data, Comparison (one row per pair)
        summary     — {"Total Pairs": n, "Matched": n, "Not Matched": n}
    """
    fip_keys: set[str] = set(
        fip_df[CONCAT_COL].dropna().astype(str).str.strip()
    ) - {""}

    result_rows = []
    matched = 0
    not_matched = 0

    for _, row in working_df.iterrows():
        for cond_col in CONDITION_COLS:
            concat_col = cond_col + " (Concat)"
            ebx_val = str(row.get(concat_col, "")).strip()

            if not ebx_val:
                continue  # no condition value for this X-Check/column — skip

            found = ebx_val in fip_keys
            result_rows.append({
                "EBX Data":   ebx_val,
                "FIP Data":   ebx_val if found else "",
                "Comparison": "Matched" if found else "Not Matched",
            })
            if found:
                matched += 1
            else:
                not_matched += 1

    results_df = pd.DataFrame(result_rows, columns=["EBX Data", "FIP Data", "Comparison"])
    results_df["FIP Data"] = results_df["FIP Data"].fillna("")

    summary = {
        "Total Pairs": len(results_df),
        "Matched": matched,
        "Not Matched": not_matched,
    }

    return results_df, summary
