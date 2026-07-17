"""
Process the FIP ZQ9_VALMETH extract.

Expected layout — raw ZQ9_VALMETH (8 columns):
  0  MethC                → MethC
  1  MK                   → MK
  2  Medium Text          → Medium Text MK
  3  ValidRule            → Normal X-Check No
  4  Medium Text          → X-Check Medium Text
  5  UCFV20G-TRUE_BRANCH  → UCFV20G-TRUE_BRANCH
  6  ValidRule            → Condition No
  7  Medium Text          → Condition Medium Text
  Concatenated built as: "Normal X-Check No|Condition No"

Files with a leading "Key" column are rejected upstream in conditions.py.
"""

import pandas as pd

_RENAMED_COLS_8 = [
    "MethC",
    "MK",
    "Medium Text MK",
    "Normal X-Check No",
    "X-Check Medium Text",
    "UCFV20G-TRUE_BRANCH",
    "Condition No",
    "Condition Medium Text",
]

CONCAT_COL = "Key (Concatenated)"


def process_fip(fip_df: pd.DataFrame) -> pd.DataFrame:
    """
    Rename columns and build the Concatenated key column from the raw
    8-column ZQ9_VALMETH extract. Returns a new DataFrame.
    """
    df = fip_df.copy()

    if len(df.columns) < 8:
        raise ValueError(
            f"FIP sheet has only {len(df.columns)} columns; expected at least 8. "
            "Check that you loaded the correct ZQ9_VALMETH sheet."
        )

    new_names = list(_RENAMED_COLS_8) + list(df.columns[8:])
    df.columns = new_names

    # Drop rows where both Normal X-Check No and Condition No are blank
    df = df[~(df["Normal X-Check No"].isna() & df["Condition No"].isna())].copy()
    df = df[~((df["Normal X-Check No"].astype(str).str.strip() == "") &
              (df["Condition No"].astype(str).str.strip() == ""))].copy()

    def _make_key(row):
        xc = str(row["Normal X-Check No"]).strip() if pd.notna(row["Normal X-Check No"]) else ""
        cond = str(row["Condition No"]).strip() if pd.notna(row["Condition No"]) else ""
        if xc and cond:
            return f"{xc}|{cond}"
        return ""

    df[CONCAT_COL] = df.apply(_make_key, axis=1)

    # Drop rows with no usable Concatenated key
    df = df[df[CONCAT_COL].notna() & (df[CONCAT_COL].astype(str).str.strip() != "")].copy()

    return df.reset_index(drop=True)
