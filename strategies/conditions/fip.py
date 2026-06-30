"""
Process the FIP ZQ9_VALMETH extract.

Two supported layouts:

Layout A — raw ZQ9_VALMETH (8 columns, no Key column):
  0  MethC                → MethC
  1  MK                   → MK
  2  Medium Text          → Medium Text MK
  3  ValidRule            → Normal X-Check No
  4  Medium Text          → X-Check Medium Text
  5  UCFV20G-TRUE_BRANCH  → UCFV20G-TRUE_BRANCH
  6  ValidRule            → Condition No
  7  Medium Text          → Condition Medium Text
  Concatenated built as: "Normal X-Check No|Condition No"

Layout B — pre-processed file with a leading Key column (11 columns):
  0  Key                  (already "MK|ValidRule" — used as Concatenated directly)
  1  MethC
  2  MK                   → MK
  3  Medium Text          → Medium Text MK
  4  ValidRule            → Normal X-Check No
  5  Long Text / Medium Text → X-Check Medium Text
  6  UCFV20G-TRUE_BRANCH
  7+ remaining columns kept as-is
  Concatenated = Key column value

The Concatenated column is the lookup key matched against the publication
working sheet's concat columns (XCheck|ConditionValue format).
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

CONCAT_COL = "Concatenated"


def process_fip(fip_df: pd.DataFrame) -> pd.DataFrame:
    """
    Rename columns and add/expose a Concatenated key column.

    Accepts both Layout A (raw 8-column ZQ9_VALMETH) and Layout B
    (pre-processed with a leading Key column).

    Returns a new DataFrame (does not mutate the input).
    """
    df = fip_df.copy()

    if len(df.columns) < 8:
        raise ValueError(
            f"FIP sheet has only {len(df.columns)} columns; expected at least 8. "
            "Check that you loaded the ZQ9_VALMETH sheet (FIP Methods Rules and Condition)."
        )

    # Detect Layout B: first column named "Key"
    has_key_col = str(df.columns[0]).strip() == "Key"

    if has_key_col:
        # Layout B — Key column already built; rename remaining cols where possible
        # Keep column names as-is (already meaningful); just expose Key as Concatenated
        df = df.rename(columns={"Key": CONCAT_COL})
    else:
        # Layout A — rename first 8 columns per spec
        new_names = list(_RENAMED_COLS_8) + list(df.columns[8:])
        df.columns = new_names

        # Drop rows where both Normal X-Check No and Condition No are blank
        df = df[~(df["Normal X-Check No"].isna() & df["Condition No"].isna())].copy()
        df = df[~((df["Normal X-Check No"].astype(str).str.strip() == "") &
                  (df["Condition No"].astype(str).str.strip() == ""))].copy()

        # Build concatenated lookup key
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
