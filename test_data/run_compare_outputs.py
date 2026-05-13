import sys
sys.path.insert(0, r"C:\Users\NICK.PULLAR\OneDrive - Zurich Insurance\Projects\Testing Automation\Python Files\Full_X-Checks")

import pandas as pd

OLD_FILE = r"C:\Users\NICK.PULLAR\OneDrive - Zurich Insurance\Projects\Testing Automation\Python Files\Full_X-Checks\test_data\X-Checks Output\ComparedFiles 20260512 112837.xlsx"
NEW_FILE = r"C:\Users\NICK.PULLAR\OneDrive - Zurich Insurance\Projects\Testing Automation\Python Files\Full_X-Checks\test_data\X-Checks Output\20260513_140429_X-Checks Comparison.xlsx"

# --- Load old output ("All Data" sheet, standard header at row 0) ---
df_old = pd.read_excel(OLD_FILE, sheet_name="All Data")
df_old = df_old.astype(str).rename(columns={"X-Check": "X-Check Number"})

# --- Load new output (no summary block, header at row 0) ---
df_new = pd.read_excel(NEW_FILE, sheet_name="X-Checks Comparison")
df_new = df_new.astype(str)

# Index both by X-Check Number
old = df_old.set_index("X-Check Number")
new = df_new.set_index("X-Check Number")

cols = ["Formula Match", "EBX Formula", "FIP Formula", "Variables Match", "EBX Variables", "FIP Variables"]

all_keys = sorted(set(old.index) | set(new.index))

differences = []
only_in_old = []
only_in_new = []

for xcheck in all_keys:
    if xcheck not in old.index:
        only_in_new.append(xcheck)
        continue
    if xcheck not in new.index:
        only_in_old.append(xcheck)
        continue

    for col in cols:
        old_val = old.loc[xcheck, col] if col in old.columns else "N/A"
        new_val = new.loc[xcheck, col] if col in new.columns else "N/A"
        if str(old_val).strip() != str(new_val).strip():
            differences.append({
                "X-Check Number": xcheck,
                "Column":         col,
                "Old":            old_val,
                "New":            new_val,
            })

# --- Print results ---
print(f"\nTotal X-Checks in OLD: {len(old)}")
print(f"Total X-Checks in NEW: {len(new)}")
print(f"\nOnly in OLD ({len(only_in_old)}): {only_in_old[:20]}")
print(f"Only in NEW ({len(only_in_new)}): {only_in_new[:20]}")
print(f"\nTotal field-level differences: {len(differences)}")

if differences:
    df_diff = pd.DataFrame(differences)
    print(f"\nDifferences by column:")
    print(df_diff["Column"].value_counts().to_string())
    print(f"\nFirst 50 differences:")
    pd.set_option("display.max_colwidth", 80)
    pd.set_option("display.width", 200)
    print(df_diff.head(50).to_string(index=False))
else:
    print("\nNo differences found — outputs are identical.")
