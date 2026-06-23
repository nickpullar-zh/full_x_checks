import sys
sys.path.insert(0, r"C:\Users\NICK.PULLAR\OneDrive - Zurich Insurance\Projects\Testing Automation\Python Files\Full_X-Checks")
sys.path.insert(0, r"C:\Users\NICK.PULLAR\OneDrive - Zurich Insurance\Projects\Testing Automation\Python Files\X-Checks")

import pandas as pd
from EBXExtraction1 import EBXExtraction1
import globals
from datetime import datetime

EBX_FILE = r"C:\Users\NICK.PULLAR\OneDrive - Zurich Insurance\Projects\Testing Automation\Python Files\Full_X-Checks\test_data\20251205 EPM X-Checks - Original.xlsx"
OUT_DIR  = r"C:\Users\NICK.PULLAR\OneDrive - Zurich Insurance\Projects\Testing Automation\Python Files\Full_X-Checks\test_data\X-Checks Output"
TARGETS  = ['A800_00', 'AE100_70', 'AE101_70', 'AE102_70', 'AL005_17', 'AL006_17', 'LR048_17']

# --- Run old EBXExtraction1 exactly as main.py does ---
globals.TIMESTAMP = datetime.now().strftime("%Y%m%d %H%M%S")
old_path = EBXExtraction1([EBX_FILE], OUT_DIR)
print(f"Old EBX output: {old_path}")

old_df = pd.read_excel(old_path, sheet_name='Sheet').astype(str)
print("\nOLD results for target X-Checks:")
for xcheck in TARGETS:
    rows = old_df[old_df['X-Check Number'] == xcheck]
    if len(rows):
        print(f"\n  {xcheck}:")
        print(f"    Formula  : {rows.iloc[0]['EBX Formula']}")
        print(f"    Variables: {rows.iloc[0]['EBX Variables']}")
    else:
        print(f"\n  {xcheck}: NOT FOUND")

# --- Run new extract_ebx on the same file loaded the same way ---
from strategies.x_checks.ebx_extraction import extract_ebx

new_df = pd.read_excel(EBX_FILE, sheet_name='cross checks all')
new_df = new_df.astype(str).reset_index().fillna('')
new_results = extract_ebx(new_df)
new_by_xcheck = {r['X-Check Number']: r for r in new_results}

print("\nNEW results for target X-Checks:")
for xcheck in TARGETS:
    if xcheck in new_by_xcheck:
        r = new_by_xcheck[xcheck]
        print(f"\n  {xcheck}:")
        print(f"    Formula  : {r['EBX Formula']}")
        print(f"    Variables: {r['EBX Variables']}")
    else:
        print(f"\n  {xcheck}: NOT FOUND")
