import sys
sys.path.insert(0, r"C:\Users\NICK.PULLAR\OneDrive - Zurich Insurance\Projects\Testing Automation\Python Files\Full_X-Checks")
sys.path.insert(0, r"C:\Users\NICK.PULLAR\OneDrive - Zurich Insurance\Projects\Testing Automation\Python Files\X-Checks")

import pandas as pd
from strategies.x_checks.ebx_extraction import extract_ebx

EBX_FILE = r"C:\Users\NICK.PULLAR\OneDrive - Zurich Insurance\Projects\Testing Automation\Python Files\Full_X-Checks\test_data\20260313 Cross Checks All.xlsx"
TARGET = 'CON_UK_CH'

# --- Check raw file ---
raw = pd.read_excel(EBX_FILE, sheet_name='cross checks all')
raw_str = raw.astype(str)

target_rows = raw_str[raw_str['X-Check No.'] == TARGET]
print(f"Rows in raw file for '{TARGET}': {len(target_rows)}")
if len(target_rows):
    print(target_rows[['X-Check No.', 'Account No.', 'SubA No.', 'Operator (X-Check Term)',
                        'Operator 1', 'Operator 2', 'Limit 1', 'Limit 2', 'Absolute (result)']].to_string(index=True))

# --- Check new extract_ebx result ---
df = raw.copy()
new_results = extract_ebx(df)
new_xchecks = [r['X-Check Number'] for r in new_results]
print(f"\nTotal X-Checks from new extract_ebx: {len(new_results)}")
print(f"'{TARGET}' in new results: {TARGET in new_xchecks}")

# --- Check old EBX extraction output ---
OLD_EBX = r"C:\Users\NICK.PULLAR\OneDrive - Zurich Insurance\Projects\Testing Automation\Python Files\Full_X-Checks\test_data\X-Checks Output\EBXExtraction 20260512 112837.xlsx"
old_df = pd.read_excel(OLD_EBX, sheet_name='Sheet').astype(str)
print(f"\nTotal X-Checks from old EBXExtraction: {len(old_df)}")
print(f"'{TARGET}' in old results: {TARGET in old_df['X-Check Number'].values}")
if TARGET in old_df['X-Check Number'].values:
    row = old_df[old_df['X-Check Number'] == TARGET].iloc[0]
    print(f"  Formula  : {row['EBX Formula']}")
    print(f"  Variables: {row['EBX Variables']}")

# --- Show what account rows look like for CON_UK_CH in the preprocessed df ---
print(f"\nPreprocessed rows for '{TARGET}':")
df2 = raw.copy().astype(str)
df2 = df2.reset_index().fillna('')
target_pre = df2[df2['X-Check No.'] == TARGET]
print(f"  Count: {len(target_pre)}")
if len(target_pre):
    print(target_pre[['index', 'X-Check No.', 'Account No.', 'SubA No.', 'Operator (X-Check Term)']].to_string(index=False))
