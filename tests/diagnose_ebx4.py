import sys
sys.path.insert(0, r"C:\Users\NICK.PULLAR\OneDrive - Zurich Insurance\Projects\Testing Automation\Python Files\Full_X-Checks")
import pandas as pd

EBX_FILE = r"C:\Users\NICK.PULLAR\OneDrive - Zurich Insurance\Projects\Testing Automation\Python Files\Full_X-Checks\test_data\20251205 EPM X-Checks - Original.xlsx"
TARGET = 'AE100_70'

df = pd.read_excel(EBX_FILE, sheet_name='cross checks all')
df = df.astype(str).reset_index().fillna('')

# Find all rows where X-Check No. == TARGET
target_rows = df[df['X-Check No.'] == TARGET]
first_idx = target_rows.index.min()
last_idx  = target_rows.index.max()

print(f"TARGET: {TARGET}")
print(f"Total rows where X-Check No. == '{TARGET}': {len(target_rows)}")
print(f"Row range: {first_idx} to {last_idx}")
print()

# Print the full row sequence from 5 before to 5 after, showing X-Check No. and Account No.
start = max(0, first_idx - 3)
end   = min(len(df) - 1, last_idx + 3)

print(f"Row sequence (index {start} to {end}):")
print(f"{'idx':>5}  {'X-Check No.':<15}  {'Account No.':<15}  trigger?")
print("-" * 60)
for idx in range(start, end + 1):
    row = df.iloc[idx]
    xcheck = row['X-Check No.']
    acct   = row['Account No.']
    # What would the look-ahead trigger show at this index?
    if idx < len(df) - 1:
        next_xcheck = str(df['X-Check No.'][idx + 1])
        trigger = "(TRIGGER)" if xcheck != 'nan' and acct != 'nan' and xcheck != next_xcheck else ""
    else:
        trigger = "(LAST ROW)"
    marker = "***" if xcheck == TARGET else "   "
    print(f"{marker}{idx:>5}  {xcheck:<15}  {acct:<15}  {trigger}")
