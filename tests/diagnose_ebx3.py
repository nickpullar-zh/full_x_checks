"""
Traces the dict_account built inside extract_ebx for target X-Checks
by temporarily patching _group_accounts to log its input.
"""
import sys
import copy
sys.path.insert(0, r"C:\Users\NICK.PULLAR\OneDrive - Zurich Insurance\Projects\Testing Automation\Python Files\Full_X-Checks")
sys.path.insert(0, r"C:\Users\NICK.PULLAR\OneDrive - Zurich Insurance\Projects\Testing Automation\Python Files\X-Checks")

import pandas as pd
import strategies.x_checks.ebx_extraction as ebx_mod
from EBXExtraction1 import group_accounts, create_variable

EBX_FILE = r"C:\Users\NICK.PULLAR\OneDrive - Zurich Insurance\Projects\Testing Automation\Python Files\Full_X-Checks\test_data\20251205 EPM X-Checks - Original.xlsx"
TARGETS  = ['A800_00', 'AE100_70', 'AE101_70', 'AE102_70', 'AL005_17', 'AL006_17', 'LR048_17']

captured = {}

# Patch _group_accounts to log what dict_account it receives for target X-Checks
original_group = ebx_mod._group_accounts
current_xcheck = [None]

def patched_group(dict_account):
    xcheck = current_xcheck[0]
    if xcheck in TARGETS:
        captured[xcheck] = copy.deepcopy(dict_account)
    return original_group(dict_account)

ebx_mod._group_accounts = patched_group

# Also patch the X-Check name tracker — hook into extract_ebx by rebuilding it inline
# Instead, just run extract_ebx and rely on the patch above being called at the right time.
# We need to track which X-Check is "current" — do this by also patching _create_formula.
original_formula = ebx_mod._create_formula
def patched_formula(dvars, absX, row):
    return original_formula(dvars, absX, row)
ebx_mod._create_formula = patched_formula

# Run new extract_ebx with raw DataFrame (no pre-processing — extract_ebx does it internally)
raw_df = pd.read_excel(EBX_FILE, sheet_name='cross checks all')

# We need to capture str_name at the time _group_accounts is called.
# Simplest: re-implement the loop capture inline.
# Reset patch and do it properly by re-reading extract_ebx source logic.
ebx_mod._group_accounts = original_group

# Manual trace: reproduce the dict_account-building loop for target X-Checks
df = raw_df.copy()
df = df.astype(str).reset_index().fillna('')

str_name = ''
dict_account = {}
dict_sub_accounts = {'SubAccounts': [], 'Operators': []}

for index, row in df.iterrows():
    if row['Account No.'] == 'nan':
        continue

    if str_name != row['X-Check No.']:
        # Before resetting, save dict_account for previous X-Check if it's a target
        if str_name in TARGETS:
            captured[str_name] = copy.deepcopy(dict_account)

        dict_account = {}
        dict_sub_accounts = {'SubAccounts': [], 'Operators': []}
        str_name = row['X-Check No.']
        dict_sub_accounts['SubAccounts'] = [[row['SubA No.'], row['Operator (X-Check Term)']]]
        dict_sub_accounts['Operators'] = [row['Operator (X-Check Term)']]
        dict_account[row['Account No.']] = dict_sub_accounts
    else:
        if row['Account No.'] not in dict_account:
            dict_sub_accounts = {'SubAccounts': [], 'Operators': []}
        dict_sub_accounts['SubAccounts'].append([row['SubA No.'], row['Operator (X-Check Term)']])
        if row['Operator (X-Check Term)'] not in dict_sub_accounts['Operators']:
            dict_sub_accounts['Operators'].append(row['Operator (X-Check Term)'])
        dict_account[row['Account No.']] = dict_sub_accounts

    if len(df) - 1 >= index:
        if len(df) - 1 == index or str_name != str(df['X-Check No.'][index + 1]):
            if str_name in TARGETS:
                captured[str_name] = copy.deepcopy(dict_account)

# Last row
if str_name in TARGETS:
    captured[str_name] = copy.deepcopy(dict_account)

print("dict_account captured from new extract_ebx loop:\n")
for xcheck in TARGETS:
    if xcheck in captured:
        da = captured[xcheck]
        print(f"{xcheck}: {len(da)} accounts")
        for acct, v in da.items():
            print(f"  {acct}: SubAccounts={v['SubAccounts'][:3]}{'...' if len(v['SubAccounts'])>3 else ''}, Ops={v['Operators']}")

        old_groups = group_accounts(copy.deepcopy(da))
        old_vars   = create_variable(old_groups)
        old_string = '|'.join(v['Variable'] for v in old_vars.values())
        print(f"  -> OLD string: {old_string[:120]}")

        new_groups = ebx_mod._group_accounts(copy.deepcopy(da))
        new_vars   = ebx_mod._create_variable(new_groups)
        new_string = '|'.join(v['Variable'] for v in new_vars.values())
        print(f"  -> NEW string: {new_string[:120]}")
        print()
    else:
        print(f"{xcheck}: NOT CAPTURED")
