import sys
import copy
sys.path.insert(0, r"C:\Users\NICK.PULLAR\OneDrive - Zurich Insurance\Projects\Testing Automation\Python Files\Full_X-Checks")
sys.path.insert(0, r"C:\Users\NICK.PULLAR\OneDrive - Zurich Insurance\Projects\Testing Automation\Python Files\X-Checks")

import pandas as pd
from EBXExtraction1 import group_accounts, create_variable
from strategies.x_checks.ebx_extraction import _group_accounts, _create_variable

EBX_FILE = r"C:\Users\NICK.PULLAR\OneDrive - Zurich Insurance\Projects\Testing Automation\Python Files\Full_X-Checks\test_data\20251205 EPM X-Checks - Original.xlsx"
TARGETS  = ['A800_00', 'AE100_70', 'AE101_70', 'AE102_70', 'AL005_17', 'AL006_17', 'LR048_17']

df = pd.read_excel(EBX_FILE, sheet_name='cross checks all')
df = df.astype(str).reset_index().fillna('')

for xcheck in TARGETS:
    rows = df[df['X-Check No.'] == xcheck]
    print(f"\n{'='*80}")
    print(f"X-Check: {xcheck}  ({len(rows)} rows)")
    print(f"{'='*80}")

    # Build dictAccount exactly as both pipelines do
    dict_account = {}
    for _, row in rows.iterrows():
        if row['Account No.'] == 'nan':
            continue
        acct = row['Account No.']
        sub  = row['SubA No.']
        op   = row['Operator (X-Check Term)']
        if acct not in dict_account:
            dict_account[acct] = {'SubAccounts': [[sub, op]], 'Operators': [op]}
        else:
            dict_account[acct]['SubAccounts'].append([sub, op])
            if op not in dict_account[acct]['Operators']:
                dict_account[acct]['Operators'].append(op)

    # Deep copy for each call so mutations don't carry over
    old_groups = group_accounts(copy.deepcopy(dict_account))
    new_groups = _group_accounts(copy.deepcopy(dict_account))

    old_vars   = create_variable(old_groups)
    new_vars   = _create_variable(new_groups)

    old_string = '|'.join(v['Variable'] for v in old_vars.values())
    new_string = '|'.join(v['Variable'] for v in new_vars.values())

    print(f"\nOLD groups ({len(old_groups)}):")
    for k, v in old_groups.items():
        print(f"  {k}: Accounts={v['Accounts']}, SubAccounts={v['SubAccounts']}, Ops={v['Operators']}")

    print(f"\nNEW groups ({len(new_groups)}):")
    for k, v in new_groups.items():
        print(f"  {k}: Accounts={v['Accounts']}, SubAccounts={v['SubAccounts']}, Ops={v['Operators']}")

    print(f"\nOLD: {old_string}")
    print(f"NEW: {new_string}")
    print(f"\nMATCH: {old_string == new_string}")
