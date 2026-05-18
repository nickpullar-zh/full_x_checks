"""
EBX Excel File Extraction

Reads the EBX 'cross checks all' sheet and extracts:
- X-Check Number
- EBX Formula
- EBX Variables (pipe-delimited string of variable definitions)
"""

import pandas as pd

from .variable_builder import build_variables_string


def extract_ebx(df: pd.DataFrame, qu_accounts: set | None = None) -> list[dict]:
    """
    Main entry point. Processes the EBX DataFrame and returns a list of dicts:
    [{"X-Check Number": ..., "EBX Formula": ..., "EBX Variables": ...}, ...]

    Args:
        df: DataFrame from the 'cross checks all' sheet

    Returns:
        List of dicts, one per X-Check number found in the EBX file
    """
    # Normalise — fillna before astype so NaN becomes '' not the string 'nan'
    df = df.fillna('')
    df = df.astype(str)
    df = df.reset_index()

    results = []
    str_name = ''
    dict_account = {}
    dict_sub_accounts = {'SubAccounts': [], 'Operators': []}
    bool_absolute_x = False

    for index, row in df.iterrows():
        # Skip rows without an account number
        if row['Account No.'] == '':
            continue

        # New X-Check number detected
        if str_name != row['X-Check No.']:
            bool_absolute_x = row['Absolute (result)'] == 'X'
            dict_account = {}
            dict_sub_accounts = {'SubAccounts': [], 'Operators': []}
            str_name = row['X-Check No.']
            dict_sub_accounts['SubAccounts'] = [[row['SubA No.'], row['Operator (X-Check Term)']]]
            dict_sub_accounts['Operators'] = [row['Operator (X-Check Term)']]
            dict_account[row['Account No.']] = dict_sub_accounts
        else:
            # New account number within same X-Check
            if row['Account No.'] not in dict_account:
                dict_sub_accounts = {'SubAccounts': [], 'Operators': []}
            dict_sub_accounts['SubAccounts'].append([row['SubA No.'], row['Operator (X-Check Term)']])
            if row['Operator (X-Check Term)'] not in dict_sub_accounts['Operators']:
                dict_sub_accounts['Operators'].append(row['Operator (X-Check Term)'])
            dict_account[row['Account No.']] = dict_sub_accounts

        # Process when we reach the last row of the current X-Check
        if len(df) - 1 == index or str_name != str(df['X-Check No.'][index + 1]):
            # Group accounts into variables
            dict_variables = _group_accounts(dict_account)
            # Create variable definitions from groups
            dict_variables_output = _create_variable(dict_variables)

            dict_formula_variables = []

            for value in dict_variables_output.values():
                dict_formula_variables.append({
                    'Name':          str_name,
                    'Variable-Name': value['Variable-Name'],
                    'Operator':      value['Operator']
                })

            use_lc      = _should_use_lc(row)
            use_qu      = _should_use_qu(dict_account, qu_accounts)
            use_pct     = _should_use_pct(row)
            str_formula = _create_formula(dict_formula_variables, bool_absolute_x, row, use_lc, use_qu, use_pct)

            raw_variables = [
                {'fs_accounts': item['Accounts'], 'movement_types': item['SubAccounts']}
                for item in dict_variables.values()
            ]
            str_output_string = build_variables_string(raw_variables)

            results.append({
                "X-Check Number": str_name,
                "EBX Formula":    str_formula,
                "EBX Variables":  str_output_string,
            })

    return results


def _should_use_qu(dict_account: dict, qu_accounts: set | None) -> bool:
    """Returns True if any account for this X-Check has Data type QU in the GCoA file."""
    if not qu_accounts:
        return False
    return any(acct in qu_accounts for acct in dict_account)


def _should_use_pct(row) -> bool:
    """Returns True when the '%' column is marked X, meaning the limit is a percentage."""
    return str(row.get('%', '')).strip() == 'X'


def _should_use_lc(row) -> bool:
    """
    Returns True when the formula should use LC_YTD/CONST_LC instead of VAL_YTD/CONST.
    Triggered by Category = "Shareholders' Equity".

    Note: Version Spanning Validation is NOT a reliable trigger — it is populated for
    Reinsurance Asset Check and SST-only categories that correctly use VAL_YTD in FIP.
    """
    category = str(row.get('Category', '')).strip()
    return category == "Shareholders' Equity"


def _create_formula(dict_formula_variables: list, bool_absolute_x: bool, row,
                    use_lc: bool = False, use_qu: bool = False, use_pct: bool = False) -> str:
    """
    Builds the formula string from the variable list and row operators/limits.
    use_qu takes priority: QU accounts use QU_YTD.
    use_lc (Shareholders' Equity) uses LC_YTD and CONST_LC.
    use_pct ('%%' column == X): right-hand side formatted as '<limit>,000000%%' instead of CONST().
    Default uses VAL_YTD and CONST.
    """
    val_fn   = 'QU_YTD'  if use_qu else ('LC_YTD'   if use_lc else 'VAL_YTD')
    const_fn = 'CONST_LC' if use_lc else 'CONST'

    str_left_side = ''
    str_right_side = ''
    str_comparator = ''
    log_left_side_abs = False

    # Build left hand side
    for item in dict_formula_variables:
        operator = item.get('Operator', '+')
        variable_name = item.get('Variable-Name', '')
        if str_left_side == '':
            str_left_side = val_fn + '(' + variable_name + ')'
        else:
            str_left_side += operator + val_fn + '(' + variable_name + ')'

    # Build right hand side
    if row['Operator 2'] == '':
        str_comparator = row['Operator 1']
    else:
        str_comparator = row['Operator 2']
        log_left_side_abs = True

    if row['Limit 2'] != '':
        str_right_side = str(int(float(row['Limit 2'])))
    else:
        if row['Limit 1'] == '':
            str_right_side = '0'
        else:
            str_right_side = str(int(float(row['Limit 1'])))

    str_right_side = str_right_side.replace(',', '')

    if use_pct:
        str_right_side = f"'{str_right_side},000000%'"
    elif str_right_side != '0':
        str_right_side = const_fn + "(" + str_right_side + ",'USD','E')"

    if log_left_side_abs:
        str_left_side = 'ABS(' + str_left_side + ')'

    return str_left_side + str_comparator + str_right_side


def _group_accounts(dict_account: dict) -> dict:
    """
    Groups accounts by their sub-assignment patterns.
    Preserves original logic from group_accounts() exactly.
    """
    counter = 0
    bool_found = False
    dict_groups = {}
    dict_one_group = {'Accounts': [], 'SubAccounts': [], 'Operators': []}
    bool_two_operators = False

    # Handle empty subassignment with +
    if any(['', '+'] in value['SubAccounts'] for value in list(dict_account.values())):
        for key, value in dict_account.items():
            if value['SubAccounts'].__contains__(['', '+']):
                dict_one_group['Accounts'].append(key)
                dict_one_group['SubAccounts'] = []
                dict_one_group['Operators'].append('+')
                value['SubAccounts'].remove(['', '+'])
                if not any('+' in x for x in value['SubAccounts']):
                    value['Operators'].remove('+')
        dict_one_group['SubAccounts'] = [*set(sublist for sublist in dict_one_group['SubAccounts'])]
        dict_one_group['Operators'] = '+'
        dict_groups[counter] = dict_one_group
        dict_one_group = {'Accounts': [], 'SubAccounts': [], 'Operators': []}
        counter += 1

    # Handle empty subassignment with -
    if any(['', '-'] in value['SubAccounts'] for value in list(dict_account.values())):
        for key, value in dict_account.items():
            if value['SubAccounts'].__contains__(['', '-']):
                dict_one_group['Accounts'].append(key)
                dict_one_group['SubAccounts'] = []
                dict_one_group['Operators'].append('-')
                value['SubAccounts'].remove(['', '-'])
                if not any('-' in x for x in value['SubAccounts']):
                    value['Operators'].remove('-')
        dict_one_group['SubAccounts'] = [*set(sublist for sublist in dict_one_group['SubAccounts'])]
        dict_one_group['Operators'] = '-'
        dict_groups[counter] = dict_one_group
        dict_one_group = {'Accounts': [], 'SubAccounts': [], 'Operators': []}
        counter += 1

    # All other cases
    for key, value in dict_account.items():
        bool_found = False
        if dict_one_group['Accounts'] == [] and not value['SubAccounts'] == []:
            if value['Operators'].__contains__('+') and value['Operators'].__contains__('-'):
                bool_two_operators = True
                for mtype in value['SubAccounts']:
                    if mtype[1] == '-':
                        dict_one_group['Accounts'].append(key)
                        dict_one_group['SubAccounts'].append(mtype[0])
                        dict_one_group['Operators'] = ['-']
                        value['SubAccounts'].remove(mtype)
                dict_one_group['SubAccounts'] = [*set(sublist for sublist in dict_one_group['SubAccounts'])]
                dict_groups[counter] = dict_one_group
                dict_one_group = {'Accounts': [], 'SubAccounts': [], 'Operators': []}
                counter += 1
            dict_one_group['Accounts'] = [key]
            dict_one_group['SubAccounts'] = sorted([sublist[0] for sublist in value['SubAccounts']])
            arr_operators = [*set(sublist[1] for sublist in value['SubAccounts'])]
            dict_one_group['Operators'] = arr_operators

        elif not value['SubAccounts'] == []:
            if bool_two_operators:
                for mtype in value['SubAccounts']:
                    if dict_one_group['SubAccounts'].__contains__(mtype[0]):
                        dict_one_group['Accounts'].append(key)
                        dict_one_group['SubAccounts'].append(mtype[0])
                        dict_one_group['Operators'].append(mtype[1])
                    else:
                        for group_value in dict_groups.values():
                            if group_value['SubAccounts'].__contains__(mtype[0]):
                                group_value['Accounts'].append(key)
                                group_value['SubAccounts'].append(mtype[0])
                                group_value['Operators'].append(mtype[1])
                                group_value['SubAccounts'] = list(dict.fromkeys(group_value['SubAccounts']))
            else:
                if sorted(dict_one_group['SubAccounts']) == sorted([sublist[0] for sublist in value['SubAccounts']]) and dict_one_group['Operators'] == value['Operators']:
                    dict_one_group['Accounts'].append(key)
                    bool_found = True
                for group_key, group_value in dict_groups.items():
                    if sorted(group_value['SubAccounts']) == sorted([sublist[0] for sublist in value['SubAccounts']]) and group_value['Operators'] == value['Operators']:
                        dict_groups[group_key]['Accounts'].append(key)
                        bool_found = True
                if not bool_found:
                    dict_one_group['SubAccounts'] = [*set(sublist for sublist in dict_one_group['SubAccounts'])]
                    dict_groups[counter] = dict_one_group
                    dict_one_group = {'Accounts': [], 'SubAccounts': [], 'Operators': []}
                    counter += 1
                    dict_one_group['Accounts'] = [key]
                    dict_one_group['SubAccounts'] = sorted([sublist[0] for sublist in value['SubAccounts']])
                    arr_operators = [*set(sublist[1] for sublist in value['SubAccounts'])]
                    dict_one_group['Operators'] = arr_operators

    # Add last group if not empty
    dict_one_group['SubAccounts'] = [*set(sublist for sublist in dict_one_group['SubAccounts'])]
    if dict_one_group['Accounts'] != []:
        bool_found = False
        for key, value in dict_groups.items():
            if value['SubAccounts'] == dict_one_group['SubAccounts'] and dict_one_group['Operators'] == value['Operators']:
                value['Accounts'] = value['Accounts'] + dict_one_group['Accounts']
                bool_found = True
                break
        if not bool_found:
            dict_groups[counter] = dict_one_group

    # Remove empty groups
    dict_groups = {
        k: v for k, v in dict_groups.items()
        if v != {'Accounts': [], 'SubAccounts': [], 'Operators': []}
    }

    return dict_groups


def _create_variable(dict_groups: dict) -> dict:
    """
    Creates variable definitions from grouped accounts.
    Preserves original logic from create_variable() exactly.
    """
    dict_output = {}
    this_variable = {
        'Variable': '', 'Variable-Name': '', 'Variable-Output': '',
        'Accounts': [], 'Accounts-Output': [],
        'SubAccounts': [], 'SubAccounts-Output': [], 'Operator': ''
    }
    counter = 0

    for item in dict_groups.values():
        str_variable_name = sorted(item['Accounts'])[0]

        # Add ff if more than one account
        if len(item['Accounts']) > 1:
            str_variable_name = str_variable_name + 'ff'

        # Add ToM if subassignment exists
        if item['SubAccounts'] != ['']:
            if item['SubAccounts'] != []:
                str_variable_name = str_variable_name + 'ToM' + sorted(item['SubAccounts'])[0].replace('.0', '')
            if len(item['SubAccounts']) > 1:
                str_variable_name = str_variable_name + 'ff'

        this_variable['Variable-Name'] = str_variable_name
        this_variable['Variable-Output'] = 'Name:' + str_variable_name
        this_variable['Accounts'] = item['Accounts']

        if this_variable['Accounts']:
            this_variable['Accounts-Output'] = 'FS Account:' + '^'.join(sorted(item['Accounts']))
        else:
            this_variable['Accounts-Output'] = '<blank>'

        this_variable['SubAccounts'] = item['SubAccounts']

        if this_variable['SubAccounts']:
            this_variable['SubAccounts-Output'] = 'Movement Types:' + '^'.join(sorted(item['SubAccounts']))
        else:
            this_variable['SubAccounts-Output'] = 'Movement Types:<blank>'

        if item['Operators']:
            this_variable['Operator'] = str(item['Operators'][0])
        else:
            this_variable['Operator'] = ''

        this_variable['Variable'] = (
            this_variable['Variable-Output'] + ';' +
            this_variable['Accounts-Output'] + ';' +
            this_variable['SubAccounts-Output']
        )

        dict_output[counter] = this_variable
        this_variable = {
            'Variable': '', 'Variable-Name': '', 'Variable-Output': '',
            'Accounts': [], 'Accounts-Output': [],
            'SubAccounts': [], 'SubAccounts-Output': [], 'Operator': ''
        }
        counter += 1

    return dict(sorted(dict_output.items()))