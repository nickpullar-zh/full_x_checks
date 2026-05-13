"""
FIP Text File Extraction

Parses the raw FIP Validation Rule text output and extracts:
- X-Check Number
- FIP Formula
- FIP Variables (pipe-delimited string of variable definitions)
"""

from itertools import groupby
import re

from .variable_builder import build_variables_string


def extract_fip(fip_text: str, x_check_list: list[str]) -> list[dict]:
    """
    Main entry point. Parses FIP text and returns a list of dicts:
    [{"X-Check Number": ..., "FIP Formula": ..., "FIP Variables": ...}, ...]

    Args:
        fip_text:       The raw merged FIP text content (already in memory)
        x_check_list:   List of X-Check numbers to search for (from EBX file)

    Returns:
        List of dicts, one per X-Check found in the FIP text
    """
    clean_lines = _clean_text(fip_text)
    results = _parse_x_checks(clean_lines, x_check_list)
    return results


def _clean_text(raw_text: str) -> dict[int, str]:
    """
    Cleans the raw FIP text:
    - Removes period thousands separators between digits
    - Collapses repeated | and - characters
    - Strips leading pipe-space patterns
    - Removes empty/whitespace-only lines
    """
    counter = 0
    cleaned = {}

    for line in raw_text.splitlines():
        line = re.sub(r'(\d)\.(\d)', r'\1\2', line)
        line = _replace_multiple_occurrence(line, '-')
        line = _replace_multiple_occurrence(line, '|')
        line = line.replace('| ', '')
        line = " ".join(line.split())
        line = line.replace('- ', '-')
        if line != '' and line != '\n':
            cleaned[counter] = line
            counter += 1

    return cleaned


def _replace_multiple_occurrence(string: str, ch: str) -> str:
    """Replaces consecutive occurrences of a character with a single instance."""
    groups = groupby(string)
    new_str = ""
    for key, group in groups:
        if key == ch:
            new_str += ch
        else:
            new_str += ''.join(group)
    return new_str


def _parse_x_checks(clean_lines: dict[int, str], x_check_list: list[str]) -> list[dict]:
    """
    Iterates through cleaned lines, finds each X-Check block,
    extracts formula and variables.
    """
    results = []
    remaining_checks = list(x_check_list)

    end_of_block1 = '|-Segment @28@ * |'
    end_of_block2 = '-|'

    x_check_block = {}
    counter = 0
    message_found = False
    x_check_found = False
    evaluate_x_check_block = False
    current_x_check = ''

    for line in clean_lines.keys():
        if not evaluate_x_check_block:
            for x_check in remaining_checks:
                if x_check + ' ' + x_check + ' ' + x_check in clean_lines[line]:
                    current_x_check = x_check
                    evaluate_x_check_block = False
                    x_check_found = True
                    remaining_checks.remove(x_check)
                    break

            while x_check_found:
                try:
                    x_check_block[counter] = clean_lines[line]
                except (KeyError, IndexError):
                    break
                counter += 1

                if clean_lines[line] == end_of_block1:
                    message_found = True
                if clean_lines[line] == end_of_block2 and message_found:
                    message_found = False
                    evaluate_x_check_block = True
                    x_check_found = False
                    break
                line += 1
        else:
            returned_x_check = _get_x_check_information(x_check_block)
            counter = 0
            x_check_block = {}
            evaluate_x_check_block = False

            str_output = ''
            raw_variables = []
            for value in returned_x_check['Variables'].values():
                if value['MovementTypes']:
                    movement_type = '^'.join(sorted(value['MovementTypes']))
                else:
                    movement_type = '<blank>'
                str_output += (
                    'Name:' + value['Variable'] +
                    ';FS Account:' + '^'.join(sorted(value['FSAccounts'])) +
                    ';Movement Types:' + movement_type + '|'
                )
                raw_variables.append({
                    "fs_accounts": value['FSAccounts'],
                    "movement_types": value['MovementTypes'],
                })

            results.append({
                "X-Check Number": current_x_check,
                "FIP Formula": returned_x_check['Formula'],
                "FIP Variables": str_output[:-1],
                "FIP Variable (Builder)": build_variables_string(raw_variables),
            })

    return results


def _get_x_check_information(x_check_block: dict) -> dict:
    """
    Parses a single X-Check block to extract Formula and Variables
    (with FS Accounts and Movement Types for each variable).
    """
    bool_formula = False
    bool_variable = False
    bool_fs_account = False
    bool_movement_type = False
    bool_alc = False
    bool_mat = False
    str_formula = ''
    str_variables = ''
    arr_fs_accounts = []
    arr_movement_types = []

    dict_all_data = {'Formula': '', 'Variables': {}}
    dict_information = {'Variable': '', 'FSAccounts': [], 'MovementTypes': []}
    counter = 0

    for line in x_check_block:
        if x_check_block[line] == '|Formula String |':
            bool_formula = True
            continue
        elif x_check_block[line] == '|-Characteristic Sel Opt Attributes Node Characteristic From To |':
            bool_variable = True
            continue

        if bool_formula:
            if x_check_block[line] != '|-|':
                if x_check_block[line] == '-|':
                    bool_formula = False
                    str_formula = str_formula.replace(' ', '').replace('MAT', 'ToM').replace('ALC', 'ToM').replace('REX', 'ToM')
                    continue
                str_formula += x_check_block[line].replace('|', '').replace('MAT', 'ToM').replace('ALC', 'ToM').replace('REX', 'ToM')

        if bool_variable:
            if x_check_block[line] != '|-|':
                if x_check_block[line] == '|':
                    bool_variable = False
                    bool_fs_account = True
                    continue
                else:
                    x_check_block[line] = x_check_block[line].replace('MAT', 'ToM').replace('ALC', 'ToM').replace('REX', 'ToM')
                    str_variables = x_check_block[line].replace(' |', '').replace(' ', '')

        if bool_fs_account:
            if x_check_block[line].__contains__('Ass. / liab.category'):
                bool_movement_type = True
                bool_alc = True
            elif _safe_split(x_check_block[line], 0).__contains__('Maturity'):
                bool_movement_type = True
                bool_mat = True
            elif x_check_block[line] == '|-FS Account |':
                break
            elif x_check_block[line] == '-|':
                dict_information['Variable'] = str_variables
                dict_information['FSAccounts'] = arr_fs_accounts
                dict_information['MovementTypes'] = arr_movement_types
                dict_all_data['Variables'][counter] = dict_information
                counter += 1
                arr_fs_accounts = []
                arr_movement_types = []
                dict_information = {'Variable': '', 'FSAccounts': [], 'MovementTypes': []}
                break
            elif not x_check_block[line].__contains__('Movement Type'):
                if x_check_block[line] == '|':
                    bool_fs_account = False
                    bool_variable = True
                    dict_information['Variable'] = str_variables
                    dict_information['FSAccounts'] = arr_fs_accounts
                    dict_information['MovementTypes'] = arr_movement_types
                    dict_all_data['Variables'][counter] = dict_information
                    counter += 1
                    arr_fs_accounts = []
                    arr_movement_types = []
                    dict_information = {'Variable': '', 'FSAccounts': [], 'MovementTypes': []}
                    continue
                else:
                    if (_should_skip_line(x_check_block[line])):
                        continue
                    if x_check_block[line] != '|-Segment @28@ * |':
                        if _safe_split(x_check_block[line], 3) == 'FS' or _safe_split(x_check_block[line], 3) == 'Business':
                            arr_fs_accounts.append(_safe_split(x_check_block[line], 5))
                        elif _safe_split(x_check_block[line], 3) == 'Account' or _safe_split(x_check_block[line], 3).__contains__('@'):
                            continue
                        elif _safe_split(x_check_block[line], 0).__contains__('Rev./Exp.'):
                            arr_fs_accounts.append(_safe_split(x_check_block[line], 2))
                        else:
                            arr_fs_accounts.append(
                                _safe_split(x_check_block[line], 3).replace('MAT', 'ToM').replace('ALC', 'ToM').replace('REX', 'ToM')
                            )
            else:
                bool_fs_account = False
                bool_movement_type = True

        if bool_movement_type:
            if _should_skip_line(x_check_block[line]):
                continue

            if bool_alc:
                arr_movement_types.append(_safe_split(x_check_block[line], 4))
                bool_movement_type = False
            elif bool_mat:
                arr_movement_types.append(_safe_split(x_check_block[line], 2))
                bool_movement_type = False
            elif x_check_block[line] == '-|':
                dict_information['Variable'] = str_variables
                dict_information['FSAccounts'] = arr_fs_accounts
                dict_information['MovementTypes'] = arr_movement_types
                dict_all_data['Variables'][counter] = dict_information
                counter += 1
                dict_information = {'Variable': '', 'FSAccounts': [], 'MovementTypes': []}
                arr_fs_accounts = []
                arr_movement_types = []
                break
            elif x_check_block[line] == '|':
                bool_movement_type = False
                bool_variable = True
                dict_information['Variable'] = str_variables
                dict_information['FSAccounts'] = arr_fs_accounts
                dict_information['MovementTypes'] = arr_movement_types
                dict_all_data['Variables'][counter] = dict_information
                counter += 1
                dict_information = {'Variable': '', 'FSAccounts': [], 'MovementTypes': []}
                arr_fs_accounts = []
                arr_movement_types = []
                continue
            else:
                if x_check_block[line] != '|-Segment @28@ * |':
                    arr_movement_types.append(_safe_split(x_check_block[line], 3))

    dict_all_data['Formula'] = str_formula
    return dict_all_data


def _should_skip_line(line: str) -> bool:
    """Returns True if the line contains metadata that should be ignored."""
    return (
        'Partner Unit' in line or
        'ConsGroup' in line or
        ('*' in line and 'Segment' not in line) or
        _safe_split(line, 0).__contains__('Version') or
        _safe_split(line, 0).__contains__('GAAP') or
        'Posting Level' in line or
        'Doc Type' in line or
        'Line of Business' in line
    )


def _safe_split(line: str, index: int, default: str = '') -> str:
    """Returns token at position index, or default if the line has fewer tokens."""
    tokens = line.split()
    return tokens[index] if len(tokens) > index else default