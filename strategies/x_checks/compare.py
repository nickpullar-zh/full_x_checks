"""
X-Checks Comparison

Compares EBX and FIP extraction results by X-Check Number.
Returns a list of row dicts ready to be turned into a DataFrame and written
to Excel via BaseStrategy.write_excel_output().

Columns produced:
    X-Check Number, Formula Match,
    EBX Formula, FIP Formula,
    Variables Match, EBX Variables, FIP Variables,
    Variables Match (Builder), FIP Variable (Builder)

Variables Match (Builder) compares EBX Variables against FIP Variable (Builder),
providing a fallback check when the primary variables comparison does not match.
"""

import re


def compare(ebx_results: list[dict], fip_results: list[dict]) -> list[dict]:
    """
    Compares EBX and FIP extraction results.

    For each X-Check:
    - Compares formulas (with variable-reorder fallback)
    - Compares variables: primary (EBX vs FIP) and builder fallback (EBX vs FIP Builder)

    X-Checks present in only one file are included with 'Not Found' for the
    missing side.
    """
    rows = []
    # Keep first occurrence per X-Check — matches old Compare_Files.py behaviour (EBXFile.index[...][0])
    ebx_by_xcheck = {}
    for r in ebx_results:
        if r['X-Check Number'] not in ebx_by_xcheck:
            ebx_by_xcheck[r['X-Check Number']] = r
    matched_xchecks = set()

    for fip in fip_results:
        xcheck      = fip['X-Check Number']
        fip_formula = fip['FIP Formula'].replace('TOM', 'ToM')
        fip_vars    = fip['FIP Variables']
        fip_builder = fip.get('FIP Variable (Builder)', '')

        if xcheck not in ebx_by_xcheck:
            rows.append(_not_found_row(
                xcheck,
                ebx_formula='Not Found', fip_formula=fip_formula,
                ebx_vars='Not Found',    fip_vars=fip_vars,
                fip_builder=fip_builder,
                missing_side='EBX',
            ))
            continue

        matched_xchecks.add(xcheck)
        ebx = ebx_by_xcheck[xcheck]

        ebx_formula = ebx['EBX Formula']
        ebx_vars    = ebx['EBX Variables']

        formula_match, normalised_fip_formula = _compare_formulas(fip_formula, ebx_formula)

        fip_formula_excl = fip.get('FIP Formula (Excl)', fip_formula).replace('TOM', 'ToM')
        ebx_formula_excl = ebx.get('EBX Formula (Excl)', ebx_formula)
        excl_match, _ = _compare_formulas(fip_formula_excl, ebx_formula_excl)

        rows.append({
            'X-Check Number':            xcheck,
            'Formula Match':             'Match' if formula_match else 'MisMatch',
            'EBX Formula':               ebx_formula,
            'FIP Formula':               normalised_fip_formula,
            'Formula Match (Excl)':      'Match' if excl_match else 'MisMatch',
            'EBX Formula (Excl)':        ebx_formula_excl,
            'FIP Formula (Excl)':        fip_formula_excl,
            'Variables Match':           'Match' if _compare_variables(fip_vars, ebx_vars) else 'MisMatch',
            'EBX Variables':             ebx_vars,
            'FIP Variables':             fip_vars,
            'Variables Match (Builder)': 'Match' if _compare_variables(fip_builder, ebx_vars) else 'MisMatch',
            'FIP Variable (Builder)':    fip_builder,
        })

    for ebx in ebx_results:
        xcheck = ebx['X-Check Number']
        if xcheck not in matched_xchecks:
            rows.append(_not_found_row(
                xcheck,
                ebx_formula=ebx['EBX Formula'], fip_formula='Not Found',
                ebx_vars=ebx['EBX Variables'],  fip_vars='Not Found',
                fip_builder='Not Found',
                missing_side='FIP',
            ))

    return rows


def _not_found_row(
    xcheck: str,
    ebx_formula: str, fip_formula: str,
    ebx_vars: str,    fip_vars: str,
    fip_builder: str,
    missing_side: str,
) -> dict:
    return {
        'X-Check Number':            xcheck,
        'Formula Match':             'Not Found',
        'EBX Formula':               ebx_formula,
        'FIP Formula':               fip_formula,
        'Formula Match (Excl)':      'Not Found',
        'EBX Formula (Excl)':        ebx_formula,
        'FIP Formula (Excl)':        fip_formula,
        'Variables Match':           'Not Found',
        'EBX Variables':             ebx_vars,
        'FIP Variables':             fip_vars,
        'Variables Match (Builder)': 'Not Found',
        'FIP Variable (Builder)':    fip_builder,
    }


def _compare_formulas(fip_formula: str, ebx_formula: str) -> tuple[bool, str]:
    """
    Returns (match, normalised_fip_formula).

    Attempts to reorder FIP variables to align with EBX variable order before
    comparing. The normalised formula (post-reorder attempt) is returned so the
    caller can store it in the output, matching the behaviour of the original
    Compare_Files.py which wrote the modified formula to the output file.
    """
    if fip_formula.lower() == ebx_formula.lower():
        return True, fip_formula

    fip_vars = re.findall(r'VAL_YTD\((.*?)\)', fip_formula)
    ebx_vars = re.findall(r'VAL_YTD\((.*?)\)', ebx_formula)

    # Addition-only formula: reorder FIP variables to match EBX order
    if ')-V' not in fip_formula and sorted(fip_vars) == sorted(ebx_vars) and '+' in fip_formula:
        bracket_vars = re.findall(r'\(VAL_YTD\((.*?)\)\)', fip_formula)
        if not bracket_vars:
            bracket_vars = re.findall(r'VAL_YTD\((.*?)\)', fip_formula)
        if bracket_vars:
            new_fip = '+'.join(f'VAL_YTD({v})' for v in ebx_vars)
            fip_formula = fip_formula.replace('VAL_YTD(' + str(bracket_vars[0]) + ')', new_fip, 1)

    # Single-minus formula: reorder FIP variables to match EBX order
    elif fip_formula.count(')-V') == 1 and sorted(fip_vars) == sorted(ebx_vars):
        bracket_vars = re.findall(
            r'(VAL_YTD\((.*?)\))',
            fip_formula.replace('P_VAL_PER', 'VAL_YTD').replace(",'0','1'", '')
        )
        if bracket_vars and len(ebx_vars) >= 2:
            new_fip = f'VAL_YTD({ebx_vars[0]})-VAL_YTD({ebx_vars[1]})'
            fip_formula = fip_formula.replace('VAL_YTD(' + str(bracket_vars[0]) + ')', new_fip, 1)

    return fip_formula.lower() == ebx_formula.lower(), fip_formula


def _compare_variables(fip_vars: str, ebx_vars: str) -> bool:
    """Returns True if the variable sets match (order-insensitive, case-insensitive)."""
    fip_parts = [v for v in fip_vars.lower().split('|') if v]
    ebx_parts = [v for v in ebx_vars.lower().split('|') if v]
    return sorted(fip_parts) == sorted(ebx_parts)
