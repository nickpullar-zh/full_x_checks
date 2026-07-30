"""
Logic-layer verification of all Logic test cases in the Fixture UAT plan.
Run from repo root: python test_data/run_logic_tests.py
"""
import sys, pandas as pd
sys.path.insert(0, '.')
from pathlib import Path
from strategies.x_checks.x_checks import XChecks

F = Path('test_data/fixtures')
results = []

def chk(case_id, desc, passed, detail=''):
    results.append((case_id, desc, 'PASS' if passed else 'FAIL', detail))


# ── FX-05: X-Checks comparison output ────────────────────────────────────────
from strategies.x_checks.ebx_extraction import extract_ebx
from strategies.x_checks.fip_extraction import extract_fip
from strategies.x_checks.compare import compare as xc_compare

ebx_df = pd.read_excel(F / 'xc_pub.xlsx', sheet_name='cross checks all')
ebx_results = extract_ebx(ebx_df)
xc_list = sorted(set(str(x) for x in ebx_df['X-Check No.'].tolist()
                     if str(x) not in ('nan', '', 'None')))
fip_text = (F / 'fip_xc.txt').read_text(encoding='utf-8')
fip_results = extract_fip(fip_text, xc_list)
xc_df = pd.DataFrame(xc_compare(ebx_results, fip_results)).sort_values('X-Check No.').reset_index(drop=True)

expected_xc = {
    'XC_ABS_FORMULA':       'Match',    # Operator 2 → ABS() wrapping
    'XC_ALL_MATCH':         'Match',    # standard full match
    'XC_ALL_MISMATCH':      'MisMatch', # formula AND variables wrong
    'XC_DIFF_EXCLUDED':     'Not Found',# Exclude Z-Core=X → no FIP block
    'XC_DIFF_GREEN':        'Match',    # Type of change=New (green) → in scope; FIP matches
    'XC_DIFF_INACTIVE':     'Not Found',# INACTIVE → no FIP block
    'XC_DIFF_IN_SCOPE':     'Match',    # Type of change=Changed (yellow) → in scope
    'XC_DIFF_NO_TOC':       'Not Found',# blank Type of change → no FIP block
    'XC_DIFF_ORANGE':       'Not Found',# Type of change=Removed → no FIP block
    'XC_DIFF_YELLOW':       'Match',    # Type of change=Changed (yellow fill) → in scope
    'XC_DIFF_YELLOW_CAT':   'Not Found',# yellow Category → no FIP block
    'XC_EXCL_MATCH':        'Match',    # @2A@ in FIP + excl in EBX → Excl=Match
    'XC_EXCL_MISMATCH':     'Match',    # EBX has excl, FIP has none → Formula=Match, Excl=MisMatch
    'XC_FF_SUFFIX':         'Match',    # two accounts → ff suffix
    'XC_FORMULA_MISMATCH':  'MisMatch', # operator differs (<=0 vs >=0); vars same
    'XC_GTE_OPERATOR':      'Match',    # >= operator
    'XC_KEL_MISMATCH':      'MisMatch', # MisMatch; annotated with reason when KEL supplied
    'XC_KEL_NO_MATCH':      'MisMatch', # KEL entry exists but wrong fingerprint → no annotation
    'XC_LC_CONST':          'Match',    # LC_YTD + CONST_LC
    'XC_LC_YTD':            'Match',    # Shareholders Equity → LC_YTD
    'XC_NONZERO_LIMIT':     'Match',    # non-zero limit → CONST(100,...)
    'XC_NOT_IN_EBX':        'Not Found',# FIP only
    'XC_NOT_IN_FIP':        'Not Found',# EBX only
    'XC_PCT_FORMAT':        'Match',    # % column → percentage format
    'XC_REORDER_MATCH':     'MisMatch', # known edge case in reorder logic
    'XC_REX_CORRECTION':    'Match',    # FIP uses REX → ToM
    'XC_SUBTRACT':          'Match',    # subtraction formula
    'XC_THOUSANDS_CORR':    'Match',    # 1.000 → 1000
    'XC_TOM_CORRECTION':    'Match',    # FIP uses TOM → ToM
    'XC_VARIABLE_MISMATCH': 'Match',    # formula matches, vars differ
}
expected_row_count = len(expected_xc)
chk('FX-05a', f'X-Checks row count = {expected_row_count}',
    len(xc_df) == expected_row_count, f'got {len(xc_df)}')
for xc_id, exp in expected_xc.items():
    rows = xc_df[xc_df['X-Check No.'] == xc_id]
    got = rows.iloc[0]['Formula Match'] if len(rows) else 'MISSING'
    chk('FX-05b', f'X-Checks {xc_id} Formula Match = {exp}', got == exp, got)

# Specific column checks
def _xc(xc_id):
    r = xc_df[xc_df['X-Check No.'] == xc_id]
    return r.iloc[0] if len(r) else None

# ── FX-06: column-level checks ────────────────────────────────────────────────
r = _xc('XC_VARIABLE_MISMATCH')
chk('FX-06a', 'XC_VARIABLE_MISMATCH Variables Match = MisMatch',
    r is not None and r['Variables Match'] == 'MisMatch', r['Variables Match'] if r is not None else 'MISSING')

r = _xc('XC_FORMULA_MISMATCH')
chk('FX-06b', 'XC_FORMULA_MISMATCH Variables Match = Match (operator differs, account same)',
    r is not None and r['Variables Match'] == 'Match', r['Variables Match'] if r is not None else 'MISSING')

r = _xc('XC_EXCL_MISMATCH')
chk('FX-06c', 'XC_EXCL_MISMATCH Formula Match = Match, Formula Match (Excl) = MisMatch',
    r is not None and r['Formula Match'] == 'Match' and r['Formula Match (Excl)'] == 'MisMatch',
    f"Formula={r['Formula Match']} Excl={r['Formula Match (Excl)']}" if r is not None else 'MISSING')

r = _xc('XC_EXCL_MATCH')
chk('FX-06d', 'XC_EXCL_MATCH Formula Match (Excl) = Match',
    r is not None and r['Formula Match (Excl)'] == 'Match',
    r['Formula Match (Excl)'] if r is not None else 'MISSING')

# ── FX-10: X-Check No Selection + Comparison filtering (differences mode) ────────
from strategies.x_checks.x_check_no_selection import select_x_check_nos

# Selection filter (drives .txt file)
selected = select_x_check_nos(ebx_df, str(F / 'xc_pub.xlsx'), 'cross checks all')
expected_selected = {'XC_DIFF_IN_SCOPE', 'XC_DIFF_YELLOW', 'XC_DIFF_GREEN'}
chk('FX-10a', 'X-Check No Selection: exactly 3 results', len(selected) == 3, str(selected))
chk('FX-10b', 'X-Check No Selection: XC_DIFF_IN_SCOPE present', 'XC_DIFF_IN_SCOPE' in selected, str(selected))
chk('FX-10c', 'X-Check No Selection: XC_DIFF_YELLOW present (Changed)', 'XC_DIFF_YELLOW' in selected, str(selected))
chk('FX-10d', 'X-Check No Selection: XC_DIFF_GREEN present (New)', 'XC_DIFF_GREEN' in selected, str(selected))
chk('FX-10e', 'X-Check No Selection: XC_DIFF_ORANGE absent (Removed)', 'XC_DIFF_ORANGE' not in selected, str(selected))
chk('FX-10f', 'X-Check No Selection: XC_DIFF_NO_TOC absent (blank TOC)', 'XC_DIFF_NO_TOC' not in selected, str(selected))
chk('FX-10g', 'X-Check No Selection: XC_DIFF_EXCLUDED absent (Z-Core)', 'XC_DIFF_EXCLUDED' not in selected, str(selected))

# Comparison sheet filtering (when diff=True, Comparison shows only in-scope rows)
in_scope_set = set(selected)
xc_df_diff = xc_df[xc_df['X-Check No.'].isin(in_scope_set)].reset_index(drop=True)
chk('FX-10h', 'Diff Comparison: XC_DIFF_YELLOW present with Match',
    'XC_DIFF_YELLOW' in xc_df_diff['X-Check No.'].values and
    xc_df_diff[xc_df_diff['X-Check No.'] == 'XC_DIFF_YELLOW'].iloc[0]['Formula Match'] == 'Match', '')
chk('FX-10i', 'Diff Comparison: XC_DIFF_GREEN present with Match',
    'XC_DIFF_GREEN' in xc_df_diff['X-Check No.'].values and
    xc_df_diff[xc_df_diff['X-Check No.'] == 'XC_DIFF_GREEN'].iloc[0]['Formula Match'] == 'Match', '')
chk('FX-10j', 'Diff Comparison: XC_DIFF_ORANGE absent', 'XC_DIFF_ORANGE' not in xc_df_diff['X-Check No.'].values, '')
chk('FX-10k', 'Diff Comparison: XC_DIFF_NO_TOC absent', 'XC_DIFF_NO_TOC' not in xc_df_diff['X-Check No.'].values, '')
chk('FX-10l', 'Diff Comparison: XC_ALL_MATCH absent (no Type of change)', 'XC_ALL_MATCH' not in xc_df_diff['X-Check No.'].values, '')

# ── FX-12: Grouping By comparison ─────────────────────────────────────────────
from strategies.grouping_by.grouping_by import GroupingBy
from task_configs import GROUPING_BY_UPLOAD_CONFIG
gb = GroupingBy(GROUPING_BY_UPLOAD_CONFIG)
gb.log = []
mapping_txt = (F / 'mapping.txt').read_text()
fip_gb = pd.read_excel(F / 'fip_ZQ9_VALFLDGR.xlsx', sheet_name='Sheet1')
loaded_gb = {
    GROUPING_BY_UPLOAD_CONFIG.file_fields[0].label: fip_gb,
    GROUPING_BY_UPLOAD_CONFIG.file_fields[1].label: ebx_df.copy(),
    GROUPING_BY_UPLOAD_CONFIG.file_fields[2].label: mapping_txt,
}
_, _, df_fip_proc = gb._process_fip(loaded_gb)
_, df_ebx_proc = gb._process_ebx(loaded_gb)
df_gb = gb._process_compare(df_fip_proc, df_ebx_proc)
chk('FX-12a', 'Grouping By row count = 2', len(df_gb) == 2, f'got {len(df_gb)}')
for key, exp in [('GB_MATCHED|GB_GROUPING_ITEM', 'Matched'),
                 ('GB_NOT_IN_FIP|GB_GROUPING_ITEM', 'Not in FIP')]:
    rows = df_gb[df_gb['EBX Key'] == key]
    got = rows.iloc[0]['Result'] if len(rows) else 'MISSING'
    chk('FX-12b', f'Grouping By {key} = {exp}', got == exp, got)

# ── FX-16: AP comparison ──────────────────────────────────────────────────────
from strategies.accounting_principles.validation_methods import parse_method_bindings
from strategies.accounting_principles.compare import compare_with_bindings
from strategies.accounting_principles.accounting_principles import DEFAULT_EVENTS
bindings = parse_method_bindings(str(F / 'validation_methods.xlsx'), DEFAULT_EVENTS)
fip_ap = pd.read_excel(F / 'fip_ZQ9_VALMSG.xlsx', sheet_name='FIP Methods Rules and Condition')
fip_ap['Key'] = (fip_ap['MK'].astype(str).str.strip()
                 + '|' + fip_ap['ValidRule'].astype(str).str.strip())
xchecks = [str(x).strip() for x in ebx_df['X-Check No.'].tolist()
           if str(x).strip() not in ('nan', '', 'None')]
ap_df = pd.DataFrame(compare_with_bindings(bindings, ebx_df, xchecks, fip_ap))
chk('FX-16a', 'AP row count = 2', len(ap_df) == 2, f'got {len(ap_df)}')
for xc, exp_match in [('AP_MATCH', 'Match'), ('AP_MISMATCH', 'MisMatch')]:
    rows = ap_df[ap_df['X-Check No.'] == xc]
    if len(rows) == 0:
        chk('FX-16b', f'AP {xc}', False, 'MISSING')
    else:
        got_m = rows.iloc[0]['Match']
        got_e = rows.iloc[0]['Event']
        chk('FX-16b', f'AP {xc} Match={exp_match} Event=IFRS New RFD',
            got_m == exp_match and got_e == 'IFRS New RFD',
            f'Match={got_m} Event={got_e}')

# ── FX-20: Conditions full file ───────────────────────────────────────────────
from strategies.conditions.extract import extract_conditions
from strategies.conditions.fip import process_fip
from strategies.conditions.compare import compare as cond_compare
fip_cond = pd.read_excel(F / 'fip_ZQ9_VALMETH.xlsx', sheet_name='FIP Conditions')
fip_proc = process_fip(fip_cond)
working_full, _ = extract_conditions(str(F / 'xc_pub.xlsx'), 'cross checks all',
                                     process_only_differences=False)
cond_full, _ = cond_compare(working_full, fip_proc)
chk('FX-20a', 'Conditions full row count = 7', len(cond_full) == 7, f'got {len(cond_full)}')
for ebx_key, exp in [
    ('COND_MATCHED|Q1',     'Matched'),
    ('COND_NOT_MATCHED|Q2', 'Not Matched'),
    ('COND_BASE|COND_BASE', 'Not Matched'),
    ('COND_BASE|Q1',        'Matched'),
    ('COND_DIFF_YELLOW|Q1', 'Matched'),
    ('COND_DIFF_GREEN|Q1',  'Matched'),
    ('COND_DIFF_WHITE|Q1',  'Not Matched'),
]:
    rows = cond_full[cond_full['EBX Data'] == ebx_key]
    got = rows.iloc[0]['Comparison'] if len(rows) else 'MISSING'
    chk('FX-20b', f'Conditions full {ebx_key} = {exp}', got == exp, got)

# ── FX-21: Conditions differences mode ───────────────────────────────────────
working_diff, _ = extract_conditions(str(F / 'xc_pub.xlsx'), 'cross checks all',
                                     process_only_differences=True)
cond_diff, _ = cond_compare(working_diff, fip_proc)
chk('FX-21a', 'Conditions diff row count = 2', len(cond_diff) == 2, f'got {len(cond_diff)}')
chk('FX-21b', 'Conditions diff COND_DIFF_YELLOW|Q1 Matched',
    len(cond_diff[cond_diff['EBX Data'] == 'COND_DIFF_YELLOW|Q1']) > 0 and
    cond_diff[cond_diff['EBX Data'] == 'COND_DIFF_YELLOW|Q1'].iloc[0]['Comparison'] == 'Matched', '')
chk('FX-21c', 'Conditions diff COND_DIFF_GREEN|Q1 Matched',
    len(cond_diff[cond_diff['EBX Data'] == 'COND_DIFF_GREEN|Q1']) > 0 and
    cond_diff[cond_diff['EBX Data'] == 'COND_DIFF_GREEN|Q1'].iloc[0]['Comparison'] == 'Matched', '')
chk('FX-21d', 'Conditions diff COND_DIFF_WHITE not collected',
    'COND_DIFF_WHITE|Q1' not in cond_diff['EBX Data'].values, '')

# ── FX-25: Full Run row counts ────────────────────────────────────────────────
# ── FX-09: X-Checks — Known Exception annotation ─────────────────────────────
from strategies.base_strategy import BaseStrategy
from unittest.mock import MagicMock
from task_configs import X_CHECKS_UPLOAD_CONFIG

xc_strategy = XChecks(X_CHECKS_UPLOAD_CONFIG)
xc_strategy.log = []
kel_path = str(F / 'known_exception_list.xlsx')

# Annotate the comparison df with the KEL
annotated = xc_strategy._annotate_known_exceptions(
    xc_df.copy(), kel_path, sheet_name='X-Checks',
    fingerprint_columns=[
        "X-Check No.", "EBX Formula", "FIP Formula",
        "EBX Formula (Excl)", "FIP Formula (Excl)",
        "EBX Variables", "FIP Variables", "FIP Variable (Builder)",
    ]
)
kel_row = annotated[annotated['X-Check No.'] == 'XC_KEL_MISMATCH']
kel_val = kel_row.iloc[0]['Known Exception'] if len(kel_row) else 'MISSING'
kel_formula = kel_row.iloc[0]['Formula Match'] if len(kel_row) else 'MISSING'

chk('FX-09a', 'KEL: XC_KEL_MISMATCH Formula Match still = MisMatch',
    kel_formula == 'MisMatch', kel_formula)
chk('FX-09b', 'KEL: XC_KEL_MISMATCH Known Exception reason populated',
    bool(kel_val) and kel_val not in ('', 'nan', 'MISSING'),
    repr(kel_val))
chk('FX-09c', 'KEL: XC_ALL_MATCH Known Exception is blank (not a mismatch)',
    annotated[annotated['X-Check No.'] == 'XC_ALL_MATCH'].iloc[0].get('Known Exception', '') == '',
    '')

# XC_KEL_NO_MATCH: KEL entry exists with wrong fingerprint → no annotation
no_match_row = annotated[annotated['X-Check No.'] == 'XC_KEL_NO_MATCH']
no_match_formula = no_match_row.iloc[0]['Formula Match'] if len(no_match_row) else 'MISSING'
no_match_kel = no_match_row.iloc[0].get('Known Exception', '') if len(no_match_row) else 'MISSING'
chk('FX-09d', 'KEL: XC_KEL_NO_MATCH Formula Match still = MisMatch',
    no_match_formula == 'MisMatch', no_match_formula)
chk('FX-09e', 'KEL: XC_KEL_NO_MATCH Known Exception is blank (fingerprint mismatch)',
    no_match_kel in ('', 'nan'), repr(no_match_kel))

# ── FX-25: Full Run row counts ────────────────────────────────────────────────
chk('FX-25', f'Full Run row counts: XC={expected_row_count} GB=2 AP=2 Cond=7',
    len(xc_df) == expected_row_count and len(df_gb) == 2 and len(ap_df) == 2 and len(cond_full) == 7,
    f'XC={len(xc_df)} GB={len(df_gb)} AP={len(ap_df)} Cond={len(cond_full)}')

# ── Summary ───────────────────────────────────────────────────────────────────
passes = sum(1 for r in results if r[2] == 'PASS')
fails  = sum(1 for r in results if r[2] == 'FAIL')
print(f'Results: {passes} PASS  {fails} FAIL\n')
for case_id, desc, status, detail in results:
    detail_str = f'  ({detail})' if detail and status == 'FAIL' else ''
    print(f'  {status}  {case_id:<8}  {desc}{detail_str}')
