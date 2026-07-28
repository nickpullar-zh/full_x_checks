"""
Logic-layer verification of all Logic test cases in the Fixture UAT plan.
Run from repo root: python test_data/run_logic_tests.py
"""
import sys, pandas as pd
sys.path.insert(0, '.')
from pathlib import Path

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
    'XC_ALL_MATCH':         'Match',
    'XC_DIFF_EXCLUDED':     'Not Found',  # has Account No in EBX but no FIP block
    'XC_DIFF_IN_SCOPE':     'Match',
    'XC_FORMULA_MISMATCH':  'MisMatch',
    'XC_NOT_IN_EBX':        'Not Found',
    'XC_NOT_IN_FIP':        'Not Found',
    'XC_REORDER_MATCH':     'MisMatch',
    'XC_THOUSANDS_CORR':    'Match',
    'XC_TOM_CORRECTION':    'Match',
    'XC_VARIABLE_MISMATCH': 'Match',
}
chk('FX-05a', 'X-Checks row count = 10', len(xc_df) == 10, f'got {len(xc_df)}')
for xc_id, exp in expected_xc.items():
    rows = xc_df[xc_df['X-Check No.'] == xc_id]
    got = rows.iloc[0]['Formula Match'] if len(rows) else 'MISSING'
    chk('FX-05b', f'X-Checks {xc_id} Formula Match = {exp}', got == exp, got)

# ── FX-06: XC_VARIABLE_MISMATCH Variables Match ───────────────────────────────
rows = xc_df[xc_df['X-Check No.'] == 'XC_VARIABLE_MISMATCH']
got_vm = rows.iloc[0]['Variables Match'] if len(rows) else 'MISSING'
chk('FX-06', 'XC_VARIABLE_MISMATCH Variables Match = MisMatch', got_vm == 'MisMatch', got_vm)

# ── FX-10: X-Check No Selection (differences mode) ────────────────────────────
from strategies.x_checks.x_check_no_selection import select_x_check_nos
selected = select_x_check_nos(ebx_df, str(F / 'xc_pub.xlsx'), 'cross checks all')
chk('FX-10a', 'X-Check No Selection: exactly 1 result', len(selected) == 1, str(selected))
chk('FX-10b', 'X-Check No Selection: XC_DIFF_IN_SCOPE present', 'XC_DIFF_IN_SCOPE' in selected, str(selected))
chk('FX-10c', 'X-Check No Selection: XC_DIFF_EXCLUDED absent', 'XC_DIFF_EXCLUDED' not in selected, str(selected))

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
chk('FX-25', 'Full Run row counts: XC=10 GB=2 AP=2 Cond=7',
    len(xc_df) == 10 and len(df_gb) == 2 and len(ap_df) == 2 and len(cond_full) == 7,
    f'XC={len(xc_df)} GB={len(df_gb)} AP={len(ap_df)} Cond={len(cond_full)}')

# ── Summary ───────────────────────────────────────────────────────────────────
passes = sum(1 for r in results if r[2] == 'PASS')
fails  = sum(1 for r in results if r[2] == 'FAIL')
print(f'Results: {passes} PASS  {fails} FAIL\n')
for case_id, desc, status, detail in results:
    detail_str = f'  ({detail})' if detail and status == 'FAIL' else ''
    print(f'  {status}  {case_id:<8}  {desc}{detail_str}')
