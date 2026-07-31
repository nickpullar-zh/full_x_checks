"""
Conditions logic test suite.
Run from repo root:  python test_data/run_cond_logic_tests.py
"""
import sys, pandas as pd
sys.path.insert(0, '.')
from pathlib import Path
from strategies.conditions.extract import extract_conditions
from strategies.conditions.fip import process_fip
from strategies.conditions.compare import compare as cond_compare
from strategies.conditions.conditions import Conditions
from task_configs import CONDITIONS_UPLOAD_CONFIG

F = FCOND = Path('test_data/fixtures')
results = []

def chk(case_id, desc, passed, detail=''):
    results.append((case_id, desc, 'PASS' if passed else 'FAIL', detail))

# ── Setup: full file mode ──────────────────────────────────────────────────
fip_df   = pd.read_excel(F / 'cond_fip_ZQ9_VALMETH.xlsx', sheet_name='FIP Conditions')
fip_proc = process_fip(fip_df)
working_df, warnings = extract_conditions(str(F / 'xc_pub.xlsx'), 'cross checks all',
                                           process_only_differences=False)
df, summary = cond_compare(working_df, fip_proc)

def result(ebx_key):
    r = df[df['EBX Data'] == ebx_key]
    return r.iloc[0]['Comparison'] if len(r) else 'MISSING'

def fip_data(ebx_key):
    r = df[df['EBX Data'] == ebx_key]
    return str(r.iloc[0]['FIP Data']) if len(r) else 'MISSING'

# ── COND-01: All 5 CONDITION_COLS produce rows ────────────────────────────
chk('COND-01a', 'Applicable Quarters: COND_APPL_QTRS|Q1 = Matched',
    result('COND_APPL_QTRS|Q1') == 'Matched', result('COND_APPL_QTRS|Q1'))
chk('COND-01b', 'Included RUs: COND_INCL_RUS|RU_NORTH = Matched',
    result('COND_INCL_RUS|RU_NORTH') == 'Matched', result('COND_INCL_RUS|RU_NORTH'))
chk('COND-01c', 'Excluded RUs: COND_EXCL_RUS|RU_SOUTH = Matched',
    result('COND_EXCL_RUS|RU_SOUTH') == 'Matched', result('COND_EXCL_RUS|RU_SOUTH'))
chk('COND-01d', 'Reference X-Check (Limit, %): COND_LIMIT_PCT|10.5 = Matched',
    result('COND_LIMIT_PCT|10.5') == 'Matched', result('COND_LIMIT_PCT|10.5'))

# ── COND-02: Reference X-Check (Condition) override ──────────────────────
# COND_REF_XC row has ref_xc="COND_REF_XC", app_qtrs="Q1"
# effective_xc = "COND_REF_XC" → key = COND_REF_XC|Q1 (Matched)
# The ref_xc col value itself = "COND_REF_XC" → key = COND_REF_XC|COND_REF_XC (Not Matched)
chk('COND-02a', 'Ref override: COND_REF_XC|Q1 = Matched (effective_xc from ref col)',
    result('COND_REF_XC|Q1') == 'Matched', result('COND_REF_XC|Q1'))
chk('COND-02b', 'Ref col itself: COND_REF_XC|COND_REF_XC = Not Matched (no FIP entry for ref col value)',
    result('COND_REF_XC|COND_REF_XC') == 'Not Matched', result('COND_REF_XC|COND_REF_XC'))

# ── COND-03: Not Matched ──────────────────────────────────────────────────
chk('COND-03a', 'COND_NOT_MATCHED|Q2 = Not Matched', result('COND_NOT_MATCHED|Q2') == 'Not Matched', result('COND_NOT_MATCHED|Q2'))
chk('COND-03b', 'FIP Data blank when Not Matched', fip_data('COND_NOT_MATCHED|Q2') in ('', 'nan'), repr(fip_data('COND_NOT_MATCHED|Q2')))
chk('COND-03c', 'FIP Data = EBX Data when Matched', fip_data('COND_APPL_QTRS|Q1') == 'COND_APPL_QTRS|Q1', fip_data('COND_APPL_QTRS|Q1'))

# ── COND-04: Multiple condition cols per row ──────────────────────────────
chk('COND-04a', 'COND_MULTI_COL|Q1 = Matched',       result('COND_MULTI_COL|Q1')    == 'Matched',     result('COND_MULTI_COL|Q1'))
chk('COND-04b', 'COND_MULTI_COL|RU_IN = Matched',    result('COND_MULTI_COL|RU_IN') == 'Matched',     result('COND_MULTI_COL|RU_IN'))
chk('COND-04c', 'COND_MULTI_COL|RU_OUT = Not Matched', result('COND_MULTI_COL|RU_OUT') == 'Not Matched', result('COND_MULTI_COL|RU_OUT'))

# ── COND-05: Total row count ──────────────────────────────────────────────
chk('COND-05', 'Full file row count = 16', len(df) == 16, f'got {len(df)}')

# ── COND-06: Differences mode — yellow cell ───────────────────────────────
working_diff, _ = extract_conditions(str(F / 'xc_pub.xlsx'), 'cross checks all',
                                      process_only_differences=True)
df_diff, _ = cond_compare(working_diff, fip_proc)
diff_keys = set(df_diff['EBX Data'].tolist())
chk('COND-06a', 'Diff mode: COND_DIFF_YELLOW|Q1 collected (yellow cell)',
    'COND_DIFF_YELLOW|Q1' in diff_keys, str(diff_keys))
chk('COND-06b', 'Diff mode: COND_DIFF_YELLOW|Q1 = Matched',
    df_diff[df_diff['EBX Data'] == 'COND_DIFF_YELLOW|Q1'].iloc[0]['Comparison'] == 'Matched'
    if 'COND_DIFF_YELLOW|Q1' in diff_keys else False, '')

# ── COND-07: Differences mode — green cell ───────────────────────────────
chk('COND-07a', 'Diff mode: COND_DIFF_GREEN|RU_NORTH collected (green cell)',
    'COND_DIFF_GREEN|RU_NORTH' in diff_keys, str(diff_keys))
chk('COND-07b', 'Diff mode: COND_DIFF_GREEN|RU_NORTH = Matched',
    df_diff[df_diff['EBX Data'] == 'COND_DIFF_GREEN|RU_NORTH'].iloc[0]['Comparison'] == 'Matched'
    if 'COND_DIFF_GREEN|RU_NORTH' in diff_keys else False, '')

# ── COND-08: Differences mode — white cell not collected ──────────────────
chk('COND-08', 'Diff mode: COND_DIFF_WHITE|Q1 NOT collected (white cell)',
    'COND_DIFF_WHITE|Q1' not in diff_keys, '')

# ── COND-09: FIP column renaming (process_fip) ───────────────────────────
chk('COND-09a', 'FIP: Normal X-Check No column present after renaming',
    'Normal X-Check No' in fip_proc.columns, str(list(fip_proc.columns)))
chk('COND-09b', 'FIP: Condition No column present after renaming',
    'Condition No' in fip_proc.columns, '')
chk('COND-09c', 'FIP: Key (Concatenated) column built',
    'Key (Concatenated)' in fip_proc.columns, '')
sample = fip_proc['Key (Concatenated)'].iloc[0]
chk('COND-09d', 'FIP key format = XCheck|ConditionNo',
    '|' in str(sample), repr(sample))

# ── COND-10: Known Exception annotation ──────────────────────────────────
strat = Conditions(CONDITIONS_UPLOAD_CONFIG)
strat.log = []
annotated = strat._annotate_known_exceptions(
    df.copy(), str(F / 'cond_kel.xlsx'),
    sheet_name='Conditions', fingerprint_columns=['EBX Data', 'FIP Data']
)
# KEL annotates COND_APPL_QTRS|Q1 (Matched — both fingerprint cols non-blank)
kel_row  = annotated[annotated['EBX Data'] == 'COND_APPL_QTRS|Q1']
# COND_INCL_RUS|RU_NORTH has a KEL entry with wrong FIP Data → no annotation
no_match = annotated[annotated['EBX Data'] == 'COND_INCL_RUS|RU_NORTH']
chk('COND-10a', 'KEL: COND_APPL_QTRS|Q1 Comparison still = Matched',
    len(kel_row) > 0 and kel_row.iloc[0]['Comparison'] == 'Matched',
    kel_row.iloc[0]['Comparison'] if len(kel_row) else 'MISSING')
chk('COND-10b', 'KEL: COND_APPL_QTRS|Q1 Known Exception reason populated',
    len(kel_row) > 0 and bool(kel_row.iloc[0].get('Known Exception', '')) and
    kel_row.iloc[0].get('Known Exception', '') not in ('', 'nan'),
    repr(kel_row.iloc[0].get('Known Exception', '')) if len(kel_row) else 'MISSING')
chk('COND-10c', 'KEL: COND_INCL_RUS|RU_NORTH Known Exception blank (wrong FIP Data fingerprint)',
    len(no_match) > 0 and no_match.iloc[0].get('Known Exception', '') in ('', 'nan', None),
    repr(no_match.iloc[0].get('Known Exception', '')) if len(no_match) else 'MISSING')
chk('COND-10d', 'KEL note: Not Matched rows cannot be annotated (blank FIP Data fails fingerprint)',
    True, 'documented design limitation')  # informational only

# ── Summary ───────────────────────────────────────────────────────────────
passes = sum(1 for r in results if r[2] == 'PASS')
fails  = sum(1 for r in results if r[2] == 'FAIL')
print(f'Results: {passes} PASS  {fails} FAIL\n')
for case_id, desc, status, detail in results:
    detail_str = f'  ({detail})' if detail and status == 'FAIL' else ''
    print(f'  {status}  {case_id:<8}  {desc}{detail_str}')
