"""
Grouping By logic test suite.
Run from repo root:  python test_data/run_gb_logic_tests.py
"""
import sys, pandas as pd
sys.path.insert(0, '.')
from pathlib import Path
from strategies.grouping_by.grouping_by import GroupingBy
from task_configs import GROUPING_BY_UPLOAD_CONFIG

F = Path('test_data/fixtures/gb')
results = []

def chk(case_id, desc, passed, detail=''):
    results.append((case_id, desc, 'PASS' if passed else 'FAIL', detail))

# ── Setup ──────────────────────────────────────────────────────────────────
gb = GroupingBy(GROUPING_BY_UPLOAD_CONFIG)
gb.log = []
mapping_txt = (F / 'gb_mapping.txt').read_text()
fip_gb   = pd.read_excel(F / 'gb_fip_ZQ9_VALFLDGR.xlsx', sheet_name='Sheet1')
ebx_gb   = pd.read_excel(F / 'gb_pub.xlsx', sheet_name='cross checks all')
loaded = {
    GROUPING_BY_UPLOAD_CONFIG.file_fields[0].label: fip_gb,
    GROUPING_BY_UPLOAD_CONFIG.file_fields[1].label: ebx_gb,
    GROUPING_BY_UPLOAD_CONFIG.file_fields[2].label: mapping_txt,
}
_, _, df_fip = gb._process_fip(loaded)
_, df_ebx    = gb._process_ebx(loaded)
df_cmp       = gb._process_compare(df_fip, df_ebx)

def result(key):
    r = df_cmp[df_cmp['EBX Key'] == key]
    return r.iloc[0]['Result'] if len(r) else 'MISSING'

def present(key):
    return key in df_cmp['EBX Key'].values

# ── GB-01: Standard match ──────────────────────────────────────────────────
chk('GB-01', 'GB_MATCHED|ITEM_A = Matched', result('GB_MATCHED|ITEM_A') == 'Matched', result('GB_MATCHED|ITEM_A'))

# ── GB-02: Not in FIP ─────────────────────────────────────────────────────
chk('GB-02', 'GB_NOT_IN_FIP|ITEM_A = Not in FIP', result('GB_NOT_IN_FIP|ITEM_A') == 'Not in FIP', result('GB_NOT_IN_FIP|ITEM_A'))

# ── GB-03: Reference X-Check (Condition) overrides base key ───────────────
chk('GB-03', 'REF_BASE|ITEM_A = Matched (reference XC override)', result('REF_BASE|ITEM_A') == 'Matched', result('REF_BASE|ITEM_A'))
chk('GB-03b', 'GB_REF_XC_KEY|ITEM_A absent (key uses ref base, not X-Check No.)', not present('GB_REF_XC_KEY|ITEM_A'), '')

# ── GB-04: Multiple comma-separated Grouping By values ────────────────────
chk('GB-04a', 'GB_MULTI|ITEM_A = Matched (first split value, FIP entry present)', result('GB_MULTI|ITEM_A') == 'Matched', result('GB_MULTI|ITEM_A'))
chk('GB-04b', 'GB_MULTI|ITEM_B = Not in FIP (second split value, no FIP entry)',  result('GB_MULTI|ITEM_B') == 'Not in FIP', result('GB_MULTI|ITEM_B'))

# ── GB-05: Deduplication (only first EBX row per X-Check No. used) ────────
chk('GB-05a', 'GB_DEDUP|ITEM_A present (first row kept)',   present('GB_DEDUP|ITEM_A'), '')
chk('GB-05b', 'GB_DEDUP|ITEM_B absent (second row dropped)', not present('GB_DEDUP|ITEM_B'), '')

# ── GB-06: FIP filtering — mapped to "ignore" ─────────────────────────────
chk('GB-06', 'GB_IGNORE_FIELD|ITEM_A = Not in FIP (FIP field mapped to ignore → dropped)', result('GB_IGNORE_FIELD|ITEM_A') == 'Not in FIP', result('GB_IGNORE_FIELD|ITEM_A'))

# ── GB-07: FIP filtering — field name not in mapping ──────────────────────
chk('GB-07', 'GB_UNMAPPED|ITEM_A = Not in FIP (FIP field not in mapping → dropped)', result('GB_UNMAPPED|ITEM_A') == 'Not in FIP', result('GB_UNMAPPED|ITEM_A'))

# ── GB-08: FIP filtering — blank ValidRule ────────────────────────────────
chk('GB-08', 'GB_BLANK_VR|ITEM_A = Not in FIP (FIP row with blank ValidRule → dropped)', result('GB_BLANK_VR|ITEM_A') == 'Not in FIP', result('GB_BLANK_VR|ITEM_A'))

# ── GB-09: Total row count ─────────────────────────────────────────────────
chk('GB-09', 'Comparison row count = 11', len(df_cmp) == 11, f'got {len(df_cmp)}')

# ── GB-10: Known Exception annotation ─────────────────────────────────────
from strategies.grouping_by.grouping_by import GroupingBy as _GB
strat = _GB(GROUPING_BY_UPLOAD_CONFIG)
strat.log = []
annotated = strat._annotate_known_exceptions(
    df_cmp.copy(), str(F / 'gb_kel.xlsx'),
    sheet_name='Grouping By', fingerprint_columns=['EBX Key']
)
kel_val       = annotated[annotated['EBX Key'] == 'GB_KEL_MATCH|ITEM_A'].iloc[0].get('Known Exception', '')
no_match_val  = annotated[annotated['EBX Key'] == 'GB_KEL_NO_MATCH|ITEM_A'].iloc[0].get('Known Exception', '')
chk('GB-10a', 'KEL: GB_KEL_MATCH result unchanged = Not in FIP',
    annotated[annotated['EBX Key'] == 'GB_KEL_MATCH|ITEM_A'].iloc[0]['Result'] == 'Not in FIP', '')
chk('GB-10b', 'KEL: GB_KEL_MATCH Known Exception reason populated',
    bool(kel_val) and kel_val not in ('', 'nan'), repr(kel_val))
chk('GB-10c', 'KEL: GB_KEL_NO_MATCH Known Exception blank (wrong fingerprint)',
    no_match_val in ('', 'nan'), repr(no_match_val))

# ── Summary ───────────────────────────────────────────────────────────────
passes = sum(1 for r in results if r[2] == 'PASS')
fails  = sum(1 for r in results if r[2] == 'FAIL')
print(f'Results: {passes} PASS  {fails} FAIL\n')
for case_id, desc, status, detail in results:
    detail_str = f'  ({detail})' if detail and status == 'FAIL' else ''
    print(f'  {status}  {case_id:<8}  {desc}{detail_str}')
