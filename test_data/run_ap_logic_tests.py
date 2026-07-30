"""
Accounting Principles logic test suite.
Run from repo root:  python test_data/run_ap_logic_tests.py
"""
import sys, pandas as pd
sys.path.insert(0, '.')
from pathlib import Path
from strategies.accounting_principles.validation_methods import parse_method_bindings
from strategies.accounting_principles.compare import compare_with_bindings
from strategies.accounting_principles.accounting_principles import (
    AccountingPrinciples, DEFAULT_EVENTS
)
from task_configs import ACCOUNTING_PRINCIPLES_UPLOAD_CONFIG

F   = Path('test_data/fixtures/ap')
VM  = Path('test_data/fixtures/validation_methods.xlsx')
results = []

def chk(case_id, desc, passed, detail=''):
    results.append((case_id, desc, 'PASS' if passed else 'FAIL', detail))

# ── Setup: run full comparison (no in-scope filter) ───────────────────────
bindings = parse_method_bindings(str(VM), DEFAULT_EVENTS)
cc_df    = pd.read_excel(F / 'ap_pub.xlsx', sheet_name='cross checks all')
fip_df   = pd.read_excel(F / 'ap_fip_ZQ9_VALMSG.xlsx',
                          sheet_name='FIP Methods Rules and Condition')
fip_df['Key'] = (fip_df['MK'].astype(str).str.strip()
                 + '|' + fip_df['ValidRule'].astype(str).str.strip())
all_xchecks = [str(x).strip() for x in cc_df['X-Check No.'].tolist()
               if str(x).strip() not in ('nan', '', 'None')]
df = pd.DataFrame(compare_with_bindings(bindings, cc_df, all_xchecks, fip_df))

def row(xc):
    r = df[df['X-Check No.'] == xc]
    return r.iloc[0] if len(r) else None

# ── AP-01: Standard Warning match ─────────────────────────────────────────
r = row('AP_MATCH_W')
chk('AP-01a', 'AP_MATCH_W Match = Match',   r is not None and r['Match'] == 'Match',   r['Match'] if r is not None else 'MISSING')
chk('AP-01b', 'AP_MATCH_W Expected = Warning', r is not None and r['Expected'] == 'Warning', r['Expected'] if r is not None else 'MISSING')
chk('AP-01c', 'AP_MATCH_W FIP = w',         r is not None and r['FIP'] == 'w',         r['FIP'] if r is not None else 'MISSING')
chk('AP-01d', 'AP_MATCH_W Actual = w',      r is not None and r['Actual'] == 'w',       r['Actual'] if r is not None else 'MISSING')

# ── AP-02: MisMatch ───────────────────────────────────────────────────────
r = row('AP_MISMATCH')
chk('AP-02a', 'AP_MISMATCH Match = MisMatch', r is not None and r['Match'] == 'MisMatch', r['Match'] if r is not None else 'MISSING')
chk('AP-02b', 'AP_MISMATCH FIP = e, Actual = w', r is not None and r['FIP'] == 'e' and r['Actual'] == 'w', f"FIP={r['FIP']} Actual={r['Actual']}" if r is not None else 'MISSING')

# ── AP-03: Both severity ──────────────────────────────────────────────────
r_w = row('AP_BOTH_W')
r_e = row('AP_BOTH_E')
chk('AP-03a', 'AP_BOTH_W Match = Match (severity=Both, FIP=w)',
    r_w is not None and r_w['Match'] == 'Match' and r_w['Expected'] == 'Both',
    f"Match={r_w['Match']} Exp={r_w['Expected']}" if r_w is not None else 'MISSING')
chk('AP-03b', 'AP_BOTH_E Match = Match (severity=Both, FIP=e)',
    r_e is not None and r_e['Match'] == 'Match' and r_e['Expected'] == 'Both',
    f"Match={r_e['Match']} Exp={r_e['Expected']}" if r_e is not None else 'MISSING')

# ── AP-04: Grey binding fallback ──────────────────────────────────────────
r = row('AP_GREY_WINS')
chk('AP-04a', 'AP_GREY_WINS produces a row (grey binding fired)',
    r is not None, 'row missing')
chk('AP-04b', 'AP_GREY_WINS Match = Match',
    r is not None and r['Match'] == 'Match', r['Match'] if r is not None else 'MISSING')

# ── AP-05: Unknown method — row skipped ───────────────────────────────────
r = row('AP_NO_BINDING')
chk('AP-05', 'AP_NO_BINDING absent from output (unknown method - skipped)',
    r is None, 'row present — should be absent')

# ── AP-06: Blank event column — row skipped ───────────────────────────────
r = row('AP_NO_ACTUAL')
chk('AP-06', 'AP_NO_ACTUAL absent from output (blank event column - no winning binding)',
    r is None, 'row present — should be absent')

# ── AP-07: Key built from MK + ValidRule (raw ZQ9_VALMSG) ─────────────────
chk('AP-07', 'Key column built at load time from MK + ValidRule',
    'Key' in fip_df.columns, 'Key column missing')
sample_key = fip_df[fip_df['ValidRule'] == 'AP_MATCH_W']['Key'].iloc[0] if len(fip_df[fip_df['ValidRule'] == 'AP_MATCH_W']) else ''
chk('AP-07b', 'Key format is MK|ValidRule',
    '|' in str(sample_key) and 'AP_MATCH_W' in str(sample_key), repr(sample_key))

# ── AP-08: process_only_differences selection filter ──────────────────────
from strategies.accounting_principles.accounting_principles import AccountingPrinciples
strat = AccountingPrinciples(ACCOUNTING_PRINCIPLES_UPLOAD_CONFIG)
strat.log = []
selected = strat._select_in_scope_x_checks(cc_df, str(F / 'ap_pub.xlsx'), 'cross checks all')
chk('AP-08a', 'AP_MATCH_W in scope', 'AP_MATCH_W' in selected, str(selected[:5]))
chk('AP-08b', 'AP_NOT_SCOPE_TOC excluded (blank Type of change)', 'AP_NOT_SCOPE_TOC' not in selected, '')
chk('AP-08c', 'AP_NOT_SCOPE_INA excluded (Status=INACTIVE)', 'AP_NOT_SCOPE_INA' not in selected, '')
chk('AP-08d', 'AP_EXCL_ZCORE excluded (Exclude Z-Core=X)', 'AP_EXCL_ZCORE' not in selected, '')
chk('AP-08e', 'AP_YELLOW_CAT excluded (Category cell yellow)', 'AP_YELLOW_CAT' not in selected, '')

# ── AP-09: Known Exception annotation ────────────────────────────────────
annotated = strat._annotate_known_exceptions(
    df.copy(), str(F / 'ap_kel.xlsx'),
    sheet_name='Accounting Principles',
    fingerprint_columns=["X-Check No.", "Event", "Expected", "FIP", "Actual", "Method"]
)
kel_row = annotated[annotated['X-Check No.'] == 'AP_MISMATCH']
chk('AP-09a', 'KEL: AP_MISMATCH Match still = MisMatch',
    len(kel_row) > 0 and kel_row.iloc[0]['Match'] == 'MisMatch',
    kel_row.iloc[0]['Match'] if len(kel_row) else 'MISSING')
chk('AP-09b', 'KEL: AP_MISMATCH Known Exception reason populated',
    len(kel_row) > 0 and bool(kel_row.iloc[0].get('Known Exception', '')) and
    kel_row.iloc[0].get('Known Exception', '') not in ('', 'nan'),
    repr(kel_row.iloc[0].get('Known Exception', '')) if len(kel_row) else 'MISSING')
match_row = annotated[annotated['X-Check No.'] == 'AP_MATCH_W']
chk('AP-09c', 'KEL: AP_MATCH_W Known Exception blank (no mismatch, no annotation)',
    len(match_row) > 0 and match_row.iloc[0].get('Known Exception', '') in ('', 'nan', None),
    repr(match_row.iloc[0].get('Known Exception', '')) if len(match_row) else 'MISSING')

# ── Summary ───────────────────────────────────────────────────────────────
passes = sum(1 for r in results if r[2] == 'PASS')
fails  = sum(1 for r in results if r[2] == 'FAIL')
print(f'Results: {passes} PASS  {fails} FAIL\n')
for case_id, desc, status, detail in results:
    detail_str = f'  ({detail})' if detail and status == 'FAIL' else ''
    print(f'  {status}  {case_id:<8}  {desc}{detail_str}')
