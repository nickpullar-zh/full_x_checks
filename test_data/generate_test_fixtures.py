"""
Generate minimal test fixture files for all four strategies.

Run:  python test_data/generate_test_fixtures.py

Produces test_data/fixtures/:
  xc_pub.xlsx                 shared EBX publication file (all strategies)
  fip_xc.txt                  FIP X-Checks text (SAP Validation Rule export format)
  fip_ZQ9_VALFLDGR.xlsx       FIP Grouping By (ZQ9_VALFLDGR extract)
  mapping.txt                 Grouping By mapping file
  validation_methods.xlsx     AP Validation Methods
  fip_ZQ9_VALMSG.xlsx         FIP Accounting Principles (ZQ9_VALMSG extract)
  fip_ZQ9_VALMETH.xlsx        FIP Conditions (ZQ9_VALMETH extract)

Expected Comparison results
===========================
X-Checks:
  XC_ALL_MATCH          all 4 comparison cols = Match
  XC_FORMULA_MISMATCH   all 4 = MisMatch
  XC_VARIABLE_MISMATCH  Formula Match = Match, Variables Match = MisMatch
  XC_NOT_IN_FIP         all 4 = Not Found  (EBX only)
  XC_NOT_IN_EBX         all 4 = Not Found  (FIP only)
  XC_TOM_CORRECTION     all 4 = Match  (FIP uses TOM; parser rewrites to ToM)
  XC_THOUSANDS_CORR     all 4 = Match  (FIP formula has 1.000; _clean_text strips to 1000)
  XC_REORDER_MATCH      Formula Match = MisMatch  (reorder logic fires but produces
                        invalid formula for simple addition; see compare.py note)

Grouping By:
  GB_MATCHED            Result = Matched
  GB_NOT_IN_FIP         Result = Not in FIP

Accounting Principles:
  AP_MATCH              Match = Match
  AP_MISMATCH           Match = MisMatch

Conditions:
  COND_MATCHED          Comparison = Matched
  COND_NOT_MATCHED      Comparison = Not Matched
  COND_REF_XCHECK       Comparison = Matched  (Reference X-Check overrides key prefix)
"""

from pathlib import Path
import openpyxl
from openpyxl.styles import Font

OUT = Path(__file__).parent / "fixtures"
OUT.mkdir(exist_ok=True)


def _write_rows(ws, header, rows):
    ws.append(header)
    for r in rows:
        ws.append(r)


# ---------------------------------------------------------------------------
# 1. Shared EBX Publication file
# ---------------------------------------------------------------------------

def _make_xc_pub():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "cross checks all"

    headers = [
        "X-Check No.",                   # A
        "Account No.",                   # B
        "SubA No.",                      # C
        "Operator (X-Check Term)",       # D
        "Absolute (result)",             # E
        "Category",                      # F
        "Operator 1",                    # G
        "Operator 2",                    # H
        "Limit 1",                       # I
        "Limit 2",                       # J
        "%",                             # K
        "Exclude Account Type",          # L
        "Version Spanning Validation",   # M
        "Ending Balance Prior Year",     # N
        "Grouping By",                   # O
        "Reference  X-Check (Condition)",  # P  (two spaces — real file header)
        "IFRS New RFD",                  # Q  (real AP validation event from validation_methods.xlsx)
        "Applicable Quarters",           # R
        "Included RUs",                  # S
        "Excluded RUs",                  # T
        "Reference X-Check (Limit, %)",  # U
        "Status",                        # V
        "Type of change",               # W
        "Exclude Z-Core",               # X
        "Sense Check",                  # Y
        "In Scope",                     # Z
    ]
    ws.append(headers)

    def row(xc, acct="", suba="", op_term="+", absolute="", category="",
            op1="<=", op2="", lim1="0", lim2="", pct="",
            excl_acc="", vsv="", ebpy="",
            grouping_by="", ref_xc_cond="",
            test_event="",
            app_qtrs="", incl_ru="", excl_ru="", ref_xc_lim="",
            status="ACTIVE", toc="", excl_zcore=""):
        return [
            xc, acct, suba, op_term, absolute, category,
            op1, op2, lim1, lim2, pct,
            excl_acc, vsv, ebpy,
            grouping_by, ref_xc_cond,
            test_event,
            app_qtrs, incl_ru, excl_ru, ref_xc_lim,
            status, toc, excl_zcore, "", "",
        ]

    # ------------------------------------------------------------------
    # X-Checks rows
    # EBX formula: VAL_YTD(<variable>)<op1><const>
    # Variable name built from Account No. (+ToM+SubA if SubA present)
    # ------------------------------------------------------------------

    # XC_ALL_MATCH: EBX → VAL_YTD(ACC001)<=0, FIP matches exactly
    ws.append(row("XC_ALL_MATCH", acct="ACC001", op1="<=", lim1="0"))

    # XC_FORMULA_MISMATCH: EBX → VAL_YTD(ACC001)<=0; FIP has VAL_YTD(ACC999)<=0
    ws.append(row("XC_FORMULA_MISMATCH", acct="ACC001", op1="<=", lim1="0"))

    # XC_VARIABLE_MISMATCH: EBX → VAL_YTD(ACC001)<=0; FIP formula matches
    # but FIP FS Account = ACC_WRONG → variable string differs → Variables MisMatch
    ws.append(row("XC_VARIABLE_MISMATCH", acct="ACC001", op1="<=", lim1="0"))

    # XC_NOT_IN_FIP: EBX row present; no FIP block → Not Found (FIP side missing)
    ws.append(row("XC_NOT_IN_FIP", acct="ACC001", op1="<=", lim1="0"))

    # XC_NOT_IN_EBX: no Account No. → extract_ebx skips it; still appears in
    # x_check_list (read from raw EBX DataFrame); FIP has a block → Not Found (EBX side)
    ws.append(row("XC_NOT_IN_EBX"))  # no acct → skipped by extract_ebx

    # XC_TOM_CORRECTION: SubA "AA" (non-numeric string) → no float conversion
    # EBX var = ACC001ToMAA, Movement Types:AA
    # FIP uses TOM in var name; VARIABLE state replaces TOM→ToM → ACC001ToMAA → Match
    ws.append(row("XC_TOM_CORRECTION", acct="ACC001", suba="AA", op_term="+", op1="<=", lim1="0"))

    # XC_THOUSANDS_CORR: EBX → VAL_YTD(ACC001)<=CONST(1000,'USD','E')
    # FIP formula has CONST(1.000,...); _clean_text re.sub strips 1.000 → 1000 → Match
    ws.append(row("XC_THOUSANDS_CORR", acct="ACC001", op1="<=", lim1="1000"))

    # XC_REORDER_MATCH: SubA "AA"/"BB" → vars A_ACCToMAA, B_ACCToMBB
    # FIP formula reversed; _compare_formulas addition-only reorder → Match
    ws.append(row("XC_REORDER_MATCH", acct="A_ACC", suba="AA", op_term="+", op1="<=", lim1="0"))
    ws.append(row("XC_REORDER_MATCH", acct="B_ACC", suba="BB", op_term="+", op1="<=", lim1="0"))

    # ------------------------------------------------------------------
    # Grouping By rows
    # ------------------------------------------------------------------
    ws.append(row("GB_MATCHED",    grouping_by="GB_GROUPING_ITEM"))
    ws.append(row("GB_NOT_IN_FIP", grouping_by="GB_GROUPING_ITEM"))

    # ------------------------------------------------------------------
    # Accounting Principles rows
    # TEST_EVENT = 'w' on both; FIP has w (Match) or e (MisMatch)
    # ------------------------------------------------------------------
    ws.append(row("AP_MATCH",    test_event="w"))
    ws.append(row("AP_MISMATCH", test_event="w"))

    # ------------------------------------------------------------------
    # Conditions rows
    # ------------------------------------------------------------------
    ws.append(row("COND_MATCHED",     app_qtrs="Q1"))
    ws.append(row("COND_NOT_MATCHED", app_qtrs="Q2"))
    # Reference X-Check (Condition)=COND_BASE overrides key prefix → key=COND_BASE|Q1
    ws.append(row("COND_REF_XCHECK", ref_xc_cond="COND_BASE", app_qtrs="Q1"))

    wb.save(OUT / "xc_pub.xlsx")
    print("  wrote xc_pub.xlsx")


# ---------------------------------------------------------------------------
# 2. FIP X-Checks text
# ---------------------------------------------------------------------------
# Line formats that survive _clean_text correctly:
#
# FS Account line: "| 1 | 2 | 3 | FS Account | <acct> |\n"
#   → after cleanup: "1 2 3 FS Account <acct> |"
#   → token[3]="FS" → arr_fs_accounts.append(token[5]=<acct>) ✓
#
# Movement Type line (first MT for a variable): "| |- Movement Type @20@ <mt> desc |\n"
#   → after cleanup: "|-Movement Type @20@ <mt> desc |"
#   → 'Movement Type' in line → token[3]=<mt> → arr_movement_types.append(<mt>) ✓
#
# Variable name line: just the name on its own line (no delimiters)
#   → cleanup is identity; then TOM→ToM replacement in VARIABLE state ✓

_SEGMENT_END = "|-Segment @28@ * |"
_BLOCK_END   = "-|"
_BLANK       = "|"
_FORMULA_HDR = "|Formula String |"
_VAR_HDR     = "|-Characteristic Sel Opt Attributes Node Characteristic From To |"


def _fip_block_single(xc_id, formula, var_name, fs_accounts, movement_types=None):
    """One variable, optional movement types."""
    if movement_types is None:
        movement_types = []

    fs_lines = "".join(
        f"| 1 | 2 | 3 | FS Account | {acc} |\n"
        for acc in fs_accounts
    )

    if movement_types:
        # First MT: line that triggers 'Movement Type' detection; token[3] = mt value
        mt_lines = f"| |- Movement Type @20@ {movement_types[0]} desc |\n"
        # Additional MTs in MOV_GENERAL state: token[3] = mt value
        for mt in movement_types[1:]:
            mt_lines += f"| 1 | 2 | 3 | @20@ | {mt} | desc |\n"
        end_var = _BLOCK_END + "\n"  # BLOCK_END saves variable and breaks in MOV_GENERAL
    else:
        mt_lines = ""
        end_var = _BLOCK_END + "\n"  # BLOCK_END in FS_ACCOUNT state saves and breaks

    return (
        f"{xc_id} {xc_id} {xc_id} test description\n"
        f"{_FORMULA_HDR}\n"
        f"{formula}\n"
        f"{_BLOCK_END}\n"
        f"{_VAR_HDR}\n"
        f"{var_name}\n"
        f"{_BLANK}\n"
        f"{fs_lines}"
        f"{mt_lines}"
        f"{end_var}"
        f"{_SEGMENT_END}\n"
        f"{_BLOCK_END}\n\n"
    )


def _fip_block_two_vars(xc_id, formula,
                        var1_name, var1_fs, var1_mt,
                        var2_name, var2_fs, var2_mt):
    """
    Two variables with movement types.
    After var1's movement type section: BLANK_LINE returns parser to VARIABLE state.
    After var2's movement type section: BLOCK_END saves var2 and breaks.
    """
    def fs_lines(accts):
        return "".join(f"| 1 | 2 | 3 | FS Account | {a} |\n" for a in accts)

    def mt_line(mt):
        return f"| |- Movement Type @20@ {mt} desc |\n"

    return (
        f"{xc_id} {xc_id} {xc_id} test description\n"
        f"{_FORMULA_HDR}\n"
        f"{formula}\n"
        f"{_BLOCK_END}\n"
        f"{_VAR_HDR}\n"
        # Variable 1
        f"{var1_name}\n"
        f"{_BLANK}\n"
        f"{fs_lines(var1_fs)}"
        f"{mt_line(var1_mt)}"
        f"{_BLANK}\n"           # BLANK in MOV_GENERAL → save var1, back to VARIABLE state
        # Variable 2
        f"{var2_name}\n"
        f"{_BLANK}\n"
        f"{fs_lines(var2_fs)}"
        f"{mt_line(var2_mt)}"
        f"{_BLOCK_END}\n"       # BLOCK_END in MOV_GENERAL → save var2, break
        f"{_SEGMENT_END}\n"
        f"{_BLOCK_END}\n\n"
    )


def _make_fip_xc():
    blocks = []

    # XC_ALL_MATCH: formula VAL_YTD(ACC001)<=0, FS Account ACC001 → all Match
    blocks.append(_fip_block_single(
        "XC_ALL_MATCH",
        formula="VAL_YTD(ACC001)<=0",
        var_name="ACC001",
        fs_accounts=["ACC001"],
    ))

    # XC_FORMULA_MISMATCH: formula uses ACC999 (EBX expects ACC001) → all MisMatch
    blocks.append(_fip_block_single(
        "XC_FORMULA_MISMATCH",
        formula="VAL_YTD(ACC999)<=0",
        var_name="ACC999",
        fs_accounts=["ACC999"],
    ))

    # XC_VARIABLE_MISMATCH: formula matches EBX but FS Account = ACC_WRONG
    # → Formula Match=Match, Variables Match=MisMatch
    blocks.append(_fip_block_single(
        "XC_VARIABLE_MISMATCH",
        formula="VAL_YTD(ACC001)<=0",
        var_name="ACC_WRONG",
        fs_accounts=["ACC_WRONG"],
    ))

    # XC_NOT_IN_EBX: valid block; X-Check not in EBX (no Account No. row) → Not Found (EBX side)
    blocks.append(_fip_block_single(
        "XC_NOT_IN_EBX",
        formula="VAL_YTD(ACC001)<=0",
        var_name="ACC001",
        fs_accounts=["ACC001"],
    ))

    # XC_TOM_CORRECTION: SubA "AA" → EBX var ACC001ToMAA, MT:AA
    # FIP uses TOM in var name; VARIABLE state replaces TOM→ToM → ACC001ToMAA → Match
    blocks.append(_fip_block_single(
        "XC_TOM_CORRECTION",
        formula="VAL_YTD(ACC001TOMAA)<=0",
        var_name="ACC001TOMAA",
        fs_accounts=["ACC001"],
        movement_types=["AA"],
    ))

    # XC_THOUSANDS_CORR: FIP formula contains 1.000; _clean_text strips to 1000
    # EBX Limit=1000 → formula VAL_YTD(ACC001)<=CONST(1000,'USD','E') → Match
    blocks.append(_fip_block_single(
        "XC_THOUSANDS_CORR",
        formula="VAL_YTD(ACC001)<=CONST(1.000,'USD','E')",
        var_name="ACC001",
        fs_accounts=["ACC001"],
    ))

    # XC_REORDER_MATCH: FIP has terms reversed (B first, A second)
    # EBX SubA=1 for A_ACC, SubA=2 for B_ACC → two separate variables
    # EBX formula: VAL_YTD(A_ACCToM1)+VAL_YTD(B_ACCToM2)<=0  (A sorted first)
    # FIP formula: VAL_YTD(B_ACCToM2)+VAL_YTD(A_ACCToM1)<=0  (reversed)
    # _compare_formulas: no )-V, sorted vars equal, has + → reorders FIP → Match
    # XC_REORDER_MATCH: SubA "AA"/"BB" → vars A_ACCToMAA, B_ACCToMBB
    # FIP reversed formula; _compare_formulas reorders → Match
    blocks.append(_fip_block_two_vars(
        "XC_REORDER_MATCH",
        formula="VAL_YTD(B_ACCToMBB)+VAL_YTD(A_ACCToMAA)<=0",
        var1_name="A_ACCToMAA",
        var1_fs=["A_ACC"],
        var1_mt="AA",
        var2_name="B_ACCToMBB",
        var2_fs=["B_ACC"],
        var2_mt="BB",
    ))

    # XC_NOT_IN_FIP: intentionally omitted → Not Found (FIP side missing)

    (OUT / "fip_xc.txt").write_text("".join(blocks), encoding="utf-8")
    print("  wrote fip_xc.txt")


# ---------------------------------------------------------------------------
# 3. FIP Grouping By — ZQ9_VALFLDGR format
# ---------------------------------------------------------------------------

def _make_fip_valfldgr():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    _write_rows(ws,
        ["ValidRule", "Long Text", "Field name"],
        [
            # GB_MATCHED: Field name maps via mapping.txt to GB_GROUPING_ITEM
            # FIP Key = ValidRule|EBX Item = GB_MATCHED|GB_GROUPING_ITEM → Matched
            ["GB_MATCHED", "Test rule", "GB_FIP_FIELD"],
            # GB_NOT_IN_FIP: no row → Not in FIP
        ],
    )
    wb.save(OUT / "fip_ZQ9_VALFLDGR.xlsx")
    print("  wrote fip_ZQ9_VALFLDGR.xlsx")


# ---------------------------------------------------------------------------
# 4. Mapping file
# ---------------------------------------------------------------------------

def _make_mapping():
    (OUT / "mapping.txt").write_text(
        "FIP Data,EBX item\nGB_FIP_FIELD,GB_GROUPING_ITEM\n",
        encoding="utf-8",
    )
    print("  wrote mapping.txt")


# ---------------------------------------------------------------------------
# 5. Validation Methods (AP)
# ---------------------------------------------------------------------------
# The real validation_methods.xlsx must be copied into test_data/fixtures/
# manually (it is the live Zurich reference file, not generated here).
# Event used by AP fixture rows: 'IFRS New RFD', method V900W (Warning, black font).
def _make_validation_methods():
    print("  skipped validation_methods.xlsx — use the real file copied into fixtures/")


# ---------------------------------------------------------------------------
# 6. FIP Accounting Principles — ZQ9_VALMSG raw format
# ---------------------------------------------------------------------------

def _make_fip_valmsg():
    """
    Raw ZQ9_VALMSG: strategy builds Key = MK|ValidRule at load time.
    MK = validation method code, ValidRule = X-Check No., MT = W/E letter.

    AP_MATCH:    FIP MT='w' == EBX actual 'w' (Warning binding for IFRS New RFD) → Match
    AP_MISMATCH: FIP MT='e' != EBX actual 'w' → MisMatch
    MK=V900W matches the black-font Warning binding for event 'IFRS New RFD' in the
    real validation_methods.xlsx. ValidRule = X-Check No. in the EBX pub file.
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "FIP Methods Rules and Condition"
    _write_rows(ws,
        ["MethC", "MK", "Medium Text", "ValidRule", "Long Text",
         "UCFV20G-TRUE_BRANCH", "Message class", "Msg.", "MT", "Message Text"],
        [
            ["1", "V900W", "Warning method", "AP_MATCH",    "AP Match",    "X", "CLS", "001", "w", "Test"],
            ["1", "V900W", "Warning method", "AP_MISMATCH", "AP MisMatch", "X", "CLS", "002", "e", "Test"],
        ],
    )
    wb.save(OUT / "fip_ZQ9_VALMSG.xlsx")
    print("  wrote fip_ZQ9_VALMSG.xlsx")


# ---------------------------------------------------------------------------
# 7. FIP Conditions — ZQ9_VALMETH raw format
# ---------------------------------------------------------------------------

def _make_fip_valmeth():
    """
    8-column raw ZQ9_VALMETH. conditions/fip.py renames and builds
    Key (Concatenated) = Normal X-Check No|Condition No.

    COND_MATCHED:    key COND_MATCHED|Q1 → Matched
    COND_NOT_MATCHED: no row → Not Matched
    COND_REF_XCHECK: EBX Reference X-Check (Condition)=COND_BASE → key COND_BASE|Q1 → Matched
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "FIP Conditions"
    # Raw ZQ9_VALMETH headers as exported from FIP (positional — duplicate names are normal).
    # conditions/fip.py renames by position: col3 (ValidRule) → Normal X-Check No,
    # col6 (ValidRule) → Condition No.
    _write_rows(ws,
        ["MethC", "MK", "Medium Text", "ValidRule",
         "Medium Text", "UCFV20G-TRUE_BRANCH", "ValidRule", "Medium Text"],
        [
            ["1", "MK1", "Test MK", "COND_MATCHED", "Matched test", "X", "Q1", "Quarter 1"],
            ["1", "MK1", "Test MK", "COND_BASE",    "Ref XC test",  "X", "Q1", "Quarter 1"],
            # COND_NOT_MATCHED intentionally absent → Not Matched
        ],
    )
    wb.save(OUT / "fip_ZQ9_VALMETH.xlsx")
    print("  wrote fip_ZQ9_VALMETH.xlsx")


# ---------------------------------------------------------------------------
# main
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    print(f"Writing fixtures to {OUT}/")
    _make_xc_pub()
    _make_fip_xc()
    _make_fip_valfldgr()
    _make_mapping()
    _make_validation_methods()
    _make_fip_valmsg()
    _make_fip_valmeth()
    print("Done.")
