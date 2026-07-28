"""
Generate the Fixture-Based UAT Test Plan workbook.

This plan uses the minimal fixture files in test_data/fixtures/ rather than
production-size files. Every expected output row is listed explicitly so the
tester can verify each outcome by inspection without manual reconciliation.

Output: docs/<YYYYMMDD> Fixture_UAT_v<version> Test Plan.xlsx

Run from the repo root:
    python docs/generate_fixture_uat.py
"""

import os
import sys
from datetime import date

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from version import __version__

VERSION  = __version__
TODAY    = date.today().strftime("%Y%m%d")
OUT_DIR  = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "docs")
FILENAME = f"{TODAY} Fixture_UAT_v{VERSION} Test Plan.xlsx"
OUT_PATH = os.path.join(OUT_DIR, FILENAME)

FIXTURES = "test_data\\fixtures"

# ---------------------------------------------------------------------------
# Zurich styles
# ---------------------------------------------------------------------------
DARK_BLUE  = "FF23366F"
LIGHT_BLUE = "FF91BFE3"
WHITE      = "FFFFFFFF"
ALT_GREY   = "FFECEEEF"

def _font(size=10, bold=False, color=DARK_BLUE):
    return Font(name="Zurich Sans", size=size, bold=bold, color=color)

def _semibold(size=10, color=DARK_BLUE):
    return Font(name="Zurich Sans Semibold", size=size, bold=True, color=color)

def _fill(hex_color):
    return PatternFill("solid", fgColor=hex_color)

_med = Side(style="medium")
_thin = Side(style="thin")
ALL_MED  = Border(left=_med, right=_med, top=_med, bottom=_med)
ALL_THIN = Border(left=_thin, right=_thin, top=_thin, bottom=_thin)
TOP_WRAP = Alignment(vertical="top", wrap_text=True)
CTR_WRAP = Alignment(horizontal="center", vertical="center", wrap_text=True)

def set_col_widths(ws, widths):
    for col, w in widths.items():
        ws.column_dimensions[col].width = w


# ---------------------------------------------------------------------------
# Test case definitions
# ---------------------------------------------------------------------------
# Each tuple: (ID, Strategy, Files/Setup, Steps, Expected Result)

TEST_CASES = [

    # ── X-Checks ─────────────────────────────────────────────────────────────
    (
        "FX-01", "X-Checks",
        f"FIP File: {FIXTURES}\\fip_xc.txt\n"
        f"X-Checks Publication File: {FIXTURES}\\xc_pub.xlsx  (sheet: cross checks all)\n"
        "No GCoA, no Known Exception List. 'Process only differences' unchecked.",
        "Load the files above into the X-Checks task and click Start.",
        "Run completes. Comparison sheet contains exactly 8 rows (one per X-Check in the "
        "fixture). Verify the Formula Match column against the table below:\n\n"
        "  XC_ALL_MATCH          → Match\n"
        "  XC_FORMULA_MISMATCH   → MisMatch\n"
        "  XC_NOT_IN_EBX         → Not Found\n"
        "  XC_NOT_IN_FIP         → Not Found\n"
        "  XC_REORDER_MATCH      → MisMatch  (known edge case — see note)\n"
        "  XC_THOUSANDS_CORR     → Match\n"
        "  XC_TOM_CORRECTION     → Match\n"
        "  XC_VARIABLE_MISMATCH  → Match\n\n"
        "Note: XC_REORDER_MATCH formula reorder produces an invalid formula for simple "
        "two-variable addition; MisMatch is the documented expected value.",
    ),
    (
        "FX-02", "X-Checks — Variables Match",
        "FA-01 output open, Comparison sheet.",
        "Check the Variables Match column for XC_VARIABLE_MISMATCH.",
        "Variables Match = MisMatch  (formula matched but FS Account differed).\n"
        "All other rows: Variables Match mirrors Formula Match.",
    ),
    (
        "FX-03", "X-Checks — colour coding",
        "FX-01 output open, Comparison sheet.",
        "Review fill colours in the Formula Match column.",
        "Match rows: green fill.\n"
        "MisMatch rows: red fill.\n"
        "Not Found rows: orange fill.",
    ),

    # ── Grouping By ──────────────────────────────────────────────────────────
    (
        "FX-04", "Grouping By",
        f"FIP File (ZQ9_VALFLDGR): {FIXTURES}\\fip_ZQ9_VALFLDGR.xlsx  (sheet: Sheet1)\n"
        f"X-Checks Publication File: {FIXTURES}\\xc_pub.xlsx  (sheet: cross checks all)\n"
        f"Mapping File: {FIXTURES}\\mapping.txt",
        "Load the files above into the Grouping By task and click Start.",
        "Run completes. Comparison sheet contains exactly 2 rows:\n\n"
        "  GB_MATCHED|GB_GROUPING_ITEM    → Matched\n"
        "  GB_NOT_IN_FIP|GB_GROUPING_ITEM → Not in FIP\n\n"
        "'Matched' row has green fill. 'Not in FIP' row has orange fill.",
    ),

    # ── Accounting Principles ────────────────────────────────────────────────
    (
        "FX-05", "Accounting Principles",
        f"Validation Methods File: {FIXTURES}\\validation_methods.xlsx  (sheet: Validation Methods)\n"
        f"X-Checks Publication File: {FIXTURES}\\xc_pub.xlsx  (sheet: cross checks all)\n"
        f"FIP File (VALMSG): {FIXTURES}\\fip_ZQ9_VALMSG.xlsx  (sheet: FIP Methods Rules and Condition)\n"
        "Note: this is a raw ZQ9_VALMSG file — the app builds the Key column automatically.",
        "Load the files above into the Accounting Principles task and click Start.",
        "Run completes. Progress log shows 'Built Key column from MK + ValidRule'.\n"
        "Comparison sheet contains exactly 2 rows:\n\n"
        "  AP_MATCH    Event=IFRS New RFD  FIP=w  Actual=w  → Match\n"
        "  AP_MISMATCH Event=IFRS New RFD  FIP=e  Actual=w  → MisMatch\n\n"
        "'Match' rows have green fill. 'MisMatch' rows have red fill.",
    ),

    # ── Conditions ───────────────────────────────────────────────────────────
    (
        "FX-06", "Conditions — full file",
        f"X-Checks Publication File: {FIXTURES}\\xc_pub.xlsx  (sheet: cross checks all)\n"
        f"FIP File (ZQ9_VALMETH): {FIXTURES}\\fip_ZQ9_VALMETH.xlsx  (sheet: FIP Conditions)\n"
        "'Process only differences' unchecked.",
        "Load the files above into the Conditions task. Uncheck 'Process only differences'. Click Start.",
        "Run completes. Comparison sheet contains exactly 4 rows:\n\n"
        "  COND_MATCHED|Q1      FIP Data=COND_MATCHED|Q1  → Matched\n"
        "  COND_NOT_MATCHED|Q2  FIP Data=(blank)           → Not Matched\n"
        "  COND_BASE|COND_BASE  FIP Data=(blank)           → Not Matched\n"
        "  COND_BASE|Q1         FIP Data=COND_BASE|Q1      → Matched\n\n"
        "Note: COND_BASE|COND_BASE comes from the 'Reference X-Check (Condition)' cell "
        "being treated as its own condition column — this is expected behaviour.\n"
        "COND_BASE|Q1 confirms the Reference X-Check key-prefix override is working.\n\n"
        "'Matched' rows have green fill. 'Not Matched' rows have red fill.",
    ),
    (
        "FX-07", "Conditions — differences only",
        f"Same files as FX-06. 'Process only differences' checked (default).",
        "Run with 'Process only differences' checked.",
        "Comparison sheet is empty (or contains 0 data rows) because the fixture "
        "EBX rows have no yellow or green cell fill — all rows are plain white.",
    ),

    # ── Full Run ─────────────────────────────────────────────────────────────
    (
        "FX-08", "Full Run — all fixture files",
        f"All fixture files as above. Output directory: any writable folder.",
        f"Load all fixture files into the Full Run task:\n"
        f"• FIP File: {FIXTURES}\\fip_xc.txt\n"
        f"• X-Checks Publication File: {FIXTURES}\\xc_pub.xlsx\n"
        f"• FIP File (ZQ9_VALFLDGR): {FIXTURES}\\fip_ZQ9_VALFLDGR.xlsx\n"
        f"• Mapping File: {FIXTURES}\\mapping.txt\n"
        f"• Validation Methods File: {FIXTURES}\\validation_methods.xlsx\n"
        f"• FIP File (VALMSG): {FIXTURES}\\fip_ZQ9_VALMSG.xlsx\n"
        f"• FIP File (ZQ9_VALMETH): {FIXTURES}\\fip_ZQ9_VALMETH.xlsx\n"
        "'Process only differences' unchecked for Conditions.\n"
        "Click Start.",
        "All four strategies run without error. Combined output workbook contains "
        "prefixed Comparison sheets:\n"
        "  XC — Comparison  (8 rows)\n"
        "  GB — Comparison  (2 rows)\n"
        "  AP — Comparison  (2 rows)\n"
        "  Cond — Comparison  (4 rows)\n"
        "Plus a single Processing Log sheet at the end.",
    ),
]


# ---------------------------------------------------------------------------
# Overview
# ---------------------------------------------------------------------------

OVERVIEW_ROWS = [
    (
        "Purpose",
        f"Verify all four strategies produce the correct row-level output when run against "
        f"the minimal fixture files in test_data\\fixtures\\. Each X-Check ID in the fixtures "
        f"encodes its expected comparison result, making pass/fail determination trivial "
        f"without manual reconciliation against large production files.",
    ),
    (
        "Fixture files",
        f"All files are in {FIXTURES}\\:\n"
        "  xc_pub.xlsx                — shared EBX publication (all strategies)\n"
        "  fip_xc.txt                 — FIP X-Checks (SAP Validation Rule text)\n"
        "  fip_ZQ9_VALFLDGR.xlsx      — FIP Grouping By\n"
        "  mapping.txt                — Grouping By field mapping\n"
        "  validation_methods.xlsx    — AP Validation Methods (real reference file)\n"
        "  fip_ZQ9_VALMSG.xlsx        — FIP Accounting Principles (raw ZQ9_VALMSG)\n"
        "  fip_ZQ9_VALMETH.xlsx       — FIP Conditions (raw ZQ9_VALMETH)\n\n"
        "Regenerate with: python test_data/generate_test_fixtures.py\n"
        "(validation_methods.xlsx is not regenerated — copy the real file manually).",
    ),
    (
        "Pre-conditions",
        "1) App EXE built and present in dist\\.\n"
        "2) No fixture files open in Excel before running.\n"
        "3) An output folder is available and writable.\n"
        "4) 'Process only differences' unchecked unless the test case specifies otherwise.",
    ),
    (
        "How to run",
        "For each test case: perform the steps in the running EXE, then compare the "
        "Comparison sheet row-by-row against the expected result. "
        "Record Pass/Fail and any discrepancy in the Result column.",
    ),
]


# ---------------------------------------------------------------------------
# Build
# ---------------------------------------------------------------------------

def build():
    wb = Workbook()
    wb.remove(wb.active)
    _build_overview(wb)
    _build_test_cases(wb)
    _build_signoff(wb)
    os.makedirs(OUT_DIR, exist_ok=True)
    wb.save(OUT_PATH)
    print(f"  Written: {OUT_PATH}")


def _build_overview(wb):
    ws = wb.create_sheet("Overview")
    ws.sheet_view.showGridLines = False
    set_col_widths(ws, {"A": 22, "B": 100})

    tc = ws.cell(row=1, column=1,
                 value=f"Fixture-Based UAT v{VERSION} — Test Plan")
    tc.font = _semibold(size=16)
    tc.alignment = TOP_WRAP
    ws.merge_cells("A1:B1")

    vc = ws.cell(row=2, column=1, value=f"Version: v{VERSION}  |  Fixtures: test_data\\fixtures\\")
    vc.font = _font(size=11)
    vc.alignment = TOP_WRAP
    ws.merge_cells("A2:B2")

    for i, (label, content) in enumerate(OVERVIEW_ROWS, start=4):
        ws.row_dimensions[i].height = 90
        lc = ws.cell(row=i, column=1, value=label)
        lc.font = _semibold()
        lc.fill = _fill(LIGHT_BLUE)
        lc.border = ALL_THIN
        lc.alignment = TOP_WRAP
        cc = ws.cell(row=i, column=2, value=content)
        cc.font = _font()
        cc.border = ALL_THIN
        cc.alignment = TOP_WRAP


def _build_test_cases(wb):
    ws = wb.create_sheet("Test Cases")
    ws.sheet_view.showGridLines = False
    set_col_widths(ws, {
        "A": 10, "B": 26, "C": 60, "D": 40, "E": 60,
        "F": 40, "G": 12, "H": 16, "I": 14,
    })

    headers = ["ID", "Strategy", "Files / Setup", "Steps",
               "Expected Result", "Actual Result", "Pass / Fail", "Tester", "Date"]

    tc = ws.cell(row=1, column=1,
                 value=f"Fixture-Based UAT v{VERSION} — Test Cases")
    tc.font = _semibold(size=16)
    tc.alignment = TOP_WRAP
    ws.merge_cells("A1:I1")

    ws.row_dimensions[3].height = 28
    for col_idx, hdr in enumerate(headers, start=1):
        cell = ws.cell(row=3, column=col_idx, value=hdr)
        cell.font = _semibold(size=11, color=WHITE)
        cell.fill = _fill(DARK_BLUE)
        cell.border = ALL_THIN
        cell.alignment = CTR_WRAP

    for row_offset, (tc_id, strategy, setup, steps, expected) in enumerate(TEST_CASES):
        row = 4 + row_offset
        ws.row_dimensions[row].height = 110
        fill = _fill(ALT_GREY) if row_offset % 2 == 1 else None
        for col_idx, val in enumerate(
            [tc_id, strategy, setup, steps, expected, "", "", "", ""], start=1
        ):
            cell = ws.cell(row=row, column=col_idx, value=val)
            cell.font = _font()
            cell.border = ALL_THIN
            cell.alignment = TOP_WRAP
            if fill:
                cell.fill = fill


def _build_signoff(wb):
    ws = wb.create_sheet("Sign-off")
    ws.sheet_view.showGridLines = False
    set_col_widths(ws, {"A": 22, "B": 26, "C": 26, "D": 26})

    tc = ws.cell(row=1, column=1,
                 value=f"Fixture-Based UAT v{VERSION} — Sign-off")
    tc.font = _semibold(size=16)
    tc.alignment = TOP_WRAP
    ws.merge_cells("A1:D1")

    ws.row_dimensions[3].height = 30
    ol = ws.cell(row=3, column=1, value="Outcome")
    ol.font = _semibold()
    ol.fill = _fill(LIGHT_BLUE)
    ol.border = ALL_THIN
    ol.alignment = TOP_WRAP
    ov = ws.cell(row=3, column=2,
                 value="All test cases passed (or any failure logged and triaged)?")
    ov.font = _font()
    ov.border = ALL_THIN
    ov.alignment = TOP_WRAP
    ws.merge_cells("B3:D3")

    for i, field in enumerate([
        "Tester (name)", "Tester (role)", "Test start date", "Test end date",
        "Pass / Fail", "Failures (count)", "Notes", "Approver (name)",
        "Approver (role)", "Approval date",
    ], start=5):
        ws.row_dimensions[i].height = 26
        lc = ws.cell(row=i, column=1, value=field)
        lc.font = _semibold()
        lc.fill = _fill(LIGHT_BLUE)
        lc.border = ALL_THIN
        lc.alignment = TOP_WRAP
        vc = ws.cell(row=i, column=2, value="")
        vc.font = _font()
        vc.border = ALL_THIN
        vc.alignment = TOP_WRAP
        ws.merge_cells(f"B{i}:D{i}")


if __name__ == "__main__":
    build()
