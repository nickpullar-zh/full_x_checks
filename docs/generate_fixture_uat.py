"""
Generate the Fixture-Based UAT Test Plan workbook.

Standalone plan covering the full app using test_data/fixtures/ files.
Each test case is marked as either:
  Logic       — verifies comparison output correctness (row-level results)
  Whole App   — verifies UI, output structure, colour coding, labelling, error handling

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

F = "test_data\\fixtures"

TASKS = [
    "Collect Live X-Checks",
    "X-Checks",
    "Grouping By",
    "Accounting Principles",
    "Conditions",
    "Full Run",
]

# ---------------------------------------------------------------------------
# Styles
# ---------------------------------------------------------------------------
DARK_BLUE  = "FF23366F"
LIGHT_BLUE = "FF91BFE3"
LOGIC_BLUE = "FFD6E4F7"   # light tint for Logic rows
WHITE      = "FFFFFFFF"
ALT_GREY   = "FFECEEEF"
GOLD       = "FFFFC000"   # whole-app row tint

def _font(size=10, bold=False, color=DARK_BLUE):
    return Font(name="Zurich Sans", size=size, bold=bold, color=color)

def _semibold(size=10, color=DARK_BLUE):
    return Font(name="Zurich Sans Semibold", size=size, bold=True, color=color)

def _fill(hex_color):
    return PatternFill("solid", fgColor=hex_color)

_thin = Side(style="thin")
ALL_THIN  = Border(left=_thin, right=_thin, top=_thin, bottom=_thin)
TOP_WRAP  = Alignment(vertical="top", wrap_text=True)
CTR_WRAP  = Alignment(horizontal="center", vertical="center", wrap_text=True)

def set_col_widths(ws, widths):
    for col, w in widths.items():
        ws.column_dimensions[col].width = w


# ---------------------------------------------------------------------------
# Test cases
# Format: (ID, Area, Test Type, Files/Setup, Steps, Expected Result)
# Test Type: "Logic" | "Whole App"
# ---------------------------------------------------------------------------

TEST_CASES = [

    # ── Launch & version ──────────────────────────────────────────────────────
    (
        "FX-01", "Launch & version", "Whole App",
        f"Production EXE: dist\\X-Checks_FullRun_{VERSION}.exe",
        f"Double-click dist\\X-Checks_FullRun_{VERSION}.exe.",
        f"Splash screen shows 'X-Check Application v{VERSION} Loading...'. "
        f"Task selector opens with 'X-Check Application v{VERSION}' in title bar and UI label.",
    ),

    # ── Task selector ─────────────────────────────────────────────────────────
    (
        "FX-02", "Task selector", "Whole App",
        "App open at task selector.",
        "Open the task dropdown and count entries.",
        f"{len(TASKS)} tasks listed in order:\n" +
        "\n".join(f"{i+1}. {t}" for i, t in enumerate(TASKS)),
    ),
    (
        "FX-03", "Task selector", "Whole App",
        "App open at task selector.",
        "Select each task in turn and confirm the form title.",
        "• Collect Live X-Checks → 'Collect Live X-Checks'\n"
        "• X-Checks → 'X-Check Files'\n"
        "• Grouping By → 'Grouping By Files'\n"
        "• Accounting Principles → 'Accounting Principles Files'\n"
        "• Conditions → 'Conditions Files'\n"
        "• Full Run → 'Full Run — All Strategies'",
    ),

    # ── X-Checks ─────────────────────────────────────────────────────────────
    (
        "FX-04", "X-Checks — file fields", "Whole App",
        "X-Checks task selected, form open.",
        "Inspect the form fields.",
        "Four fields present:\n"
        "1. 'FIP File' (.txt)\n"
        "2. 'X-Checks Publication File' (.xlsx, default sheet: cross checks all)\n"
        "3. 'GCoA Publication File' (.xlsx, optional)\n"
        "4. 'Known Exception List' (.xlsx, optional)\n"
        "Two experimental checkboxes unchecked by default. "
        "'Process only differences' present and checked by default.",
    ),
    (
        "FX-05", "X-Checks — comparison output", "Logic",
        f"FIP File: {F}\\fip_xc.txt\n"
        f"X-Checks Publication File: {F}\\xc_pub.xlsx  (sheet: cross checks all)\n"
        "No GCoA. No Known Exception List. 'Process only differences' unchecked.",
        "Load the files above into X-Checks and click Start.",
        "Run completes. Comparison sheet contains exactly 11 rows. "
        "Formula Match column values:\n\n"
        "  XC_ALL_MATCH          → Match\n"
        "  XC_DIFF_EXCLUDED      → Not Found  (EBX row present, no FIP block)\n"
        "  XC_DIFF_IN_SCOPE      → Match  (also in-scope for differences mode)\n"
        "  XC_FORMULA_MISMATCH   → MisMatch\n"
        "  XC_KEL_MISMATCH       → MisMatch  (without KEL; annotated when KEL supplied)\n"
        "  XC_NOT_IN_EBX         → Not Found\n"
        "  XC_NOT_IN_FIP         → Not Found\n"
        "  XC_REORDER_MATCH      → MisMatch  (known edge case in reorder logic)\n"
        "  XC_THOUSANDS_CORR     → Match  (FIP '1.000' stripped to '1000')\n"
        "  XC_TOM_CORRECTION     → Match  (FIP 'TOM' normalised to 'ToM')\n"
        "  XC_VARIABLE_MISMATCH  → Match\n\n"
        "Note: XC_DIFF_EXCLUDED appears in the Comparison sheet (Not Found) even "
        "though it is excluded from the X-Check No Selection .txt output.",
    ),
    (
        "FX-06", "X-Checks — Variables Match", "Logic",
        "FX-05 output open, Comparison sheet.",
        "Check the Variables Match column.",
        "XC_VARIABLE_MISMATCH: Variables Match = MisMatch  "
        "(formula matched but FS Account differed).\n"
        "All other rows: Variables Match mirrors Formula Match.",
    ),
    (
        "FX-07", "X-Checks — output structure", "Whole App",
        "FX-05 complete. Output workbook open.",
        "Check the sheet tabs.",
        "Workbook contains:\n"
        "1. Comparison\n"
        "2. Processing Log",
    ),
    (
        "FX-08", "X-Checks — colour coding", "Whole App",
        "FX-05 output open, Comparison sheet.",
        "Review fill colours in the Formula Match column.",
        "Match rows → green fill.\n"
        "MisMatch rows → red fill.\n"
        "Not Found rows → orange fill.",
    ),
    (
        "FX-09", "X-Checks — Known Exception annotation", "Logic",
        f"Same files as FX-05. Known Exception List: {F}\\known_exception_list.xlsx  (sheet: X-Checks).\n"
        "The KEL contains one entry for XC_KEL_MISMATCH.",
        "Add the Known Exception List file and run.",
        "Run completes. Progress log shows 'Known exceptions loaded  (1)'.\n"
        "XC_KEL_MISMATCH row: Formula Match = MisMatch (unchanged), "
        "'Known Exception' column = 'Test fixture — expected mismatch' with blue fill.\n"
        "All other rows: 'Known Exception' column is blank.",
    ),

    (
        "FX-10", "X-Checks — differences mode (X-Check No Selection)", "Logic",
        f"FIP File: {F}\\fip_xc.txt\n"
        f"X-Checks Publication File: {F}\\xc_pub.xlsx  (sheet: cross checks all)\n"
        "'Process only differences' checked.",
        "Check 'Process only differences'. Load files and click Start.",
        "A .txt file is written alongside the output containing exactly 1 X-Check No:\n\n"
        "  XC_DIFF_IN_SCOPE\n\n"
        "XC_DIFF_EXCLUDED is absent (Exclude Z-Core = X). "
        "All other X-Checks are absent (blank Type of change).",
    ),

    # ── Grouping By ──────────────────────────────────────────────────────────
    (
        "FX-11", "Grouping By — file fields", "Whole App",
        "Grouping By task selected, form open.",
        "Inspect the form fields.",
        "Four fields present:\n"
        "1. 'FIP File (ZQ9_VALFLDGR)' (.xlsx, default sheet: Sheet1)\n"
        "2. 'X-Checks Publication File' (.xlsx, default sheet: cross checks all)\n"
        "3. 'Mapping File' (.csv / .txt)\n"
        "4. 'Known Exception List' (.xlsx, optional)",
    ),
    (
        "FX-12", "Grouping By — comparison output", "Logic",
        f"FIP File (ZQ9_VALFLDGR): {F}\\fip_ZQ9_VALFLDGR.xlsx  (sheet: Sheet1)\n"
        f"X-Checks Publication File: {F}\\xc_pub.xlsx  (sheet: cross checks all)\n"
        f"Mapping File: {F}\\mapping.txt",
        "Load the files above into Grouping By and click Start.",
        "Run completes. Comparison sheet contains exactly 2 rows:\n\n"
        "  GB_MATCHED|GB_GROUPING_ITEM    → Matched\n"
        "  GB_NOT_IN_FIP|GB_GROUPING_ITEM → Not in FIP",
    ),
    (
        "FX-13", "Grouping By — output structure", "Whole App",
        "FX-12 complete. Output workbook open.",
        "Check the sheet tabs.",
        "Workbook contains:\n"
        "1. Mapping File\n"
        "2. FIP - Original\n"
        "3. FIP - Processed\n"
        "4. EBX - Original\n"
        "5. EBX - Processed\n"
        "6. Comparison\n"
        "7. Processing Log",
    ),
    (
        "FX-14", "Grouping By — colour coding", "Whole App",
        "FX-12 output open, Comparison sheet.",
        "Review the Result column fill colours.",
        "'Matched' → green fill.\n"
        "'Not in FIP' → orange fill.",
    ),

    # ── Accounting Principles ─────────────────────────────────────────────────
    (
        "FX-15", "Accounting Principles — file fields", "Whole App",
        "Accounting Principles task selected, form open.",
        "Inspect the form fields.",
        "Four fields present:\n"
        "1. 'Validation Methods File' (.xlsx, default sheet: Validation Methods)\n"
        "2. 'X-Checks Publication File' (.xlsx, default sheet: cross checks all)\n"
        "3. 'FIP File (VALMSG)' (.xlsx, sheet: FIP Methods Rules and Condition)\n"
        "4. 'Known Exception List' (.xlsx, optional)",
    ),
    (
        "FX-16", "Accounting Principles — Key column built from raw VALMSG", "Logic",
        f"Validation Methods File: {F}\\validation_methods.xlsx  (sheet: Validation Methods)\n"
        f"X-Checks Publication File: {F}\\xc_pub.xlsx  (sheet: cross checks all)\n"
        f"FIP File (VALMSG): {F}\\fip_ZQ9_VALMSG.xlsx  (sheet: FIP Methods Rules and Condition)\n"
        "Note: fip_ZQ9_VALMSG.xlsx is a raw ZQ9_VALMSG export — no pre-built Key column.",
        "Load the files above into Accounting Principles and click Start.",
        "Run completes. Progress log shows 'Built Key column from MK + ValidRule'. "
        "Comparison sheet contains exactly 2 rows:\n\n"
        "  AP_MATCH    Event=IFRS New RFD  FIP=w  Actual=w  → Match\n"
        "  AP_MISMATCH Event=IFRS New RFD  FIP=e  Actual=w  → MisMatch",
    ),
    (
        "FX-17", "Accounting Principles — output structure", "Whole App",
        "FX-16 complete. Output workbook open.",
        "Check the sheet tabs.",
        "Workbook contains:\n"
        "1. EBX\n"
        "2. FIP\n"
        "3. Comparison\n"
        "4. Processing Log",
    ),
    (
        "FX-18", "Accounting Principles — colour coding", "Whole App",
        "FX-16 output open, Comparison sheet.",
        "Review the Match column fill colours.",
        "'Match' → green fill.\n"
        "'MisMatch' → red fill.",
    ),

    # ── Conditions ────────────────────────────────────────────────────────────
    (
        "FX-19", "Conditions — file fields", "Whole App",
        "Conditions task selected, form open.",
        "Inspect the form fields.",
        "Three fields present:\n"
        "1. 'X-Checks Publication File' (.xlsx, default sheet: cross checks all)\n"
        "2. 'FIP File (ZQ9_VALMETH)' (.xlsx, default sheet: FIP Conditions)\n"
        "3. 'Known Exception List' (.xlsx, optional)\n"
        "'Process only differences' checkbox present and checked by default.",
    ),
    (
        "FX-20", "Conditions — full file run", "Logic",
        f"X-Checks Publication File: {F}\\xc_pub.xlsx  (sheet: cross checks all)\n"
        f"FIP File (ZQ9_VALMETH): {F}\\fip_ZQ9_VALMETH.xlsx  (sheet: FIP Conditions)\n"
        "'Process only differences' unchecked.",
        "Uncheck 'Process only differences'. Load the files and click Start.",
        "Run completes. Comparison sheet contains exactly 5 rows:\n\n"
        "  COND_MATCHED|Q1          FIP Data=COND_MATCHED|Q1      → Matched\n"
        "  COND_NOT_MATCHED|Q2      FIP Data=(blank)               → Not Matched\n"
        "  COND_BASE|COND_BASE      FIP Data=(blank)               → Not Matched\n"
        "  COND_BASE|Q1             FIP Data=COND_BASE|Q1          → Matched\n"
        "  COND_DIFF_YELLOW|Q1      FIP Data=COND_DIFF_YELLOW|Q1   → Matched\n"
        "  COND_DIFF_WHITE|Q1       FIP Data=(blank)               → Not Matched\n\n"
        "Note: COND_BASE|COND_BASE is expected — the 'Reference X-Check (Condition)' "
        "column is itself one of the 5 condition columns.",
    ),
    (
        "FX-21", "Conditions — differences mode (yellow + green cells)", "Logic",
        f"Same files as FX-20. 'Process only differences' checked (default).",
        "Run with 'Process only differences' checked (default).",
        "Comparison sheet contains exactly 2 rows:\n\n"
        "  COND_DIFF_GREEN|Q1   FIP Data=COND_DIFF_GREEN|Q1   → Matched\n"
        "  COND_DIFF_YELLOW|Q1  FIP Data=COND_DIFF_YELLOW|Q1  → Matched\n\n"
        "Yellow-filled and green-filled Applicable Quarters cells are both collected. "
        "COND_DIFF_WHITE (plain white cell) is not collected — no output row.",
    ),
    (
        "FX-22", "Conditions — output structure", "Whole App",
        "FX-20 complete. Output workbook open.",
        "Check the sheet tabs and first column header of the FIP Data sheet.",
        "Workbook contains:\n"
        "1. Working Sheet\n"
        "2. FIP Data  (first column header: 'Key (Concatenated)')\n"
        "3. Comparison\n"
        "4. Processing Log",
    ),
    (
        "FX-23", "Conditions — colour coding", "Whole App",
        "FX-20 output open, Comparison sheet.",
        "Review the Comparison column fill colours.",
        "'Matched' → green fill.\n"
        "'Not Matched' → red fill.",
    ),

    # ── Full Run ──────────────────────────────────────────────────────────────
    (
        "FX-24", "Full Run — file fields", "Whole App",
        "Full Run task selected, form open.",
        "Count the file fields and verify no duplicates.",
        "All unique fields from every strategy are merged into one form. "
        "Fields include: FIP File, X-Checks Publication File, GCoA Publication File, "
        "FIP File (ZQ9_VALFLDGR), Mapping File, Validation Methods File, "
        "FIP File (VALMSG), FIP File (ZQ9_VALMETH), Known Exception List. "
        "No field label appears twice.",
    ),
    (
        "FX-25", "Full Run — all strategies", "Logic",
        f"All fixture files. 'Process only differences' unchecked.\n"
        f"• FIP File: {F}\\fip_xc.txt\n"
        f"• X-Checks Publication File: {F}\\xc_pub.xlsx\n"
        f"• FIP File (ZQ9_VALFLDGR): {F}\\fip_ZQ9_VALFLDGR.xlsx\n"
        f"• Mapping File: {F}\\mapping.txt\n"
        f"• Validation Methods File: {F}\\validation_methods.xlsx\n"
        f"• FIP File (VALMSG): {F}\\fip_ZQ9_VALMSG.xlsx\n"
        f"• FIP File (ZQ9_VALMETH): {F}\\fip_ZQ9_VALMETH.xlsx",
        "Load all fixture files into Full Run. Uncheck 'Process only differences'. Click Start.",
        "All four strategies run without error. Combined output contains:\n\n"
        "  XC — Comparison    11 rows\n"
        "  GB — Comparison    2 rows\n"
        "  AP — Comparison    2 rows\n"
        "  Cond — Comparison  6 rows\n\n"
        "Row counts and values match the individual strategy results in FX-05, FX-12, FX-16, FX-20.",
    ),
    (
        "FX-26", "Full Run — combined output structure", "Whole App",
        "FX-25 complete. Combined output workbook open.",
        "Check all sheet tabs and tab colours.",
        "One workbook with all strategy sheets prefixed by strategy name "
        "(e.g. 'XC — Comparison', 'GB — Comparison', 'AP — Comparison', 'Cond — Comparison'). "
        "Tabs are colour-coded by strategy. Single 'Processing Log' sheet at the end.",
    ),
    (
        "FX-27", "Full Run — abort on strategy failure", "Whole App",
        "Full Run form open.",
        "Set an incorrect sheet name for one file, then click Start.",
        "Failing strategy logs a clear error. Full Run aborts immediately — "
        "does not continue to the next strategy. 'Return to Form' is available.",
    ),

    # ── Settings / Known Exception Builder ────────────────────────────────────
    (
        "FX-28", "Settings — gear menu", "Whole App",
        "App open at task selector.",
        "Click the ⚙ gear button at the bottom-right.",
        "A popup menu appears with at least 'Build Known Exception List…'. "
        "No dialog opens directly.",
    ),
    (
        "FX-29", "Settings — Known Exception Builder", "Whole App",
        "Settings popup open (FX-28).",
        "Click 'Build Known Exception List…'.",
        "Modal dialog opens with: 'Save as' hint text "
        "'Click Browse and select a folder, then type the filename', "
        "Browse button, optional comparison import section, "
        "'Open file after building' checkbox (checked by default), Build button.",
    ),
    (
        "FX-30", "Settings — build and open KEL", "Whole App",
        "Known Exception Builder dialog open. Output folder available.",
        "Click Browse, select a folder, type a filename. Leave 'Open file after building' checked. Click Build.",
        "File created at the chosen path. Dialog closes. File opens in Excel. "
        "Contains sheets: X-Checks, Grouping By, Accounting Principles, Conditions, Instructions. "
        "Row 2 of each strategy sheet is a guidance row. "
        "File carries the 'Internal Use Only' sensitivity label.",
    ),

    # ── Processing Log ────────────────────────────────────────────────────────
    (
        "FX-31", "Processing Log — content", "Whole App",
        "Any completed run. Output workbook open, Processing Log sheet.",
        "Review the log entries.",
        f"First entry shows v{VERSION}. "
        "Log includes: files loaded, strategy steps, output path, expected sensitivity label. "
        "All entries have Timestamp, File, Step, Count columns.",
    ),
    (
        "FX-32", "Processing Log — output path entry", "Whole App",
        "Any completed run. Processing Log sheet open.",
        "Find the 'Output written to' entry.",
        "Entry with File='Output' and Step starting 'Output written to:' is present.",
    ),
    (
        "FX-33", "Processing Log — sensitivity label entry", "Whole App",
        "Any completed run. Processing Log sheet open.",
        "Find the sensitivity label entry.",
        "Entry with File='Sensitivity' and Step='Expected label: Internal_Use_Only' is present.",
    ),

    # ── Sensitivity label ──────────────────────────────────────────────────────
    (
        "FX-34", "Sensitivity label", "Whole App",
        "Any completed run. Output .xlsx saved to disk.",
        "Right-click the output file in Explorer → Properties → Details, "
        "or open in Excel and check the sensitivity bar.",
        "File carries the 'Internal Use Only' MIP label. "
        "Progress dialog shows 'Applied label: Internal_Use_Only'.",
    ),

    # ── Stop / error handling ──────────────────────────────────────────────────
    (
        "FX-35", "Stop / Return to Form", "Whole App",
        "Any task started.",
        "Click Stop during processing.",
        "Processing halts. Dialog shows 'Processing halted by user'. "
        "'Return to Form' reopens the form with previously chosen files pre-filled.",
    ),
    (
        "FX-36", "Error — wrong sheet name", "Whole App",
        "Any task's file-selection form open.",
        "Set a sheet name to 'does_not_exist', then click Start.",
        "Run aborts with a clear error identifying the missing sheet. "
        "App returns to form — does not crash.",
    ),
    (
        "FX-37", "Error — missing required file", "Whole App",
        "Any task's file-selection form open.",
        "Click Start without selecting any required files.",
        "Start does not begin processing. Form indicates missing required fields.",
    ),
]

LOGIC_COUNT    = sum(1 for t in TEST_CASES if t[2] == "Logic")
WHOLE_APP_COUNT = sum(1 for t in TEST_CASES if t[2] == "Whole App")

# ---------------------------------------------------------------------------
# Overview
# ---------------------------------------------------------------------------

OVERVIEW_ROWS = [
    (
        "Purpose",
        f"Standalone UAT for X-Checks Full Application v{VERSION} using the minimal "
        f"fixture files in test_data\\fixtures\\. "
        f"Covers all strategies, UI, output structure, colour coding, sensitivity labelling, "
        f"and error handling.\n\n"
        f"Test cases are categorised:\n"
        f"  Logic ({LOGIC_COUNT} cases)     — verifies row-level comparison output. "
        f"Run after every code change to confirm strategy logic is correct.\n"
        f"  Whole App ({WHOLE_APP_COUNT} cases) — verifies UI, structure, labelling, error handling. "
        f"Run for full release sign-off.",
    ),
    (
        "Fixture files",
        f"All in test_data\\fixtures\\:\n"
        "  xc_pub.xlsx               — shared EBX publication (all strategies)\n"
        "  fip_xc.txt                — FIP X-Checks text\n"
        "  fip_ZQ9_VALFLDGR.xlsx     — FIP Grouping By\n"
        "  mapping.txt               — Grouping By mapping\n"
        "  validation_methods.xlsx   — AP Validation Methods (real reference file)\n"
        "  fip_ZQ9_VALMSG.xlsx       — FIP AP (raw ZQ9_VALMSG)\n"
        "  fip_ZQ9_VALMETH.xlsx      — FIP Conditions (raw ZQ9_VALMETH)\n\n"
        "Regenerate (except validation_methods.xlsx) with:\n"
        "  python test_data/generate_test_fixtures.py",
    ),
    (
        "Test executable",
        f"dist\\X-Checks_FullRun_{VERSION}.exe",
    ),
    (
        "Pre-conditions",
        "1) EXE present in dist\\.\n"
        "2) No fixture files open in Excel.\n"
        "3) An output folder is available and writable.\n"
        "4) MIP client installed (for sensitivity-label checks).\n"
        "5) 'Process only differences' unchecked unless the test case specifies otherwise.",
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

    vc = ws.cell(row=2, column=1,
                 value=f"Version: v{VERSION}  |  {LOGIC_COUNT} Logic cases  |  {WHOLE_APP_COUNT} Whole App cases")
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
        "A": 10, "B": 26, "C": 14, "D": 55, "E": 38,
        "F": 55, "G": 38, "H": 12, "I": 16, "J": 14,
    })

    headers = ["ID", "Area", "Test Type", "Files / Setup", "Steps",
               "Expected Result", "Actual Result", "Pass / Fail", "Tester", "Date"]

    tc = ws.cell(row=1, column=1,
                 value=f"Fixture-Based UAT v{VERSION} — Test Cases")
    tc.font = _semibold(size=16)
    tc.alignment = TOP_WRAP
    ws.merge_cells("A1:J1")

    # Legend row
    ws.row_dimensions[2].height = 18
    leg = ws.cell(row=2, column=1,
                  value="Test Type:   Logic = verify comparison output logic     Whole App = verify UI, structure, labelling, error handling")
    leg.font = _font(size=9, color="FF444444")
    leg.alignment = TOP_WRAP
    ws.merge_cells("A2:J2")

    ws.row_dimensions[3].height = 28
    for col_idx, hdr in enumerate(headers, start=1):
        cell = ws.cell(row=3, column=col_idx, value=hdr)
        cell.font = _semibold(size=11, color=WHITE)
        cell.fill = _fill(DARK_BLUE)
        cell.border = ALL_THIN
        cell.alignment = CTR_WRAP

    for row_offset, (tc_id, area, test_type, setup, steps, expected) in enumerate(TEST_CASES):
        row = 4 + row_offset
        ws.row_dimensions[row].height = 100
        # Logic rows: light blue tint; Whole App rows: plain white
        row_fill = _fill(LOGIC_BLUE) if test_type == "Logic" else None
        for col_idx, val in enumerate(
            [tc_id, area, test_type, setup, steps, expected, "", "", "", ""], start=1
        ):
            cell = ws.cell(row=row, column=col_idx, value=val)
            cell.font = _font()
            cell.border = ALL_THIN
            cell.alignment = TOP_WRAP
            if row_fill:
                cell.fill = row_fill
        # Bold the test-type cell
        ws.cell(row=row, column=3).font = _semibold()


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
                 value="All in-scope cases passed (or any failure logged and triaged)?")
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
