"""
Generates the functional test plan workbook for the X-Check Application v1.0.0.

Output: docs/<YYYYMMDD> X-Checks_v<version> Test Plan.xlsx

Run:  python docs/generate_uat.py
"""
import os
import sys
from datetime import date

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from version import __version__

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, Side

APP_NAME = "X-Checks"

# ---------------------------------------------------------------------------
# Styles
# ---------------------------------------------------------------------------
NARRATIVE_FONT = Font(name="Aptos Narrow", size=11)
NARRATIVE_BOLD = Font(name="Aptos Narrow", size=11, bold=True)
ZS_LIGHT       = Font(name="Zurich Sans Light",  size=10, color="FF4D4D4D")
ZS_MEDIUM      = Font(name="Zurich Sans Medium", size=20, color="FF2167AE")
ZS_HEADER      = Font(name="Zurich Sans Light",  size=10, bold=True, color="FF4D4D4D")

MEDIUM         = Side(style="medium", color="FF000000")
TOP_BOT_BORDER = Border(top=MEDIUM, bottom=MEDIUM)
BOT_BORDER     = Border(bottom=MEDIUM)
WRAP_TOP       = Alignment(vertical="top",    wrap_text=True)
WRAP_CENTER    = Alignment(vertical="center", wrap_text=True)

# ---------------------------------------------------------------------------
# Sheet 1 — Instructions (narrative walkthrough)
# ---------------------------------------------------------------------------
FEATURES = [
    "Task selector launcher — dropdown lists all five strategies; Start button disabled until selection made.",
    "X-Checks strategy — compares EBX cross-check formulas and variables against FIP validation rules; flags differences and known exceptions.",
    "File upload dialog — dynamically built from config; sheet name fields enabled on file select.",
    "Sheet name advisory label — displayed when any field has a sheet name input.",
    "Grouping By strategy — compares X-Check grouping data (EBX publication) against FIP ZQ9_VALFLDGR extract.",
    "Accounting Principles strategy — compares severity letters (W/E) from EBX publication against FIP VALMSG.",
    "Conditions strategy — matches X-Check|Condition pairs from EBX publication against FIP ZQ9_VALMETH.",
    "Full Run strategy — runs all three strategies sequentially and combines output into one colour-coded workbook.",
    "Progress dialog — real-time scrollable log; error lines in bold red; Stop / Return to Form / Exit buttons.",
    "Sensitivity labelling — Microsoft Information Protection label applied automatically to every output workbook.",
    "Output to ~/Downloads/Output — all runs write to the same folder for easy access.",
]

# (Test, Action, Expected)  — empty Test/Action = continuation row
PLAN_ROWS = [
    # --- Launch ---
    ("App launches",
     "Double-click X-Checks_v{ver}.exe.",
     "Splash screen appears showing version v{ver}."),
    ("", "",
     "Splash closes and the task selector dialog opens showing 'X-Check Application v{ver}'."),
    ("", "",
     "Dropdown lists exactly five options: X-Checks, Grouping By, Accounting Principles, Conditions, Full Run."),
    ("", "",
     "Start button is disabled until an option is selected."),

    # --- Task selector ---
    ("Task selector — select task",
     "Select any task from the dropdown.",
     "Start button becomes enabled."),
    ("Task selector — close",
     "Click the red X on the task selector.",
     "Application closes cleanly with no error."),

    # --- File upload dialog ---
    ("File upload dialog opens",
     "Select 'Grouping By' and click Start.",
     "File upload dialog opens titled 'Grouping By Files'."),
    ("", "",
     "Required field labels show a trailing asterisk (*); optional fields show '(optional)'."),
    ("", "",
     "All file path labels show 'No file selected' in grey."),
    ("", "",
     "Sheet name fields are greyed out / disabled."),
    ("", "",
     "Advisory note 'If the sheet name changes from the current default in the Excel workbook, update it manually.' is visible, centred."),
    ("", "",
     "Proceed button is disabled."),

    ("Browse — file select",
     "Click Browse next to 'FIP File (ZQ9_VALFLDGR)' and select the VALFLDGR test file.",
     "File path label updates to the filename (not full path)."),
    ("", "",
     "Sheet name field for that row becomes enabled."),
    ("", "",
     "Proceed button remains disabled (other required files not yet provided)."),

    ("Proceed enables",
     "Select valid files for all required fields and an output directory.",
     "Proceed button becomes enabled."),

    ("Cancel dialog",
     "Close the upload dialog using the red X.",
     "Task selector reappears."),

    # --- X-Checks ---
    ("X-Checks — run (debug EXE)",
     "Run X-Checks_Debug_XChecks_v{ver}.exe.",
     "App launches in debug mode; progress dialog opens immediately with pre-loaded test files."),
    ("", "",
     "Progress log shows 'Loading files into memory...' then 'Files loaded successfully'."),
    ("", "",
     "Log shows EBX extraction steps, FIP extraction steps, then comparison."),
    ("", "",
     "Log ends with 'Processing complete. You may close this window.'"),
    ("", "",
     "Output workbook written to ~/Downloads/Output/ with timestamp prefix."),
    ("", "",
     "Workbook contains 2 sheets: X-Checks Comparison, Processing Log."),
    ("", "",
     "X-Checks Comparison sheet has a 'Result' column with colour-coded values (Matched / MisMatch / Not Found etc.)."),

    # --- Grouping By ---
    ("Grouping By — run (debug EXE)",
     "Run X-Checks_Debug_GroupingBy_v{ver}.exe.",
     "App launches in debug mode; progress dialog opens immediately with pre-loaded test files."),
    ("", "",
     "Progress log shows 'Loading files into memory...' then 'Files loaded successfully'."),
    ("", "",
     "Log shows EBX processing steps: Original file loaded → Filtered → Deduplicated → Split → Stacked."),
    ("", "",
     "Log shows Compare result: 'Matched: N | Not in FIP: M' with N > 0."),
    ("", "",
     "Log ends with 'Processing complete. You may close this window.'"),
    ("", "",
     "Output workbook written to ~/Downloads/Output/ with timestamp prefix."),
    ("", "",
     "Workbook contains 7 sheets: Mapping File, FIP - Original, FIP - Processed, EBX - Original, EBX - Processed, Compare, Processing Log."),
    ("", "",
     "Compare sheet has a 'Result' column with values 'Matched' and/or 'Not in FIP'."),

    # --- Accounting Principles ---
    ("Accounting Principles — run (debug EXE)",
     "Run X-Checks_Debug_AccountingPrinciples_v{ver}.exe.",
     "App launches in debug mode; progress dialog opens with pre-loaded test files."),
    ("", "",
     "Log shows 'Validation Methods — Method bindings extracted N' with N > 0."),
    ("", "",
     "Log shows 'EBX — In-scope X-Check Nos N'."),
    ("", "",
     "Log shows 'Compare — Comparison rows produced N'."),
    ("", "",
     "Log ends with 'Processing complete. You may close this window.'"),
    ("", "",
     "Output workbook written to ~/Downloads/Output/."),
    ("", "",
     "Workbook contains 4 sheets: EBX, FIP, Comparison, Processing Log."),
    ("", "",
     "Comparison sheet has green rows (Match) and/or red rows (MisMatch)."),

    # --- Conditions ---
    ("Conditions — run (debug EXE)",
     "Run X-Checks_Debug_Conditions_v{ver}.exe.",
     "App launches in debug mode; progress dialog opens with pre-loaded test files."),
    ("", "",
     "Log shows 'Publication — Extracted N unique X-Check No. entries'."),
    ("", "",
     "Log shows 'Comparison — Pairs checked: N, Matched: M, Not matched: P'."),
    ("", "",
     "Log ends with 'Processing complete. You may close this window.'"),
    ("", "",
     "Output workbook written to ~/Downloads/Output/."),
    ("", "",
     "Workbook contains 4 sheets: Conditions, Working Sheet, FIP Data, Processing Log."),
    ("", "",
     "Conditions sheet has a summary block at the top and TRUE/FALSE rows below."),

    # --- Full Run ---
    ("Full Run — run (debug EXE)",
     "Run X-Checks_Debug_FullRun_v{ver}.exe.",
     "App launches in debug mode; progress dialog opens with pre-loaded test files."),
    ("", "",
     "Log shows '— Starting: Grouping By —', then '— Starting: Accounting Principles —', then '— Starting: Conditions —' in order."),
    ("", "",
     "Log shows 'Writing combined workbook (12 sheets)'."),
    ("", "",
     "Log ends with 'Processing complete. You may close this window.'"),
    ("", "",
     "Single output workbook written to ~/Downloads/Output/."),
    ("", "",
     "Workbook contains 13 sheets: 6 GB tabs (orange), 3 AP tabs (dark blue), 3 Cond tabs (blue), 1 Processing Log (grey)."),
    ("", "",
     "Tab colours match Zurich brand palette: Grouping By = orange, Accounting Principles = dark blue, Conditions = Zurich blue."),

    # --- Error handling ---
    ("Wrong sheet name",
     "In the upload dialog, change a sheet name to a name that does not exist in the file, then click Proceed.",
     "Progress dialog opens and immediately shows an error: 'Could not find sheet … in …'."),
    ("", "",
     "'Return to Form' button appears; clicking it re-opens the upload dialog with previous inputs pre-filled."),

    ("Stop mid-run",
     "Start a run (e.g. Full Run) and click Stop during processing.",
     "Processing halts after the current step completes."),
    ("", "",
     "Log shows 'Processing halted by user.'"),
    ("", "",
     "'Return to Form' button appears."),

    # --- Version ---
    ("Version displayed",
     "Inspect the title bar of the task selector and the first line of any processing log.",
     "Title bar reads 'X-Check Application v{ver}'."),
    ("", "",
     "First log entry reads 'System — X-Check Application v{ver}'."),
]


# ---------------------------------------------------------------------------
# Sheet 2 — Files Required
# (CaseUI, Field Label, Filename, Sheet, Required)
# ---------------------------------------------------------------------------
FILES_ROWS = [
    # X-Checks
    ("X-Checks", "FIP File",                  "20260318 FIP X-Checks.txt",              "N/A",              "Yes"),
    ("X-Checks", "X-Checks Publication File", "20260313 Cross Checks All.xlsx",          "cross checks all", "Yes"),
    ("X-Checks", "GCoA Publication File",     "GCoA Publication file (optional)",        "GCoA Base account table", "No"),
    ("X-Checks", "Known Exception List",      "Known Exception List (optional)",         "Known Exceptions", "No"),
    # Grouping By
    ("Grouping By", "FIP File (ZQ9_VALFLDGR)",   "VALFLDGR file with 12348 Data rows on sheet Sheet1.XLSX", "Sheet1",           "Yes"),
    ("Grouping By", "X-Checks Publication File",  "20260313 Cross Checks All.xlsx",                         "cross checks all", "Yes"),
    ("Grouping By", "Mapping File",               "Mapping Table with 20 rows.txt",                         "N/A",              "Yes"),
    # Accounting Principles
    ("Accounting Principles", "Validation Methods File",   "validation methods.xlsx",                                "Validation Methods",            "Yes"),
    ("Accounting Principles", "X-Checks Publication File", "20260602 VALMSG (Accounting Principle).xlsx",            "cross checks all",              "Yes"),
    ("Accounting Principles", "FIP File (VALMSG)",         "20260602 VALMSG (Accounting Principle).xlsx",            "FIP Methods Rules and Condition","Yes"),
    # Conditions
    ("Conditions", "X-Checks Publication File", "20260313 Cross Checks All - Copy.xlsx", "cross checks all", "Yes"),
    ("Conditions", "FIP File (ZQ9_VALMETH)",     "20260602 VALMETH (Conditions).xlsx",    "FIP Conditions",   "Yes"),
    # Full Run
    ("Full Run", "FIP File",                   "20260318 FIP X-Checks.txt",                              "N/A",                           "Yes"),
    ("Full Run", "FIP File (ZQ9_VALFLDGR)",    "VALFLDGR file with 12348 Data rows on sheet Sheet1.XLSX", "Sheet1",                        "Yes"),
    ("Full Run", "X-Checks Publication File",  "20260313 Cross Checks All.xlsx",                         "cross checks all",              "Yes"),
    ("Full Run", "Mapping File",               "Mapping Table with 20 rows.txt",                         "N/A",                           "Yes"),
    ("Full Run", "Validation Methods File",    "validation methods.xlsx",                                "Validation Methods",            "Yes"),
    ("Full Run", "FIP File (VALMSG)",          "20260602 VALMSG (Accounting Principle).xlsx",            "FIP Methods Rules and Condition","Yes"),
    ("Full Run", "FIP File (ZQ9_VALMETH)",      "20260602 VALMETH (Conditions).xlsx",                     "FIP Conditions",                "Yes"),
]


# ---------------------------------------------------------------------------
# Sheet 2 — General UI cases  (Test, Action, Expected)
# ---------------------------------------------------------------------------
GENERAL_UI_CASES = [
    ("Splash screen on launch",
     "Double-click X-Checks_v{ver}.exe.",
     "Splash screen displays with version v{ver} and closes automatically when the main window is ready."),
    ("Task selector title",
     "Observe the task selector window title and heading label.",
     "Title bar and heading both read 'X-Check Application v{ver}'."),
    ("Dropdown population",
     "Open the task selector dropdown.",
     "Five options present: X-Checks, Grouping By, Accounting Principles, Conditions, Full Run — in that order."),
    ("Start button disabled by default",
     "Open the app without selecting a task.",
     "Start button is disabled (greyed out)."),
    ("Start button enables on selection",
     "Select any task from the dropdown.",
     "Start button becomes enabled immediately."),
    ("Close task selector",
     "Click the red X on the task selector.",
     "Application exits cleanly — no error dialogs, no orphaned processes."),
    ("Upload dialog title",
     "Select each strategy in turn and click Start.",
     "Upload dialog title matches the strategy: 'X-Check Files', 'Grouping By Files', 'Accounting Principles Files', 'Conditions Files', 'Full Run — All Strategies'."),
    ("Advisory sheet-name label",
     "Open any upload dialog that has sheet name fields.",
     "Centred black label reads 'If the sheet name changes from the current default in the Excel workbook, update it manually.'"),
    ("Proceed disabled until all files provided",
     "Open upload dialog; fill in only some required fields.",
     "Proceed button stays disabled until all required fields and output directory are provided."),
    ("Cancel upload dialog returns to selector",
     "Open upload dialog and close it with the red X.",
     "Task selector reappears; Start button returns to disabled state."),
    ("Progress dialog opens on Proceed",
     "Complete all required fields and click Proceed.",
     "Progress dialog opens with scrollable log; Stop button visible."),
    ("Progress dialog — error style",
     "Trigger a file-not-found error (rename a file after selecting it).",
     "Error line appears in bold red text; Windows error chime plays; Exit Application button appears."),
    ("Return to Form after error",
     "After an error, click 'Return to Form'.",
     "Upload dialog re-opens with all previously entered paths pre-filled."),
    ("Stop button halts processing",
     "Start a run and click Stop.",
     "Processing halts cleanly; log shows 'Processing halted by user.'; Return to Form button appears."),
]


# ---------------------------------------------------------------------------
# Sheet 2 — Field interaction cases  (Test, Action, Expected)
# ---------------------------------------------------------------------------
FIELD_CASES = [
    ("Required field asterisk label",
     "Open any upload dialog.",
     "Required fields show label text ending in ' *'; output directory shows 'Output Directory *'."),
    ("Browse button opens file picker",
     "Click Browse next to any file field.",
     "OS file picker dialog opens filtered to the correct file type (e.g. Excel Files *.xlsx)."),
    ("File selected — path label updates",
     "Select a file via Browse.",
     "Path label updates to the filename (not the full path); text turns black."),
    ("Sheet name field enables on file select",
     "Select a file for a field that has a sheet name input.",
     "Sheet name entry becomes editable; text turns black."),
    ("Sheet name retains default until edited",
     "Select a file without changing the sheet name.",
     "Sheet name field shows the configured default (e.g. 'cross checks all', 'Sheet1')."),
    ("Output directory Browse",
     "Click Browse next to Output Directory.",
     "Folder picker opens; selected path shows in full in the label."),
    ("Process only differences — default ON",
     "Open any upload dialog.",
     "'Process only differences' checkbox is ticked by default."),
    ("Dialog position",
     "Move the task selector to a corner of the screen, then open the upload dialog.",
     "Upload dialog appears at the same position as the task selector, clamped within the usable screen area."),
]


# ---------------------------------------------------------------------------
# Sheet 2 — Workflow cases  (Workflow, Test, Action, Expected)
# ---------------------------------------------------------------------------
WORKFLOW_CASES = [
    # X-Checks
    ("X-Checks", "End-to-end run completes",
     "Run debug EXE with test data.",
     "Output workbook produced in ~/Downloads/Output; 2 sheets present; no error lines in log."),
    ("X-Checks", "X-Checks Comparison sheet present",
     "Open output workbook.",
     "Sheet named 'X-Checks Comparison' present with rows of comparison data."),
    ("X-Checks", "Result column colour-coded",
     "Open X-Checks Comparison sheet.",
     "Result column contains formatted values (Matched / MisMatch / Not Found); conditional formatting applied."),
    ("X-Checks", "Optional files handled gracefully",
     "Run with only required files (no GCoA or Known Exception List).",
     "Run completes without error; log notes optional files were skipped."),
    ("X-Checks", "Experimental checkboxes visible",
     "Open X-Checks upload dialog.",
     "Two experimental checkboxes present: 'Apply Version Spanning Validation' and 'Apply Prior Year Balance Formula', both unchecked by default."),

    # Grouping By
    ("Grouping By", "End-to-end run completes",
     "Run debug EXE with test data.",
     "Output workbook produced in ~/Downloads/Output; 7 sheets present; no error lines in log."),
    ("Grouping By", "Compare sheet contains results",
     "Open output workbook → Compare sheet.",
     "Result column contains 'Matched' and/or 'Not in FIP' values; row count > 0."),
    ("Grouping By", "FIP - Processed has Key column",
     "Open FIP - Processed sheet.",
     "Column 'Key' present; values follow 'ValidRule|EBX Item' format."),
    ("Grouping By", "EBX - Processed stacks Grouping By values",
     "Open EBX - Processed sheet.",
     "Each Grouping By value appears as a separate row with a Key column."),
    ("Grouping By", "Mapping File sheet matches input",
     "Open Mapping File sheet.",
     "Two columns: FIP Data and EBX item; row count matches input mapping file."),

    # Accounting Principles
    ("Accounting Principles", "End-to-end run completes",
     "Run debug EXE with test data.",
     "Output workbook produced in ~/Downloads/Output; 4 sheets present; no error lines in log."),
    ("Accounting Principles", "Comparison sheet formatting",
     "Open Comparison sheet.",
     "Match rows have green fill + green text; MisMatch rows have red fill + red text."),
    ("Accounting Principles", "Method bindings extracted",
     "Check progress log.",
     "Log line 'Validation Methods — Method bindings extracted N' with N > 0."),
    ("Accounting Principles", "In-scope X-Checks logged",
     "Check progress log.",
     "Log line 'EBX — In-scope X-Check Nos N' with N > 0."),

    # Conditions
    ("Conditions", "End-to-end run completes",
     "Run debug EXE with test data.",
     "Output workbook produced in ~/Downloads/Output; 4 sheets present; no error lines in log."),
    ("Conditions", "Summary block present",
     "Open Conditions sheet.",
     "Top rows show summary: Total Pairs, Matched (TRUE), Not Matched (FALSE) before data rows."),
    ("Conditions", "Conditions sheet has TRUE and FALSE rows",
     "Inspect Comparison column in Conditions sheet.",
     "Column contains TRUE and/or FALSE values; EBX Data column contains XCheck|Condition format keys."),
    ("Conditions", "Working Sheet has concat columns",
     "Open Working Sheet.",
     "Columns present for each of the 5 condition types plus corresponding '(Concat)' columns."),

    # Full Run
    ("Full Run", "End-to-end run completes",
     "Run debug EXE with test data.",
     "Single combined output workbook produced in ~/Downloads/Output; 13 sheets present; no error lines in log."),
    ("Full Run", "All four strategies execute in order",
     "Check progress log.",
     "Log shows '— Starting: X-Checks —', then '— Starting: Grouping By —', then '— Starting: Accounting Principles —', then '— Starting: Conditions —' in sequence."),
    ("Full Run", "Sheet tab colours",
     "Open combined workbook and inspect tab colours.",
     "XC tabs = green; GB tabs = orange; AP tabs = dark blue; Cond tabs = Zurich blue; Processing Log = grey."),
    ("Full Run", "Tab prefixes",
     "Inspect sheet tab names.",
     "X-Checks prefixed 'XC — '; Grouping By prefixed 'GB — '; Accounting Principles prefixed 'AP — '; Conditions prefixed 'Cond — '."),
    ("Full Run", "Combined workbook sheet count",
     "Count sheets in combined workbook.",
     "15 sheets: 1 (XC) + 6 (GB) + 3 (AP) + 3 (Cond) + 1 (Processing Log) = 14 data sheets + 1 log."),
    ("Full Run", "Processing Log spans all strategies",
     "Open Processing Log sheet.",
     "Log contains entries from all four strategies in sequence — no gaps or missing strategy sections."),

    # Sensitivity label
    ("All strategies", "Sensitivity label applied",
     "Open any output workbook in Excel, check File → Info → Protect Workbook.",
     "Microsoft Information Protection label 'Internal Use Only' is applied to the workbook."),

    # Version
    ("All strategies", "Version stamp in log",
     "Open Processing Log sheet of any output workbook.",
     "First row of log reads 'System — X-Check Application v{ver}'."),
]


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
def _v(text, version):
    return text.replace("{ver}", version) if text else text


def write_section_title(ws, row, text):
    c = ws.cell(row=row, column=1, value=text)
    c.font = ZS_MEDIUM
    c.alignment = WRAP_CENTER
    ws.row_dimensions[row].height = 25.5


def write_section_header(ws, row, headers):
    for i, h in enumerate(headers, start=1):
        c = ws.cell(row=row, column=i, value=h)
        c.font = ZS_HEADER
        c.alignment = WRAP_CENTER
        c.border = TOP_BOT_BORDER


def write_section_row(ws, row, values, n_cols, height=51):
    for i in range(1, n_cols + 1):
        v = values[i - 1] if i - 1 < len(values) else None
        c = ws.cell(row=row, column=i, value=v)
        c.font = ZS_LIGHT
        c.alignment = WRAP_CENTER
        c.border = BOT_BORDER
    ws.row_dimensions[row].height = height


# ---------------------------------------------------------------------------
# Sheet 1 — Instructions
# ---------------------------------------------------------------------------
def write_instructions(ws, version):
    ws.title = "Instructions"
    ws.sheet_view.showGridLines = False

    ws["A1"] = f"Test plan for {APP_NAME} Application v{version}"
    ws["A1"].font = NARRATIVE_FONT
    ws["A2"] = "Implemented features:"
    ws["A2"].font = NARRATIVE_FONT

    for i, feat in enumerate(FEATURES, start=3):
        c = ws.cell(row=i, column=2, value=feat)
        c.font = NARRATIVE_FONT
        c.alignment = WRAP_TOP

    header_row = len(FEATURES) + 4
    ws.cell(row=header_row - 1, column=1, value="Test Plan").font = NARRATIVE_BOLD

    headers = ["Test", "Action", "Expected result", "Tested By", "Tested Date", "Result"]
    for i, h in enumerate(headers, start=1):
        ws.cell(row=header_row, column=i, value=h).font = NARRATIVE_BOLD

    for idx, (test, action, expected) in enumerate(PLAN_ROWS):
        r = header_row + 1 + idx
        if test:
            ws.cell(row=r, column=1, value=test).font = NARRATIVE_FONT
        if action:
            ws.cell(row=r, column=2, value=_v(action, version)).font = NARRATIVE_FONT
        if expected:
            ws.cell(row=r, column=3, value=_v(expected, version)).font = NARRATIVE_FONT
        for col in range(1, 7):
            ws.cell(row=r, column=col).alignment = WRAP_TOP

    for col, w in {"A": 45.7, "B": 45.7, "C": 60.0, "D": 9.7, "E": 11.9, "F": 6.9}.items():
        ws.column_dimensions[col].width = w


# ---------------------------------------------------------------------------
# Sheet 2 — UAT Tests
# ---------------------------------------------------------------------------
def write_uat_tests(ws, version):
    ws.title = "UAT Tests"
    ws.sheet_view.showGridLines = False
    r = 1

    # 1. Files Required
    write_section_title(ws, r, "Files Required")
    r += 1
    write_section_header(ws, r, ["Strategy", "Field Label", "Filename", "Sheet (if applicable)", "Required"])
    r += 1
    for vals in FILES_ROWS:
        write_section_row(ws, r, vals, 5, height=15.75)
        r += 1

    # 2. General UI
    r += 1
    write_section_title(ws, r, "General UI and Dialog Behavior Test Cases")
    r += 1
    write_section_header(ws, r, ["Test", "Action", "Expected Result", "Tested By", "Tested Date", "Result"])
    r += 1
    for case in GENERAL_UI_CASES:
        write_section_row(ws, r, [case[0], _v(case[1], version), _v(case[2], version), "", "", ""], 6)
        r += 1

    # 3. Field Interaction
    r += 1
    write_section_title(ws, r, "Detailed Field Interaction Test Cases")
    r += 1
    write_section_header(ws, r, ["Test", "Action", "Expected Result", "Tested By", "Tested Date", "Result"])
    r += 1
    for case in FIELD_CASES:
        write_section_row(ws, r, [case[0], _v(case[1], version), _v(case[2], version), "", "", ""], 6)
        r += 1

    # 4. Workflow-Specific
    r += 1
    write_section_title(ws, r, "Workflow-Specific Test Cases")
    r += 1
    write_section_header(ws, r, ["Workflow", "Test", "Action", "Expected Result", "Tested By", "Tested Date", "Result"])
    r += 1
    for case in WORKFLOW_CASES:
        write_section_row(ws, r, [case[0], case[1], _v(case[2], version), _v(case[3], version), "", "", ""], 7)
        r += 1

    for col, w in {"A": 19.4, "B": 23.7, "C": 60.3, "D": 22.9, "E": 11.1, "F": 36.1, "G": 29.7, "H": 53.4}.items():
        ws.column_dimensions[col].width = w


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
def main():
    wb = Workbook()
    write_instructions(wb.active, __version__)
    write_uat_tests(wb.create_sheet("UAT Tests"), __version__)

    out_dir = os.path.dirname(os.path.abspath(__file__))
    today = date.today().strftime("%Y%m%d")
    out_path = os.path.join(out_dir, f"{today} {APP_NAME}_v{__version__} Test Plan.xlsx")
    wb.save(out_path)
    print(f"Wrote: {out_path}")


if __name__ == "__main__":
    main()
