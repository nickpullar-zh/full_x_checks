"""
Generate the X-Checks UAT workbook in the same format as
`20260410 X-Checks_v0.1 Test Plan.xlsx`.

Run: python docs/generate_uat.py
Output: docs/<YYYYMMDD> X-Checks_v<version> Test Plan.xlsx
"""
import os
import sys
from datetime import date
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, Side
from openpyxl.utils import get_column_letter

from version import __version__

# Matches the reference template
NARRATIVE_FONT = Font(name="Aptos Narrow", size=11)
NARRATIVE_BOLD = Font(name="Aptos Narrow", size=11, bold=True)

ZS_LIGHT  = Font(name="Zurich Sans Light",  size=10, color="FF4D4D4D")
ZS_MEDIUM = Font(name="Zurich Sans Medium", size=20, color="FF2167AE")
ZS_HEADER = Font(name="Zurich Sans Light",  size=10, bold=True, color="FF4D4D4D")

MEDIUM = Side(style="medium", color="FF000000")
TOP_BORDER    = Border(top=MEDIUM)
BOT_BORDER    = Border(bottom=MEDIUM)
TOP_BOT_BORDER = Border(top=MEDIUM, bottom=MEDIUM)

WRAP_TOP    = Alignment(vertical="top",    wrap_text=True)
WRAP_CENTER = Alignment(vertical="center", wrap_text=True)


# -----------------------------------------------------------------------------
# Sheet1 — Test Plan (narrative walkthrough, plain Aptos Narrow)
# -----------------------------------------------------------------------------

# Sequential walkthrough rows. Empty 'test' means continuation row of the previous test
# (Action/Expected may be blank, allowing C-only continuation as in the reference).
PLAN_ROWS = [
    # (Test, Action, Expected)
    ("Application starts",
     "Double click on X-Checks_v{ver}.exe",
     "Splash dialog shows with version v{ver}, then Application Selector dialog opens"),

    ("Application exits from Application Selector",
     "Click X on top right of Application Selector",
     "Application closes, all dialog boxes close"),

    ("User cannot start process without selection in dropdown",
     'With no selection in the dropdown\nClick "Start"',
     "Nothing happens"),

    ("X-Checks Process Starts",
     'Choose "X-Checks" in the dropdown\nClick "Start"',
     '"X-Check Files" dialog appears'),

    ('X-Checks process "FIP file" upload',
     'Click on "Browse" button in "FIP file" row',
     '"Select FIP file" file picker opens'),

    ("",
     'Select the "20251205 FIP X-Checks - Original.txt" file\nClick "Open"',
     '"20251205 FIP X-Checks - Original.txt" appears in the "FIP file" row'),

    ("", "", '"Proceed" button is inactive'),

    ('X-Checks process "X-Checks Publication file" upload',
     'Click on "Browse" button in "X-Checks Publication file" row',
     '"X-Check Files" file picker opens'),

    ("",
     'Select the "20251205 EPM X-Checks - Original.xlsx" file\nClick "Open"',
     '"20251205 EPM X-Checks - Original.xlsx" appears in the "X-Checks Publication file" row'),

    ("", "", 'Sheet name dialog box becomes active and is populated with "cross checks all"'),
    ("", "", '"Proceed" button is inactive'),

    ('X-Checks process "Known Exception List" upload (optional)',
     'Click on "Browse" button in "Known Exception List" row',
     '"Known Exception List" file picker opens'),

    ("",
     'Select the "Known_Exception_List.xlsx" file\nClick "Open"',
     '"Known_Exception_List.xlsx" appears in the "Known Exception List" row'),

    ('X-Checks process "Output Directory"',
     'Click on "Browse" button in "Output Directory" row',
     '"Select Output Directory" folder picker appears'),

    ("",
     'Click "Select Folder"',
     'The path that was active in the Folder picker should appear in the "Output Directory" row'),

    ("", "", '"Proceed" button is active'),

    ("Run the X-Checks process with optional Known Exception file",
     'Click "Proceed"',
     "X-Check Files dialog disappears"),

    ("", "",
     'Progress dialog "X-Checks - Processing" opens.\n'
     'FIP file:\n  File   : 20251205 FIP X-Checks - Original.txt\n  Content: loaded OK (text file)\n\n'
     'EBX (X-Checks Publication) file:\n  File   : 20251205 EPM X-Checks - Original.xlsx\n  Sheet  : cross checks all\n\n'
     'Known Exception List:\n  File   : Known_Exception_List.xlsx\n  Sheet  : Known Exceptions\n\n'
     'Run completes; "Close" button shown.'),

    ("Verify output workbook is created",
     'Open "test_data\\X-Checks Output\\<timestamp>_X-Checks Comparison.xlsx"',
     '"All Data" sheet contains 13 columns: X-Check Number, Formula Match, EBX Formula, '
     'FIP Formula, Formula Match (Excl), EBX Formula (Excl), FIP Formula (Excl), '
     'Variables Match, EBX Variables, FIP Variables, Variables Match (Builder), '
     'FIP Variable (Builder), Known Exception. Rows sorted by X-Check Number.'),

    ("Close the progress dialog",
     'Click "Close"',
     "The dialog box closes. The process is finished and the app is closed."),

    ("Cancel during processing returns to form",
     'Re-launch the app, choose "X-Checks", select files, click Proceed, then click Stop / X '
     "before completion.",
     'Progress dialog button changes to "Return to Form". Clicking returns to the file '
     "selection form with previously selected files pre-filled. App does NOT exit."),

    ("Error during processing returns to form",
     "Re-launch, select an invalid file (e.g. a non-Excel file as EBX), click Proceed.",
     'Run errors out with a clear message in the progress log. Button shows "Return to Form". '
     "Clicking returns to the file selection form. App does NOT exit. (v0.3.13 fix.)"),

    ("Re-run with second test pair (regression)",
     "Restart, select EBX = 20260313 Cross Checks All.xlsx and FIP = 20260318 FIP X-Checks.txt. "
     "Run.",
     "New comparison file produced. Output structure identical to first pair. No crashes."),

    # =========================================================================
    # v0.4.x — Collect Live X-Checks task
    # =========================================================================
    ("Collect Live X-Checks Process Starts",
     'Re-launch the app, choose "Collect Live X-Checks" in the dropdown\nClick "Start"',
     '"Collect Live X-Checks" dialog appears with only two fields: '
     'X-Checks Publication File * and Output Directory *.'),

    ("",
     'Verify "Collect Live X-Checks" is the FIRST entry in the dropdown',
     'Application Selector dropdown order: Collect Live X-Checks, X-Checks, '
     'X-Checks Grouping By.'),

    ('Collect Live process "X-Checks Publication file" upload',
     'Click on "Browse" button in "X-Checks Publication file" row.\n'
     'Select "20260313 Cross Checks All.xlsx"\nClick "Open"',
     '"20260313 Cross Checks All.xlsx" appears in the row. Sheet name '
     'becomes active and is populated with "cross checks all".'),

    ('Collect Live process "Output Directory"',
     'Click "Browse" next to Output Directory; pick a folder; click "Select Folder"',
     'Folder path appears. "Proceed" button is now active.'),

    ("Run Collect Live X-Checks",
     'Click "Proceed"',
     "Dialog disappears. Progress dialog opens."),

    ("", "",
     'Log shows:\n'
     '  [System] X-Check Application v{ver}\n'
     '  [System] Loading files into memory...\n'
     '  [System] Files loaded successfully\n'
     '  [System] Starting Collect Live X-Checks\n'
     '  [Collect] Wrote <timestamp>_X-Check_Nos.txt\n'
     '  [Collect] Copied to clipboard'),

    ("Verify .txt output file",
     'Open the chosen Output Directory and locate '
     '<timestamp>_X-Check_Nos.txt. Open it in Notepad.',
     'File contains one X-Check Number per line, in order of first appearance, '
     'with no duplicates. Rows where Status = INACTIVE, Type of Change is blank, '
     'Exclude Z-Core = X, or Category cell is yellow are NOT in the file.'),

    ("Verify clipboard contents",
     "Open Excel/Notepad. Press Ctrl+V to paste.",
     "Same X-Check Numbers as the .txt, one per line. Identical content."),

    ("Close progress dialog",
     'Click "Close"',
     "App exits cleanly."),

    # =========================================================================
    # v0.4.3–v0.4.5 — Error visibility + exit options
    # =========================================================================
    ("Error styling: bold red + chime",
     'Re-launch. Choose Collect Live X-Checks. Pick an EBX file that is currently OPEN '
     'in Excel (so pandas cannot read it). Click Proceed.',
     'Progress log line "[System] Error loading files: ..." appears in BOLD RED. '
     'A Windows error chime ("bonk") plays at the same time.'),

    ("Exit Application button",
     'After the error above, two buttons are visible at the bottom of the progress '
     'dialog: "Return to Form" and "Exit Application". Click "Exit Application".',
     "App exits cleanly. No file selection form is shown."),

    ("Exit Application button visible during a healthy run",
     "Re-launch and start any task. Observe the progress dialog while it runs.",
     '"Exit Application" button is visible alongside the Stop button at all times.'),
]


def write_test_plan(ws, version):
    ws.title = "Instructions"
    ws.sheet_view.showGridLines = False

    # Lead-in (top of sheet)
    ws["A1"] = "Test plan for X-Checks Application"
    ws["A1"].font = NARRATIVE_FONT
    ws["A2"] = "Implemented features:"
    ws["A2"].font = NARRATIVE_FONT
    ws["B3"] = "FIP file + EBX file uploads with Known Exception (optional) and GCoA (optional)"
    ws["B3"].font = NARRATIVE_FONT
    ws["B4"] = "Validate Files"
    ws["B4"].font = NARRATIVE_FONT
    ws["B5"] = ("Collect Live X-Checks task — produces .txt of in-scope X-Check Nos "
                "and copies to clipboard")
    ws["B5"].font = NARRATIVE_FONT

    # Test Plan section
    ws["A6"] = "Test Plan"
    ws["A6"].font = NARRATIVE_BOLD

    # Header row 7
    headers = ["Test", "Action", "Expected result", "Tested By", "Tested Date", "Result"]
    for i, h in enumerate(headers, start=1):
        c = ws.cell(row=7, column=i, value=h)
        c.font = NARRATIVE_BOLD

    # Body
    for idx, (test, action, expected) in enumerate(PLAN_ROWS):
        r = 8 + idx
        if test:
            ws.cell(row=r, column=1, value=test).font = NARRATIVE_FONT
        if action:
            ws.cell(row=r, column=2, value=action.replace("{ver}", version)).font = NARRATIVE_FONT
        if expected:
            ws.cell(row=r, column=3, value=expected.replace("{ver}", version)).font = NARRATIVE_FONT
        for col in range(1, 7):
            ws.cell(row=r, column=col).alignment = WRAP_TOP

    # Column widths from the reference
    widths = {"A": 45.7, "B": 45.7, "C": 60.0, "D": 9.7, "E": 11.9, "F": 6.9}
    for col, w in widths.items():
        ws.column_dimensions[col].width = w


# -----------------------------------------------------------------------------
# Sheet2 — sections with Zurich-brand styling
# -----------------------------------------------------------------------------

# Files Required table (CaseUI / Field Label / Filename / Sheet / Required)
FILES_ROWS = [
    ("X-Checks", "FIP file",
     "20251205 FIP X-Checks - Original.txt",
     "N/A", "Yes"),
    ("X-Checks", "X-Checks Publication File (EBX)",
     "20251205 EPM X-Checks - Original.xlsx",
     "cross checks all", "Yes"),
    ("X-Checks", "GCoA Publication File",
     "GCoA file with 13106 Data rows on sheet GCoA Base account table.xlsx",
     "GCoA Base account table", "No"),
    ("X-Checks", "Known Exception List",
     "Known_Exception_List.xlsx",
     "Known Exceptions", "No"),
    ("Collect Live X-Checks", "X-Checks Publication File (EBX)",
     "20260313 Cross Checks All.xlsx",
     "cross checks all", "Yes"),
]

# General UI and Dialog Behavior test cases (Test/Action/Expected/Tested By/Date/Result)
GENERAL_UI_CASES = [
    ("Splash + version on launch",
     "Double-click X-Checks_v{ver}.exe.",
     "Splash dialog appears with text 'X-Check Application v{ver} Loading...'. "
     "Splash auto-dismisses after the Application Selector opens."),

    ("Application Selector title",
     "Inspect the Application Selector window after splash closes.",
     "Window title and version label both display v{ver} (matches splash)."),

    ("Dialog opens with correct title and fields",
     "Pick 'X-Checks' from the dropdown, click Start.",
     "Dialog window displays the correct title and task name. All fields per config "
     "appear in the correct order: FIP file *, X-Checks Publication File *, "
     "GCoA Publication File (optional), Known Exception List (optional), Output Directory *."),

    ("Dialog is modal",
     "Attempt to interact with the parent window while the dialog is open.",
     "Parent window is unresponsive; dialog maintains focus lock (modal behavior)."),

    ("Dialog is not resizable",
     "Attempt to resize dialog window by dragging edges.",
     "Dialog window cannot be resized."),

    ("Dialog closes via X button",
     "Click the red X on the dialog.",
     "Dialog closes cleanly. Application returns to the Application Selector."),

    ("Proceed button initial state",
     "Open the X-Checks dialog and observe the Proceed button.",
     "Proceed button is visible and disabled until all required fields are filled."),

    # ---- v0.4.x additions ------------------------------------------------------
    ("Dropdown order — Collect Live X-Checks first",
     "Open the Application Selector dropdown.",
     "Dropdown lists tasks in this order: Collect Live X-Checks, X-Checks, "
     "X-Checks Grouping By."),

    ("Progress dialog has Exit Application button",
     "Run any task. While the progress dialog is open, look at the bottom button bar.",
     'Two buttons are visible: the existing Stop / Close / Return to Form button '
     'AND a permanent "Exit Application" button to its right.'),

    ("Exit Application button shuts down everything",
     'Click "Exit Application" while a run is in progress (or after it errors).',
     "App exits immediately. No further dialogs appear. The Application Selector "
     "is also closed."),

    ("Error log lines styled bold red",
     "Trigger an error (e.g. supply an EBX file that is open in Excel). Observe the "
     "progress dialog log.",
     "Lines containing 'error', 'failed', 'failure', 'exception' or 'traceback' "
     "(case-insensitive) appear in BOLD RED. Non-error lines remain in the default "
     "dark-blue Courier."),

    ("Error chime plays on error log line",
     "Trigger any error (audio on).",
     "Standard Windows critical-stop chime plays at the moment the red line is "
     "added to the log."),
]

# Detailed Field Interaction
FIELD_CASES = [
    ("Required file field label rendering",
     "Open dialog with X-Checks task selected.",
     "Required file labels render as '<Label> *'."),
    ("Optional file field label rendering",
     "Inspect GCoA / Known Exception field labels.",
     "Optional labels render as '<Label> (optional)'."),
    ("File picker opens and filters types",
     "Click Browse next to a file field.",
     "OS file picker opens with title matching field label, filtered by file_types."),
    ("File selection updates path label",
     "Select a file and confirm.",
     "Path label updates to the filename (not full path), in black font. The internal "
     "variable stores the full path."),
    ("Cancel file picker leaves field empty",
     "Open file picker and cancel.",
     "Path label remains 'No file selected' in grey. Internal variable is empty."),
    ("Optional file field left empty",
     "Leave optional field empty, fill all required fields.",
     "Proceed button becomes enabled. On submit, optional file value is None."),
    ("Sheet name entry default",
     "Select EBX file with show_sheet=True.",
     "Sheet name entry becomes enabled and shows default 'cross checks all'."),
    ("Output directory selection and label update",
     "Click Browse for output directory, select a folder.",
     "Label updates to the full path, in black font. Internal variable set."),
    ("Proceed enables when required fields filled",
     "Fill all required file fields and output directory.",
     "Proceed button becomes enabled."),
    ("Form pre-fill on return after error / cancel",
     "Cancel mid-run, then 'Return to Form'.",
     "All previously selected file paths and the sheet name are restored. Tester does "
     "not need to re-pick anything."),
]

# Workflow-specific test cases (Workflow/Test/Action/Expected/Tested By/Date/Result)
WORKFLOW_CASES = [
    ("X-Checks", "All required fields only",
     "Upload FIP and EBX files only. Leave optional GCoA and Known Exception empty. "
     "Select output directory.",
     "Proceed enables. Run completes. Output workbook written. 'Known Exception' column "
     "in output is empty for all rows."),

    ("X-Checks", "With Known Exception List",
     "Upload FIP, EBX and Known Exception. Run.",
     "Output 'Known Exception' column populated for any X-Check whose number appears in "
     "the list. Match columns for those rows show 'Mismatch - Known Exception' "
     "(blue-highlighted) instead of red 'MisMatch'."),

    ("X-Checks", "With GCoA file (QU substitution)",
     "Upload FIP, EBX and GCoA. Run. Find an X-Check whose account is Data type = QU in GCoA.",
     "EBX Formula uses QU_YTD(...) instead of VAL_YTD(...) for that account."),

    ("X-Checks", "Shareholders' Equity → LC_YTD",
     "Pick an X-Check where Category = \"Shareholders' Equity\" (e.g. S380_00).",
     "EBX Formula uses LC_YTD and CONST_LC instead of VAL_YTD/CONST. Formula Match = Match."),

    ("X-Checks", "Percentage limit format",
     "Pick an X-Check with the '%' column flagged X (e.g. S002_00).",
     "EBX right-hand side is a quoted percent literal like \"'1,500000%'\". Formula Match = Match."),

    ("X-Checks (Excl)", "FIP plain name + @2A@ Account Type",
     "Pick A159_09 (variable name has no 'excl' token but FIP source has @2A@ Account Type 2).",
     "FIP Formula (Excl) shows excl.acc.type=2 inserted before ToM/TOM. EBX Formula (Excl) "
     "matches: ABS(VAL_YTD(OAN_00277ffexcl.acc.type=2ToM660ff))<=CONST(5,'USD','E'). "
     "Formula Match (Excl) = Match."),

    ("X-Checks (Excl)", "Multi-type @2A@ rows",
     "Pick an X-Check with two @2A@ Account Type rows on a single variable.",
     "Suffix is excl.acc.type=N,M (sorted numerically, comma-joined)."),

    ("X-Checks (Excl)", "Strip pre-written excl text",
     "Pick an X-Check where the FIP variable name has pre-written excl text (e.g. excl.2-Aff) "
     "without a backing @2A@ row.",
     "FIP Formula (Excl) has the messy text removed. ToM/TOM suffix preserved."),

    ("X-Checks (Excl)", "LA006_09 per-variable scoping",
     "Find LA006_09. Inspect EBX Formula (Excl).",
     "Only the variable whose row had 'Exclude Account Type' set shows excl.acc.type=N. "
     "The other variable is unchanged."),

    ("X-Checks (Excl)", "Literal-then-constructed fallback",
     "Find an X-Check where the literal FIP Formula already equals EBX Formula (Excl) but "
     "the constructed FIP Formula (Excl) differs.",
     "Formula Match (Excl) = Match (via the literal path)."),

    ("X-Checks", "Error returns to form (v0.3.13)",
     "Pick a non-Excel file as EBX. Click Proceed.",
     "Run errors out, button changes to 'Return to Form', clicking it returns to the file "
     "selection form with previous selections pre-filled. App does NOT exit."),

    ("X-Checks", "Pair 2 regression",
     "Run with 20260313 + 20260318 inputs.",
     "Output produced with same structure. No crashes, no missing X-Checks."),

    # ---- Collect Live X-Checks -------------------------------------------------
    ("Collect Live X-Checks", "Smoke run — full pipeline",
     "Pick Collect Live X-Checks. Use 20260313 Cross Checks All.xlsx. Output dir = "
     "anywhere writable. Click Proceed.",
     "Progress log shows '[Collect] Wrote ..._X-Check_Nos.txt' and '[Collect] "
     "Copied to clipboard'. .txt file exists in the output folder. Ctrl+V into "
     "Notepad pastes the same list."),

    ("Collect Live X-Checks", ".txt content matches selection rules",
     "Open the .txt produced above. Spot-check 5 X-Check Nos against the source "
     "Cross Checks All sheet.",
     "Each X-Check in the .txt has at least one row with Status = ACTIVE and a "
     "non-blank Type of Change. None have any row with Exclude Z-Core = X. None "
     "have a yellow (#FFFF00) Category cell."),

    ("Collect Live X-Checks", "INACTIVE rows filtered out",
     "Pick a known X-Check from the source whose only rows are Status = INACTIVE.",
     "That X-Check Number does NOT appear in the output .txt or on the clipboard."),

    ("Collect Live X-Checks", "Exclude Z-Core filter",
     "Pick an X-Check from the source where any active row has Exclude Z-Core = X.",
     "That X-Check Number does NOT appear in the output."),

    ("Collect Live X-Checks", "Yellow Category filter",
     "Pick an X-Check whose Category cell is filled with standard Excel yellow #FFFF00.",
     "That X-Check Number does NOT appear in the output."),

    ("Collect Live X-Checks", "Order is first-occurrence, no duplicates",
     "Inspect the .txt: confirm an X-Check that appears multiple times in the source "
     "appears exactly once, in the order of its first occurrence.",
     "One line per unique X-Check, source-order preserved."),

    ("Collect Live X-Checks", "Case-insensitive column headers",
     "Optional: clone the EBX file and rename headers to mixed case (e.g. 'status', "
     "'Type of change', 'EXCLUDE Z-CORE', 'category'). Run the task on the clone.",
     "Output identical to the original — header lookup is case-insensitive."),

    ("Collect Live X-Checks", "Error visibility — file open in Excel",
     "Open the EBX file in Excel. Run Collect Live X-Checks against it.",
     "Log line '[System] Error loading files: ... [Errno 13] Permission denied: ...' "
     "appears in BOLD RED, error chime plays. No .txt is written and the clipboard "
     "is unchanged."),

    ("Collect Live X-Checks", "Exit cleanly with Exit Application",
     "After a successful run, click Exit Application instead of Close.",
     "App exits immediately. .txt remains on disk."),
]


def write_section_title(ws, row, text, end_col):
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


def write_section_row(ws, row, values, n_cols):
    for i in range(1, n_cols + 1):
        v = values[i - 1] if i - 1 < len(values) else None
        c = ws.cell(row=row, column=i, value=v)
        c.font = ZS_LIGHT
        c.alignment = WRAP_CENTER
        c.border = BOT_BORDER


def write_sheet2(ws, version):
    ws.title = "UAT Tests"
    ws.sheet_view.showGridLines = False

    # Files Required
    write_section_title(ws, 1, "Files Required", end_col=5)
    write_section_header(ws, 2, ["CaseUI", "Field Label", "Filename", "Sheet (if applicable)", "Required"])
    r = 3
    for vals in FILES_ROWS:
        write_section_row(ws, r, vals, 5)
        ws.row_dimensions[r].height = 15.75
        r += 1

    # General UI section (3 blank rows separator → next section header at r+3)
    title_row = r + 2
    header_row = title_row + 1
    write_section_title(ws, title_row, "General UI and Dialog Behavior Test Cases", end_col=6)
    write_section_header(ws, header_row, ["Test", "Action", "Expected Result", "Tested By", "Tested Date", "Result"])
    r = header_row + 1
    for case in GENERAL_UI_CASES:
        vals = [case[0], case[1].replace("{ver}", version), case[2].replace("{ver}", version), "", "", ""]
        write_section_row(ws, r, vals, 6)
        ws.row_dimensions[r].height = 51
        r += 1

    # Detailed Field Interaction
    title_row = r + 1
    header_row = title_row + 1
    write_section_title(ws, title_row, "Detailed Field Interaction Test Cases", end_col=6)
    write_section_header(ws, header_row, ["Test", "Action", "Expected Result", "Tested By", "Tested Date", "Result"])
    r = header_row + 1
    for case in FIELD_CASES:
        vals = [case[0], case[1], case[2], "", "", ""]
        write_section_row(ws, r, vals, 6)
        ws.row_dimensions[r].height = 51
        r += 1

    # Workflow-Specific
    title_row = r + 1
    header_row = title_row + 1
    write_section_title(ws, title_row, "Workflow-Specific Test Cases", end_col=7)
    write_section_header(ws, header_row, ["Workflow", "Test", "Action", "Expected Result", "Tested By", "Tested Date", "Result"])
    r = header_row + 1
    for case in WORKFLOW_CASES:
        vals = [case[0], case[1], case[2], case[3], "", "", ""]
        write_section_row(ws, r, vals, 7)
        ws.row_dimensions[r].height = 51
        r += 1

    # Column widths from the reference
    widths = {"A": 19.4, "B": 23.7, "C": 60.3, "D": 22.9, "E": 11.1, "F": 36.1, "G": 29.7, "H": 53.4}
    for col, w in widths.items():
        ws.column_dimensions[col].width = w


def main():
    wb = Workbook()
    write_test_plan(wb.active, __version__)
    write_sheet2(wb.create_sheet("UAT Tests"), __version__)

    out_dir = os.path.dirname(os.path.abspath(__file__))
    today = date.today().strftime("%Y%m%d")
    out_path = os.path.join(out_dir, f"{today} X-Checks_v{__version__} Test Plan.xlsx")
    wb.save(out_path)
    print(f"Wrote: {out_path}")


if __name__ == "__main__":
    main()
