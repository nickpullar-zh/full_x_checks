"""
Generate the Accounting Principles UAT workbook for the
v0.5-Accounting-Principles branch in the standard reference format.

Run: python docs/generate_uat.py
Output: docs/<YYYYMMDD> X-Checks_AccountingPrinciples_v<version> Test Plan.xlsx
"""
import os
import sys
from datetime import date
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, Side

from version import __version__

APP_NAME = "X-Checks_AccountingPrinciples"

# ---- Styles (mirroring the reference workbook) ---------------------------
NARRATIVE_FONT = Font(name="Aptos Narrow", size=11)
NARRATIVE_BOLD = Font(name="Aptos Narrow", size=11, bold=True)

ZS_LIGHT  = Font(name="Zurich Sans Light",  size=10, color="FF4D4D4D")
ZS_MEDIUM = Font(name="Zurich Sans Medium", size=20, color="FF2167AE")
ZS_HEADER = Font(name="Zurich Sans Light",  size=10, bold=True, color="FF4D4D4D")

MEDIUM = Side(style="medium", color="FF000000")
TOP_BOT_BORDER = Border(top=MEDIUM, bottom=MEDIUM)
BOT_BORDER     = Border(bottom=MEDIUM)

WRAP_TOP    = Alignment(vertical="top",    wrap_text=True)
WRAP_CENTER = Alignment(vertical="center", wrap_text=True)


# =========================================================================
# Sheet 1 — Instructions (linear narrative walkthrough)
# =========================================================================
# Each tuple = (Test, Action, Expected). Empty Test/Action = continuation row.
PLAN_ROWS = [
    # --- Launch & version ---
    ("Application starts",
     "Double click on X-Checks_AccountingPrinciples_v{ver}.exe",
     "Splash dialog shows with version v{ver}, then Application Selector dialog opens"),

    ("Application Selector shows correct task",
     "Read the dropdown in the Application Selector",
     '"Accounting Principles" is the only task in the dropdown'),

    ("Application exits from Application Selector",
     "Click X on top right of Application Selector",
     "Application closes, all dialog boxes close"),

    ("User cannot start process without selection in dropdown",
     'With no selection in the dropdown, click "Start"',
     "Nothing happens"),

    # --- Open the AP form ---
    ("Accounting Principles process starts",
     'Choose "Accounting Principles" in the dropdown\nClick "Start"',
     '"Accounting Principles Files" dialog appears'),

    ("",
     "Inspect the dialog",
     "Three file fields: 'Validation Methods File *', 'X-Checks Publication File *', "
     "'FIP File (VALMSG) *', plus 'Output Directory *' and 'Process only differences' "
     "checkbox (default ON)"),

    ("",
     "Look at the field label 'X-Checks Publication File'",
     "The * is on the SAME LINE as the label — not wrapped to row 2 (v0.5.8 fix)"),

    # --- Validation Methods file ---
    ('Validation Methods File upload',
     'Click "Browse" next to "Validation Methods File"',
     '"Select Validation Methods File" file picker opens'),

    ("",
     'Select "validation methods.xlsx", click Open',
     'Filename appears next to the field; sheet name "Validation Methods" appears'),

    ("", "", '"Proceed" button remains inactive (other required fields empty)'),

    # --- X-Checks Publication File (EBX) ---
    ('X-Checks Publication File upload',
     'Click "Browse" next to "X-Checks Publication File"',
     '"Select X-Checks Publication File" file picker opens'),

    ("",
     'Select "20260313 Cross Checks All.xlsx", click Open',
     'Filename appears next to the field; sheet name "cross checks all" appears'),

    # --- FIP File ---
    ('FIP File (VALMSG) upload',
     'Click "Browse" next to "FIP File (VALMSG)"',
     '"Select FIP File (VALMSG)" file picker opens'),

    ("",
     'Select "20260602 VALMSG (Accounting Principle).xlsx", click Open',
     'Filename appears next to the field; sheet name "FIP Methods Rules and Condition" appears'),

    # --- Output directory ---
    ('Output Directory selection',
     'Click "Browse" next to "Output Directory"',
     '"Select Output Directory" folder picker appears'),

    ("",
     'Pick a folder, click "Select Folder"',
     'Path appears next to the field. "Proceed" button is now active'),

    # --- Run with Process only differences ON ---
    ('Run with "Process only differences" ON (default)',
     'Confirm checkbox is ticked; click "Proceed"',
     "Form closes, Processing Log dialog opens"),

    ("", "",
     'Log shows entries:\n'
     '  [System] X-Check Application v{ver}\n'
     '  [Validation Methods] Method bindings extracted\n'
     '  [EBX] In-scope X-Check Nos (filtered)\n'
     '  [Compare] Comparison rows produced\n'
     '  [EBX] Rows kept for output sheet\n'
     '  [FIP] Rows kept for output sheet\n'
     '  [System] Processing complete'),

    ("", "",
     'Bottom of dialog shows TWO buttons: "Close" and "Exit Application"'),

    # --- Verify output file ---
    ("Verify output workbook",
     "Open <timestamp>_Accounting Principles Comparison.xlsx in the chosen output folder",
     "Workbook contains FOUR sheets in this order: EBX, FIP, Comparison, Processing Log"),

    ("", "",
     "EBX sheet: only the in-scope X-Check rows (filtered by Status≠INACTIVE, "
     "non-blank Type of change, Exclude Z-Core≠X, non-yellow Category)"),

    ("", "",
     "FIP sheet: only rows whose V-code (left of the | in 'Key') belongs to the "
     "validation-methods subset"),

    ("", "",
     "Comparison sheet: 7 columns — X-Check No., Event, Expected, FIP, Actual, "
     "Method, Match. Match column has green Match / red MisMatch conditional formatting."),

    ("", "",
     "EXACTLY one row per (X-Check No., Method) — no duplicates "
     "(v0.5.10 black/grey font rule)"),

    # --- Close → selector ---
    ("Clean Close returns to selector",
     'On the Processing Log dialog, click "Close"',
     "Dialog closes; Application Selector reappears (same session). Selector dropdown "
     "is empty so the user can pick a task again. (v0.5.9)"),

    # --- Run with Process only differences OFF ---
    ("Run with checkbox OFF",
     'Re-pick Accounting Principles, fill the same files, untick "Process only '
     'differences", click Proceed',
     'Log line "[EBX] In-scope X-Check Nos" (no "(filtered)" suffix) shows the '
     "full unfiltered count — every unique X-Check No. in the EBX file"),

    # --- Error path ---
    ("Error path — file open in Excel",
     "Re-pick Accounting Principles. Open the EBX file in Excel BEFORE clicking "
     "Proceed. Click Proceed.",
     'Log line "[System] Error loading files: ... Permission denied" appears in BOLD '
     'RED. Windows error chime plays. Button changes to "Return to Form".'),

    ("Return to Form pre-fills",
     'Click "Return to Form"',
     "File upload form reappears with all previously selected files and the "
     '"Process only differences" state pre-filled. (v0.5.9 + earlier)'),

    # --- Exit Application ---
    ("Exit Application during a run",
     "Re-pick AP, fill files, click Proceed. While the run is processing, click "
     '"Exit Application"',
     "Application exits cleanly (no return to selector, no return to form)"),

    # --- Stop ---
    ("Stop button",
     'Run again. Click "Stop" mid-run',
     'Log shows "[---] User requested stop ..." and the button changes to "Return to '
     'Form". Click Return to Form returns to the file upload form pre-filled.'),
]


def write_instructions(ws, version):
    ws.title = "Instructions"
    ws.sheet_view.showGridLines = False

    ws["A1"] = f"Test plan for {APP_NAME} Application"
    ws["A1"].font = NARRATIVE_FONT
    ws["A2"] = "Implemented features:"
    ws["A2"].font = NARRATIVE_FONT
    ws["B3"] = "Accounting Principles strategy — compares cross-checks-all w/e against FIP W/E"
    ws["B3"].font = NARRATIVE_FONT
    ws["B4"] = "Validation Methods parser with rows 4–6 + black/grey font priority"
    ws["B4"].font = NARRATIVE_FONT
    ws["B5"] = "Process only differences filter (Status / Type of change / Exclude Z-Core / yellow Category)"
    ws["B5"].font = NARRATIVE_FONT
    ws["B6"] = "Header row auto-detection (rows 1–6); punctuation-insensitive event matching"
    ws["B6"].font = NARRATIVE_FONT
    ws["B7"] = "Output workbook with EBX, FIP and Comparison sheets + conditional formatting"
    ws["B7"].font = NARRATIVE_FONT
    ws["B8"] = "Error UX: bold-red lines + Windows chime; Exit Application button; Close → selector"
    ws["B8"].font = NARRATIVE_FONT

    ws["A10"] = "Test Plan"
    ws["A10"].font = NARRATIVE_BOLD

    headers = ["Test", "Action", "Expected result", "Tested By", "Tested Date", "Result"]
    for i, h in enumerate(headers, start=1):
        ws.cell(row=11, column=i, value=h).font = NARRATIVE_BOLD

    for idx, (test, action, expected) in enumerate(PLAN_ROWS):
        r = 12 + idx
        if test:
            ws.cell(row=r, column=1, value=test).font = NARRATIVE_FONT
        if action:
            ws.cell(row=r, column=2, value=action.replace("{ver}", version)).font = NARRATIVE_FONT
        if expected:
            ws.cell(row=r, column=3, value=expected.replace("{ver}", version)).font = NARRATIVE_FONT
        for c in range(1, 7):
            ws.cell(row=r, column=c).alignment = WRAP_TOP

    for col, w in {"A": 45.7, "B": 45.7, "C": 60.0, "D": 9.7, "E": 11.9, "F": 6.9}.items():
        ws.column_dimensions[col].width = w


# =========================================================================
# Sheet 2 — UAT Tests (Zurich-branded sectioned tables)
# =========================================================================

# Files Required: (CaseUI, Field Label, Filename, Sheet, Required)
FILES_ROWS = [
    ("Accounting Principles", "Validation Methods File",
     "validation methods.xlsx", "Validation Methods", "Yes"),
    ("Accounting Principles", "X-Checks Publication File",
     "20260313 Cross Checks All.xlsx", "cross checks all", "Yes"),
    ("Accounting Principles", "FIP File (VALMSG)",
     "20260602 VALMSG (Accounting Principle).xlsx", "FIP Methods Rules and Condition", "Yes"),
]

# General UI: (Test, Action, Expected)
GENERAL_UI_CASES = [
    ("Splash + version on launch",
     "Double-click X-Checks_AccountingPrinciples_v{ver}.exe.",
     "Splash dialog appears with text 'X-Check Application v{ver} Loading...' on "
     "the Zurich Blue→Dark Blue gradient with brand fonts. Splash auto-dismisses "
     "after the Application Selector opens."),

    ("Application Selector title",
     "Inspect the Application Selector window after splash closes.",
     "Window title and version label both display v{ver}."),

    ("Dropdown order",
     "Open the Application Selector dropdown.",
     "Only entry is 'Accounting Principles' (this branch ships only one strategy)."),

    ("Dialog opens with correct title and fields",
     "Pick 'Accounting Principles' from the dropdown, click Start.",
     "Dialog window displays 'Accounting Principles Files'. Three file fields visible: "
     "Validation Methods File *, X-Checks Publication File *, FIP File (VALMSG) *. "
     "Plus Output Directory * and 'Process only differences' checkbox (default ON)."),

    ("Field labels render * on same row (v0.5.8)",
     "Inspect the file-field labels.",
     "'X-Checks Publication File *' renders on a single row (no wrap). The * is "
     "always immediately after the label, never on a new line."),

    ("Dialog is modal",
     "Attempt to interact with the Application Selector while the dialog is open.",
     "Selector is unresponsive; dialog maintains focus lock."),

    ("Dialog is not resizable",
     "Attempt to resize the dialog window by dragging its edges.",
     "Dialog window cannot be resized."),

    ("Proceed button initial state",
     "Open the dialog without filling anything.",
     "Proceed button is visible and disabled until all required fields and the "
     "output directory are set."),

    ("Process only differences default",
     "Open the dialog.",
     "'Process only differences' checkbox is TICKED by default."),

    # Error UX
    ("Error log lines styled bold red (v0.4.3)",
     "Trigger any error (e.g. EBX file open in Excel). Observe the Processing Log.",
     "Log lines whose text contains error / failed / failure / exception / "
     "traceback / aborting / aborted / cannot / invalid / missing / not found "
     "are rendered in BOLD RED. Default lines are dark-blue Courier."),

    ("Error chime plays (v0.4.4)",
     "Same trigger, with audio on.",
     "Standard Windows critical-stop chime plays at the moment the red line "
     "is added to the log."),

    ("Exit Application button always visible (v0.4.5)",
     "Run any task. Look at the bottom button bar of the Processing Log.",
     "Two buttons visible: action button (Stop / Close / Return to Form) AND a "
     "permanent 'Exit Application' button."),

    ("Exit Application shuts down everything",
     "Click 'Exit Application' during a run, after success, or after an error.",
     "App exits immediately. No dialog reappears."),

    ("Close returns to selector (v0.5.9)",
     "After a successful run, click 'Close'.",
     "Processing Log dismisses. Application Selector becomes visible again. "
     "Tester can pick another task in the same session."),

    ("Return to Form pre-fills (v0.5.9)",
     "After Stop or any error, click 'Return to Form'.",
     "File upload form reappears with all previously chosen files, sheet names "
     "and 'Process only differences' state pre-filled."),
]

# Detailed Field Interaction: (Test, Action, Expected)
FIELD_CASES = [
    ("Required file field label rendering",
     "Open the Accounting Principles dialog.",
     "All three file labels render as '<Label> *'."),

    ("File picker opens and filters types",
     "Click Browse next to any file field.",
     "OS file picker opens with title matching the field label, filtered to *.xlsx."),

    ("File selection updates path label",
     "Select an .xlsx and confirm.",
     "Path label updates to the filename (not full path), in black font. The "
     "internal variable stores the full path."),

    ("Cancel file picker leaves field empty",
     "Open file picker and cancel.",
     "Path label remains 'No file selected' in grey. Internal variable empty."),

    ("Sheet-name field becomes active after file pick",
     "Pick the validation methods file.",
     "The sheet-name entry next to the field becomes enabled and shows default "
     "'Validation Methods'. Tester can override the sheet name by typing."),

    ("Output directory selection and label update",
     "Click Browse for output directory, select a folder.",
     "Label updates to the full path, in black font. Internal variable set."),

    ("Output directory picker cancel",
     "Click Browse and cancel in directory picker.",
     "Label remains 'No directory selected'; variable empty."),

    ("Proceed enables when all required fields filled",
     "Fill all 3 file fields and the output directory.",
     "Proceed button becomes enabled."),

    ("Header row auto-detection (v0.5.3)",
     "Pick an EBX file whose 'cross checks all' header is on row 2 (e.g. the "
     "VALMSG workbook's bundled sheet). Run.",
     "Strategy detects the correct header row automatically; comparison runs "
     "without 'X-Check No. column not found' errors."),
]

# Workflow-Specific: (Workflow, Test, Action, Expected)
WORKFLOW_CASES = [
    # Smoke / structure
    ("Accounting Principles", "Smoke run with default 27-event subset",
     "Upload the three files (validation methods.xlsx, 20260313 Cross Checks "
     "All.xlsx, 20260602 VALMSG (Accounting Principle).xlsx). Tick 'Process "
     "only differences'. Run.",
     "Run completes. Output workbook produced. Comparison sheet has ~318 rows "
     "(filtered by the 27-event subset). All four output sheets present: EBX, "
     "FIP, Comparison, Processing Log."),

    ("Accounting Principles", "Output sheet structure",
     "Open the produced .xlsx.",
     "EBX = filtered cross-checks-all rows. FIP = filtered VALMSG rows whose "
     "V-code is in the validation methods subset. Comparison = 7-column "
     "summary with conditional formatting on Match (green/red)."),

    ("Accounting Principles", "Comparison columns",
     "Open the Comparison sheet.",
     "Columns in order: X-Check No. | Event | Expected | FIP | Actual | "
     "Method | Match. Sorted by X-Check No., Event, Method."),

    # Match rules
    ("Accounting Principles (Match rules)", "Warning ↔ w",
     "Find a row where Expected=Warning and Actual=w.",
     "Match column reads 'Match'."),

    ("Accounting Principles (Match rules)", "Error ↔ e",
     "Find a row where Expected=Error and Actual=e.",
     "Match column reads 'Match'."),

    ("Accounting Principles (Match rules)", "Both ↔ w (Stammhaus etc.)",
     "Find a row where Expected=Both and Actual=w.",
     "Match column reads 'Match'."),

    ("Accounting Principles (Match rules)", "Both ↔ e",
     "Find a row where Expected=Both and Actual=e.",
     "Match column reads 'Match'."),

    ("Accounting Principles (Match rules)", "Mismatched letters",
     "Find a row where Expected=Warning and Actual=e (or any non-matching combo).",
     "Match column reads 'MisMatch' and renders in red."),

    # Black/grey font priority (v0.5.10)
    ("Accounting Principles (Font priority)", "V900W appears once",
     "Open Comparison; filter Method=V900W.",
     "Every V900W row is attributed to event 'IFRS New RFD'. None to 'IFRS New "
     "SFD' (the SFD copy is grey-font and is suppressed when RFD has data)."),

    ("Accounting Principles (Font priority)", "One row per (X-Check, Method)",
     "Open Comparison; pivot or count by (X-Check No., Method).",
     "Every (X-Check No., Method) pair appears EXACTLY once. No duplicates "
     "regardless of how many events the V-code legitimately serves (v0.5.10)."),

    ("Accounting Principles (Font priority)", "Grey-only fallback",
     "Find a V-code that exists ONLY in grey-font cells (e.g. a copy with no "
     "black-font home). Confirm output uses the leftmost grey column with non-"
     "empty data.",
     "Output row attributes the V-code to the leftmost grey-font event whose "
     "cross-checks-all column has a non-empty actual letter."),

    # Punctuation-insensitive matching (v0.5.4)
    ("Accounting Principles", "DE-GAAP vs DE GAAP column matching",
     "Confirm rows are emitted for events 'DE-GAAP RFD', 'DE-GAAP SFD' even "
     "though cross-checks-all uses 'DE GAAP RFD' (space).",
     "Output Event column shows 'DE-GAAP RFD'/'DE-GAAP SFD' (validation-methods "
     "spelling). Rows are correctly populated via punctuation-insensitive match."),

    # Process only differences filter (v0.5.5)
    ("Accounting Principles", "Filter ON narrows X-Checks",
     "Run with 'Process only differences' ON. Note the [EBX] In-scope count.",
     "Count is smaller than the full file's unique X-Check count. Live data: "
     "189 unique → ~152 in-scope after filter."),

    ("Accounting Principles", "Filter OFF processes all X-Checks",
     "Re-run with checkbox OFF.",
     "[EBX] In-scope count == every unique non-blank X-Check No. in the file."),

    # Output sheet contents
    ("Accounting Principles", "EBX sheet content",
     "Open EBX sheet.",
     "Rows for the in-scope X-Checks only. Columns mirror the source 'cross "
     "checks all' sheet layout (header auto-detected)."),

    ("Accounting Principles", "FIP sheet content",
     "Open FIP sheet.",
     "Only rows whose Key starts with one of the V-codes from the validation "
     "methods subset (e.g. V900W, V900A, V901A ..., V791A ..., V600A ...). "
     "Roughly 5,000 rows for the live data."),

    # Sensitivity label (v0.5.12)
    ("Accounting Principles", "Output file is labelled Internal_Use_Only",
     "Run the strategy. After it completes, open the produced "
     "<timestamp>_Accounting Principles Comparison.xlsx in Excel and inspect "
     "the sensitivity label (File → Info → Sensitivity, or the bar at the top of the workbook).",
     "Sensitivity label reads 'Internal_Use_Only'. The processing log shows "
     "'[Sensitivity] Applied label: Internal_Use_Only' as the final step before "
     "'Processing complete'."),

    ("Accounting Principles", "Sensitivity failure does not abort run",
     "Trigger a labelling failure (e.g. close Excel forcibly during the labelling step, "
     "or run on a machine where Excel is not installed). Run the strategy.",
     "The .xlsx is still produced (just unlabelled). The processing log shows "
     "'[Sensitivity] Could not apply label: <reason>' in BOLD RED but the run "
     "completes with the 'Processing complete' status. Tester can apply the label "
     "manually."),

    # Error visibility
    ("Accounting Principles", "EBX file locked in Excel",
     "Open the EBX .xlsx in Excel. Run the strategy.",
     "Bold-red error line: '[System] Error loading files: ... Permission "
     "denied: ...'. Chime plays. Action button shows 'Return to Form'."),

    ("Accounting Principles", "Bad sheet name",
     "Type a non-existent sheet name in the EBX sheet field. Run.",
     "Bold-red error line: '[System] Error loading files: Could not find sheet "
     "'<typed>'...'. Chime plays."),

    ("Accounting Principles", "Required column missing — abort",
     "Edit a copy of the EBX file to remove the 'X-Check No.' column. Run.",
     "Bold-red line: '[EBX] Required column 'X-Check No.' not found — aborting'. "
     "Chime plays. Action button shows 'Return to Form'."),
]


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


def write_uat_tests(ws, version):
    ws.title = "UAT Tests"
    ws.sheet_view.showGridLines = False

    # 1. Files Required
    write_section_title(ws, 1, "Files Required")
    write_section_header(ws, 2, ["CaseUI", "Field Label", "Filename",
                                  "Sheet (if applicable)", "Required"])
    r = 3
    for vals in FILES_ROWS:
        write_section_row(ws, r, vals, 5, height=15.75)
        r += 1

    # 2. General UI
    title_row = r + 2
    write_section_title(ws, title_row, "General UI and Dialog Behavior Test Cases")
    write_section_header(ws, title_row + 1,
        ["Test", "Action", "Expected Result", "Tested By", "Tested Date", "Result"])
    r = title_row + 2
    for case in GENERAL_UI_CASES:
        vals = [case[0], case[1].replace("{ver}", version),
                case[2].replace("{ver}", version), "", "", ""]
        write_section_row(ws, r, vals, 6)
        r += 1

    # 3. Detailed Field Interaction
    title_row = r + 1
    write_section_title(ws, title_row, "Detailed Field Interaction Test Cases")
    write_section_header(ws, title_row + 1,
        ["Test", "Action", "Expected Result", "Tested By", "Tested Date", "Result"])
    r = title_row + 2
    for case in FIELD_CASES:
        write_section_row(ws, r, [case[0], case[1], case[2], "", "", ""], 6)
        r += 1

    # 4. Workflow-Specific
    title_row = r + 1
    write_section_title(ws, title_row, "Workflow-Specific Test Cases")
    write_section_header(ws, title_row + 1,
        ["Workflow", "Test", "Action", "Expected Result",
         "Tested By", "Tested Date", "Result"])
    r = title_row + 2
    for case in WORKFLOW_CASES:
        write_section_row(ws, r, [case[0], case[1], case[2], case[3], "", "", ""], 7)
        r += 1

    for col, w in {"A": 19.4, "B": 23.7, "C": 60.3, "D": 22.9,
                   "E": 11.1, "F": 36.1, "G": 29.7, "H": 53.4}.items():
        ws.column_dimensions[col].width = w


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
