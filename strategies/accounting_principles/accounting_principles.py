"""
Accounting Principles strategy.

Compares the severity letter recorded for each X-Check on the EBX 'cross
checks all' sheet against the W/E recorded in the 'FIP Methods Rules and
Condition' sheet (the VALMSG dump). The Validation Methods xlsx supplies
the candidate methods per Validation Event and their expected severity.
"""
from __future__ import annotations

import pandas as pd

from strategies.base_strategy import BaseStrategy, UploadTaskConfig
from .validation_methods import parse_validation_methods
from .compare import compare


# Default subset of Validation Events used on first run (before the user's
# choices have been persisted to %APPDATA%/X-Checks/accounting_principles.json).
# Names must match the row-1 headers of the validation methods file exactly.
DEFAULT_EVENTS = [
    "IFRS New RFD", "IFRS New SFD", "IFRS New CFD",
    "Stammhaus SLST RFD", "Stammhaus SLST SFD", "Stammhaus SLST CFD",
    "SST RFD", "SST SFD", "SST CFD",
    "SII RFD", "SII SFD", "SII CFD",
    "RI Assets IFRSN RFD", "RI Assets IFRSN SFD", "RI Assets IFRSN CFD",
    "Equity IFRSN RFD", "Equity IFRSN SFD", "Equity IFRSN CFD",
    "I/C IFRSN RFD", "I/C IFRSN SFD", "I/C IFRSN CFD",
    "Tax IFRSN RFD", "Tax IFRSN SFD", "Tax IFRSN CFD",
    "DE-GAAP RFD", "DE-GAAP SFD", "DE-GAAP CFD",
]

OUTPUT_COLUMNS = [
    "X-Check No.", "Event", "Expected", "FIP", "Actual", "Method", "Match",
]


class AccountingPrinciples(BaseStrategy):

    def __init__(self, config: UploadTaskConfig):
        super().__init__(config)

    def process(self, loaded_files: dict, files: dict):
        self.log_step(self.log, "System", "Starting Accounting Principles processing", 0)

        vm_path = files["files"].get("Validation Methods File")
        if not vm_path:
            self.log_step(self.log, "System",
                          "Missing required file: Validation Methods File", 0)
            return

        # 1. Validation Methods: parse expected severity per event
        subset = files.get("validation_events_subset") or DEFAULT_EVENTS
        self.log_step(self.log, "Validation Methods",
                      "Parsing validation methods file...", len(subset))
        defs = parse_validation_methods(vm_path, subset)
        self.log_step(self.log, "Validation Methods",
                      "Definitions extracted", len(defs))

        # 2. Cross Checks All — already loaded by BaseStrategy._load_files,
        #    using header_signals to detect the right header row automatically.
        cc_df = loaded_files.get("X-Checks Publication File")
        if cc_df is None:
            self.log_step(self.log, "System",
                          "Missing required file: X-Checks Publication File", 0)
            return

        # 3. FIP Methods Rules and Condition — also already loaded.
        fip_df = loaded_files.get("FIP File (VALMSG)")
        if fip_df is None:
            self.log_step(self.log, "System",
                          "Missing required file: FIP File (VALMSG)", 0)
            return

        # 4. In-scope X-Checks: every unique non-blank X-Check No. for now.
        # NOTE: a future revision can plug in select_x_check_nos here.
        x_check_col = "X-Check No."
        if x_check_col not in cc_df.columns:
            self.log_step(self.log, "EBX",
                          f"Required column '{x_check_col}' not found — aborting", 0)
            return
        xchecks: list[str] = []
        seen: set = set()
        for v in cc_df[x_check_col].tolist():
            s = str(v).strip()
            if s in ("", "nan", "None") or s in seen:
                continue
            seen.add(s)
            xchecks.append(s)
        self.log_step(self.log, "EBX", "In-scope X-Check Nos", len(xchecks))

        # 5. Compare
        rows = compare(defs, cc_df, xchecks, fip_df)
        if not rows:
            self.log_step(self.log, "Compare",
                          "No comparable rows produced — aborting output", 0)
            return
        df = pd.DataFrame(rows, columns=OUTPUT_COLUMNS)
        self.log_step(self.log, "Compare", "Comparison rows produced", len(df))

        # 6. Write output
        out_path = self.build_output_path(
            files["output_directory"],
            "Accounting Principles Comparison",
            files["timestamp"],
        )
        self.write_excel_output(
            output_path=out_path,
            sheets={"Accounting Principles": df},
            log=self.log,
        )
        return True

    def apply_output_formatting(self, workbook):
        from openpyxl.styles import PatternFill, Font

        sheet_name = "Accounting Principles"
        if sheet_name not in workbook.sheetnames:
            return

        green_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
        green_font = Font(color="276221")
        red_fill   = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
        red_font   = Font(color="9C0006")

        ws = workbook[sheet_name]
        self.apply_conditional_formatting(
            worksheet=ws,
            column_name="Match",
            rules={
                "Match":    (green_fill, green_font),
                "MisMatch": (red_fill,   red_font),
            },
        )
