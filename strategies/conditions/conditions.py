"""
Conditions strategy — compares X-Check condition data from the publication
file against the FIP ZQ9_VALMETH extract.

Output workbook (4 sheets + Processing Log):
  Conditions    — one row per pair: EBX Data | FIP Data | Comparison (True/False)
  Working Sheet — deduplicated X-Check list with condition values + concat keys
  FIP Data      — renamed FIP columns with Concatenated key column
  Processing Log — standard BaseStrategy log (always written automatically)
"""

import os
from collections import OrderedDict

import pandas as pd

from strategies.base_strategy import BaseStrategy
from strategies.conditions.extract import extract_conditions
from strategies.conditions.fip import process_fip
from strategies.conditions.compare import compare


class Conditions(BaseStrategy):

    TAB_COLOUR        = "2167AE"  # Zurich Blue
    TAB_COLOUR_PASTEL = "BDD7EE"

    def process(self, loaded_files: dict, files: dict) -> bool:
        pub_path = files["files"]["X-Checks Publication File"]
        pub_sheet = files["sheet_names"].get("X-Checks Publication File", "cross checks all")
        fip_df = loaded_files["FIP File (ZQ9_VALMETH)"]

        # ------------------------------------------------------------------
        # 1. Extract yellow/green conditions from publication file
        # ------------------------------------------------------------------
        process_only_differences = files.get("process_only_differences", True)
        mode = "changed/new rows only" if process_only_differences else "full file"
        self.log_step(self.log, "Publication", f"Extracting condition cells ({mode})", 0)
        working_df, warnings = extract_conditions(pub_path, pub_sheet, process_only_differences)

        for w in warnings:
            self.log_step(self.log, "Publication", f"Warning: {w}", 0)

        self.log_step(
            self.log, "Publication",
            f"Extracted {len(working_df)} unique X-Check No. entries",
            len(working_df),
        )

        # ------------------------------------------------------------------
        # 2. Process FIP
        # ------------------------------------------------------------------
        self.log_step(self.log, "FIP", "Renaming columns and building concatenation key", 0)
        fip_processed = process_fip(fip_df)
        self.log_step(
            self.log, "FIP",
            f"FIP processed: {len(fip_processed)} rows, {len(fip_processed.columns)} columns",
            len(fip_processed),
        )

        # ------------------------------------------------------------------
        # 3. Compare
        # ------------------------------------------------------------------
        self.log_step(
            self.log, "Comparison",
            "Matching X-Check|Condition pairs against FIP", 0,
        )
        results_df, summary = compare(working_df, fip_processed)

        self.log_step(
            self.log, "Comparison",
            f"Pairs checked: {summary['Total Pairs']}, "
            f"Matched: {summary['Matched']}, "
            f"Not matched: {summary['Not Matched']}",
            summary["Total Pairs"],
        )

        # ------------------------------------------------------------------
        # 3b. Apply known exceptions (annotation only — Comparison column unchanged)
        # ------------------------------------------------------------------
        result = self._annotate_known_exceptions(
            results_df, files["files"].get("Known Exception List"),
            sheet_name="Conditions", fingerprint_columns=["EBX Data", "FIP Data"]
        )
        if result is False:
            return False
        results_df = result

        # ------------------------------------------------------------------
        # 4. Write output
        # ------------------------------------------------------------------
        output_path = self.build_output_path(
            files["output_directory"],
            "Conditions Comparison",
            files["timestamp"],
        )

        sheets = OrderedDict([
            ("Working Sheet", working_df),
            ("FIP Data",      fip_processed),
            ("Comparison",    results_df),
        ])

        summary_od = OrderedDict(summary)

        self.write_excel_output(output_path, sheets, self.log, summaries={"Comparison": summary_od})
        return True

    def apply_output_formatting(self, workbook):
        if "Comparison" not in workbook.sheetnames:
            return
        self.apply_conditional_formatting(
            worksheet=workbook["Comparison"],
            column_name="Comparison",
            rules={
                "Matched":     (self.FILL_GREEN, self.FONT_GREEN),
                "Not Matched": (self.FILL_RED,   self.FONT_RED),
            },
        )
        self._apply_tab_colours(workbook)
