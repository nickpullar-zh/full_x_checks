import os
import pandas as pd
from strategies.base_strategy import BaseStrategy, UploadTaskConfig
from .ebx_extraction import extract_ebx
from .fip_extraction import extract_fip
from .compare import compare
from .x_check_no_selection import select_x_check_nos


class XChecks(BaseStrategy):

    def __init__(self, config: UploadTaskConfig):
        super().__init__(config)

    def process(self, loaded_files: dict, files: dict):
        self.log_step(self.log, "System", "Starting X-Checks processing", 0)

        # 0. X-Check No Selection — write the list of changed X-Check Nos to a .txt
        #    file the user can paste into FIP. Only runs when "Process only differences"
        #    is enabled, since otherwise every X-Check would be listed.
        if files.get("process_only_differences", False):
            self._write_x_check_no_list(loaded_files, files)

        # 1. Load GCoA QU accounts (optional)
        qu_accounts: set = set()
        gcoa_df = loaded_files.get("GCoA Publication File")
        if gcoa_df is not None:
            qu_mask = gcoa_df["Data type"].astype(str).str.strip() == "QU"
            qu_accounts = set(gcoa_df.loc[qu_mask, "Account ID"].astype(str).str.strip())
            self.log_step(self.log, "GCoA", "QU accounts loaded", len(qu_accounts))
        else:
            self.log_step(self.log, "GCoA", "No GCoA file provided — QU_YTD substitution skipped", 0)

        # 2. Extract EBX
        self.log_step(self.log, "EBX", "Extracting from publication file...", 0)
        ebx_results = extract_ebx(
            loaded_files["X-Checks Publication File"],
            qu_accounts=qu_accounts,
            apply_version_spanning=files.get("apply_version_spanning", False),
            apply_prior_year_balance=files.get("apply_prior_year_balance", False),
        )
        self.log_step(self.log, "EBX", "X-Checks extracted", len(ebx_results))

        # 3. Extract FIP — x_check_list from all unique X-Check No. values in the raw file,
        #    matching old FIPExtraction.py which used the EBX file directly rather than
        #    extraction results (ensures X-Checks with no Account No. rows are still searched in FIP)
        ebx_df = loaded_files["X-Checks Publication File"]
        if "X-Check No." not in ebx_df.columns:
            self.log_step(self.log, "EBX", "Required column 'X-Check No.' not found — aborting", 0)
            return
        x_check_list = sorted(set(
            str(x) for x in ebx_df["X-Check No."].tolist()
            if str(x) not in ("nan", "", "NaN", "None")
        ))
        self.log_step(self.log, "FIP", "Extracting from FIP text...", len(x_check_list))
        fip_results = extract_fip(loaded_files["FIP File"], x_check_list)
        self.log_step(self.log, "FIP", "X-Checks extracted", len(fip_results))

        # 4. Compare and sort — matches old Compare_Files.py "All Data" sheet sort order
        self.log_step(self.log, "Comparison", "Comparing EBX and FIP...", 0)
        comparison_rows = compare(ebx_results, fip_results)
        if not comparison_rows:
            self.log_step(self.log, "Comparison", "No X-Checks to compare — aborting output", 0)
            return
        df_comparison = pd.DataFrame(comparison_rows)
        df_comparison = df_comparison.sort_values("X-Check No.").reset_index(drop=True)

        # 5. Apply known exceptions if file was provided
        _XC_FINGERPRINT = [
            "X-Check No.", "EBX Formula", "FIP Formula",
            "EBX Formula (Excl)", "FIP Formula (Excl)",
            "EBX Variables", "FIP Variables", "FIP Variable (Builder)",
        ]
        result = self._annotate_known_exceptions(
            df_comparison, files["files"].get("Known Exception List"),
            sheet_name="X-Checks", fingerprint_columns=_XC_FINGERPRINT
        )
        if result is False:
            return
        df_comparison = result

        # 6. Write Excel output — no summary, headers at row 1
        self.write_excel_output(
            output_path=self.build_output_path(
                files["output_directory"], "Comparison", files["timestamp"]
            ),
            sheets={"Comparison": df_comparison},
            log=self.log,
        )
        return True

    def _write_x_check_no_list(self, loaded_files: dict, files: dict) -> None:
        """
        Runs the full Cross Checks All selection pipeline (Status / Type of Change /
        Exclude Z-Core / yellow Category) and writes the surviving X-Check Nos to
        <timestamp>_X-Check_Nos.txt in the output directory, one per line.

        Operates on a copy of the EBX DataFrame; the loaded df is unchanged so
        downstream comparison still runs against the full dataset.
        """
        ebx_df = loaded_files.get("X-Checks Publication File")
        if ebx_df is None:
            self.log_step(self.log, "X-Check No Selection", "Skipped — EBX not loaded", 0)
            return

        ebx_path = files["files"].get("X-Checks Publication File")
        ebx_sheet = files["sheet_names"].get("X-Checks Publication File")
        if not ebx_path or not ebx_sheet:
            self.log_step(self.log, "X-Check No Selection",
                          "Skipped — EBX file path or sheet not provided", 0)
            return

        x_check_nos = select_x_check_nos(ebx_df, ebx_path, ebx_sheet)
        if not x_check_nos:
            self.log_step(self.log, "X-Check No Selection",
                          "No X-Check Nos in scope after pipeline", 0)
            return

        out_path = self.build_output_path(
            files["output_directory"], "X-Check_Nos", files["timestamp"], extension=".txt"
        )
        with open(out_path, "w", encoding="utf-8") as f:
            f.write("\n".join(x_check_nos))
        self.log_step(self.log, "X-Check No Selection",
                      f"Wrote {os.path.basename(out_path)}", len(x_check_nos),
                      notes=out_path)

    def apply_output_formatting(self, workbook):
        if "Comparison" not in workbook.sheetnames:
            return
        ws = workbook["Comparison"]

        COMPARISON_COLS = (
            "Formula Match", "Formula Match (Excl)",
            "Variables Match", "Variables Match (Builder)",
        )

        header_values = [cell.value for cell in ws[1]]

        def _col(name):
            return header_values.index(name) + 1 if name in header_values else None

        xc_col_idx  = _col("X-Check No.")
        kel_col_idx = _col("Known Exception")
        cmp_col_idxs = [_col(c) for c in COMPARISON_COLS if _col(c) is not None]

        for row in ws.iter_rows(min_row=2, min_col=1, max_col=ws.max_column):

            # Is there a valid Known Exception annotation on this row?
            kel_val = ""
            if kel_col_idx:
                raw = row[kel_col_idx - 1].value
                kel_val = str(raw).strip() if raw and str(raw).strip() not in ("", "nan") else ""

            has_excepted  = False   # any MisMatch turned Excepted
            has_mismatch  = False   # any MisMatch remaining (not excepted)
            has_not_found = False

            # Style each comparison cell; rewrite MisMatch → MisMatch (Excepted) when KEL present
            for idx in cmp_col_idxs:
                cell = row[idx - 1]
                val  = str(cell.value).strip() if cell.value is not None else ""

                if val == "MisMatch" and kel_val:
                    cell.value = "MisMatch (Excepted)"
                    cell.fill  = self.FILL_BLUE
                    cell.font  = self.FONT_BLUE
                    has_excepted = True
                elif val == "MisMatch":
                    cell.fill = self.FILL_RED
                    cell.font = self.FONT_RED
                    has_mismatch = True
                elif val == "Not Found":
                    cell.fill = self.FILL_ORANGE
                    cell.font = self.FONT_ORANGE
                    has_not_found = True
                elif val == "Match":
                    cell.fill = self.FILL_GREEN
                    cell.font = self.FONT_GREEN
                elif val == "MisMatch (Excepted)":
                    # Already written (e.g. re-formatting after reload)
                    cell.fill = self.FILL_BLUE
                    cell.font = self.FONT_BLUE
                    has_excepted = True

            # Style Known Exception cell
            if kel_col_idx and kel_val:
                kel_cell = row[kel_col_idx - 1]
                kel_cell.fill = self.FILL_BLUE
                kel_cell.font = self.FONT_BLUE

            # Style X-Check No. cell to reflect overall row result
            if xc_col_idx:
                xc_cell = row[xc_col_idx - 1]
                if has_mismatch:
                    # Any un-excepted MisMatch → red (red beats Not Found)
                    xc_cell.fill = self.FILL_RED
                    xc_cell.font = self.FONT_RED
                elif has_excepted and not has_mismatch and not has_not_found:
                    # All mismatches are excepted, nothing else bad → blue
                    xc_cell.fill = self.FILL_BLUE
                    xc_cell.font = self.FONT_BLUE
                elif has_not_found:
                    xc_cell.fill = self.FILL_ORANGE
                    xc_cell.font = self.FONT_ORANGE
                else:
                    # All Match (possibly with Excepted on some cols — show blue)
                    if has_excepted:
                        xc_cell.fill = self.FILL_BLUE
                        xc_cell.font = self.FONT_BLUE
                    else:
                        xc_cell.fill = self.FILL_GREEN
                        xc_cell.font = self.FONT_GREEN
