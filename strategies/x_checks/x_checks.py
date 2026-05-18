import pandas as pd
from strategies.base_strategy import BaseStrategy, UploadTaskConfig
from .ebx_extraction import extract_ebx
from .fip_extraction import extract_fip
from .compare import compare


class XChecks(BaseStrategy):

    def __init__(self, config: UploadTaskConfig):
        super().__init__(config)

    def process(self, loaded_files: dict, files: dict):
        self.log_step(self.log, "System", "Starting X-Checks processing", 0)

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
        fip_results = extract_fip(loaded_files["FIP file"], x_check_list)
        self.log_step(self.log, "FIP", "X-Checks extracted", len(fip_results))

        # 4. Compare and sort — matches old Compare_Files.py "All Data" sheet sort order
        self.log_step(self.log, "Compare", "Comparing EBX and FIP...", 0)
        comparison_rows = compare(ebx_results, fip_results)
        if not comparison_rows:
            self.log_step(self.log, "Compare", "No X-Checks to compare — aborting output", 0)
            return
        df_comparison = pd.DataFrame(comparison_rows)
        df_comparison = df_comparison.sort_values("X-Check Number").reset_index(drop=True)

        # 5. Apply known exceptions if file was provided
        exc_path = files["files"].get("Known Exception List")
        if exc_path:
            try:
                known_exceptions = self._load_known_exceptions(exc_path)
            except (ValueError, KeyError) as e:
                self.log_step(self.log, "Exceptions", f"Known Exception List is invalid — aborting: {e}", 0)
                return
            self.log_step(self.log, "Exceptions", "Known exceptions loaded", len(known_exceptions))
            df_comparison["Known Exception"] = df_comparison["X-Check Number"].map(
                lambda x: known_exceptions.get(x, "")
            )
            known_mask = df_comparison["Known Exception"].str.strip() != ""
            for col in ("Formula Match", "Variables Match", "Variables Match (Builder)"):
                if col in df_comparison.columns:
                    df_comparison.loc[known_mask & (df_comparison[col] == "MisMatch"), col] = "Mismatch - Known Exception"
        else:
            self.log_step(self.log, "Exceptions", "No Known Exception List provided — skipping", 0)

        # 6. Write Excel output — no summary, headers at row 1
        self.write_excel_output(
            output_path=self.build_output_path(
                files["output_directory"], "X-Checks Comparison", files["timestamp"]
            ),
            sheets={"X-Checks Comparison": df_comparison},
            log=self.log,
        )

    def _load_known_exceptions(self, path: str) -> dict:
        """
        Reads the 'Known Exceptions' sheet from the given file.
        Returns a dict of {X-Check Number: Reason}.
        Row 2 of the sheet is a guidance row and is skipped.
        Raises ValueError if any data row is missing X-Check Number or Reason.
        """
        try:
            df = pd.read_excel(path, sheet_name="Known Exceptions", skiprows=[1])
        except Exception as e:
            self.log_step(self.log, "Exceptions", f"Could not read Known Exceptions sheet: {e}", 0)
            return {}

        # Empty sheet is acceptable
        df = df.dropna(how="all")
        if df.empty:
            return {}

        missing_cols = [c for c in ("X-Check Number", "Reason") if c not in df.columns]
        if missing_cols:
            raise ValueError(f"Known Exception List is missing required columns: {missing_cols}")

        invalid_rows = []
        exceptions = {}
        for i, row in df.iterrows():
            xc = str(row["X-Check Number"]).strip()
            reason = str(row["Reason"]).strip()
            xc_missing = xc in ("", "nan", "NaN", "None")
            reason_missing = reason in ("", "nan", "NaN", "None")

            if xc_missing or reason_missing:
                label = xc if not xc_missing else f"row {i + 2}"
                missing = []
                if xc_missing:
                    missing.append("X-Check Number")
                if reason_missing:
                    missing.append("Reason")
                invalid_rows.append(f"{label} (missing: {', '.join(missing)})")
                continue

            exceptions[xc] = reason

        if invalid_rows:
            raise ValueError(
                f"Known Exception List has {len(invalid_rows)} incomplete row(s): {'; '.join(invalid_rows)}"
            )

        return exceptions

    def apply_output_formatting(self, workbook):
        from openpyxl.styles import PatternFill, Font

        if "X-Checks Comparison" not in workbook.sheetnames:
            return

        green_fill  = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
        green_font  = Font(color="276221")
        red_fill    = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
        red_font    = Font(color="9C0006")
        orange_fill = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
        orange_font = Font(color="9C6500")

        blue_fill = PatternFill(start_color="DDEEFF", end_color="DDEEFF", fill_type="solid")
        blue_font = Font(color="003399")

        ws = workbook["X-Checks Comparison"]
        for col in ("Formula Match", "Variables Match", "Variables Match (Builder)"):
            self.apply_conditional_formatting(
                worksheet=ws,
                column_name=col,
                rules={
                    "Match":                      (green_fill,  green_font),
                    "MisMatch":                   (red_fill,    red_font),
                    "Not Found":                  (orange_fill, orange_font),
                    "Mismatch - Known Exception": (blue_fill,   blue_font),
                }
            )

        # Highlight known exceptions in blue — applied to every non-blank cell in the column
        header_values = [cell.value for cell in ws[1]]
        if "Known Exception" in header_values:
            col_idx = header_values.index("Known Exception") + 1
            for row in ws.iter_rows(min_row=2, min_col=col_idx, max_col=col_idx):
                cell = row[0]
                if cell.value and str(cell.value).strip() not in ("", "nan"):
                    cell.fill = blue_fill
                    cell.font = blue_font
