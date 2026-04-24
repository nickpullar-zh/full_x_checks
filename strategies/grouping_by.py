from csv import writer
import os
import tkinter as tk
from tkinter import ttk
import pandas as pd
import re
import shutil
import app_state
from datetime import datetime
from strategies.base_strategy import BaseStrategy
from task_configs import GROUPING_BY_UPLOAD_CONFIG, UploadTaskConfig

class GroupingBy(BaseStrategy):

    def __init__(self, config: UploadTaskConfig):  # Accept config
        super().__init__(config)                   # Pass up to BaseStrategy
    
    def process(self, loaded_files, files, output_directory):
        print(loaded_files.keys())  

        log = []  # Shared log passed through all steps

        df_mapping_file, df_fip_original, df_fip_processed = self._process_fip(loaded_files, log)
        if df_fip_original is None:
            print("FIP processing failed — aborting.")
            return

        df_ebs_original, df_ebs_processed = self._process_ebs(loaded_files, log)
        if df_ebs_original is None:
            print("EBS processing failed — aborting.")
            return

        self.write_excel_output(
            output_path=self.build_output_path(
                output_directory,
                "X-Check Grouping By Results",
                app_state.timestamp
            ),
            sheets={
                "Mapping File":    df_mapping_file,
                "FIP - Original":  df_fip_original,
                "FIP - Processed": df_fip_processed,
                "EBS - Original":  df_ebs_original,
                "EBS - Processed": df_ebs_processed,
            },
            log=log
        )

    def _process_fip(self, loaded_files, log) -> tuple[pd.DataFrame|None, pd.DataFrame|None, pd.DataFrame|None]:
        try:
            """Returns (mapping, original, processed) DataFrames and appends to log."""
            print("Processing FIP file...")

            self.log_step(log, "FIP", "Started processing", 0)
            df_original = loaded_files[GROUPING_BY_UPLOAD_CONFIG.file_fields[0].label].copy()
            self.log_step(log, "FIP", "Original File copied", len(df_original), "Copied original DataFrame")

            self.log_step(log, "Mapping File", "Started processing", 0)
            mapping_file_content = loaded_files[GROUPING_BY_UPLOAD_CONFIG.file_fields[2].label]  # Mapping File
            self.log_step(log, "Mapping File", "Loaded", len(mapping_file_content.splitlines()),"Including header row")
            mapping_dict = {}
            for line in mapping_file_content.splitlines()[1:]:  # [1:] skips the header row "FIP Data,EBS item"
                line = line.strip()
                if not line:
                    continue  # Skip empty lines
                parts = line.split(",", maxsplit=1)  # maxsplit=1 protects against commas in values
                if len(parts) == 2:
                    key, value = parts[0].strip(), parts[1].strip()
                    mapping_dict[key] = value
            # Split each line on comma into two values, then build the DataFrame
            df_mapping_file = pd.DataFrame(
                [line.split(",", maxsplit=1) for line in mapping_file_content.splitlines()[1:] if line.strip()],
                columns=["FIP Data", "EBS item"]
            )
            self.log_step(log, "Mapping File", "Mapping dictionary created", len(mapping_dict))

            # Access directly from loaded_files
            df_fip = df_original.copy()  # 'FIP File (ZQ9_VALFLDGR)'
            self.log_step(log, "FIP", "Original file", len(df_original), "FIP Dataframe ready for processing")
            
            # --- Lookup: map "Field name" to EBS item ---
            df_fip["EBS Item"] = df_fip["Field name"].map(mapping_dict)
            self.log_step(log, "FIP", "Mapped 'Field name' to 'EBS Item'", len(df_fip), "Mapped using mapping dictionary")

            # --- Build Key column ---+
            df_fip["Key"] = df_fip.apply(
                lambda row: f"{row['ValidRule']}|{row['EBS Item']}"
                if pd.notna(row["ValidRule"]) and str(row["ValidRule"]).strip() != ""
                else "",
                axis=1
            )
            self.log_step(log, "FIP", "Constructed 'Key' column", len(df_fip), "Concatenated 'ValidRule' and mapped 'EBS Item'")

            # Remove all rows where Key is empty or whitespace
            df_fip = df_fip[df_fip["Key"].str.strip() != ""]    
            self.log_step(log, "FIP", "Removed blank rows", len(df_fip), "Removed rows where 'Key' is empty or whitespace")
            
            self.log_step(log, "FIP", "Finished processing", len(df_fip), "Returned processed DataFrame for comparison")
            return df_mapping_file, df_original, df_fip # Return for use in compare
        except Exception as e:
            import traceback
            print(f"  ERROR inside _process_fip: {e}")
            print(traceback.format_exc())
            return None, None, None  # ← Prevents the unpack error masking the real one

    def _process_ebs(self, loaded_files, log) -> tuple[pd.DataFrame|None, pd.DataFrame|None]:
        """Processes the EBS file and writes output."""
        print("Processing EBS file...")
        self.log_step(log, "EBS", "Started processing", 0)

        # Step 1 — Load the EBS file
        df_ebs_original = loaded_files[GROUPING_BY_UPLOAD_CONFIG.file_fields[1].label].copy()
        df_ebs_loaded = df_ebs_original.copy()
        self.log_step(log, "EBS", "Original file", len(df_ebs_original))

        # Step 2 — Strip out rows that don't have a 'Grouping By' element
        df_ebs_loaded = df_ebs_loaded[df_ebs_loaded["Grouping By"].notna() & (df_ebs_loaded["Grouping By"].str.strip() != "")]
        self.log_step(log, "EBS", "Filtered to rows with 'Grouping By'", len(df_ebs_loaded), "Removed rows without 'Grouping By' value")

        # Step 3 — Remove duplicate rows based on "X-Check No." — keep first occurrence
        df_ebs_loaded = df_ebs_loaded.drop_duplicates(subset=["X-Check No."], keep="first").reset_index(drop=True)
        self.log_step(log, "EBS", "Removed duplicate 'X-Check No.' rows", len(df_ebs_loaded), "Kept first occurrence of duplicates")

        # Step 4 — Split "Grouping By" column on comma into separate columns
        split_cols = df_ebs_loaded["Grouping By"].str.split(",", expand=True)
        split_cols.columns = [f"Grouping By {i + 1}" for i in range(split_cols.shape[1])]
        split_cols = split_cols.apply(lambda col: col.str.strip())
        self.log_step(log, "EBS", "Split 'Grouping By' into separate columns", split_cols.notna().sum().sum(), f"Created {split_cols.shape[1]} 'Grouping By n' columns")    

        # Step 5 — Drop original "Grouping By" column and insert split columns in its place
        col_position = df_ebs_loaded.columns.get_loc("Grouping By")
        df_ebs_loaded = df_ebs_loaded.drop(columns=["Grouping By"])
        for i, col in enumerate(reversed(split_cols.columns.tolist())):
            df_ebs_loaded.insert(col_position, col, split_cols[col])
        self.log_step(log, "EBS", "Inserted split 'Grouping By n' columns", len(df_ebs_loaded), "Replaced original 'Grouping By' column with split columns")    

        # Step 6 — Build the base key value
        df_ebs_loaded["_base_key"] = df_ebs_loaded["Reference  X-Check (Condition)"].where(
            df_ebs_loaded["Reference  X-Check (Condition)"].notna() &
            (df_ebs_loaded["Reference  X-Check (Condition)"].str.strip() != ""),
            other=df_ebs_loaded["X-Check No."].astype(str).str.strip()
        )
        self.log_step(log, "EBS", "Constructed base key column", df_ebs_loaded["_base_key"].notna().sum(), "Used 'Reference  X-Check (Condition)' where available, otherwise 'X-Check No.'")    

        # Step 7 — For each "Grouping By n" column, build key columns
        grouping_by_cols = [col for col in df_ebs_loaded.columns if col.startswith("Grouping By ")]
        for col in grouping_by_cols:
            key_col_name = col.replace("Grouping By ", "Grouping By Key ")
            df_ebs_loaded[key_col_name] = df_ebs_loaded.apply(
                lambda row, c=col: (
                    f"{row['_base_key']}|{str(row[c]).strip()}"
                    if pd.notna(row[c]) and str(row[c]).strip() != ""
                    else ""
                ),
                axis=1
            )
        self.log_step(log, "EBS", "Constructed 'Grouping By Key n' columns", df_ebs_loaded[[col for col in df_ebs_loaded.columns if col.startswith("Grouping By Key ")]].notna().sum().sum(), "Concatenated base key with each 'Grouping By n' value to create 'Grouping By Key n' columns")    

        # Step 8 — Drop the temporary base key column
        df_ebs = df_ebs_loaded.drop(columns=["_base_key"])

        # Step 9 — Stack
        key_cols = [col for col in df_ebs.columns if col.startswith("Grouping By Key ")]
        stacked = (
            df_ebs.set_index([c for c in df_ebs.columns if c not in key_cols])
            .stack()
            .reset_index()
            .rename(columns={0: "Key"})
            .drop(columns=[f"level_{len(df_ebs.columns) - len(key_cols)}"])
        )
        self.log_step(log, "EBS", "Stacked 'Grouping By Key n' columns into single 'Key' column", len(stacked), "Transformed from wide to long format based on 'Grouping By Key n' columns")    

        # Step 10 — Remove rows where Key is empty or whitespace
        df_ebs = stacked[stacked["Key"].str.strip() != ""]

        # Step 11 — Reset index cleanly
        df_ebs = stacked.reset_index(drop=True)
        
        self.log_step(log, "EBS", "Finished processing", len(df_ebs), "Returned processed DataFrame for comparison")
        return df_ebs_original, df_ebs # Return for use in compare

    def _process_compare(self, df_fip: pd.DataFrame, df_ebs: pd.DataFrame, output_directory: str):
        """Compares FIP and EBS processed data and writes output."""
        print("Running comparison...")
        # ... compare logic ...

    # def _build_output_path(self, output_directory: str, label: str, timestamp: str) -> str:
    #     """
    #     Builds a safe, timestamped output file path from a label.
    #     Replaces characters that are invalid in Windows filenames.
    #     Returns the full output path including filename.
    #     """
    #     safe_label = re.sub(r'[<>:"/\\|?*()]', '_', label)
    #     filename = f"{timestamp}_{safe_label}.xlsx"
    #     return os.path.join(output_directory, filename)
    
    # def _write_output(self, output_directory, df_mapping_file, df_fip_original, df_fip_processed,        
    #                 df_ebs_original, df_ebs_processed, log):
    #     #def _write_output(self, output_directory, df_fip_original, df_fip_processed,
    #     #              df_ebs_original, df_ebs_processed, df_compare, log):        
        
    #     import shutil
    #     from config import OUTPUT_TEMPLATE

    #     """Writes all sheets to a single timestamped Excel workbook."""
    #     output_path = self.build_output_path(
    #         output_directory,
    #         "X-Check Grouping By Results",
    #         app_state.timestamp
    #     )

    #     # Copy pre-labelled template to output path
    #     shutil.copy(OUTPUT_TEMPLATE, output_path)

    #     df_log = pd.DataFrame(log)  # Convert log list to DataFrame

    #     with pd.ExcelWriter(output_path, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
    #         # df_compare.to_excel(writer, sheet_name="Compare",               index=False)
    #         df_mapping_file.to_excel(writer, sheet_name="Mapping File",     index=False)
    #         df_fip_original.to_excel(writer, sheet_name="FIP - Original",   index=False)
    #         df_fip_processed.to_excel(writer, sheet_name="FIP - Processed", index=False)
    #         df_ebs_original.to_excel(writer, sheet_name="EBS - Original",   index=False)
    #         df_ebs_processed.to_excel(writer, sheet_name="EBS - Processed", index=False)
    #         df_log.to_excel(writer, sheet_name="Processing Log",            index=False)
        
    #         # Remove default Sheet1 from template
    #         if "Sheet1" in writer.book.sheetnames:
    #             del writer.book["Sheet1"]
            
    #         # Auto-fit all sheets after writing
    #         for sheet_name in writer.book.sheetnames:
    #             print("Autofitting columns in sheet:", sheet_name) # Debug statement to confirm we're in the loop
    #             self.autofit_columns(writer.book[sheet_name])

    #     print(f"Output written to: {output_path}")

    # def _autofit_columns(self, worksheet, max_width: int = 60):
    #     """Auto-fits all columns in a worksheet to their content width, capped at max_width."""
    #     for column in worksheet.columns:
    #         max_length = 0
    #         col_letter = column[0].column_letter
    #         for cell in column:
    #             if cell.value:
    #                 max_length = max(max_length, len(str(cell.value)))
    #         worksheet.column_dimensions[col_letter].width = min(max_length + 2, max_width)  # Capped

    def _log(self, log: list, file: str, step: str, rows: int, notes: str = ""):
        """Helper to append a timestamped entry to the processing log."""
        log.append({
            "Timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "File": file,
            "Step": step,
            "Rows": rows,
            "Notes": notes
        })