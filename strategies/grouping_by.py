from csv import writer
from logging import log
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

        df_mapping_file, df_fip_original, df_fip_processed = self._process_fip(loaded_files)
        if df_fip_original is None:
            print("FIP processing failed — aborting.")
            return

        df_ebs_original, df_ebs_processed = self._process_ebs(loaded_files)
        if df_ebs_original is None:
            print("EBS processing failed — aborting.")
            return
        
        assert df_fip_processed is not None
        assert df_ebs_processed is not None
        df_comparison = self._process_compare(df_fip_processed, df_ebs_processed)

        assert df_comparison is not None
        matched     = (df_comparison["Result"] == "Matched").sum()
        not_matched = (df_comparison["Result"] == "Not in FIP").sum()
        fip_path    = files["files"][GROUPING_BY_UPLOAD_CONFIG.file_fields[0].label]  # ← add here
        ebs_path    = files["files"][GROUPING_BY_UPLOAD_CONFIG.file_fields[1].label]  # ← add here

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
                "Compare":      df_comparison
            },
            log=self.log,
            summaries={
                "FIP - Original":  {"Source filename:": fip_path, "Number of rows:": len(df_fip_original)},
                "FIP - Processed": {"Source filename:": fip_path, "Number of rows:": len(df_fip_processed)},
                "EBS - Original":  {"Source filename:": ebs_path, "Number of rows:": len(df_ebs_original)},
                "EBS - Processed": {"Source filename:": ebs_path, "Number of rows:": len(df_ebs_processed)},
                "Compare":         {"Number of rows:": len(df_comparison), "Number matched:": matched, "Number of errors:": not_matched},            
            }
        )

    def _process_fip(self, loaded_files) -> tuple[pd.DataFrame|None, pd.DataFrame|None, pd.DataFrame|None]:
        try:
            """Returns (mapping, original, processed) DataFrames and appends to log."""
            self.log_step(self.log, "Mapping File", "Started processing", 0)
            mapping_file_content = loaded_files[GROUPING_BY_UPLOAD_CONFIG.file_fields[2].label]  # Mapping File
            self.log_step(self.log, "Mapping File", "Loaded", len(mapping_file_content.splitlines()),"Including header row")
            mapping_dict = {}
            for line in mapping_file_content.splitlines()[1:]:  # [1:] skips the header row "FIP Data,EBS item"
                line = line.strip()
                if not line:
                    continue  # Skip empty lines
                parts = line.split(",", maxsplit=1)  # maxsplit=1 protects against commas in values
                if len(parts) == 2:
                    key, value = parts[0].strip(), parts[1].strip()
                    mapping_dict[key] = value
            self.log_step(self.log, "Mapping File", "Mapping dictionary created", len(mapping_dict))
            # Split each line on comma into two values, then build the DataFrame
            df_mapping_file = pd.DataFrame(
                [line.split(",", maxsplit=1) for line in mapping_file_content.splitlines()[1:] if line.strip()],
                columns=["FIP Data", "EBS item"]
            )
            self.log_step(self.log, "Mapping File", "df_mapping_data created", len(df_mapping_file), "DataFrame created from mapping dictionary for output")
            self.log_step(self.log, "Mapping File", "Finished processing", len(mapping_dict))

            # Access directly from loaded_files
            self.log_step(self.log, "FIP", "Started processing", 0)
            df_original = loaded_files[GROUPING_BY_UPLOAD_CONFIG.file_fields[0].label].copy()
            self.log_step(self.log, "FIP", "Original File copied", len(df_original), "Copied original DataFrame for output comparison")            
            df_fip = df_original.copy()  # 'FIP File (ZQ9_VALFLDGR)'
            self.log_step(self.log, "FIP", "Original file", len(df_original), "FIP Dataframe ready for processing")
            
            # --- Lookup: map "Field name" to EBS item ---
            df_fip["EBS Item"] = df_fip["Field name"].map(mapping_dict)
            self.log_step(self.log, "FIP", "Mapped 'Field name' to 'EBS Item'", len(df_fip), "Mapped using mapping dictionary")

            # Remove all rows where EBS Item is empty or whitespace
            df_fip = df_fip[
                df_fip["EBS Item"].notna() &   
                (df_fip["EBS Item"].str.strip() != "") & 
                (df_fip["EBS Item"].str.strip().str.lower() != "ignore")  # Also exclude rows where mapping resulted in "ignore"
            ]
            # Remove all rows where ValidRule Item is empty or whitespace
            df_fip = df_fip[
                df_fip["ValidRule"].notna() &   
                (df_fip["ValidRule"].str.strip() != "")
            ]
            # Debug — remove once fixed
            self.log_step(self.log, "FIP", "Removed blank rows", len(df_fip), "Removed rows where 'Key' is empty or 'ignore'")

            # --- Build Key column ---+
            df_fip["Key"] = df_fip.apply(
                lambda row: f"{row['ValidRule']}|{row['EBS Item']}"
                if pd.notna(row["ValidRule"]) and str(row["ValidRule"]).strip() != ""
                else "",
                axis=1
            )
            self.log_step(self.log, "FIP", "Constructed 'Key' column", len(df_fip), "Concatenated 'ValidRule' and mapped 'EBS Item'")
            
            self.log_step(self.log, "FIP", "Finished processing", len(df_fip), "Returned processed DataFrame for comparison")
            return df_mapping_file, df_original, df_fip # Return for use in compare
        except Exception as e:
            import traceback
            print(f"  ERROR inside _process_fip: {e}")
            print(traceback.format_exc())
            return None, None, None  # ← Prevents the unpack error masking the real one

    def _process_ebs(self, loaded_files) -> tuple[pd.DataFrame|None, pd.DataFrame|None]:
        """Processes the EBS file and writes output."""
        self.log_step(self.log, "EBS", "Started processing", 0)

        # Step 1 — Load the EBS file
        df_ebs_original = loaded_files[GROUPING_BY_UPLOAD_CONFIG.file_fields[1].label].copy()
        df_ebs_loaded = df_ebs_original.copy()
        self.log_step(self.log, "EBS", "Original file copied", len(df_ebs_original), "Copied original DataFrame for output comparison")

        # Step 2 — Strip out rows that don't have a 'Grouping By' element
        df_ebs_loaded = df_ebs_loaded[df_ebs_loaded["Grouping By"].notna() & (df_ebs_loaded["Grouping By"].str.strip() != "")]
        self.log_step(self.log, "EBS", "Filtered to rows with 'Grouping By'", len(df_ebs_loaded), "Removed rows without 'Grouping By' value")

        # Step 3 — Remove duplicate rows based on "X-Check No." — keep first occurrence
        df_ebs_loaded = df_ebs_loaded.drop_duplicates(subset=["X-Check No."], keep="first").reset_index(drop=True)
        self.log_step(self.log, "EBS", "Removed duplicate 'X-Check No.' rows", len(df_ebs_loaded), "Kept first occurrence of duplicates")

        # Step 4 — Split "Grouping By" column on comma into separate columns
        split_cols = df_ebs_loaded["Grouping By"].str.split(",", expand=True)
        split_cols.columns = [f"Grouping By {i + 1}" for i in range(split_cols.shape[1])]
        split_cols = split_cols.apply(lambda col: col.str.strip())
        self.log_step(self.log, "EBS", "Split 'Grouping By' into separate columns", split_cols.notna().sum().sum(), f"Created {split_cols.shape[1]} 'Grouping By n' columns")    

        # Step 5 — Drop original "Grouping By" column and insert split columns in its place
        col_position = df_ebs_loaded.columns.get_loc("Grouping By")
        df_ebs_loaded = df_ebs_loaded.drop(columns=["Grouping By"])
        for i, col in enumerate(reversed(split_cols.columns.tolist())):
            df_ebs_loaded.insert(col_position, col, split_cols[col])
        self.log_step(self.log, "EBS", "Inserted split 'Grouping By n' columns", len(df_ebs_loaded), "Replaced original 'Grouping By' column with split columns")    

        # Step 6 — Build the base key value
        df_ebs_loaded["_base_key"] = df_ebs_loaded["Reference  X-Check (Condition)"].where(
            df_ebs_loaded["Reference  X-Check (Condition)"].notna() &
            (df_ebs_loaded["Reference  X-Check (Condition)"].str.strip() != ""),
            other=df_ebs_loaded["X-Check No."].astype(str).str.strip()
        )
        self.log_step(self.log, "EBS", "Constructed base key column", df_ebs_loaded["_base_key"].notna().sum(), "Used 'Reference  X-Check (Condition)' where available, otherwise 'X-Check No.'")    

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
        self.log_step(self.log, "EBS", "Constructed 'Grouping By Key n' columns", df_ebs_loaded[[col for col in df_ebs_loaded.columns if col.startswith("Grouping By Key ")]].notna().sum().sum(), "Concatenated base key with each 'Grouping By n' value to create 'Grouping By Key n' columns")    

        # Step 8 — Drop the temporary base key column
        df_ebs = df_ebs_loaded.drop(columns=["_base_key"])

        # Step 9 — Stack
        key_cols = [col for col in df_ebs.columns if col.startswith("Grouping By Key ")]
        index_cols = [c for c in df_ebs.columns if c not in key_cols]
        stacked = (
            df_ebs.set_index(index_cols)
            .stack()
            .reset_index()
            .rename(columns={0: "Key"})
        )
        # Drop the level column that stack() adds (named after the column position)
        level_col = f"level_{len(index_cols)}"
        if level_col in stacked.columns:
            stacked = stacked.drop(columns=[level_col])
        self.log_step(self.log, "EBS", "Stacked 'Grouping By Key n' columns into single 'Key' column", len(stacked), "Transformed from wide to long format based on 'Grouping By Key n' columns")    

        # Step 10 — Remove rows where Key is empty or whitespace
        df_ebs = stacked[stacked["Key"].str.strip() != ""].reset_index(drop=True)  # Fix: filter AND reset on same variable
        
        self.log_step(self.log, "EBS", "Finished processing", len(df_ebs), "Returned processed DataFrame for comparison")
        return df_ebs_original, df_ebs # Return for use in compare

    def _process_compare(self, df_fip: pd.DataFrame, df_ebs: pd.DataFrame) -> pd.DataFrame|None:
        """Compares FIP and EBS processed data and writes output."""
        
        self.log_step(self.log, "Compare", "Started comparison", 0)

        # Build a lookup set from FIP keys
        fip_keys = df_fip[["Key"]].drop_duplicates().copy()
        fip_keys["In FIP"] = True
        self.log_step(self.log, "Compare", "FIP key lookup built", len(fip_keys), f"{len(fip_keys)} unique keys extracted from FIP")

        # Get unique EBS keys
        ebs_keys = df_ebs[["Key"]].drop_duplicates().copy()
        self.log_step(self.log, "Compare", "EBS keys extracted", len(ebs_keys), f"{len(ebs_keys)} unique keys extracted from EBS")

        # Left join EBS onto FIP keys
        df_compare = ebs_keys.merge(
            fip_keys,
            on="Key",
            how="left"
        )
        self.log_step(self.log, "Compare", "Merge completed", len(df_compare), "Left join of EBS keys onto FIP keys")

        # Fill unmatched rows and add readable Match column
        df_compare["In FIP"] = df_compare["In FIP"].fillna(False)
        df_compare["Result"] = df_compare["In FIP"].map({True: "Matched", False: "Not in FIP"})

        # Summary counts
        matched = df_compare["Result"].eq("Matched").sum()
        not_matched = df_compare["Result"].eq("Not in FIP").sum()
        self.log_step(self.log, "Compare", "Match results calculated", len(df_compare), f"Matched: {matched} | Not in FIP: {not_matched}")

        # Rename columns and sort
        df_compare = (
            df_compare
            .drop(columns=["In FIP"])
            .rename(columns={"Key": "EBS Key"})
            .sort_values("EBS Key")
            .reset_index(drop=True)
        )
        self.log_step(self.log, "Compare", "Finished comparison", len(df_compare))

        return df_compare
