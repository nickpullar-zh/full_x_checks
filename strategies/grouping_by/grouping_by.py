"""
Grouping By strategy — compares X-Check grouping data from the publication
file against the FIP ZQ9_VALFLDGR extract.

Output workbook (6 sheets + Processing Log):
  Mapping File    — the CSV mapping loaded as a DataFrame
  FIP - Original  — raw FIP data as loaded
  FIP - Processed — FIP with EBS Item mapped + Key column built
  EBS - Original  — raw EBS (publication) data as loaded
  EBS - Processed — EBS filtered, split, stacked, with Key column
  Compare         — EBS Key vs FIP lookup: Matched / Not in FIP
  Processing Log  — standard BaseStrategy log (always written automatically)
"""

import pandas as pd
from collections import OrderedDict

from strategies.base_strategy import BaseStrategy
from task_configs import GROUPING_BY_UPLOAD_CONFIG


class GroupingBy(BaseStrategy):

    def process(self, loaded_files: dict, files: dict) -> bool:

        df_mapping_file, df_fip_original, df_fip_processed = self._process_fip(loaded_files)
        if df_fip_original is None:
            self.log_step(self.log, "Grouping By", "FIP processing failed — aborting.", 0)
            return False

        df_ebs_original, df_ebs_processed = self._process_ebs(loaded_files)
        if df_ebs_original is None:
            self.log_step(self.log, "Grouping By", "EBS processing failed — aborting.", 0)
            return False

        df_comparison = self._process_compare(df_fip_processed, df_ebs_processed)

        matched     = (df_comparison["Result"] == "Matched").sum()
        not_matched = (df_comparison["Result"] == "Not in FIP").sum()

        fip_path = files["files"][GROUPING_BY_UPLOAD_CONFIG.file_fields[0].label]
        ebs_path = files["files"][GROUPING_BY_UPLOAD_CONFIG.file_fields[1].label]

        output_path = self.build_output_path(
            files["output_directory"],
            "Grouping By Comparison",
            files["timestamp"],
        )

        sheets = OrderedDict([
            ("Mapping File",    df_mapping_file),
            ("FIP - Original",  df_fip_original),
            ("FIP - Processed", df_fip_processed),
            ("EBS - Original",  df_ebs_original),
            ("EBS - Processed", df_ebs_processed),
            ("Compare",         df_comparison),
        ])

        summaries = {
            "FIP - Original":  OrderedDict([("Source filename:", fip_path), ("Number of rows:", len(df_fip_original))]),
            "FIP - Processed": OrderedDict([("Source filename:", fip_path), ("Number of rows:", len(df_fip_processed))]),
            "EBS - Original":  OrderedDict([("Source filename:", ebs_path), ("Number of rows:", len(df_ebs_original))]),
            "EBS - Processed": OrderedDict([("Source filename:", ebs_path), ("Number of rows:", len(df_ebs_processed))]),
            "Compare":         OrderedDict([("Number of rows:", len(df_comparison)), ("Matched:", matched), ("Not in FIP:", not_matched)]),
        }

        self.write_excel_output(output_path, sheets, self.log, summaries=summaries)
        return True

    # ------------------------------------------------------------------
    # FIP processing
    # ------------------------------------------------------------------

    def _process_fip(self, loaded_files) -> tuple:
        try:
            self.log_step(self.log, "Mapping File", "Started processing", 0)
            mapping_content = loaded_files[GROUPING_BY_UPLOAD_CONFIG.file_fields[2].label]
            self.log_step(self.log, "Mapping File", "Loaded", len(mapping_content.splitlines()), "Including header row")

            mapping_dict = {}
            for line in mapping_content.splitlines()[1:]:
                line = line.strip()
                if not line:
                    continue
                parts = line.split(",", maxsplit=1)
                if len(parts) == 2:
                    mapping_dict[parts[0].strip()] = parts[1].strip()
            self.log_step(self.log, "Mapping File", "Mapping dictionary created", len(mapping_dict))

            df_mapping_file = pd.DataFrame(
                [line.split(",", maxsplit=1) for line in mapping_content.splitlines()[1:] if line.strip()],
                columns=["FIP Data", "EBS item"]
            )
            self.log_step(self.log, "Mapping File", "Finished processing", len(mapping_dict))

            self.log_step(self.log, "FIP", "Started processing", 0)
            df_original = loaded_files[GROUPING_BY_UPLOAD_CONFIG.file_fields[0].label].copy()
            df_fip = df_original.copy()
            self.log_step(self.log, "FIP", "Original file loaded", len(df_original))

            df_fip["EBS Item"] = df_fip["Field name"].map(mapping_dict)
            self.log_step(self.log, "FIP", "Mapped 'Field name' to 'EBS Item'", len(df_fip))

            df_fip = df_fip[
                df_fip["EBS Item"].notna() &
                (df_fip["EBS Item"].str.strip() != "") &
                (df_fip["EBS Item"].str.strip().str.lower() != "ignore")
            ]
            df_fip = df_fip[
                df_fip["ValidRule"].notna() &
                (df_fip["ValidRule"].str.strip() != "")
            ]
            self.log_step(self.log, "FIP", "Removed unmapped and blank rows", len(df_fip))

            df_fip["Key"] = df_fip.apply(
                lambda row: f"{row['ValidRule']}|{row['EBS Item']}"
                if pd.notna(row["ValidRule"]) and str(row["ValidRule"]).strip() != ""
                else "",
                axis=1
            )
            self.log_step(self.log, "FIP", "Finished processing", len(df_fip))
            return df_mapping_file, df_original, df_fip

        except Exception as exc:
            import traceback
            self.log_step(self.log, "FIP", f"Exception: {exc}", 0)
            self.log_step(self.log, "FIP", traceback.format_exc(), 0)
            return None, None, None

    # ------------------------------------------------------------------
    # EBS processing
    # ------------------------------------------------------------------

    def _process_ebs(self, loaded_files) -> tuple:
        try:
            self.log_step(self.log, "EBS", "Started processing", 0)

            df_ebs_original = loaded_files[GROUPING_BY_UPLOAD_CONFIG.file_fields[1].label].copy()
            df_ebs = df_ebs_original.copy()
            self.log_step(self.log, "EBS", "Original file loaded", len(df_ebs_original))

            df_ebs = df_ebs[df_ebs["Grouping By"].notna() & (df_ebs["Grouping By"].str.strip() != "")]
            self.log_step(self.log, "EBS", "Filtered to rows with 'Grouping By'", len(df_ebs))

            df_ebs = df_ebs.drop_duplicates(subset=["X-Check No."], keep="first").reset_index(drop=True)
            self.log_step(self.log, "EBS", "Removed duplicate X-Check No. rows", len(df_ebs))

            split_cols = df_ebs["Grouping By"].str.split(",", expand=True)
            split_cols.columns = [f"Grouping By {i + 1}" for i in range(split_cols.shape[1])]
            split_cols = split_cols.apply(lambda col: col.str.strip())
            self.log_step(self.log, "EBS", f"Split 'Grouping By' into {split_cols.shape[1]} columns", split_cols.notna().sum().sum())

            col_position = df_ebs.columns.get_loc("Grouping By")
            df_ebs = df_ebs.drop(columns=["Grouping By"])
            for col in reversed(split_cols.columns.tolist()):
                df_ebs.insert(col_position, col, split_cols[col])
            self.log_step(self.log, "EBS", "Replaced 'Grouping By' with split columns", len(df_ebs))

            df_ebs["_base_key"] = df_ebs["Reference  X-Check (Condition)"].where(
                df_ebs["Reference  X-Check (Condition)"].notna() &
                (df_ebs["Reference  X-Check (Condition)"].str.strip() != ""),
                other=df_ebs["X-Check No."].astype(str).str.strip()
            )
            self.log_step(self.log, "EBS", "Constructed base key column", df_ebs["_base_key"].notna().sum())

            for col in [c for c in df_ebs.columns if c.startswith("Grouping By ")]:
                key_col = col.replace("Grouping By ", "Grouping By Key ")
                df_ebs[key_col] = df_ebs.apply(
                    lambda row, c=col: (
                        f"{row['_base_key']}|{str(row[c]).strip()}"
                        if pd.notna(row[c]) and str(row[c]).strip() != ""
                        else ""
                    ),
                    axis=1
                )
            self.log_step(self.log, "EBS", "Constructed 'Grouping By Key n' columns", len(df_ebs))

            df_ebs = df_ebs.drop(columns=["_base_key"])

            key_cols = [c for c in df_ebs.columns if c.startswith("Grouping By Key ")]
            index_cols = [c for c in df_ebs.columns if c not in key_cols]
            stacked = (
                df_ebs.set_index(index_cols)
                .stack()
                .reset_index()
                .rename(columns={0: "Key"})
            )
            level_col = f"level_{len(index_cols)}"
            if level_col in stacked.columns:
                stacked = stacked.drop(columns=[level_col])
            self.log_step(self.log, "EBS", "Stacked key columns into single 'Key' column", len(stacked))

            df_ebs = stacked[stacked["Key"].str.strip() != ""].reset_index(drop=True)
            self.log_step(self.log, "EBS", "Finished processing", len(df_ebs))
            return df_ebs_original, df_ebs

        except Exception as exc:
            import traceback
            self.log_step(self.log, "EBS", f"Exception: {exc}", 0)
            self.log_step(self.log, "EBS", traceback.format_exc(), 0)
            return None, None

    # ------------------------------------------------------------------
    # Compare
    # ------------------------------------------------------------------

    def _process_compare(self, df_fip: pd.DataFrame, df_ebs: pd.DataFrame) -> pd.DataFrame:
        self.log_step(self.log, "Compare", "Started comparison", 0)

        fip_keys = df_fip[["Key"]].drop_duplicates().copy()
        fip_keys["In FIP"] = True
        self.log_step(self.log, "Compare", "FIP key lookup built", len(fip_keys))

        ebs_keys = df_ebs[["Key"]].drop_duplicates().copy()
        self.log_step(self.log, "Compare", "EBS keys extracted", len(ebs_keys))

        df_compare = ebs_keys.merge(fip_keys, on="Key", how="left")
        df_compare["In FIP"] = df_compare["In FIP"].fillna(False)
        df_compare["Result"] = df_compare["In FIP"].map({True: "Matched", False: "Not in FIP"})

        matched     = df_compare["Result"].eq("Matched").sum()
        not_matched = df_compare["Result"].eq("Not in FIP").sum()
        self.log_step(self.log, "Compare", f"Matched: {matched} | Not in FIP: {not_matched}", len(df_compare))

        df_compare = (
            df_compare
            .drop(columns=["In FIP"])
            .rename(columns={"Key": "EBS Key"})
            .sort_values("EBS Key")
            .reset_index(drop=True)
        )
        self.log_step(self.log, "Compare", "Finished comparison", len(df_compare))
        return df_compare
