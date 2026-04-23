from abc import ABC, abstractmethod
from tkinter import messagebox
from typing import Optional
from file_upload_config import UploadTaskConfig
import pandas as pd
import openpyxl
import os
import re
import shutil
import app_state
from datetime import datetime
from config import OUTPUT_TEMPLATE


class BaseStrategy(ABC):
    """
    Handles everything that is ALWAYS the same:
    - Loading files into memory
    - Calling the use-case-specific processing
    - Writing Excel output
    - Logging processing steps
    """
    
    def __init__(self, config: UploadTaskConfig):
        self.config = config

    def execute(self, files: dict, output_directory: str):
        """
        Entry point called by main.py.
        Loads all files then hands off to the subclass.
        """
        print("Loading files into memory...")
        loaded_files = self._load_files(files["files"], files["sheet_names"], self.config.file_fields)

        if loaded_files is None:
            return

        print("Files loaded successfully:")
        for label, data in loaded_files.items():
            print(f"  {label}: {type(data)}")

        self.process(loaded_files, files, output_directory)

    # -------------------------------------------------------------------------
    # Excel output utilities — available to all strategies
    # -------------------------------------------------------------------------

    def build_output_path(self, output_directory: str, label: str, timestamp: str) -> str:
        """
        Builds a safe, timestamped output file path from a label.
        Replaces characters that are invalid in Windows filenames.
        """
        safe_label = re.sub(r'[<>:"/\\|?*()]', '_', label)
        filename = f"{timestamp}_{safe_label}.xlsx"
        return os.path.join(output_directory, filename)

    def log_step(self, log: list, file: str, step: str, rows: int, notes: str = ""):
        """Appends a timestamped entry to the processing log."""
        log.append({
            "Timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "File": file,
            "Step": step,
            "Rows": rows,
            "Notes": notes
        })

    def autofit_columns(self, worksheet, max_width: int = 90):
        """Auto-fits all columns in a worksheet to their content width, capped at max_width."""
        for column in worksheet.columns:
            max_length = 0
            col_letter = column[0].column_letter
            for cell in column:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            worksheet.column_dimensions[col_letter].width = min(max_length + 2, max_width)

    def write_excel_output(self, output_path: str, sheets: dict, log: list, df_compare: pd.DataFrame|None = None):
        """
        Writes a dictionary of DataFrames to a single timestamped Excel workbook.
        Copies the pre-labelled template, writes all sheets, auto-fits columns.

        Args:
            output_path:  Full path to the output file
            sheets:       Ordered dict of {sheet_name: DataFrame}
            log:          List of log entry dicts
            df_compare:   Optional compare DataFrame — written to "Compare" sheet if provided
        """
        # Copy pre-labelled template to output path
        shutil.copy(OUTPUT_TEMPLATE, output_path)

        df_log = pd.DataFrame(log)

        with pd.ExcelWriter(output_path, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:

            # Write all provided sheets
            for sheet_name, df in sheets.items():
                df.to_excel(writer, sheet_name=sheet_name, index=False)

            # Write compare sheet if provided
            if df_compare is not None:
                df_compare.to_excel(writer, sheet_name="Compare", index=False)

            # Write log sheet
            df_log.to_excel(writer, sheet_name="Processing Log", index=False)

            # Remove default Sheet1 from template
            if "Sheet1" in writer.book.sheetnames:
                del writer.book["Sheet1"]

            # Auto-fit all sheets
            for sheet_name in writer.book.sheetnames:
                self.autofit_columns(writer.book[sheet_name])

        print(f"Output written to: {output_path}")

    # -------------------------------------------------------------------------
    # File loading utilities — available to all strategies
    # -------------------------------------------------------------------------

    def _load_files(self, files: dict, sheet_names: dict, file_fields: list) -> Optional[dict]:
        """
        Reads each file into memory based on its extension.
        """
        loaded = {}

        column_map = {
            f.label: f.required_columns
            for f in file_fields
        }

        for label, path in files.items():
            if path is None:
                continue
            try:
                ext = os.path.splitext(path)[1].lower()

                if ext in (".xlsx", ".xls"):
                    sheet = sheet_names.get(label, "Sheet1")
                    try:
                        excel_file = pd.ExcelFile(path)
                    except Exception as e:
                        messagebox.showerror(
                            "Excel File Error",
                            f"Could not open '{os.path.basename(path)}' as an Excel file.\n\n"
                            f"{str(e)}"
                        )
                        return None

                    if sheet not in excel_file.sheet_names:
                        messagebox.showerror(
                            "Sheet Not Found",
                            f"Could not find sheet '{sheet}' in '{os.path.basename(path)}'.\n\n"
                            f"Please check the file and sheet name then try again."
                        )
                        return None

                    df = pd.read_excel(path, sheet_name=sheet)
                    df = self._select_columns(df, label, column_map.get(label), path)
                    if df is None:
                        return None

                    if app_state.process_only_differences:
                        df = self._filter_coloured_rows(df, path, sheet)

                    loaded[label] = df

                elif ext == ".csv":
                    df = pd.read_csv(path)
                    df = self._select_columns(df, label, column_map.get(label), path)
                    if df is None:
                        return None
                    loaded[label] = df

                elif ext == ".txt":
                    with open(path, "r") as f:
                        loaded[label] = f.read()
                else:
                    raise ValueError(f"Unsupported file type for: {path}")

            except PermissionError:
                messagebox.showerror(
                    "File In Use",
                    f"'{os.path.basename(path)}' is currently open in another application.\n\n"
                    f"Please close it and try again.\n\n"
                    f"The application will now stop."
                )
                return None

            except ValueError as e:
                if "Worksheet" in str(e) or "sheet" in str(e).lower():
                    messagebox.showerror(
                        "Sheet Not Found",
                        f"Could not find sheet '{sheet_names.get(label)}' "
                        f"in '{os.path.basename(path)}'.\n\n"
                        f"Please check the sheet name and try again."
                    )
                else:
                    messagebox.showerror(
                        "Unsupported File Type",
                        f"'{os.path.basename(path)}' is not a supported file type.\n\n"
                        f"Please select an xlsx, csv or txt file."
                    )
                    return None

        return loaded

    def _filter_coloured_rows(self, df: pd.DataFrame, filepath: str, sheet_name: str) -> pd.DataFrame:
        """
        Uses openpyxl to detect rows with any background fill colour,
        then filters the DataFrame to those rows only.
        """
        ext = os.path.splitext(filepath)[1].lower()
        if ext != ".xlsx":
            print(f"  [process_only_differences] Skipping colour filter for {os.path.basename(filepath)} — .xls not supported by openpyxl")
            return df

        try:
            wb = openpyxl.load_workbook(filepath, data_only=True, read_only=False)
            ws = wb[sheet_name]
        except Exception as e:
            messagebox.showerror(
                "Colour Filter Error",
                f"Could not read cell colours from '{os.path.basename(filepath)}'.\n\n"
                f"{str(e)}\n\n"
                f"The full dataset will be used instead."
            )
            return df

        coloured_row_indices = []

        for row in ws.iter_rows(min_row=2):
            for cell in row:
                fill = cell.fill
                if (
                    fill.fill_type not in (None, "none")
                    and fill.fgColor.type != "auto"
                ):
                    if cell.row is not None:
                        coloured_row_indices.append(cell.row - 2)
                    break

        wb.close()

        if not coloured_row_indices:
            print(f"  [process_only_differences] No coloured rows found in {os.path.basename(filepath)} — returning empty DataFrame")
            return df.iloc[[]]

        filtered = df.iloc[coloured_row_indices].reset_index(drop=True)
        print(f"  [process_only_differences] {len(filtered)} coloured row(s) retained from {os.path.basename(filepath)}")
        return filtered

    def _select_columns(self, df: pd.DataFrame, label: str, required_columns: Optional[list[str]], filepath: str) -> Optional[pd.DataFrame]:
        """
        Reduces the DataFrame to only the required columns.
        Returns None if any required column is missing.
        """
        if required_columns is None:
            return df

        missing = [col for col in required_columns if col not in df.columns]
        if missing:
            messagebox.showerror(
                "Missing Columns",
                f"'{os.path.basename(filepath)}' is missing the following required column(s):\n\n"
                + "\n".join(f"  • {col}" for col in missing)
                + f"\n\nPlease check the file and try again."
            )
            return None

        return df[required_columns]

    @abstractmethod
    def process(self, loaded_files: dict, files: dict, output_directory: str):
        """
        Subclasses implement THIS — not execute().
        By the time this is called, all files are already in memory.
        """
        pass