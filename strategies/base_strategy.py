import pandas as pd
import openpyxl
import os
import re
import shutil
import threading
from abc import ABC, abstractmethod
from tkinter import messagebox
from typing import Optional
from file_upload_config import UploadTaskConfig
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
        self._progress_dialog = None          # Set via set_progress_dialog()
        self._stop_event: threading.Event | None = None    

    def execute(self, files: dict):
        """
        Entry point called by main.py.
        Loads all files then hands off to the subclass.
        """
        from exceptions import FileLoadError, SheetNotFoundError, MissingColumnsError, UnsupportedFileTypeError

        self.log = []  # Initialised here — available to all strategies via self.log
        self.process_only_differences = files.get("process_only_differences", False)

        # Store process flag on instance so _load_files can access it
        self.process_only_differences = files.get("process_only_differences", False)

        try:
            self.log_step(self.log, "System", "Loading files into memory...", 0)
            try:
                loaded_files = self._load_files(files["files"], files["sheet_names"], self.config.file_fields)
            except (FileLoadError, SheetNotFoundError, MissingColumnsError, UnsupportedFileTypeError) as e:
                self.log_step(self.log, "System", f"Error loading files: {e}", 0)
                return

            if loaded_files is None:
                return

            self.log_step(self.log, "System", "Files loaded successfully:", len(loaded_files))
            for label, data in loaded_files.items():
                self.log_step(self.log, "    " + label, f"Loaded {type(data).__name__}", len(data))

            self.process(loaded_files, files)

        except StopIteration:
            # User pressed Stop — log it, then return cleanly
            timestamp = __import__("datetime").datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            print(f"  [{timestamp}] System — Processing halted by user.")
            self.log.append({
                "Timestamp": timestamp,
                "File": "System",
                "Step": "Processing halted by user.",
                "Count": 0,
                "Notes": ""
            })
            if self._progress_dialog is not None:
                self._progress_dialog.append_entry("System", "Processing halted by user. You may now close this window.")

#        self.log_step(self.log, "System", "Loading files into memory...", 0)
#        try:
#            loaded_files = self._load_files(files["files"], files["sheet_names"], self.config.file_fields)
#
#        except (FileLoadError, SheetNotFoundError, MissingColumnsError, UnsupportedFileTypeError) as e:
#            self.log_step(self.log, "System", f"Error loading files: {e}", 0)
#            messagebox.showerror("File Loading Error", str(e))
#            return
#
#        if loaded_files is None:
#            return
#
#        self.log_step(self.log, "System", "Files loaded successfully:", len(loaded_files))
#        for label, data in loaded_files.items():
#            self.log_step(self.log, "    " + label, f"Loaded {type(data).__name__}", len(data))
#
#        self.process(loaded_files, files)

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

    def set_progress_dialog(self, dialog):
        """
        Called by main.py (debug mode only) to attach the ProgressDialog.
        Also stores a reference to its stop event for checkpoint polling.
        """
        self._progress_dialog = dialog
        self._stop_event = dialog.stop_event

    def log_step(self, log: list, file: str, step: str, count: int, notes: str = ""):
        """
        Appends a timestamped entry to the processing log and prints to console.
        If a ProgressDialog is attached, pushes the entry to it.
        If the stop event has been set, raises StopIteration to unwind processing.
        The stop check happens AFTER the current step is recorded, so the step
        that was already running completes before the halt takes effect.
        """
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        print(f"  [{timestamp}] {file} — {step} ({count})")
        log.append({
            "Timestamp": timestamp,
            "File": file,
            "Step": step,
            "Count": count,
            "Notes": notes
        })

        # Push to UI dialog if attached
        if self._progress_dialog is not None:
            self._progress_dialog.append_entry(file, step, count, notes)

        # Check stop event AFTER completing this step
        if self._stop_event is not None and self._stop_event.is_set():
            raise StopIteration("Processing stopped by user.")

    def autofit_columns(self, worksheet, max_width: int = 90,  skip_rows: int = 0):
        """Auto-fits all columns in a worksheet to their content width, capped at max_width."""
        for column in worksheet.columns:
            max_length = 0
            col_letter = column[0].column_letter
            for cell in column:
                if cell.row <= skip_rows:
                    continue
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            worksheet.column_dimensions[col_letter].width = min(max_length + 2, max_width)

    def write_excel_output(self, output_path: str, sheets: dict, log: list, summaries: dict | None = None):
        """
        Writes a dictionary of DataFrames to a single timestamped Excel workbook.
        Copies the pre-labelled template, writes all sheets, auto-fits columns.

        Args:
            output_path:  Full path to the output file
            sheets:       Ordered dict of {sheet_name: DataFrame}
            log:          List of log entry dicts
            summaries:    Optional dict of {label: value} for summary blocks
        """
        # Copy pre-labelled template to output path
        shutil.copy(OUTPUT_TEMPLATE, output_path)
        df_log = pd.DataFrame(log)

        with pd.ExcelWriter(output_path, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:

            # Write all provided sheets
            for sheet_name, df in sheets.items():
                if summaries and sheet_name in summaries:
                    self.write_sheet_with_summary(writer, sheet_name, df, summaries[sheet_name])
                else:
                    df.to_excel(writer, sheet_name=sheet_name, index=False)

            # Write log sheet
            df_log.to_excel(writer, sheet_name="Processing Log", index=False)

            # Remove default Sheet1 from template
            if "Sheet1" in writer.book.sheetnames:
                del writer.book["Sheet1"]

            # Auto-fit all sheets
            for sheet_name in writer.book.sheetnames:
                if summaries and sheet_name in summaries:
                    skip = len(summaries[sheet_name]) + 2  # skip summary rows + blank row
                    self.autofit_columns(writer.book[sheet_name], skip_rows=skip)
                else:
                    self.autofit_columns(writer.book[sheet_name])

            self.log_step(self.log, "Output", f"  Sheets in workbook: {writer.book.sheetnames}", 0)

            # Allow subclasses to apply strategy-specific formatting
            self.apply_output_formatting(writer.book)

        self.log_step(self.log, "Output", f"Output written to: {output_path}", 0)
    
    def apply_output_formatting(self, workbook):
        """
        Hook for subclasses to apply strategy-specific formatting.
        Override in subclass — default does nothing.
        """
        pass

    # -------------------------------------------------------------------------
    # File loading utilities — available to all strategies
    # -------------------------------------------------------------------------

    def _load_files(self, files: dict, sheet_names: dict, file_fields: list) -> Optional[dict]:
        """
        Reads each file into memory based on its extension.
        """
        from exceptions import FileLoadError, SheetNotFoundError, MissingColumnsError, UnsupportedFileTypeError

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
                        raise FileLoadError(
                            f"Could not open '{os.path.basename(path)}' as an Excel file.\n\n"
                            f"{str(e)}"
                        )

                    if sheet not in excel_file.sheet_names:
                        raise SheetNotFoundError(
                            f"Could not find sheet '{sheet}' in '{os.path.basename(path)}'.\n\n"
                            f"Please check the file and sheet name then try again."
                        )

                    df = pd.read_excel(path, sheet_name=sheet)
                    df = self._select_columns(df, label, column_map.get(label), path)
                    if df is None:
                        return None

                    if self.process_only_differences:
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
                    raise UnsupportedFileTypeError(
                        f"'{os.path.basename(path)}' is not a supported file type.\n\n"
                        f"Please select a csv or txt file."
                    )

            except PermissionError:
                raise FileLoadError(
                    f"'{os.path.basename(path)}' is currently open in another application.\n\n"
                    f"Please close it and try again."
                )

            except ValueError as e:
                # Catch unexpected pandas errors (e.g. malformed CSV, encoding issues)
                raise FileLoadError(
                    f"Error reading '{os.path.basename(path)}':\n\n{str(e)}"
                )
            
        return loaded

    def _filter_coloured_rows(self, df: pd.DataFrame, filepath: str, sheet_name: str) -> pd.DataFrame:
        """
        Uses openpyxl to detect rows with any background fill colour,
        then filters the DataFrame to those rows only.

        NOTE: This loads the workbook a second time (first load is via pd.read_excel).
        This trade-off keeps data loading and colour detection cleanly separated.
        If performance becomes an issue with large files, consider loading openpyxl
        first and extracting both data and colours in a single pass.
        """
        ext = os.path.splitext(filepath)[1].lower()
        if ext != ".xlsx":
            #print(f"  [process_only_differences] Skipping colour filter for {os.path.basename(filepath)} — .xls not supported by openpyxl")
            self.log_step(self.log, os.path.basename(filepath), "Colour filter skipped — .xls not supported by openpyxl", 0)
            return df

        try:
            wb = openpyxl.load_workbook(filepath, data_only=True, read_only=False)
            ws = wb[sheet_name]
        except Exception as e:
            # Non-fatal: fall back to full dataset if colour detection fails
            self.log_step(self.log, os.path.basename(filepath),
                        "Colour filter failed — using full dataset", len(df),
                        notes=str(e))
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
            self.log_step(self.log, os.path.basename(filepath), "No coloured rows found — returning empty DataFrame", 0)
            return df.iloc[[]]

        filtered = df.iloc[coloured_row_indices].reset_index(drop=True)
        self.log_step(self.log, os.path.basename(filepath),
                "Coloured rows retained", len(filtered),
                notes=f"Filtered from {len(df)} total rows") 
        return filtered

    def _select_columns(self, df: pd.DataFrame, label: str, required_columns: Optional[list[str]], filepath: str) -> Optional[pd.DataFrame]:
        """
        Reduces the DataFrame to only the required columns.
        Returns None if any required column is missing.
        """
        from exceptions import MissingColumnsError

        if required_columns is None:
            return df

        missing = [col for col in required_columns if col not in df.columns]
        if missing:
            raise MissingColumnsError(
                f"'{os.path.basename(filepath)}' is missing the following required column(s):\n\n"
                + "\n".join(f"  • {col}" for col in missing)
                + f"\n\nPlease check the file and try again."
            )
            
        return df[required_columns]
    
    def apply_conditional_formatting(self, worksheet, column_name: str, rules: dict):
        """
        Applies conditional formatting to a named column in a worksheet.

        Args:
            worksheet:    The openpyxl worksheet object
            column_name:  The header name of the column to format
            rules:        Dict of {cell_value: (PatternFill, Font)} or {cell_value: PatternFill}
        """
        from openpyxl.formatting.rule import CellIsRule

        self.log_step(self.log, "Formatting",
                    f"Applying to '{worksheet.title}', column '{column_name}'", 0)

        # Find the column letter by scanning all rows for the header
        target_col = None
        header_row = None
        for row in worksheet.iter_rows():
            for cell in row:
                if cell.value == column_name:
                    target_col = cell.column_letter
                    header_row = cell.row
                    break
            if target_col:
                break

        if target_col is None:
            self.log_step(self.log, "Formatting",
                        f"Column '{column_name}' not found — skipping", 0)
            return

        last_row = worksheet.max_row
        data_start = header_row + 1
        target_range = f"{target_col}{data_start}:{target_col}{last_row}"

        for value, formatting in rules.items():
            # Accept either a plain fill or a (fill, font) tuple
            if isinstance(formatting, tuple):
                fill, font = formatting
            else:
                fill, font = formatting, None

            worksheet.conditional_formatting.add(
                target_range,
                CellIsRule(
                    operator="equal",
                    formula=[f'"{value}"'],
                    fill=fill,
                    font=font
                )
            )

        self.log_step(self.log, "Formatting",
                    f"Applied {len(rules)} rule(s) to {target_range}", 0)

    def write_sheet_with_summary(self, writer, sheet_name: str, df: pd.DataFrame, summary: dict):
        """
        Writes a summary block followed by a blank row then the main DataFrame.

        Args:
            writer:       The pd.ExcelWriter instance
            sheet_name:   Name of the sheet to write to
            df:           The main DataFrame
            summary:      Ordered dict of {label: value} for the summary block
        """
        # Write main DataFrame first so the sheet is created, leaving room for summary
        start_row = len(summary) + 2  # +1 for blank row, +1 for header
        df.to_excel(writer, sheet_name=sheet_name, index=False, startrow=start_row)

        # Write summary rows directly into cells
        worksheet = writer.sheets[sheet_name]
        for i, (label, value) in enumerate(summary.items()):
            worksheet.cell(row=i + 1, column=1, value=label)
            worksheet.cell(row=i + 1, column=2, value=value)

    @abstractmethod
    def process(self, loaded_files: dict, files: dict):
        """
        Subclasses implement THIS — not execute().
        By the time this is called, all files are already in memory.
        Access output_directory via files["output_directory"].
        Access timestamp via files["timestamp"].        
        """
        pass
