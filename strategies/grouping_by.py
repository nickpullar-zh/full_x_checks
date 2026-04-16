import os
import tkinter as tk
from tkinter import ttk
import pandas as pd
import re
from strategies.base_strategy import BaseStrategy
from task_configs import GROUPING_BY_UPLOAD_CONFIG
from datetime import datetime

class GroupingBy(BaseStrategy):

    def process(self, loaded_files: dict, files: dict, output_directory: str):

        print(loaded_files.keys()) 
    #     lines = ["=== Grouping By: Processing ===\n"]

    #     for label, data in loaded_files.items():
    #         filename = os.path.basename(files["files"][label])
    #         if isinstance(data, pd.DataFrame):
    #             lines.append(f"{label}:")
    #             lines.append(f"  File    : {filename}")
    #             lines.append(f"  Rows    : {len(data)}")
    #             lines.append(f"  Columns : {len(data.columns)}")
    #             lines.append(f"  Columns : {list(data.columns)}\n")
    #         else:
    #             lines.append(f"{label}:")
    #             lines.append(f"  File    : {filename}")
    #             lines.append(f"  Content : loaded OK (text file)\n")

    #     output = "\n".join(lines)
    #     print(output)
    #     self._show_output(output)

        # Turn the Mapping text file into a dictionary
        mapping_file_content = loaded_files[GROUPING_BY_UPLOAD_CONFIG.file_fields[2].label]  # Mapping File
        mapping_dict = {}
        for line in mapping_file_content.splitlines()[1:]:  # [1:] skips the header row "FIP Data,EBS item"
            line = line.strip()
            if not line:
                continue  # Skip empty lines
            parts = line.split(",", maxsplit=1)  # maxsplit=1 protects against commas in values
            if len(parts) == 2:
                key, value = parts[0].strip(), parts[1].strip()
                mapping_dict[key] = value

        # Access directly from loaded_files
        df_fip = loaded_files[GROUPING_BY_UPLOAD_CONFIG.file_fields[0].label]  # 'FIP File (ZQ9_VALFLDGR)'
        
        # --- Lookup: map "Field name" to EBS item ---
        df_fip["EBS Item"] = df_fip["Field name"].map(mapping_dict)

        # --- Build Key column ---
        df_fip["Key"] = df_fip.apply(
            lambda row: f"{row['ValidRule']}|{row['EBS Item']}"
            if pd.notna(row["ValidRule"]) and str(row["ValidRule"]).strip() != ""
            else "",
            axis=1
        )

        # --- Generate timestamped filename ---
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        safe_label = re.sub(r'[<>:"/\\|?*()]', '_', GROUPING_BY_UPLOAD_CONFIG.file_fields[0].label)  # "FIP File (ZQ9_VALFLDGR)" Replace problematic characters
        output_filename = f"{timestamp}_{safe_label}.xlsx"
        output_path = os.path.join(output_directory, output_filename)

        # --- Write to Excel ---
        df_fip.to_excel(output_path, index=False)

    def _show_output(self, output: str):
        """Display output in a resizable window with a Text widget."""
        window = tk.Toplevel()
        window.title("Grouping By — Processing Summary")
        window.resizable(True, True)

        # --- Text widget with scrollbar ---
        frame = ttk.Frame(window, padding="10")
        frame.grid(row=0, column=0, sticky="nsew")

        scrollbar = ttk.Scrollbar(frame, orient="vertical")
        scrollbar.grid(row=0, column=1, sticky="ns")

        text = tk.Text(
            frame,
            width=120,
            height=20,
            font=("Courier New", 10),
            yscrollcommand=scrollbar.set,
            wrap="none"
        )
        text.grid(row=0, column=0, sticky="nsew")
        scrollbar.config(command=text.yview)

        # --- Horizontal scrollbar ---
        h_scrollbar = ttk.Scrollbar(frame, orient="horizontal", command=text.xview)
        h_scrollbar.grid(row=1, column=0, sticky="ew")
        text.config(xscrollcommand=h_scrollbar.set)

        # --- Insert output and lock editing ---
        text.insert("1.0", output)
        text.config(state="disabled")

        # --- OK button ---
        ttk.Button(
            window,
            text="OK",
            command=window.destroy
        ).grid(row=1, column=0, pady=10)

        # --- Make window and frame resizable ---
        window.columnconfigure(0, weight=1)
        window.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(0, weight=1)

        window.wait_window()