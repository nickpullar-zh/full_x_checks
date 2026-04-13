import os
import tkinter as tk
from tkinter import ttk
import pandas as pd
from strategies.base_strategy import BaseStrategy

class XChecks(BaseStrategy):

    def process(self, loaded_files: dict, files: dict, output_directory: str):
        lines = ["=== X-Checks: Processing ===\n"]

        for label, data in loaded_files.items():
            filename = os.path.basename(files["files"][label])
            if isinstance(data, pd.DataFrame):
                lines.append(f"{label}:")
                lines.append(f"  File    : {filename}")
                lines.append(f"  Rows    : {len(data)}")
                lines.append(f"  Columns : {len(data.columns)}")
                lines.append(f"  Columns : {list(data.columns)}\n")
            else:
                lines.append(f"{label}:")
                lines.append(f"  File    : {filename}")
                lines.append(f"  Content : loaded OK (text file)\n")

        output = "\n".join(lines)
        print(output)
        self._show_output(output)

    def _show_output(self, output: str):
        """Display output in a resizable window with a Text widget."""
        window = tk.Toplevel()
        window.title("X-Checks — Processing Summary")
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