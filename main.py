import sys
import tkinter as tk
import os
from tkinter import ttk
from file_upload_ui import FileUploadUI
from task_registry import TASK_REGISTRY
from version import __version__

# ==========================================
# Debug Configuration
# ← Set to True to skip UI during debugging
# ← Set to False for normal UI operation
# ==========================================
DEBUG_MODE = False
# Clears terminal on every run (cross-platform)
os.system('cls' if os.name == 'nt' else 'clear')

def _get_base_path() -> str:
    """Returns the base path for test data — works both in dev and bundled exe."""
    if getattr(sys, 'frozen', False):
        # Running as bundled exe
        return sys._MEIPASS
    else:
        # Running in development
        return os.path.dirname(os.path.abspath(__file__))

_BASE = _get_base_path()

DEBUG_FILES = {
    "files": {
        "FIP File (ZQ9_VALFLDGR)":    os.path.join(_BASE, "test_data", "VALFLDGR file with 12348 Data rows on sheet Sheet1.XLSX"),
        "X-Checks Publication File":   os.path.join(_BASE, "test_data", "EPM X-Checks file with 3345 Data rows on sheet cross checks all.xlsx"),
        "Mapping File":                os.path.join(_BASE, "test_data", "Mapping Table with 20 rows.txt"),
    },
    "sheet_names": {
        "FIP File (ZQ9_VALFLDGR)":    "Sheet1",
        "X-Checks Publication File":   "cross checks all",
        "Mapping File":                "Sheet1",
    },
    "output_directory": os.path.join(os.path.expanduser("~"), "Downloads", "Output"),
    "process_only_differences": False
}

DEBUG_FILES_X_CHECKS = {
    "files": {
        "FIP file":                  os.path.join(_BASE, "test_data", "<YOUR_FIP_FILE.txt>"),
        "X-Checks Publication File": os.path.join(_BASE, "test_data", "EPM X-Checks file with 3345 Data rows on sheet cross checks all.xlsx"),
    },
    "sheet_names": {
        "X-Checks Publication File": "cross checks all",
    },
    "output_directory": os.path.join(os.path.expanduser("~"), "Downloads", "Output"),
    "process_only_differences":   False,
    "apply_version_spanning":     False,
    "apply_prior_year_balance":   False,
}

DEBUG_TASK = "X-Checks Grouping By"  # ← Must match a key in TASK_REGISTRY

class TaskSelectorUI:
    """
    Simple launcher window — user picks a task,
    then the appropriate upload dialog is shown.
    """

    def __init__(self):
        self.root = tk.Tk()
        self.root.title(f"X-Check Application v{__version__}")
        self.root.resizable(False, False)
        self.root.config(cursor="watch")   # busy cursor while UI builds
        self._build_ui()
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self.root.update_idletasks()
        self.root.config(cursor="")        # restore normal cursor
        # Close the PyInstaller splash screen once the window is ready
        try:
            import pyi_splash
            pyi_splash.close()
        except ImportError:
            pass

    def _on_close(self):
        """
        Called when the user clicks the red X on the launcher.
        Exits the application cleanly.
        """
        self.root.destroy()
        sys.exit()

    def _on_task_selected(self, *args):
        """
        Called whenever the dropdown selection changes.
        Enables the Start button only when a task is selected.
        """
        if self.task_var.get():
            self.start_btn.config(state="normal")
        else:
            self.start_btn.config(state="disabled")

    def _build_ui(self):
        frame = ttk.Frame(self.root, padding="20")
        frame.grid(row=0, column=0)

        ttk.Label(
            frame,
            text=f"X-Check Application v{__version__}",
            font=("Helvetica", 16, "bold")
        ).grid(row=0, column=0, columnspan=2, pady=(0, 20))

        ttk.Label(
            frame,
            text="Select a task to begin:"
        ).grid(row=1, column=0, sticky="w", pady=5)

        self.task_var = tk.StringVar()
        self.task_var.trace_add("write", self._on_task_selected)
        task_dropdown = ttk.Combobox(
            frame,
            textvariable=self.task_var,
            values=list(TASK_REGISTRY.keys()),
            state="readonly",
            width=30
        )
        task_dropdown.grid(row=1, column=1, padx=10, pady=5)

        self.start_btn = ttk.Button(
            frame,
            text="Start",
            command=self._on_start
        )
        self.start_btn.grid(row=2, column=0, columnspan=2, pady=15)
        self.start_btn.config(state="disabled")  # ← Disabled by default

    def _on_start(self):
        task_name = self.task_var.get()
        if not task_name:
            return

        config, strategy_class = TASK_REGISTRY[task_name]

        from progress_dialog import ProgressDialog
        import threading

        def run_processing(strategy, dialog, files):
            try:
                strategy.execute(files)
                if not dialog.is_stopped():
                    dialog.append_entry("System", "Processing complete. You may close this window.")
                    self.root.after(0, lambda: dialog.action_btn.config(text="Close"))
                    self.root.after(0, lambda: setattr(dialog, "_stopped", True))
            except Exception as e:
                import traceback
                error_msg = traceback.format_exc()
                print(f"  [ERROR] Unhandled exception in processing thread:\n{error_msg}")
                dialog.append_entry("ERROR", f"Unhandled exception: {e}")
                self.root.after(0, lambda: dialog.action_btn.config(text="Close"))
                self.root.after(0, lambda: setattr(dialog, "_stopped", True))

        if DEBUG_MODE:
            from datetime import datetime

            required_labels = {f.label for f in config.file_fields if f.required}
            provided_labels = set(DEBUG_FILES["files"].keys())
            missing = required_labels - provided_labels
            extra = provided_labels - {f.label for f in config.file_fields}

            if missing:
                print(f"  [ERROR] DEBUG_FILES is missing required file(s):")
                for label in missing:
                    print(f"    - {label}")
                self.root.destroy()
                return

            if extra:
                print(f"  [WARNING] DEBUG_FILES contains unrecognised file(s):")
                for label in extra:
                    print(f"    - {label}")

            DEBUG_FILES["timestamp"] = datetime.now().strftime("%Y%m%d_%H%M%S")
            os.makedirs(DEBUG_FILES["output_directory"], exist_ok=True)

            strategy = strategy_class(config)
            dialog = ProgressDialog(self.root)
            self.root.withdraw()
            strategy.set_progress_dialog(dialog)

            thread = threading.Thread(target=run_processing, args=(strategy, dialog, DEBUG_FILES), daemon=True)
            thread.start()
            return

        # --- Normal (non-debug) mode ---
        self.root.withdraw()
        files = FileUploadUI(config, self.root).run()

        if files:
            strategy = strategy_class(config)
            dialog = ProgressDialog(self.root)
            self.root.withdraw()
            strategy.set_progress_dialog(dialog)

            thread = threading.Thread(target=run_processing, args=(strategy, dialog, files), daemon=True)
            thread.start()
        else:
            # User closed the upload dialog without submitting
            self.root.destroy()

    def run(self):
        if DEBUG_MODE:
            # Hide the launcher window in debug mode
            #self.root.withdraw()
            # Set task directly and start without showing launcher
            self.task_var.set(DEBUG_TASK)
            self.root.after(0, self._on_start)  # ← Call _on_start immediately after mainloop starts

        self.root.mainloop()

if __name__ == "__main__":
    TaskSelectorUI().run()