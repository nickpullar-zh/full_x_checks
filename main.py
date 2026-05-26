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
        "FIP file":                  os.path.join(_BASE, "test_data", "20260318 FIP X-Checks.txt"),
        "X-Checks Publication File": os.path.join(_BASE, "test_data", "20260313 Cross Checks All.xlsx"),
    },
    "sheet_names": {
        "X-Checks Publication File": "cross checks all",
    },
    "output_directory": os.path.join(os.path.expanduser("~"), "Downloads", "Output"),
    "process_only_differences":   False,
    "apply_version_spanning":     False,
    "apply_prior_year_balance":   False,
}

# Maps each task name to its debug file config — build.py patches DEBUG_TASK per build
DEBUG_TASK = "X-Checks Grouping By"  # ← patched by build.py; must match a key in TASK_REGISTRY
_DEBUG_FILES_MAP = {
    "X-Checks Grouping By": DEBUG_FILES,
    "X-Checks":             DEBUG_FILES_X_CHECKS,
}


def _register_fonts():
    """Register Zurich brand fonts with the OS so tkinter can use them."""
    import ctypes
    fonts_dir = os.path.join(_get_base_path(), "templates", "fonts")
    if not os.path.isdir(fonts_dir):
        return
    for fname in os.listdir(fonts_dir):
        if fname.lower().endswith(".ttf"):
            ctypes.windll.gdi32.AddFontResourceExW(
                os.path.join(fonts_dir, fname), 0x10, 0  # FR_PRIVATE
            )


class TaskSelectorUI:
    """
    Simple launcher window — user picks a task,
    then the appropriate upload dialog is shown.
    """

    def __init__(self):
        _register_fonts()
        self.root = tk.Tk()
        # Apply Zurich Sans as the default font for all widgets
        import tkinter.font as tkfont
        for fname in ("TkDefaultFont", "TkTextFont", "TkMenuFont", "TkHeadingFont",
                      "TkCaptionFont", "TkSmallCaptionFont", "TkIconFont", "TkTooltipFont"):
            try:
                tkfont.nametofont(fname).configure(family="Zurich Sans", size=10)
            except Exception:
                pass
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
            font=("Zurich Sans Semibold", 16)
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

        if DEBUG_MODE:
            self._run_debug(config, strategy_class)
        else:
            self._run_task_loop(config, strategy_class, prefill=None)

    def _run_task_loop(self, config, strategy_class, prefill=None):
        """Show the file-upload form (optionally pre-filled) then start processing.
        On cancel, returns to the task selector. On stop/error, the ProgressDialog
        calls back here with the previous inputs pre-filled."""
        self.root.withdraw()
        files = FileUploadUI(config, self.root, prefill=prefill).run()

        if not files:
            # User cancelled the form — show the task selector again
            self.root.deiconify()
            return

        self._run_task(config, strategy_class, files)

    def _run_task(self, config, strategy_class, files):
        """Create the progress dialog and run the processing thread.
        On cancel or error the dialog calls back to _run_task_loop with the
        same inputs pre-filled so the user can adjust and re-run."""
        from progress_dialog import ProgressDialog
        import threading

        def run_processing(strategy, dialog, files):
            try:
                strategy.execute(files)
                if not dialog.is_stopped():
                    dialog.mark_success()
                    dialog.append_entry("System", "Processing complete. You may close this window.")
                    self.root.after(0, lambda: dialog.action_btn.config(text="Close"))
                    self.root.after(0, lambda: setattr(dialog, "_stopped", True))
            except Exception as e:
                import traceback
                error_msg = traceback.format_exc()
                print(f"  [ERROR] Unhandled exception in processing thread:\n{error_msg}")
                dialog.append_entry("ERROR", f"Unhandled exception: {e}")
                self.root.after(0, lambda: dialog.action_btn.config(text="Return to Form"))
                self.root.after(0, lambda: setattr(dialog, "_stopped", True))

        dialog = ProgressDialog(
            self.root,
            on_dismiss=lambda: self._run_task_loop(config, strategy_class, prefill=files),
        )
        strategy = strategy_class(config)
        strategy.set_progress_dialog(dialog)

        thread = threading.Thread(target=run_processing, args=(strategy, dialog, files), daemon=True)
        thread.start()

    def _run_debug(self, config, strategy_class):
        """Debug mode: skip the UI form and run directly with hardcoded files.
        Retains original exit-on-close behaviour."""
        from progress_dialog import ProgressDialog
        from datetime import datetime
        import threading

        debug_files = _DEBUG_FILES_MAP.get(DEBUG_TASK, DEBUG_FILES)

        required_labels = {f.label for f in config.file_fields if f.required}
        provided_labels = set(debug_files["files"].keys())
        missing = required_labels - provided_labels
        extra   = provided_labels - {f.label for f in config.file_fields}

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

        debug_files["timestamp"] = datetime.now().strftime("%Y%m%d_%H%M%S")
        os.makedirs(debug_files["output_directory"], exist_ok=True)

        def run_processing(strategy, dialog, files):
            try:
                strategy.execute(files)
                if not dialog.is_stopped():
                    dialog.mark_success()
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

        strategy = strategy_class(config)
        dialog = ProgressDialog(self.root)   # no on_dismiss — exit on close
        self.root.withdraw()
        strategy.set_progress_dialog(dialog)

        thread = threading.Thread(target=run_processing, args=(strategy, dialog, debug_files), daemon=True)
        thread.start()

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