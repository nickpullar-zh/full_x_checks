import sys
import tkinter as tk
from tkinter import ttk
import threading


class ProgressDialog:
    """
    A scrollable progress log dialog for UAT/debug runs.
    Displays log entries in real time as processing occurs.
    Runs on the main thread; processing runs on a background thread.

    Stop / dismiss behaviour:
    - "Stop"           → sets stop event; button changes to "Return to Form".
    - "Return to Form" → destroys dialog; calls on_dismiss() to return to the
                         file-upload form pre-filled with the previous inputs.
    - On success       → button changes to "Close"; "Close" exits the application.
    - on_dismiss=None  → always exits (legacy / debug mode behaviour).
    """

    WINDOW_SIZE = 550  # Square dimensions in pixels

    def __init__(self, root: tk.Tk, on_dismiss=None):
        self.root = root
        self.on_dismiss = on_dismiss        # callable() → return to upload form
        self.stop_event = threading.Event()
        self._stopped = False
        self._completed_successfully = False

        self._build_ui()
        self._centre_on_screen()

    # =========================================================
    # UI Construction
    # =========================================================

    def _build_ui(self):
        self.window = tk.Toplevel(self.root)
        self.window.title("X-Checks UAT — Processing Log")
        self.window.resizable(False, False)
        self.window.protocol("WM_DELETE_WINDOW", self._on_stop_or_close)

        # Prevent interaction with the launcher behind it
        self.window.grab_set()

        outer_frame = ttk.Frame(self.window, padding="12")
        outer_frame.pack(fill="both", expand=True)

        # --- Title label ---
        ttk.Label(
            outer_frame,
            text="Processing Log",
            font=("Zurich Sans Semibold", 13)
        ).pack(pady=(0, 8))

        # --- Scrollable text area ---
        text_frame = ttk.Frame(outer_frame)
        text_frame.pack(fill="both", expand=True)

        scrollbar = ttk.Scrollbar(text_frame, orient="vertical")
        scrollbar.pack(side="right", fill="y")

        self.text_area = tk.Text(
            text_frame,
            state="disabled",       # Read-only
            wrap="word",
            font=("Courier", 9),
            bg="#ECEEEF",
            fg="#23366F",
            relief="sunken",
            borderwidth=1,
            yscrollcommand=scrollbar.set,
            width=60,
            height=28
        )
        self.text_area.pack(side="left", fill="both", expand=True)
        scrollbar.config(command=self.text_area.yview)

        # --- Stop / Close button ---
        btn_frame = ttk.Frame(outer_frame)
        btn_frame.pack(pady=(10, 0))

        self.action_btn = ttk.Button(
            btn_frame,
            text="Stop",
            width=12,
            command=self._on_stop_or_close
        )
        self.action_btn.pack()

    def _centre_on_screen(self):
        self.window.update_idletasks()
        w = self.WINDOW_SIZE
        h = self.WINDOW_SIZE
        screen_w = self.window.winfo_screenwidth()
        screen_h = self.window.winfo_screenheight()
        x = (screen_w - w) // 2
        y = (screen_h - h) // 2
        self.window.geometry(f"{w}x{h}+{x}+{y}")

    # =========================================================
    # Button handler
    # =========================================================

    def _on_stop_or_close(self):
        if not self._stopped:
            # First press — stop processing
            self._stopped = True
            self.stop_event.set()
            self.action_btn.config(text="Return to Form")
            self.append_entry("---", "User requested stop. Waiting for current step to finish...")
        else:
            # Second press — dismiss
            self.window.grab_release()
            self.window.destroy()
            if self._completed_successfully or self.on_dismiss is None:
                # Success or no callback registered → exit the application
                self.root.destroy()
                sys.exit(0)
            else:
                # Cancel or error → return to the upload form
                self.on_dismiss()

    # =========================================================
    # Public interface — called from background thread
    # =========================================================

    def append_entry(self, file: str, step: str, count: int = 0, notes: str = ""):
        """
        Thread-safe method to append a log line to the text area.
        Uses root.after() to marshal the update onto the main thread.
        """
        line = f"[{file}]  {step}"
        if count:
            line += f"  ({count})"
        if notes:
            line += f"  — {notes}"
        line += "\n"

        self.root.after(0, self._write_line, line)

    def _write_line(self, line: str):
        """Must only be called on the main thread via root.after()."""
        self.text_area.config(state="normal")
        self.text_area.insert("end", line)
        self.text_area.see("end")          # Auto-scroll to latest entry
        self.text_area.config(state="disabled")

    def is_stopped(self) -> bool:
        return self.stop_event.is_set()

    def mark_success(self):
        """Call from the processing thread when the run completes without error."""
        self._completed_successfully = True