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
        # Bold red highlighting for error / exception lines.
        self.text_area.tag_configure(
            "error",
            foreground="#C00000",
            font=("Courier", 9, "bold"),
        )
        # Green for matched/complete lines.
        self.text_area.tag_configure(
            "matched",
            foreground="#276221",
            font=("Courier", 9),
        )
        # Orange for not-found / not-matched lines.
        self.text_area.tag_configure(
            "mismatch",
            foreground="#9C6500",
            font=("Courier", 9),
        )
        # Bold dark-blue separator line (dashes).
        self.text_area.tag_configure(
            "separator",
            foreground="#23366F",
            font=("Courier", 9, "bold"),
        )

        # --- Stop / Close + Exit buttons ---
        btn_frame = ttk.Frame(outer_frame)
        btn_frame.pack(pady=(10, 0))

        self.action_btn = ttk.Button(
            btn_frame,
            text="Stop",
            width=16,
            command=self._on_stop_or_close
        )
        self.action_btn.pack(side="left", padx=(0, 8))

        self.exit_btn = ttk.Button(
            btn_frame,
            text="Exit Application",
            width=16,
            command=self._on_exit_application
        )
        self.exit_btn.pack(side="left")

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

    def _on_exit_application(self):
        """Hard exit — close dialog and shut down the entire app regardless of state."""
        self.stop_event.set()
        try:
            self.window.grab_release()
            self.window.destroy()
        except Exception:
            pass
        try:
            self.root.destroy()
        except Exception:
            pass
        sys.exit(0)

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

    _ERROR_KEYWORDS    = ("error", "failed", "failure", "traceback")
    _MATCHED_KEYWORDS  = ("matched:", "matched ", "complete", "successfully",
                          "copied to clipboard", "applied label")
    _MISMATCH_KEYWORDS = ("not in fip", "not matched", "mismatch", "not found")
    # "exception" excluded — appears in "Known Exception List" (not an error).

    def append_entry(self, file: str, step: str, count: int = 0, notes: str = "",
                     timestamp: str = ""):
        """
        Thread-safe method to append a log line to the text area.
        Format: [File]  [yyyymmdd hhmmss]  step  (count)  — notes
        """
        ts_part = f"  [{timestamp}]" if timestamp else ""
        line = f"[{file}]{ts_part}  {step}"
        if count:
            line += f"  ({count})"
        if notes:
            line += f"  — {notes}"
        line += "\n"

        haystack = f"{file} {step} {notes}".casefold()
        is_error    = any(kw in haystack for kw in self._ERROR_KEYWORDS)
        is_matched  = not is_error and any(kw in haystack for kw in self._MATCHED_KEYWORDS)
        is_mismatch = not is_error and not is_matched and any(
            kw in haystack for kw in self._MISMATCH_KEYWORDS
        )

        tag = "error" if is_error else ("matched" if is_matched else ("mismatch" if is_mismatch else None))

        if is_error:
            self._play_error_sound()
        self.root.after(0, self._write_line, line, tag)

    @staticmethod
    def _play_error_sound() -> None:
        """Plays the Windows system 'critical stop' chime (MB_ICONHAND).
        Silent no-op on non-Windows platforms."""
        try:
            import winsound
            winsound.MessageBeep(0x10)  # MB_ICONHAND
        except Exception:
            pass

    def _write_line(self, line: str, tag: str = None):
        """Must only be called on the main thread via root.after()."""
        self.text_area.config(state="normal")
        if tag:
            self.text_area.insert("end", line, tag)
        else:
            self.text_area.insert("end", line)
        self.text_area.see("end")          # Auto-scroll to latest entry
        self.text_area.config(state="disabled")

    def append_separator(self):
        """Writes two blank lines then a dashed separator line."""
        self.root.after(0, self._write_separator)

    def _write_separator(self):
        self.text_area.config(state="normal")
        self.text_area.insert("end", "\n\n")
        self.text_area.insert("end", "-" * 60 + "\n", "separator")
        self.text_area.see("end")
        self.text_area.config(state="disabled")

    def is_stopped(self) -> bool:
        return self.stop_event.is_set()

    def mark_success(self):
        """Call from the processing thread when the run completes without error."""
        self._completed_successfully = True