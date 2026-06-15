"""
Collect Live X-Checks strategy.

Runs the same X-Check selection pipeline as v0.4.1 (Status / Type of Change /
Exclude Z-Core / yellow Category), but as a standalone task: only requires the
X-Checks Publication file and an output directory. Writes the result both as a
timestamped .txt file and to the system clipboard.
"""
import os
import tkinter as tk

from strategies.base_strategy import BaseStrategy, UploadTaskConfig
from strategies.x_checks.x_check_no_selection import select_x_check_nos


class CollectLiveXChecks(BaseStrategy):

    def __init__(self, config: UploadTaskConfig):
        super().__init__(config)

    def process(self, loaded_files: dict, files: dict):
        self.log_step(self.log, "System", "Starting Collect Live X-Checks", 0)

        ebx_df    = loaded_files.get("X-Checks Publication File")
        ebx_path  = files["files"].get("X-Checks Publication File")
        ebx_sheet = files["sheet_names"].get("X-Checks Publication File")
        if ebx_df is None or not ebx_path or not ebx_sheet:
            self.log_step(self.log, "Collect", "Missing required file or sheet — aborting", 0)
            return

        x_check_nos = select_x_check_nos(ebx_df, ebx_path, ebx_sheet)
        if not x_check_nos:
            self.log_step(self.log, "Collect", "No X-Check Nos in scope after pipeline", 0)
            return

        text = "\n".join(x_check_nos)

        out_path = self.build_output_path(
            files["output_directory"], "X-Check_Nos", files["timestamp"], extension=".txt"
        )
        with open(out_path, "w", encoding="utf-8") as f:
            f.write(text)
        self.log_step(self.log, "Collect", f"Wrote {os.path.basename(out_path)}",
                      len(x_check_nos), notes=out_path)

        try:
            self._copy_to_clipboard(text)
            self.log_step(self.log, "Collect", "Copied to clipboard", len(x_check_nos))
        except Exception as e:
            # Non-fatal — the .txt is still on disk
            self.log_step(self.log, "Collect", "Clipboard copy failed", 0, notes=str(e))

    def _copy_to_clipboard(self, text: str) -> None:
        """
        Pushes `text` to the Windows clipboard via a short-lived hidden Tk root.
        Tk requires update() before destroy to actually flush the clipboard data.
        """
        root = tk.Tk()
        root.withdraw()
        try:
            root.clipboard_clear()
            root.clipboard_append(text)
            root.update()           # required: flushes clipboard to OS
        finally:
            root.destroy()
