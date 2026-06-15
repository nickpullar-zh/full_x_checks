"""
Collect Live X-Checks strategy.

Runs the same X-Check selection pipeline as v0.4.1 (Status / Type of Change /
Exclude Z-Core / yellow Category), but as a standalone task: only requires the
X-Checks Publication file and an output directory. Writes the result both as a
timestamped .txt file and to the system clipboard.
"""
import os
import sys

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
        Pushes `text` to the Windows clipboard.

        Uses the Win32 API (user32.OpenClipboard + SetClipboardData) directly via
        ctypes — Tk's clipboard_append does not reliably flush data when called
        from a worker thread, but Win32 does, and it doesn't require a running
        message loop on this thread.
        """
        if sys.platform != "win32":
            raise OSError("Clipboard copy is only implemented on Windows.")

        import ctypes
        from ctypes import wintypes

        CF_UNICODETEXT = 13
        GMEM_MOVEABLE  = 0x0002

        u32 = ctypes.windll.user32
        k32 = ctypes.windll.kernel32

        # Bind argtypes/restypes — required on 64-bit Windows so ctypes uses
        # 64-bit pointer-sized values for handles and pointers.
        k32.GlobalAlloc.argtypes  = [wintypes.UINT, ctypes.c_size_t]
        k32.GlobalAlloc.restype   = wintypes.HGLOBAL
        k32.GlobalLock.argtypes   = [wintypes.HGLOBAL]
        k32.GlobalLock.restype    = wintypes.LPVOID
        k32.GlobalUnlock.argtypes = [wintypes.HGLOBAL]
        k32.GlobalUnlock.restype  = wintypes.BOOL
        k32.GlobalFree.argtypes   = [wintypes.HGLOBAL]
        k32.GlobalFree.restype    = wintypes.HGLOBAL
        u32.OpenClipboard.argtypes  = [wintypes.HWND]
        u32.OpenClipboard.restype   = wintypes.BOOL
        u32.EmptyClipboard.restype  = wintypes.BOOL
        u32.SetClipboardData.argtypes = [wintypes.UINT, wintypes.HANDLE]
        u32.SetClipboardData.restype  = wintypes.HANDLE
        u32.CloseClipboard.restype  = wintypes.BOOL

        # Win32 expects a null-terminated UTF-16 string
        data = text + "\0"
        size = len(data) * ctypes.sizeof(ctypes.c_wchar)

        h_mem = k32.GlobalAlloc(GMEM_MOVEABLE, size)
        if not h_mem:
            raise OSError("GlobalAlloc failed")
        ptr = k32.GlobalLock(h_mem)
        if not ptr:
            k32.GlobalFree(h_mem)
            raise OSError("GlobalLock failed")
        ctypes.memmove(ptr, ctypes.create_unicode_buffer(data), size)
        k32.GlobalUnlock(h_mem)

        if not u32.OpenClipboard(None):
            k32.GlobalFree(h_mem)
            raise OSError("OpenClipboard failed")
        try:
            u32.EmptyClipboard()
            if not u32.SetClipboardData(CF_UNICODETEXT, h_mem):
                k32.GlobalFree(h_mem)
                raise OSError("SetClipboardData failed")
            # Once SetClipboardData succeeds, the OS owns h_mem — do not free it
        finally:
            u32.CloseClipboard()
