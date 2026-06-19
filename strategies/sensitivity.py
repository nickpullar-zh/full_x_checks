"""
Apply Microsoft Information Protection (MIP) sensitivity labels to Excel
workbooks via the Office COM API.

Mirrors the VBA SetLabelInfo procedure:
  Workbook.SensitivityLabel.CreateLabelInfo
  → set AssignmentMethod = PRIVILEGED, LabelName, LabelId, SiteId
  → SensitivityLabel.SetLabel(myLabelInfo, myLabelInfo)
  → Workbook.Save()

Driven from Python via pywin32. Excel must be installed on the host
(true for the project's target environment). All operations are wrapped
in best-effort error handling — failure to label a file is logged but
does not abort the strategy run.
"""
from __future__ import annotations

import logging
import os
import sys
from typing import Optional


# Site ID comes from your VBA module — same Tenant ID for every label.
SITE_ID = "473672ba-cd07-4371-a2ae-788b4c61840e"

# Level name → (LabelId, LabelName). LabelName matches Office's friendly name;
# LabelId is the GUID configured for your tenant.
_LABELS: dict[str, tuple[str, str]] = {
    "Public":
        ("9a7ed875-cb67-40d7-9ea6-a804b08b1148", "Public"),
    "Internal_Use_Only":
        ("9108d454-5c13-4905-93be-12ec8059c842", "Internal_Use_Only"),
    "Confidential_Personal_Data":
        ("588a69e3-909e-4e76-9cf8-f87e7adb3732", "Confidential_Personal_Data"),
    "Confidential_Non_Personal_Data":
        ("407e232e-ec95-40fd-afb0-3c2f06064326", "Confidential_Non_Personal_Data"),
    "Highly_Confidential_Personal_Data":
        ("3bfdccf1-4047-4a16-b0dc-edce276357e3", "Highly_Confidential_Personal_Data"),
    "Highly_Confidential_Non_Personal_Data":
        ("0fa44580-70ca-490f-9809-703eff01a2f8", "Highly_Confidential_Non_Personal_Data"),
    "Highly_Confidential_Sensitive_Personal_Data":
        ("cb65cec6-b0dd-4cff-b768-7f91ed673364",
         "Highly_Confidential_Sensitive_Personal_Data"),
}

# Office MsoAssignmentMethod.PRIVILEGED = 2 (per Office object library)
ASSIGNMENT_METHOD_PRIVILEGED = 2


def label_for(level_name: str) -> tuple[str, str]:
    """Returns (label_id, label_name) for a level name. Raises KeyError if unknown."""
    return _LABELS[level_name]


class ExcelLabeler:
    """
    Holds a single Excel.Application COM instance and labels workbooks
    in-place. Reused across multiple write_excel_output calls in the same
    strategy run so we pay the Excel-startup cost once.

    Usage:
        labeler = ExcelLabeler()           # lazy COM startup
        labeler.label_file(path, "Internal_Use_Only", justification=None)
        labeler.close()                    # quits Excel cleanly

    Failures are logged and swallowed; the caller's run is not aborted.
    """

    def __init__(self):
        self._excel = None     # win32com Excel.Application instance
        self._available = sys.platform == "win32"

    def _start_excel(self) -> bool:
        """Lazily creates the Excel.Application. Returns True on success."""
        if not self._available:
            return False
        if self._excel is not None:
            return True
        try:
            import win32com.client
            self._excel = win32com.client.DispatchEx("Excel.Application")
            self._excel.Visible = False
            self._excel.DisplayAlerts = False
            return True
        except Exception as e:
            logging.getLogger(__name__).warning(
                "Could not start Excel COM for sensitivity labelling: %s", e)
            self._excel = None
            self._available = False
            return False

    def label_file(self, path: str, level_name: str,
                   justification: Optional[str] = None) -> tuple[bool, str]:
        """
        Applies the named sensitivity label to the workbook at `path`.
        Returns (success, message_or_label_name).
        """
        if not os.path.isfile(path):
            return False, f"File not found: {path}"
        try:
            label_id, label_friendly = label_for(level_name)
        except KeyError:
            return False, f"Unknown sensitivity level: {level_name!r}"

        if not self._start_excel():
            return False, "Excel COM unavailable"

        wb = None
        try:
            wb = self._excel.Workbooks.Open(os.path.abspath(path))
            sl = wb.SensitivityLabel
            info = sl.CreateLabelInfo()
            info.AssignmentMethod = ASSIGNMENT_METHOD_PRIVILEGED
            info.LabelName = label_friendly
            info.LabelId   = label_id
            info.SiteId    = SITE_ID
            if justification:
                info.Justification = justification
            sl.SetLabel(info, info)
            wb.Save()
            return True, label_friendly
        except Exception as e:
            return False, str(e)
        finally:
            if wb is not None:
                try:
                    wb.Close(SaveChanges=False)
                except Exception:
                    pass

    def close(self) -> None:
        """Quits the cached Excel.Application, if any."""
        if self._excel is not None:
            try:
                self._excel.Quit()
            except Exception:
                pass
            self._excel = None
