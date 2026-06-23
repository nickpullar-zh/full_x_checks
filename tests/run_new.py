import sys
import os
sys.path.insert(0, r"C:\Users\NICK.PULLAR\OneDrive - Zurich Insurance\Projects\Testing Automation\Python Files\Full_X-Checks")

from datetime import datetime
from task_configs import X_CHECKS_UPLOAD_CONFIG
from strategies.x_checks import XChecks

output_dir = r"C:\Users\NICK.PULLAR\OneDrive - Zurich Insurance\Projects\Testing Automation\Python Files\Full_X-Checks\test_data\X-Checks Output"
os.makedirs(output_dir, exist_ok=True)

files = {
    "files": {
        "FIP file":                  r"C:\Users\NICK.PULLAR\OneDrive - Zurich Insurance\Projects\Testing Automation\Python Files\Full_X-Checks\test_data\20251205 FIP X-Checks - Original.txt",
        "X-Checks Publication File": r"C:\Users\NICK.PULLAR\OneDrive - Zurich Insurance\Projects\Testing Automation\Python Files\Full_X-Checks\test_data\20251205 EPM X-Checks - Original - Copy.xlsx",
        "Known Exception List":      r"C:\Users\NICK.PULLAR\OneDrive - Zurich Insurance\Projects\Testing Automation\Python Files\Full_X-Checks\test_data\Known_Exception_List.xlsx",
    },
    "sheet_names": {
        "X-Checks Publication File": "cross checks all",
        "Known Exception List":      "Known Exceptions",
    },
    "output_directory":          output_dir,
    "timestamp":                 datetime.now().strftime("%Y%m%d_%H%M%S"),
    "process_only_differences":  False,
}

strategy = XChecks(X_CHECKS_UPLOAD_CONFIG)
strategy.execute(files)
