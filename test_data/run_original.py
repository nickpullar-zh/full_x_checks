import sys
import os
sys.path.insert(0, r"C:\Users\NICK.PULLAR\OneDrive - Zurich Insurance\Projects\Testing Automation\Python Files\X-Checks")

import globals
from datetime import datetime
from FIPExtraction import FIPExtraction
from EBXExtraction1 import EBXExtraction1
from Compare_Files import Compare_Files
from OutputExcel_Formating import Format_Excel_File

arrEBXPaths = [r"C:\Users\NICK.PULLAR\OneDrive - Zurich Insurance\Projects\Testing Automation\Python Files\Full_X-Checks\test_data\20251205 EPM X-Checks - Original.xlsx"]
arrFIPPaths = [r"C:\Users\NICK.PULLAR\OneDrive - Zurich Insurance\Projects\Testing Automation\Python Files\Full_X-Checks\test_data\20251205 FIP X-Checks - Original.txt"]
output_dir  = r"C:\Users\NICK.PULLAR\OneDrive - Zurich Insurance\Projects\Testing Automation\Python Files\Full_X-Checks\test_data\X-Checks Output"

os.makedirs(output_dir, exist_ok=True)
globals.TIMESTAMP = datetime.now().strftime("%Y%m%d %H%M%S")
strMergedFIPFile = os.path.join(output_dir, "Merged FIP File.txt")

print("Running FIPExtraction...")
strFIPPath = FIPExtraction(arrFIPPaths, strMergedFIPFile, arrEBXPaths, output_dir)
print(f"FIP done: {strFIPPath}")

print("Running EBXExtraction1...")
strEBXPath = EBXExtraction1(arrEBXPaths, output_dir)
print(f"EBX done: {strEBXPath}")

print("Running Compare_Files...")
strComparePath = Compare_Files(strEBXPath, strFIPPath, output_dir)
print(f"Compare done: {strComparePath}")

print("Running Format_Excel_File...")
strFormatted = Format_Excel_File(strComparePath, 'Sheet1')
print(f"Formatted: {strFormatted}")

print("All done!")
