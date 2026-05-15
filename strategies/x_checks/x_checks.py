import pandas as pd
from strategies.base_strategy import BaseStrategy, UploadTaskConfig
from .ebx_extraction import extract_ebx
from .fip_extraction import extract_fip
from .compare import compare


class XChecks(BaseStrategy):

    def __init__(self, config: UploadTaskConfig):
        super().__init__(config)

    def process(self, loaded_files: dict, files: dict):
        self.log_step(self.log, "System", "Starting X-Checks processing", 0)

        # 1. Extract EBX
        self.log_step(self.log, "EBX", "Extracting from publication file...", 0)
        ebx_results = extract_ebx(loaded_files["X-Checks Publication File"])
        self.log_step(self.log, "EBX", "X-Checks extracted", len(ebx_results))

        # 2. Extract FIP — x_check_list from all unique X-Check No. values in the raw file,
        #    matching old FIPExtraction.py which used the EBX file directly rather than
        #    extraction results (ensures X-Checks with no Account No. rows are still searched in FIP)
        x_check_list = sorted(set(
            str(x) for x in loaded_files["X-Checks Publication File"]["X-Check No."].tolist()
            if str(x) not in ("nan", "", "NaN", "None")
        ))
        self.log_step(self.log, "FIP", "Extracting from FIP text...", len(x_check_list))
        fip_results = extract_fip(loaded_files["FIP file"], x_check_list)
        self.log_step(self.log, "FIP", "X-Checks extracted", len(fip_results))

        # 3. Compare and sort — matches old Compare_Files.py "All Data" sheet sort order
        self.log_step(self.log, "Compare", "Comparing EBX and FIP...", 0)
        df_comparison = pd.DataFrame(compare(ebx_results, fip_results))
        df_comparison = df_comparison.sort_values("X-Check Number").reset_index(drop=True)

        # 4. Write Excel output — no summary, headers at row 1
        self.write_excel_output(
            output_path=self.build_output_path(
                files["output_directory"], "X-Checks Comparison", files["timestamp"]
            ),
            sheets={"X-Checks Comparison": df_comparison},
            log=self.log,
        )

    def apply_output_formatting(self, workbook):
        from openpyxl.styles import PatternFill, Font

        if "X-Checks Comparison" not in workbook.sheetnames:
            return

        green_fill  = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
        green_font  = Font(color="276221")
        red_fill    = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
        red_font    = Font(color="9C0006")
        orange_fill = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
        orange_font = Font(color="9C6500")

        ws = workbook["X-Checks Comparison"]
        for col in ("Formula Match", "Variables Match", "Variables Match (Builder)"):
            self.apply_conditional_formatting(
                worksheet=ws,
                column_name=col,
                rules={
                    "Match":     (green_fill,  green_font),
                    "MisMatch":  (red_fill,    red_font),
                    "Not Found": (orange_fill, orange_font),
                }
            )
