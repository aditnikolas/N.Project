import pandas as pd
import win32com.client as win32

from pathlib import Path
from .comsum_addon import ComsumAddon
from controller import log_event

class ParsingKRaw:
    
    def main_parsing(self, wb_frex, excel, target_parent, poi, wb_csum, n_poi):
        
        self.comsum_addon   = ComsumAddon()
        komdigi_raw_check   = Path("Komdigi_Raw")

        if not komdigi_raw_check.exists():
            return

        # ---- get Komdigi_Raw exact file name
        poi_file_list       = (Path(target_parent) / poi).iterdir()
        file_name           = "Komdigi_Raw_TBB.xlsm"
        kraw_path           = Path(file_name)

        # ---- stop auto calculation for frex
        excel.Calculation  = -4135
        try:
            # ---- Parsing data Komdigi_Raw to Formula_Rex_V3
            self.kraw_frex_parse(kraw_path, wb_frex, excel, target_parent.name, poi, n_poi)
            # ---- Calculate excel
            excel.Calculate()
            # ---- Copy Formula_Rex_V3 result to Comsum Template
            self.apply_to_comsum(wb_frex, wb_csum)
        finally:
            excel.CalculateBeforeSave = False
            # ---- Make excel automatic calculation
            excel.Calculation = -4105

    def kraw_frex_parse(self, kraw_path, wb_frex, excel, target_parent_name, poi, n_poi):
        sheet_list = []
        wb_src = excel.Workbooks.Open(str(kraw_path), ReadOnly=True)

        for sheet_name in sheet_list:
            try:
                ws_src  = wb_src.Worksheets(sheet_name)
                ws_frex = wb_frex.Worksheets(sheet_name)
                
                # clear target content (start for row 3)
                self.comsum_addon.clean_data_content(ws_frex)
            except Exception:
                continue

            last_row = ws_src.Cells(
                ws_src.Rows.Count,
                2
            ).End(-4162).Row

            last_col = ws_src.Cells(
                2,
                ws_src.Columns.Count
            ).End(-4159).Column

            if last_row < 3:
                continue

            src_range = ws_src.Range(
                ws_src.Cells(3, 1),
                ws_src.Cells(last_row, last_col)
            )

            data = src_range.Value
            if not data:
                continue

            rows = len(data)
            cols = len(data[0])

            ws_frex.Range(
                ws_frex.Cells(3, 1),
                ws_frex.Cells(3 + rows - 1, cols)
            ).Value = data
            log_event("CSum", f"Parsing Poi-{n_poi}", f"[OK] Parsed {sheet_name}: {rows} rows")

        ws_addon                    = wb_frex.Worksheets("Addon_Sheet")
        ws_addon.Range("E3").Value  = target_parent_name
        ws_addon.Range("E4").Value  = poi
        wb_src.Close(False)

    def apply_to_comsum(self, wb_frex, wb_csum):
        log_event("CSum", "Finishing", f"Apply final data to Comsum")
        ws_csum = wb_csum.Sheets("Compile_Summary")
        ws_frex = wb_frex.Sheets("Compile_Summary")

        # ---- Get last row with data in Comsum
        last_col = ws_frex.Cells(3, ws_frex.Columns.Count).End(-4159).Column
        src_range = ws_frex.Range(
            ws_frex.Cells(3, 2),  # B3
            ws_frex.Cells(5, last_col + 1)
        )
        data      = src_range.Value2
        if not data:
            return

        # ---- Put data from Frex to Comsum
        rows = len(data)
        cols = len(data[0])

        last_row    = ws_csum.Cells(ws_csum.Rows.Count, "M").End(-4162).Row
        start_row   = max(last_row + 1, 3)

        ws_csum.Range(
            ws_csum.Cells(start_row, 2),
            ws_csum.Cells(start_row + rows - 1, 2 + cols - 1)
        ).Value = data

        wb_csum.Save()
