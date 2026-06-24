import subprocess
import json
import win32com.client as win32
from pathlib import Path
from controller import log_event
from bulk import BulkExtractFilesVer2
from examplefunc import (
    ParsingKRaw,
    QRawParsing
)

class ComsumMain:

    def main(self, json_path):
        """
        Portfolio preview version.

        This class demonstrates the main orchestration flow of a telecom
        drive test data processing automation system.

        Full production logic, internal templates, and confidential
        business rules are not included.
        """
        excel           = win32.DispatchEx("Excel.Application")
        n_poi           = 1
        temp_path       = Path("Folder_temp")
        csum_path       = Path("compile_summary_path")
        poi_list        = []
        target_parent   = Path(json.load(open(json_path))["target_parent"])

        try:
            frex_path   = temp_path.glob("Formula_Rex*.xlsx")
            wb_csum         = excel.Workbooks.Open(str(csum_path), ReadOnly=False)
            wb_frex     = excel.Workbooks.Open(str(list(frex_path)[0]), ReadOnly=False)
            for poi in poi_list:
                # ---- Parsing data Komdigi_Raw to Formula_Rex_V3
                ParsingKRaw().main_parsing(wb_frex, excel, target_parent, poi, wb_csum, n_poi)
                n_poi += 1
        finally:
            wb_csum.Close(True)
            excel.Quit()
            del wb_csum, wb_frex, excel

    def main_qraw(self, json_path, qraw_event):
        temp_ready = False
        temp_ready = temp_ready if self.temp_path.exists() else False

        if qraw_event:
            src_path = "Komdigi_Raw_TBB_path"
            if not src_path.exists():
                return

            poi_list = QRawAddon().get_collection_list(src_path)
            # ---- check if there is csv files in temp folder, if yes process ETL, if not export to csv and then process ETL
            if temp_ready:
                # ---- process ETL if DT: dt_poi_list / ST: st_poi_list, Apply Cover, Apply Highlight
                QRawParsingEvent().etl_from_temp(poi_list, template_path, target_parent, temp_path)
            
            else:
                # ---- export to csv
                BulkExtractFilesVer2().extract_xlsm_ver2(src_path, temp_path, poi_list)
                # ---- csv import to qraw template
                QRawParsingEvent().etl_from_temp(poi_list, template_path, target_parent, temp_path)
                QRawParsingEvent().clean_temp_path(temp_path)
        else:
            for method in ["DT", "ST"]:
                src_path = "Komdigi_Raw_TBB_Path_for_each_method"
                poi_list = QRawAddon().get_collection_list(src_path)
                # ---- check if there is csv files in temp folder, if yes process ETL, if not export to csv and then process ETL
                if temp_ready:
                    # ---- process ETL if DT: dt_poi_list / ST: st_poi_list, Apply Cover, Apply Highlight
                    QRawParsing().etl_from_temp(method, poi_list, template_path, target_parent, temp_path)
                
                else:
                    # ---- export to csv
                    BulkExtractFilesVer2().extract_xlsm_ver2(src_path, temp_path, poi_list)
                    # ---- csv import to qraw template
                    QRawParsing().etl_from_temp(method, poi_list, Path(template_path), target_parent, temp_path)
                    QRawParsing().clean_temp_path(temp_path)
    
    def map_plot_helper(self, ARCGISPY, json_path, parent_folder):
        mapplot_path    = "mapplot_path_folder"
        raw_path        = "RAW_folder_path"
        target_folder   = []

        if raw_path.exists():
            target_folder.append(mapplot_path)
        else:
            target_folder = [
                fname for fname in mapplot_path.iterdir() if fname.is_dir()
            ]
        
        run_file = self.resolve_run_file(parent_folder)      
        for target in target_folder:
            MAPPLOTRUN      = Path(__file__).parent / run_file #"RUN.py"
            subprocess.run(
                [ARCGISPY, str(MAPPLOTRUN), str(json_path), str(target)],
                check=True
            )




            
