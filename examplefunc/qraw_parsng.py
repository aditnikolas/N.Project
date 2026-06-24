from pathlib import Path
from collections import defaultdict
from .qraw_addon import (
    QRawAddon,
    QRawCover,
    Kabkota_list,
    QrawOptionalCsv
)
from controller import log_event

import csv
import shutil
import pythoncom


class QRawParsing:

    def etl_from_temp(self, method, poi_list, template_path, target_parent, temp_path):
        pythoncom.CoInitialize()

        parts               = target_parent.name.split(" - ", 1)
        qraw_sub_name       = f"{parts[0]} {parts[1]}"
        excel               = QRawAddon().excel_prepatarion()
        opsel_books         = {}

        opsel_map = {
            "Operator":   {"OP1", "OP2"},
            "Celullar": {"Celullar"}
        }

        # ---- FAST lookup Home Operator -> Opsel
        home_op_to_opsel = {
            v.strip().lower(): opsel
            for opsel, values in opsel_map.items()
            for v in values
        }

        poi_set = set(poi_list)
        print("-"*60)
        log_event("QRaw", "ETL", f"Starting QRaw {method}, ETL process from temp files")

        try:
            file_name = target_parent.name.lower()
            for _, kabkota in Kabkota_list:
                if kabkota.lower() in file_name:
                    qraw_template_path  = Path(r"Excel_template_thd3000.xlsm")
                    break
                else:
                    qraw_template_path  = Path(r"Excel_template_thd1000.xlsm")

            for opsel in opsel_map:
                qraw_file_name = f"Qos Raw {opsel} {method} - {qraw_sub_name}.xlsm"
                qraw_dest_path = Path(qraw_file_name)
                shutil.copy2(qraw_template_path, qraw_dest_path)

                wb                  = excel.Workbooks.Open(str(qraw_dest_path))
                opsel_books[opsel]  = wb
                log_event("QRaw", "ETL", f"Qos Raw File {method} {opsel} {qraw_file_name}")
            
            has_data        = {opsel: False for opsel in opsel_books}
            chunk_size      = 50_000
            buffers         = defaultdict(lambda: defaultdict(list))

            for csv_file in temp_path.glob("*.csv"):
                sheet_name  = csv_file.stem
                if sheet_name == "Distance":
                    continue

                with open(csv_file, newline="", encoding="utf-8-sig") as f:
                    reader = csv.DictReader(f)

                    if not {"Collection", "Home Operator"}.issubset(reader.fieldnames):
                        continue

                    for row in reader:
                        poi = row["Collection"]
                        if poi not in poi_set:
                            continue

                        home_op = row["Home Operator"]
                        if not home_op:
                            continue

                        opsel = home_op_to_opsel.get(home_op.strip().lower())
                        if not opsel:
                            continue

                        buffers[opsel][sheet_name].append(
                            [row[h] for h in reader.fieldnames]
                        )
                        has_data[opsel] = True
                        if len(buffers[opsel][sheet_name]) >= chunk_size:
                            QRawAddon().flush_to_excel(opsel_books[opsel], sheet_name, buffers[opsel][sheet_name])
                            buffers[opsel][sheet_name].clear()
                            has_data[opsel] = True


            # ---- flush remaining
            for opsel, sheets in buffers.items():
                wb = opsel_books[opsel]
                for sheet_name, rows in sheets.items():
                    if rows:
                        QRawAddon().flush_to_excel(wb, sheet_name, rows)
                        has_data[opsel] = True

            for opsel, wb in opsel_books.items():
                # ---- Applying cover for every wb
                QRawCover().apply_cover(wb, template_path, target_parent, poi_list)
                # ---- Applying highlight for every wb
                file_path = Path(wb.FullName)
                wb.Save()
                wb.Close(False)
                log_event("QRaw", "ETL", f"Saved workbook for {opsel}")
                if has_data[opsel] == False:
                    Path(file_path).unlink()
                    log_event("QRaw", "ETL", f"Deleting no data file Qos Raw {opsel}")

        finally:
            excel.Quit()
            pythoncom.CoUninitialize()

        # ---- Running Optional if signal strength data more than excel rows
        overflow_flag    = QrawOptionalCsv().detect_overflow_flag(temp_path, poi_set, home_op_to_opsel)
        QrawOptionalCsv().export_overflow_csv(temp_path, poi_set, home_op_to_opsel, overflow_flag)

        log_event("QRaw", "ETL", f"ETL process finished successfully")
