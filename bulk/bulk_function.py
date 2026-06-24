import csv
import win32com.client as win32

from pathlib import Path
from controller import log_event

class BulkExtractFilesVer2:

    def extract_xlsm_ver2(self, src_file, temp_path, poi_tbb):
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        excel.EnableEvents = False
        try:
            excel.Calculation = -4135  # xlCalculationManual
        except Exception:
            pass # xlCalculationManual

        src_wb          = None
        header_row      = 2
        data_start_row  = 3
        xlUp            = -4162
        xlToLeft        = -4159

        try:
            src_wb = excel.Workbooks.Open(src_file, ReadOnly=True)

            poi_set = {
                p.strip()
                for p in poi_tbb
                if isinstance(p, str) and p.strip()
            }

            for ws in src_wb.Worksheets:
                csv_path        = Path("ws.Name_path")
                file_exists     = csv_path.exists()
                last_col        = ws.Cells(header_row, ws.Columns.Count).End(xlToLeft).Column
                headers = ws.Range(
                    ws.Cells(header_row, 1),
                    ws.Cells(header_row, last_col)
                ).Value

                if not headers:
                    log_event("Bulk", "Parsing", f"Skipped -- {ws.Name} Sheet")
                    continue

                headers = list(headers[0]) if isinstance(headers[0], tuple) else list(headers)
                headers = [
                    str(h).strip() if h is not None else ""
                    for h in headers
                ]

                try:
                    col_idx = headers.index("Collection")  # zero-based index
                except ValueError:
                    log_event("Bulk", "Parsing", f"Skipped -- {ws.Name} Sheet")
                    continue

                excel_col_idx = col_idx + 1

                # --- Last Row from Column
                last_row = ws.Cells(ws.Rows.Count, excel_col_idx).End(xlUp).Row

                if last_row < data_start_row:
                    log_event("Bulk", "Parsing", f"Skipped -- {ws.Name} Sheet")
                    continue

                data_range = ws.Range(
                    ws.Cells(data_start_row, 1),
                    ws.Cells(last_row, last_col)
                ).Value

                if not data_range:
                    log_event("Bulk", "Parsing", f"Skipped -- {ws.Name} Sheet")
                    continue

                if not isinstance(data_range, tuple):
                    data_range = ((data_range,),)
                elif data_range and not isinstance(data_range[0], tuple):
                    data_range = (data_range,)

                if file_exists:
                    self.clean_csv_empty_rows(csv_path)

                # --- Default filter use POI Set
                target_set = poi_set

                # --- Check existing collection name in csv file and append if its new
                if file_exists and ws.Name == "Collection Name":
                    existing_poi = set()

                    with open(csv_path, newline="", encoding="utf-8-sig") as rf:
                        reader = csv.DictReader(rf)

                        if reader.fieldnames and "Collection" in reader.fieldnames:
                            for r in reader:
                                p = r.get("Collection")
                                if p:
                                    p = p.strip()
                                    if p:
                                        existing_poi.add(p)

                    target_set = poi_set - existing_poi

                    # --- If there is no new collection will skip
                    if not target_set:
                        log_event("Bulk", "Parsing", f"Skipped -- {ws.Name}.csv No new data")
                        continue

                rows_to_write = []
                for row in data_range:
                    val = row[col_idx]

                    if not isinstance(val, str):
                        continue
                    val = val.strip()
                    if val and val in target_set:
                        rows_to_write.append(row)

                if not rows_to_write:
                    log_event("Bulk", "Parsing", f"Skipped -- {ws.Name}.csv No matched data")
                    continue

                mode = "a" if file_exists else "w"
                with open(csv_path, mode, newline="", encoding="utf-8-sig") as f:
                    writer = csv.writer(f)

                    if not file_exists:
                        writer.writerow(headers)

                    writer.writerows(rows_to_write)

                action = "Appended" if file_exists else "Created"
                log_event("Bulk", "Parsing", f"{action} -- {ws.Name}.csv")

        finally:
            if src_wb is not None:
                src_wb.Close(False)

            excel.Quit()
    
    def extract_csv_ver2(self, src_file, temp_path, poi_tbb):
        csv_map = {
            "RxLev"   : "2G_Quality_Strength",
            "RSCP"    : "3G_Quality_Strength",
            "RSRP"    : "4G_Quality_Strength",
            "SS-RSRP" : "5G_Quality_Strength"
        }

        src_file  = Path(src_file)
        temp_path = Path(temp_path)

        # ---- clean poi set
        poi_set = {
            p.replace("\xa0", "").strip()
            for p in poi_tbb
            if isinstance(p, str) and p.strip()
        }

        with open(src_file, newline="", encoding="utf-8-sig") as f:
            reader = csv.reader(f)
            header = next(reader, [])

            # ---- define csv target
            dst_name = None
            for key, value in csv_map.items():
                if key in header:
                    dst_name = value
                    break

            if not dst_name:
                return

            dst_path = temp_path / f"{dst_name}.csv"
            file_exists = dst_path.exists()

            # ---- find collection column
            try:
                col_idx = header.index("Collection")
            except ValueError:
                return

            # ---- open target csv
            with open(dst_path, "a" if file_exists else "w",
                    newline="", encoding="utf-8-sig") as out:
                writer = csv.writer(out)

                # ---- write once
                if not file_exists:
                    writer.writerow(header)

                # ---- FILTER & APPEND
                for row in reader:
                    val = row[col_idx]

                    if not isinstance(val, str):
                        continue

                    val = val.replace("\xa0", "").strip()
                    if not val:
                        continue

                    if val in poi_set:
                        writer.writerow(row)