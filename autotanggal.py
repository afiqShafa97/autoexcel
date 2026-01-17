import time
import os
import win32com.client
from datetime import datetime, date

# ---------- CONFIG ----------
FILE_PATH = os.path.abspath("./test.xlsx")
SCAN_INTERVAL_SECONDS = 3
EDIT_DELAY_SECONDS = 0.2

GRAY_DAYS = {1, 3, 4, 10, 11, 16, 17, 18, 24, 25, 31}

COLOR_WHITE = 2
COLOR_GRAY = 15
# ----------------------------


def get_excel_app():
    try:
        excel = win32com.client.GetActiveObject("Excel.Application")
        print("Connected to existing Excel instance.")
    except Exception:
        excel = win32com.client.Dispatch("Excel.Application")
        print("Started new Excel instance.")

    excel.Visible = True
    excel.DisplayAlerts = False
    return excel


def excel_is_ready(excel):
    try:
        return excel.Ready
    except Exception:
        return False


def find_column_by_keyword(ws, used, keyword):
    """Find column index by header keyword (case-insensitive)"""
    for r in range(1, used.Rows.Count + 1):
        for c in range(1, used.Columns.Count + 1):
            val = ws.Cells(r, c).Value
            if isinstance(val, str) and keyword in val.lower():
                return c
    return None


def find_rightmost_column(ws, used):
    for r in range(1, used.Rows.Count + 1):
        for c in range(1, used.Columns.Count + 1):
            val = ws.Cells(r, c).Value
            if isinstance(val, str):
                t = val.lower()
                if "penerangan" in t and "ac" in t and "poin" in t:
                    return c
    return None


def find_mulai_selesai_columns(ws, used, max_col):
    cols = []
    for r in range(1, used.Rows.Count + 1):
        for c in range(1, max_col):
            val = ws.Cells(r, c).Value
            if isinstance(val, str):
                t = val.lower()
                if "mulai" in t or "selesai" in t:
                    cols.append(c)
    return list(set(cols))


def find_nearest_formula_above(ws, col, row):
    for r in range(row - 1, 0, -1):
        cell = ws.Cells(r, col)
        if cell.HasFormula:
            return r
    return None


excel = get_excel_app()
wb = None

print("Script running (Tanggal-driven logic, autofill twice with recolor)...")

# ---------- MAIN LOOP ----------
while True:
    try:
        if not os.path.exists(FILE_PATH):
            time.sleep(SCAN_INTERVAL_SECONDS)
            continue

        if wb is None:
            for book in excel.Workbooks:
                if book.FullName.lower() == FILE_PATH.lower():
                    wb = book
                    break
            if wb is None:
                wb = excel.Workbooks.Open(FILE_PATH)

        if not excel_is_ready(excel):
            time.sleep(1)
            continue

        ws = excel.ActiveSheet
        used = ws.UsedRange

        # ---- FIND REQUIRED COLUMNS ----
        tanggal_col = find_column_by_keyword(ws, used, "tanggal")
        max_col = find_rightmost_column(ws, used)

        if tanggal_col is None or max_col is None:
            time.sleep(SCAN_INTERVAL_SECONDS)
            continue

        # ---- STEP 1: CHANGE DECEMBER -> JANUARY ----
        date_rows = set()
        found_31_jan = False

        for r in range(1, used.Rows.Count + 1):
            cell = ws.Cells(r, tanggal_col)
            val = cell.Value

            if isinstance(val, (datetime, date)):
                date_rows.add(r)

                if val.month == 12:
                    cell.Value = val.replace(month=1)
                    time.sleep(EDIT_DELAY_SECONDS)

                if val.month == 1 and val.day == 31:
                    found_31_jan = True

        if not found_31_jan:
            time.sleep(SCAN_INTERVAL_SECONDS)
            continue

        # ---- STEP 2: CLASSIFY ROWS BASED ONLY ON TANGGAL ----
        row_is_gray = {}

        for r in date_rows:
            d = ws.Cells(r, tanggal_col).Value
            row_is_gray[r] = isinstance(d, (datetime, date)) and d.day in GRAY_DAYS

            for c in range(1, max_col + 1):
                ws.Cells(r, c).Interior.ColorIndex = (
                    COLOR_GRAY if row_is_gray[r] else COLOR_WHITE
                )

        # ---- STEP 3: CLEAR MULAI / SELESAI ----
        mulai_selesai_cols = find_mulai_selesai_columns(ws, used, max_col)
        for r in date_rows:
            for c in mulai_selesai_cols:
                ws.Cells(r, c).Value = None

        # ---- STEP 4: RIGHTMOST COLUMN LOGIC (RUN TWICE) ----
        for pass_index in range(2):
            for r in date_rows:
                cell = ws.Cells(r, max_col)

                # Always re-evaluate gray on second pass
                if pass_index == 1:
                    d = ws.Cells(r, tanggal_col).Value
                    row_is_gray[r] = (
                        isinstance(d, (datetime, date)) and d.day in GRAY_DAYS
                    )

                if row_is_gray[r]:
                    cell.Value = None
                    cell.Interior.ColorIndex = COLOR_GRAY
                    continue

                if cell.Value in (None, "") and not cell.HasFormula:
                    src_row = find_nearest_formula_above(ws, max_col, r)
                    if src_row:
                        src = ws.Cells(src_row, max_col)
                        dest = ws.Range(src, cell)
                        src.AutoFill(dest, 0)

                # Explicit repaint on second pass
                if pass_index == 1:
                    cell.Interior.ColorIndex = COLOR_WHITE

            time.sleep(0.2)

        # ---- STEP 5: NEXT TAB ----
        next_index = ws.Index + 1 if ws.Index < wb.Worksheets.Count else 1
        wb.Worksheets(next_index).Activate()
        time.sleep(1)

    except KeyboardInterrupt:
        print("Stopped by user.")
        break

    except Exception as e:
        print("Runtime warning:", e)
        time.sleep(1)
