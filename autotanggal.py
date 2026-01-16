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


def find_rightmost_column_by_keywords(ws, used):
    for r in range(1, used.Rows.Count + 1):
        for c in range(1, used.Columns.Count + 1):
            val = ws.Cells(r, c).Value
            if isinstance(val, str):
                t = val.lower()
                if "penerangan" in t and "ac" in t and "poin" in t:
                    return c
    return None


def find_mulai_selesai_columns(ws, used, max_col):
    cols = set()
    for r in range(1, used.Rows.Count + 1):
        for c in range(1, max_col + 1):
            val = ws.Cells(r, c).Value
            if isinstance(val, str):
                t = val.lower()
                if "mulai" in t or "selesai" in t:
                    cols.add(c)
    return sorted(cols)


def find_nearest_formula_above(ws, col, start_row):
    for r in range(start_row - 1, 0, -1):
        cell = ws.Cells(r, col)
        if cell.HasFormula:
            return r
    return None


excel = get_excel_app()
wb = None

print("Script running (RACE-SAFE logic, formula-only autofill)...")

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

        ws = wb.ActiveSheet
        used = ws.UsedRange

        max_col = find_rightmost_column_by_keywords(ws, used)
        if max_col is None:
            time.sleep(SCAN_INTERVAL_SECONDS)
            continue

        found_31_jan = False
        date_rows = set()

        # ---- STEP 1: CHANGE DECEMBER -> JANUARY ----
        for r in range(1, used.Rows.Count + 1):
            for c in range(1, max_col + 1):
                cell = ws.Cells(r, c)
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

        # ---- STEP 2: COLOR + CLEAR MULAI/SELESAI + DECISION CACHE ----
        mulai_selesai_cols = find_mulai_selesai_columns(ws, used, max_col)

        # Cache row decisions HERE (CRITICAL FIX)
        row_should_be_gray = {}

        for r in date_rows:
            gray = False
            for c in range(1, max_col + 1):
                v = ws.Cells(r, c).Value
                if isinstance(v, (datetime, date)) and v.day in GRAY_DAYS:
                    gray = True
                    break

            row_should_be_gray[r] = gray

            # Paint row
            for c in range(1, max_col + 1):
                ws.Cells(r, c).Interior.ColorIndex = COLOR_GRAY if gray else COLOR_WHITE

            # Clear mulai / selesai
            for c in mulai_selesai_cols:
                ws.Cells(r, c).Value = None

        # ---- STEP 3: RIGHTMOST COLUMN (LOGIC FROM CACHE, NOT EXCEL) ----
        for r in date_rows:
            cell = ws.Cells(r, max_col)

            if row_should_be_gray[r]:
                # Gray row → force empty
                cell.Value = None
                continue

            # White row → autofill if empty
            if cell.Value in (None, "") and not cell.HasFormula:
                src_row = find_nearest_formula_above(ws, max_col, r)
                if src_row is not None:
                    src = ws.Cells(src_row, max_col)
                    dest = ws.Range(src, cell)

                    # Autofill formula (enum 0)
                    src.AutoFill(dest, 0)

                    # Force color back to white
                    cell.Interior.ColorIndex = COLOR_WHITE

        time.sleep(SCAN_INTERVAL_SECONDS)

    except KeyboardInterrupt:
        print("Stopped by user.")
        break

    except Exception as e:
        print("Runtime warning:", e)
        time.sleep(1)
