import time
import os
import re
import win32com.client
from datetime import datetime, date

# ---------- CONFIG ----------
FILE_PATH = os.path.abspath("./test.xlsx")
SCAN_INTERVAL_SECONDS = 3
EDIT_DELAY_SECONDS = 0.5

GRAY_DAYS = {1, 3, 4, 10, 11, 16, 17, 18, 24, 25, 31}

# Regex to detect formula like =(J3+O3)*(1+1)*E3+1
FORMULA_PATTERN = re.compile(
    r"^\=\([A-Z]+[0-9]+\+[A-Z]+[0-9]+\)\*\(1\+1\)\*[A-Z]+[0-9]+\+1$",
    re.IGNORECASE
)
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


def find_rightmost_column_by_formula(ws, used):
    """
    Find the column index that contains the target formula pattern.
    """
    max_col_found = None

    for r in range(1, used.Rows.Count + 1):
        for c in range(1, used.Columns.Count + 1):
            try:
                cell = ws.Cells(r, c)
                if cell.HasFormula:
                    formula = cell.Formula
                    if isinstance(formula, str) and FORMULA_PATTERN.match(formula):
                        max_col_found = c
            except Exception:
                pass

    return max_col_found


excel = get_excel_app()
wb = None

print("Script running (formula-bounded table, NO autosave, slow edit)...")

# ---------- MAIN LOOP (NEVER STOPS) ----------
while True:
    try:
        if not os.path.exists(FILE_PATH):
            print("Excel file not found. Waiting...")
            time.sleep(SCAN_INTERVAL_SECONDS)
            continue

        if wb is None:
            for book in excel.Workbooks:
                if book.FullName.lower() == FILE_PATH.lower():
                    wb = book
                    print("Workbook already open.")
                    break

            if wb is None:
                wb = excel.Workbooks.Open(FILE_PATH)
                print("Workbook opened by script.")

        if not excel_is_ready(excel):
            time.sleep(1)
            continue

        ws = wb.ActiveSheet
        used = ws.UsedRange

        max_col = find_rightmost_column_by_formula(ws, used)
        if max_col is None:
            print("Formula boundary not found. Waiting...")
            time.sleep(SCAN_INTERVAL_SECONDS)
            continue

        found_31_jan = False

        # ---- STEP 1: CHANGE DECEMBER -> JANUARY (SLOW) ----
        for r in range(1, used.Rows.Count + 1):
            for c in range(1, max_col + 1):
                cell = ws.Cells(r, c)
                val = cell.Value

                if isinstance(val, (datetime, date)):
                    if val.month == 12:
                        cell.Select()
                        cell.Value = val.replace(month=1)
                        time.sleep(EDIT_DELAY_SECONDS)

                    if val.month == 1 and val.day == 31:
                        found_31_jan = True

        # ---- STEP 2: ROW COLORING (DATE ROWS ONLY, BOUNDED) ----
        if found_31_jan:
            print("31 Januari detected. Applying bounded row coloring...")

            for r in range(1, used.Rows.Count + 1):
                row_has_date = False
                row_should_be_gray = False

                for c in range(1, max_col + 1):
                    val = ws.Cells(r, c).Value
                    if isinstance(val, (datetime, date)):
                        row_has_date = True
                        if val.day in GRAY_DAYS:
                            row_should_be_gray = True

                if row_has_date:
                    for c in range(1, max_col + 1):
                        cell = ws.Cells(r, c)
                        if row_should_be_gray:
                            cell.Interior.ColorIndex = 15  # Gray
                        else:
                            cell.Interior.ColorIndex = 2   # White

            print("Row coloring applied (formula-bounded).")

    except KeyboardInterrupt:
        print("Script stopped by user.")
        break

    except Exception as e:
        print("Runtime warning (ignored):", e)

    time.sleep(SCAN_INTERVAL_SECONDS)
