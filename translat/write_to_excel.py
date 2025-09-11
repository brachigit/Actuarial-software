import os
from is_formula_valid import is_formula_valid
import win32com.client


def write_to_excel(formula, target_cell="A1", excel_file=None, sheet_name=None):
    valid = is_formula_valid(formula)  # <-- מחזיר True/False
    if valid:
        excel_path = os.path.abspath(os.path.join("outputs", excel_file))


        if not excel_file or not os.path.exists(excel_path):
         raise FileNotFoundError(f"Excel file not found: {excel_path}")
        if not sheet_name:
            raise ValueError("Sheet name must be specified")

        excel = win32com.client.Dispatch("Excel.Application")
        wb = excel.Workbooks.Open(excel_path)
        ws = wb.Sheets(sheet_name)
        if valid:
            ws.Range(target_cell).Formula = formula  # <-- כאן מכניסים את הנוסחה לתא


        wb.Save()
        wb.Close()
        excel.Quit()

