import logging
import win32com.client

logging.basicConfig(filename='formula_check.log', level=logging.INFO,
                    format='%(asctime)s - %(levelname)s - %(message)s')
# מילון עם קודי שגיאה נפוצים של Excel והסבר קריא
EXCEL_ERROR_CODES = {
    -2146827284: "Syntax error in the formula",
    -2146827283: "Division by zero",
    -2146826273: "Cell reference error (#REF!)",
    -2146826246: "Name not recognized (#NAME?)",
    -2146826248: "Value error (#VALUE!)",
    -2146826265: "Number error (#NUM!)",
    # אפשר להוסיף עוד קודים לפי הצורך
}

def is_formula_valid(formula: str) -> bool:

    excel = win32com.client.Dispatch("Excel.Application")
    wb = excel.Workbooks.Add()  # חוברת עבודה זמנית
    ws = wb.Sheets(1)

    try:
        # מנסים להכניס את הנוסחה עם אופציה שתמנע הערכת תאים חסרים
        ws.Cells(1, 1).Formula = formula
        valid = True
        logging.info(f'Formula "{formula}" is valid')  # <-- כאן מוסיפים רישום הצלחה
    except Exception as e:
        valid = False
        details = str(e)

        # בדיקה של קוד פנימי מתוך details
        error_code = None
        if hasattr(e, 'args') and e.args:
            # נסיון לקחת את הקוד האמיתי מתוך tuple פנימי של details
            try:
                error_code = e.args[2][5]  # כאן נמצא הקוד הפנימי (-2146827284)
            except (IndexError, TypeError):
                error_code = e.args[0] if isinstance(e.args[0], int) else None

        human_readable = EXCEL_ERROR_CODES.get(error_code, "Unknown error")

        error_message = (
            f"Formula '{formula}' is invalid. "
            f"Exception type: {type(e).__name__}, "
            f"Error code: {error_code}, "
            f"Description: {human_readable}, "
            f"Details: {details}"
        )
        logging.error(error_message)
        print(error_message)



    finally:
     wb.Close(SaveChanges=False)
     excel.Quit()
    return valid
