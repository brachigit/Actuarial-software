import fitz  # PyMuPDF
import re
import os
import pandas as pd
from pathlib import Path
from pdfminer.high_level import extract_pages
from pdfminer.layout import LTTextContainer, LTChar
from collections import defaultdict
import pdfplumber



#הפונקציה עורכת חיפוש בתוכן עניינים האם קיימת המחרוזתstart_str
#ומחזירה את מספר העמוד בו נמצאה המחרוזת
def extract_first_str_in_table_of_contents(pdf_path, start_str="Input Page"):
    doc = fitz.open(pdf_path)

    # טווח העמודים: עמודים 1 עד 5 (באינדקסים 0 עד 4)
    for page_num in range(0, 5):
        page = doc[page_num]
        text = page.get_text("text")
        lines = text.split("\n")

        for line in lines:
            if start_str in line:
                match = re.search(r"(\d+)\s*$", line)
                if match :
                    page_number = int(match.group(1))
                    return page_number  # מחזיר את הראשון שנמצא
    print("No start_str found in the range of pages 1 to 5")
    return None

#חיפוש המשתנה מסוג External Source  שהתקבל +הפונקציה שולחת לפונקציה extract_first_input_page  כדי לדעת באיזה תווך לערוך את החיפוש
def search_variable_Input_Manager_in_pdf(pdf_path, variable_name, depth=0):
    """
    סורקת את ה-PDF החל מעמוד ה-Input Page ומחפשת את המשתנה הנתון.
    
    Args:
        pdf_path: נתיב לקובץ ה-PDF
        variable_name: שם המשתנה לחיפוש
        depth: עומק הרקורסיה הנוכחי
        
    Returns:
        str: הערך המעודכן של המשתנה אם נמצא, אחרת None
    """
    start_page = extract_first_str_in_table_of_contents(pdf_path, "Input Page")
    if start_page is None:
        print("Input Page not found in the PDF.")
        return None

    print(f"Starting search from Input Manager at page {start_page} (Depth: {depth})")

    with fitz.open(pdf_path) as doc:
        # מתחילים את הסריקה מעמוד ה-Input Page
        for page_num in range(start_page - 1, doc.page_count):
            page = doc[page_num]
            text = page.get_text()

            # בדיקה אם הגענו ל-Data Page
            if "Data Page" in text:
                print(f"Reached Data Page at page {page_num + 1}. Stopping search.")
                break

            # חלוקת הטקסט לשורות
            lines = text.split('\n')

            # איתחול משתנים למעקב אחר טווחים
            in_range = False
            current_range = []

            # סריקת כל שורה בעמוד
            for line in lines:
                line = line.strip()

                # אם מצאנו התחלה של טווח חדש
                if "Associated Code Variables:" in line:
                    in_range = True
                    current_range = []
                    continue

                # אם הגענו לסוף הטווח הנוכחי
                if "Modified On:" in line and in_range:
                    in_range = False

                    # בדיקת המשתנה בטווח הנוכחי
                    for item_idx, item in enumerate(current_range):
                        # חלוקת הערכים לפי פסיקים
                        values = [v.strip() for v in item.split(',')]
                        if variable_name in values:
                            # חיפוש למעלה בשורות הקודמות את ה-Input Variable
                            current_line_text = item.strip()
                            current_line_in_page = -1
                            for i, l in enumerate(lines):
                                if l.strip() == current_line_text:
                                    current_line_in_page = i
                                    break
                            
                            if current_line_in_page >= 0:
                                # חיפוש למעלה מהשורה הנוכחית עד לתחילת העמוד
                                for i in range(current_line_in_page - 1, -1, -1):
                                    line_text = lines[i].strip()
                                    if "Input Variable:" in line_text:
                                        updated_value = line_text.split("Input Variable:", 1)[1].strip()
                                        print(f"Found matching variable '{variable_name}' on page {page_num + 1}")
                                        print(f"Input Variable value: {updated_value}")
                                        # קריאה לפונקציה עם הערך המעודכן
                                        return Search_variable_Lookup_Settings(pdf_path, updated_value, variable_name, depth+1)

                    # איפוס הטווח הנוכחי
                    current_range = []

                # אם אנחנו בתוך טווח, מוסיפים את השורה הנוכחית
                elif in_range and line:  # מוסיפים רק שורות לא ריקות
                    current_range.append(line)

    print(f"Variable '{variable_name}' was not found in the PDF.")
    return None


def parse_table_line(line, headers, prev_columns=None):
    """
    פונקציה שמנתחת שורת טבלה ומחזירה מילון של עמודה-ערך
    """
    if not line.strip() or ':' in line:
        return None

    # ניסיון ראשון: פיצול לפי רווחים כפולים
    columns = [col.strip() for col in line.split('  ') if col.strip()]
    
    # אם יש מספיק עמודות, נחזיר אותן
    if len(columns) >= len(headers):
        return dict(zip(headers, columns[:len(headers)]))
    
    # אם אין מספיק עמודות, ננסה פיצול אחר
    if len(columns) < len(headers) and prev_columns is not None:
        # מנסים למזג עם השורה הקודמת
        merged = prev_columns.copy()
        for i, col in enumerate(columns):
            if i < len(merged):
                merged[i] = (merged[i] + ' ' + col).strip()
            else:
                merged.append(col)
        if len(merged) >= len(headers):
            return dict(zip(headers, merged[:len(headers)]))
    
    return None



def Search_variable_Lookup_Settings(pdf_path, variable_Lookup_Settings, variable_name=None, depth=0, max_depth=3):
    """
    Finds and parses the Column/Row Lookup tables for a specific variable,
    using an optimized search that starts from the relevant section in the PDF.
    """
    try:
        # שלב 1: איתור מהיר של אזור החיפוש
        print("Step 1: Finding the 'Lookup Settings' section via table of contents...")
        start_page_num = extract_first_str_in_table_of_contents(pdf_path, "Lookup Settings")
        if start_page_num is None:
            print("Error: Could not find 'Lookup Settings' in the document's table of contents.")
            return None

        # שלב 2: קריאת טקסט ממוקדת
        print(f"Step 2: Starting focused text extraction from page {start_page_num} onwards (with a buffer).")
        doc = fitz.open(pdf_path)
        full_text = ""  # <-- אתחול המשתנה
        search_from_page_index = max(0, start_page_num - 1)  # <-- הגדרת המשתנה
        for i in range(search_from_page_index, len(doc)):
            full_text += doc[i].get_text() + "\n"

        # ניקוי רווחים כפולים ושורות ריקות
        lines = []
        for line in full_text.split('\n'):
            line = ' '.join(line.split())  # מסיר רווחים כפולים
            if line.strip():  # מתעלם משורות ריקות
                lines.append(line)

        # הדפסה לניפוי (ניתן להסיר אחרי הבדיקה)
        print(f"Found {len(lines)} lines of text")
        if lines:
            print("Sample of text found:", lines[:5])  # 5 השורות הראשונות

        # =================================================================================
        # שלב 3 (גמיש): מציאת העוגן בשורות עוקבות
        # =================================================================================
        print(f"Step 3 (Flexible): Finding anchor for '{variable_Lookup_Settings}' across adjacent lines.")

        start_anchor_index = -1
        end_anchor_index = -1

        # --- מציאת העוגן ההתחלתי (בדיקה על שתי שורות) ---
        for i in range(len(lines) - 1):  # לולאה עד השורה הלפני אחרונה
            current_line = lines[i]
            next_line = lines[i + 1]

            # תנאי 1: האם בשורה הנוכחית יש תבנית מספור?
            has_numbering = re.search(r'\d+(\.\d+){2,}', current_line)

            # תנאי 2: האם בשורה הבאה יש את שם המשתנה?
            has_variable_name = (variable_Lookup_Settings in current_line) or (variable_Lookup_Settings in next_line)

            if has_numbering and has_variable_name:
                start_anchor_index = i  # העוגן הוא שורת המספור
                break

        if start_anchor_index == -1:
            print(f"Error: Could not find the anchor line for '{variable_Lookup_Settings}'.")
            return None

        # --- מציאת העוגן הסופי (המשתנה הבא) ---
        for i in range(start_anchor_index + 1, len(lines)):
            line = lines[i]
            if re.search(r'\d+(\.\d+){2,}', line):
                end_anchor_index = i
                break

        # --- חיתוך אזור החיפוש המדויק ---
        if end_anchor_index != -1:
            section_lines = lines[start_anchor_index: end_anchor_index]
        else:
            section_lines = lines[start_anchor_index: start_anchor_index + 50]
        # =================================================================================
        # שלב 4: חיפוש תת-הכותרת בתוך האזור המדויק
        # =================================================================================
        print("Step 4: Extracting table from section...")

        table_lines = []
        header_found = False
        headers = []

        for line in section_lines:
            clean_line = line.strip()

            # דילוג על שורות ריקות
            if not clean_line:
                continue

            # זיהוי התחלה של טבלה לפי מילים כמו "Column" או "Key" או רווחים מרובים
            if re.search(r'\b(Column|Key|Value|Name)\b', clean_line) or re.search(r'\s{2,}', clean_line):
                header_found = True
                headers = re.split(r'\s{2,}', clean_line)
                continue

            if header_found:
                # חילוץ שורות טבלה - כל עוד זה נראה כמו שורה עם רווחים מרובים
                if re.search(r'\s{2,}', clean_line):
                    row = re.split(r'\s{2,}', clean_line)
                    # השלמת עמודות חסרות באורך הכותרת
                    if len(row) < len(headers):
                        row += [''] * (len(headers) - len(row))
                    table_lines.append(row)
                else:
                    # אם השורה כבר לא נראית כחלק מהטבלה — מפסיקים
                    if len(table_lines) > 0:
                        break

        # --- החזרת התוצאה ---
        if not table_lines:
            print("No table found in this section.")
            return None

        print(f"Extracted table with {len(table_lines)} rows and {len(headers)} columns.")
        return [headers] + table_lines

    except Exception as e:
        print(f"An unexpected error occurred: {e}")
        import traceback
        traceback.print_exc()
        return None


def return_variable_Lookup_Settings():

    base_path = r"uploads"
    # רשימת הקבצים בתיקייה
    files = os.listdir(base_path)
    # נניח שאת רוצה את הראשון
    if files:
        pdf_path = os.path.join(base_path, files[0])  #בצורה הזו שם הקובץ אינדוודואל
        pdf_num=search_variable_Input_Manager_in_pdf(pdf_path, "res_prop_old_data")
        print(pdf_num)



def main():
    base_path = r"uploads"
    # רשימת הקבצים בתיקייה
    files = os.listdir(base_path)
    # נניח שאת רוצה את הראשון
    if files:
        pdf_path = os.path.join(base_path, files[0])  # בצורה הזו שם הקובץ אינדוודואל
        pdf_num = search_variable_Input_Manager_in_pdf(pdf_path, "res_prop_old_data")
        print(pdf_num)

if __name__ == "__main__":
    main()
    # Call the lookup settings function if needed
    return_variable_Lookup_Settings()