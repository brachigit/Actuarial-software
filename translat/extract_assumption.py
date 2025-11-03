import fitz  # PyMuPDF
import re
import os
import pandas as pd
from pathlib import Path
from pdfminer.high_level import extract_pages
from pdfminer.layout import LTTextContainer, LTChar
from collections import defaultdict
import pdfplumber
from difflib import SequenceMatcher



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

def find_anchor_with_fitz(pdf_path, variable_name, start_page=1):
    """מחפש את שם המשתנה החל מעמוד נתון ומחזיר את ה־bbox ואת מספר העמוד"""
    doc = fitz.open(pdf_path)
    for page_num in range(start_page - 1, len(doc)):
        page = doc[page_num]
        text = page.get_text("text")
        if variable_name in text:
            for match in page.search_for(variable_name):
                print(f"✅ Found '{variable_name}' on page {page_num + 1}")
                print(f"   BBox: {match}")
                return match, page_num
    print(f"❌ Variable '{variable_name}' not found in any page starting from {start_page}.")
    return None, None


def find_next_variable_anchor(pdf_path, current_variable, start_page, current_bbox):
    """
    מחפשת את המשתנה הבא אחרי המשתנה הנוכחי — מאותה שורה ומטה, ואם לא נמצא, בעמודים הבאים
    """
    doc = fitz.open(pdf_path)
    pattern = r'\b\d+(?:\.\d+)*\b\s+[A-Za-z_]\w*\b'
    current_y = current_bbox.y0 if current_bbox else 0  # גובה המשתנה הנוכחי

    for page_num in range(start_page - 1, len(doc)):
        page = doc[page_num]
        blocks = page.get_text("blocks")  # [(x0, y0, x1, y1, text, block_no, block_type, block_flags)]

        for b in blocks:
            x0, y0, x1, y1, text, *_ = b

            # נתעלם מטקסטים שנמצאים *מעל* המשתנה הנוכחי באותו עמוד
            if page_num == start_page - 1 and y1 <= current_y:
                continue

            # נבדוק אם יש שם משתנה אחר (שאינו הנוכחי)
            if re.match(pattern, text.strip()) and current_variable not in text:
                bbox = fitz.Rect(x0, y0, x1, y1)
                print(f"📍 Found next variable '{text.strip()}' on page {page_num + 1}")
                print(f"   BBox: {bbox}")
                return bbox, page_num


    print(f"⚠️ No next variable found — will continue to end of document .")
    return None,None



def clean_table(table):
    """מסיר שורות ריקות ומנקה רווחים."""
    cleaned = []
    for row in table:
        if any(cell and str(cell).strip() for cell in row):
            cleaned.append([str(cell).strip() if cell else "" for cell in row])
    return cleaned



def extract_tables_from_bbox(
    pdf_path,
    start_page,
    end_page,
    bbox,
    top_offset=-30,
    bottom_extension=1000,
    left_extension=150,
    right_extension=700,
    save_debug_images=True,
    debug_prefix="debug_page",
    next_bbox=None
):
    """
    מחלץ טבלאות מטווח עמודים [start_page .. end_page] כשהחיתוך בעמוד ההתחלתי
    מבוסס על bbox (מהשורה/המילה של המשתנה), ובשאר העמודים משתמשים בכל העמוד.

    Args:
        pdf_path (str): נתיב לקובץ PDF
        start_page (int): אינדקס העמוד ההתחלתי (0-indexed)
        end_page (int): אינדקס העמוד הסופי (0-indexed, כולל)
        bbox (tuple): (x0, y0, x1, y1) כפי שמוחזר מ-fit z עבור העוגן
        top_offset (int/float): כמה להרים את הגבול העליון ב-px (יכול להיות שלילי)
        bottom_extension (int/float): כמה להרחיב למטה ב-px
        left_extension (int/float): כמה להרחיב שמאלה ב-px
        right_extension (int/float): כמה להרחיב ימינה ב-px
        save_debug_images (bool): לשמור תמונה מכל עמוד נבדק (debug)
        debug_prefix (str): קידומת לשמות הקבצים שנשמרים

    Returns:
        list[list[list[str]]] | None: רשימה של טבלאות (כל טבלה = רשימת שורות), או None אם לא נמצא דבר
    """
    all_tables = []

    found_row_marker = False  # האם כבר ראינו את המילה Row
    found_column_marker = False  # האם כבר ראינו את המילה Column
    previous_page_had_table = False  # האם בעמוד הקודם נמצאה טבלה

    # בטיחות: אם end_page קטן מ-start_page - נתקן
    if end_page < start_page:
        end_page = start_page

    with pdfplumber.open(pdf_path) as pdf:
        n_pages = len(pdf.pages)
        # הגבלה לתוך תחום הדוק של המסמך
        start_page = max(0, min(start_page, n_pages - 1))
        end_page = max(0, min(end_page, n_pages - 1))

        for page_num in range(start_page, end_page + 1):
            page = pdf.pages[page_num]
            page_width, page_height = page.width, page.height

            text = page.extract_text() or ""
            if "Column" in text and not found_column_marker:
                found_column_marker = True
                print(f"📘 Found 'Column' on page {page_num + 1} – starting new Column table")

            elif "Row" in text and not found_row_marker:
                found_row_marker = True
                print(f"📗 Found 'Row' on page {page_num + 1} – starting Row table")

            # ❌ אם כבר הופיעה בעבר אחת מהמילים — לא נמשיך לחלץ טבלאות
            elif found_column_marker or found_row_marker:
                print(f"⏭️ Skipping page {page_num + 1} – tables ignored after first 'Column'/'Row'")
                continue

            if start_page == end_page:
                # 💡 עמוד יחיד – חתוך בין המשתנה הנוכחי לזה שאחריו
                x0, y0, x1, y1 = bbox
                if next_bbox:
                    nx0, ny0, nx1, ny1 = next_bbox
                    extended_bbox = (0, y0 + top_offset, page_width, ny0)
                else:
                    extended_bbox = (0, y0 + top_offset, page_width, page_height)

                cropped_page = page.within_bbox(extended_bbox)

            elif page_num == start_page:
                # ✂️ עמוד התחלתי — חתוך מלמעלה (מהמשתנה הנוכחי)
                x0, y0, x1, y1 = bbox
                extended_y0 = max(0, y0 + top_offset)
                extended_y1 = min(page_height, y1 + bottom_extension)
                new_x0 = max(0, x0 - left_extension)
                new_x1 = min(page_width, x1 + right_extension)
                extended_bbox = (new_x0, extended_y0, new_x1, extended_y1)
                cropped_page = page.within_bbox(extended_bbox)

            elif page_num == end_page:
                # ✂️ עמוד סופי — חתוך עד למיקום המשתנה הבא בלבד
                if next_bbox:
                    nx0, ny0, nx1, ny1 = next_bbox
                    crop_box = (0, 0, page_width, ny0)  # עד תחילת המשתנה הבא
                    cropped_page = page.within_bbox(crop_box)
                    print(f"✂️ End page {page_num + 1} cropped up to next variable (y={ny0})")
                else:
                    cropped_page = page  # אין משתנה הבא — ניקח את כל העמוד

            else:
                # עמודים שבאמצע — ניקח את כולם
                cropped_page = page

            # חילוץ טבלאות מהעמוד/חתך
            try:
                tables = cropped_page.extract_tables()

                if tables:
                    for table in tables:
                        table = clean_table(table)  # ניקוי שורות ריקות ורווחים
                        if not table:
                            continue

                        if not all_tables:
                            all_tables.append(table)
                            continue

                        # השווה כותרת לטבלה האחרונה
                        header_existing = all_tables[-1][0]
                        header_new = table[0]

                        # נשתמש בהשוואה גמישה (למקרה של רווחים או הבדלים קטנים)
                        ratio = SequenceMatcher(None, ",".join(header_existing), ",".join(header_new)).ratio()

                        if ratio > 0.9:
                            # כותרות זהות מספיק — נאחד בלי לשכפל את הכותרת
                            print(f"🔗 Merging table with similar header (ratio={ratio:.2f})")
                            all_tables[-1].extend(table[1:])
                        else:
                            # כותרת שונה — נתחיל טבלה חדשה
                            print(f"➕ New table detected (header difference ratio={ratio:.2f})")
                            all_tables.append(table)


            except Exception as e:
                print(f"⚠️ extract_tables failed on page {page_num + 1}: {e}")
                tables = None

            # לשמירת תמונת בדיקה של האזור הנבדק (מאוד שימושי לדיבאג)
            if save_debug_images:
                try:
                    img_name = f"{debug_prefix}_{page_num + 1}.png"
                    cropped_page.to_image(resolution=150).save(img_name)
                    print(f"🖼️ Saved debug image: {img_name}")
                except Exception as e:
                    print(f"⚠️ Failed to save debug image for page {page_num + 1}: {e}")

            if tables:
                print(f"✅ Found {len(tables)} table(s) on page {page_num + 1}")
            else:
                print(f"⚠️ No tables found on page {page_num + 1}")

    if not all_tables:
        print("❌ No tables found in the entire range.")
        return None

    print(f"✅ Total extracted tables: {len(all_tables)} from pages {start_page + 1}–{end_page + 1}")
    all_tables = [clean_table(t) for t in all_tables if t]
    return all_tables

def extract_lookup_sections(pdf_path, start_page, end_page, main_bbox, next_bbox=None):
    """
    מחלק את טווח ה־Lookup לשני חלקים לפי עוגני:
    'Column Lookup Details' → 'Row Lookup Details'
    'Row Lookup Details' → המשתנה הבא או סוף הקובץ.
    """
    # 1️⃣ מצא את מיקום העוגנים
    column_anchor, column_page = find_anchor_with_fitz(pdf_path, "Column Lookup Details", start_page)
    row_anchor, row_page = find_anchor_with_fitz(pdf_path, "Row Lookup Details", start_page)

    if not column_anchor or not row_anchor:
        print("⚠️ Could not find both anchors — falling back to regular extraction.")
        return extract_tables_from_bbox(pdf_path, start_page, end_page, main_bbox, next_bbox=next_bbox)

    # 2️⃣ חילוץ שתי טבלאות לפי גבולות העוגנים
    print("Extracting Column Lookup Details section...")
    table_1 = extract_tables_from_bbox(
        pdf_path,
        start_page=column_page,
        end_page=row_page,
        bbox=column_anchor,
        next_bbox=row_anchor
    )

    print("Extracting Row Lookup Details section...")
    table_2 = extract_tables_from_bbox(
        pdf_path,
        start_page=row_page,
        end_page=end_page,
        bbox=row_anchor,
        next_bbox=next_bbox
    )

    return (table_1 or []) + (table_2 or [])



def Search_variable_Lookup_Settings(pdf_path, variable_Lookup_Settings, variable_name=None, depth=0, max_depth=3):
    """
    מוצא ומחלץ את טבלת ה־Lookup עבור משתנה מסוים.
    מבוסס על איתור מיקום 'Lookup Settings' בתוכן העניינים, ואז חיפוש ממוקד של המשתנה.
    """
    try:
        # שלב 1: איתור אזור "Lookup Settings" מתוך תוכן העניינים
        print("Step 1: Finding the 'Lookup Settings' section via table of contents...")
        start_page_num = extract_first_str_in_table_of_contents(pdf_path, "Lookup Settings")
        if start_page_num is None:
            print("Error: Could not find 'Lookup Settings' in the document's table of contents.")
            return None

        print(f"Starting from page {start_page_num} for anchor search...")

        # שלב 2: איתור העוגן של המשתנה בעזרת bbox
        bbox, page_num = find_anchor_with_fitz(pdf_path, variable_Lookup_Settings, start_page=start_page_num)
        if bbox is None:
            print("Anchor not found.")
            return None

        next_bbox, next_page = find_next_variable_anchor(pdf_path, variable_Lookup_Settings, page_num + 1,bbox)

        # אם לא נמצא משתנה נוסף – נגדיר שהעמוד האחרון הוא הגבול
        with fitz.open(pdf_path) as doc:
            last_page_num = len(doc) - 1
        end_page = next_page if next_page is not None else last_page_num

        print(f"Extracting tables from page {page_num + 1} up to {end_page + 1}...")

        # שלב 3: חילוץ טבלה מהאזור שסביב המשתנה
        print("Step 3: Extracting tables near the anchor...")
        table_data = extract_lookup_sections(pdf_path, page_num, end_page, bbox, next_bbox)
        if not table_data:
            print("No tables found in the anchor area.")
            return None

        # שלב 4: (אופציונלי) עיבוד הנתונים מתוך הטבלה
        print("Step 4: Processing extracted table data...")
        for i, table in enumerate(table_data, 1):
            print(f"\nTable {i}:")
            for row in table:
                print(row)

        return table_data

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
