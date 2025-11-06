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

#  הפונקציה נועדה למצוא למצוא בעזרת Input Variable: Input Variable: את Code Variable
def search_variable_Associated_Code_in_pdf(pdf_path, variable_name):
    """
    סורקת את ה-PDF ומחפשת את המשתנה שמופיע אחרי Input Variable:
    ואז מחלצת את הערך שמופיע אחרי Associated Code Variables:
    """

    start_page = extract_first_str_in_table_of_contents(pdf_path, "Data Page")
    if start_page is None:
        print("Input Page not found in the PDF.")
        return None

    print(f"Starting search from Input Variable at page {start_page}")

    with fitz.open(pdf_path) as doc:
        for page_num in range(start_page - 1, doc.page_count):
            page = doc[page_num]
            text = page.get_text()
            lines = text.split('\n')

            for i, line in enumerate(lines):
                if "Input Variable:" in line:
                    input_var = line.split("Input Variable:")[1].strip()
                    if input_var == variable_name:
                        # חיפוש שורה שבה מופיע Associated Code Variables:
                        for j in range(i + 1, len(lines)):
                            if "Associated Code Variables:" in lines[j]:
                                # בודקת האם המשתנה באותה שורה או בשורה הבאה
                                if j + 1 < len(lines) and not lines[j + 1].strip().startswith(
                                        ("Modified On:", "Input Variable:")):
                                    assoc_value = lines[j + 1].strip()
                                else:
                                    assoc_value = lines[j].split("Associated Code Variables:")[1].strip()
                                print(f"Found variable '{variable_name}' → Associated Code: {assoc_value}")
                                return assoc_value
                        print(f"Associated Code Variables not found after '{variable_name}' on page {page_num + 1}")
                        return None

    print(f"Variable '{variable_name}' not found in the PDF.")
    return None
def classify_lookup_tables(tables):
    row_table = None
    col_table = None

    for table in tables:
        if not table or not isinstance(table, list) or not table[0]:
            continue

        header_text = " ".join(table[0]).lower()
        sample_text = " ".join([" ".join(r) for r in table[:2]]).lower()

        # 🔍 ננסה לזהות לפי מילים אופייניות
        if "row" in header_text or "row lookup" in sample_text:
            row_table = table
        elif "col" in header_text or "column lookup" in sample_text:
            col_table = table

    # אם עדיין לא זוהו, ננסה לפי סוג הנתונים:
    if not row_table and not col_table and len(tables) == 2:
        # נניח שהשנייה היא Row
        col_table, row_table = tables

    return row_table, col_table



def process_lookup_logic(pdf_path, excel_input_path, excel_output_path, description_value, tables):
    """
    🧩 פונקציה זו מממשת את לוגיקת העבודה המלאה על פי האפיון.
    היא משלבת בין טבלאות ה-Row וה-Column שנשלפו מה-PDF,
    ומבצעת חיפוש, השוואות, ושליפת ערכים מקבצי האקסל.
    """

    # ---------------------------------------------------------
    # שלב 1 – פתיחת הלשונית המתאימה לפי description_value
    # ---------------------------------------------------------
    print(f"📘 Opening sheet '{description_value}' in Excel file: {excel_input_path}")
    xl = pd.ExcelFile(excel_input_path)
    if description_value not in xl.sheet_names:
        print(f"❌ Sheet '{description_value}' not found in {excel_input_path}")
        return None

    df_input = xl.parse(description_value)
    print(f"✅ Loaded sheet with {len(df_input)} rows and {len(df_input.columns)} columns")

    # ---------------------------------------------------------
    # שלב 2 – זיהוי טבלאות Row ו-Column
    # ---------------------------------------------------------
    row_table, column_table = None, None
    for t in tables:
        header = [h.lower() for h in t[0]]
        if any("row" in h for h in header):
            row_table = pd.DataFrame(t[1:], columns=t[0])
        elif any("column" in h for h in header):
            column_table = pd.DataFrame(t[1:], columns=t[0])

    print("\n🔍 Debug: Checking extracted lookup tables:")
    for idx, tbl in enumerate(tables, start=1):
        if not tbl:
            continue
        header = tbl[0]
        print(f"Table {idx} header: {header[:5]} ...")

    # כאן בד"כ נעשית הסיווג של הטבלאות:
    row_table, col_table = classify_lookup_tables(tables)

    if column_table is None or row_table is None:
        print("❌ Missing Row or Column table.")
        return None

    # ---------------------------------------------------------
    # שלב 3 – קריאת טבלת Column (עמודה אחת בלבד)
    # ---------------------------------------------------------
    print("🔹 Processing Column table")
    col_lookup_term = column_table.loc[0, "Lookup term"]
    col_target_column = column_table.loc[0, "Column"]

    print(f"   Lookup term: {col_lookup_term}")
    print(f"   Target column in Excel: {col_target_column}")

    # ---------------------------------------------------------
    # שלב 4 – ריצה על טבלת Row כדי למצוא שורה מתאימה באקסל
    # ---------------------------------------------------------
    print("🔹 Processing Row table")
    row_terms = row_table["Lookup term"].tolist()

    matched_row_idx = None
    start_col_idx = 0  # נתחיל מהעמודה הראשונה

    for idx in range(len(df_input)):  # מעבר על שורות באקסל
        match = True
        current_col_idx = start_col_idx  # בכל שורה נתחיל מהעמודה הנוכחית
        for term in row_terms:
            try:
                resolved_term = resolve_lookup_term(pdf_path, excel_output_path, term)
            except Exception as e:
                print(f"⚠️ Failed to resolve term '{term}': {e}")
                resolved_term = None

            # אם לא הצליח – דילוג לשורה הבאה בטבלת row וגם לעמודה הבאה באקסל
            if not resolved_term or resolved_term == "":
                print(f"➡️ Skipping row term '{term}' — unresolved value.")
                match = False
                start_col_idx += 1  # התקדמות לעמודה הבאה
                break

            if current_col_idx >= len(df_input.columns):
                match = False
                break

            cell_value = str(df_input.iloc[idx, current_col_idx]).strip()
            if str(cell_value).strip() != str(resolved_term).strip():
                match = False
                break
            current_col_idx += 1  # התקדמות לעמודה הבאה

        if match:
            matched_row_idx = idx
            print(f"✅ Found matching row at index {idx}")
            break

    if matched_row_idx is None:
        print("⚠️ No matching row found in Excel.")
        return None

    # ---------------------------------------------------------
    # שלב 5 – שליפת הערך מהעמודה שהתקבלה מטבלת Column
    # ---------------------------------------------------------
    if col_target_column not in df_input.columns:
        print(f"❌ Target column '{col_target_column}' not found in Excel sheet.")
        return None

    final_value = df_input.loc[matched_row_idx, col_target_column]
    print(f"🎯 Final extracted value: {final_value}")
    return final_value


# ---------------------------------------------------------
# פונקציה פנימית – פירוש הערך שבעמודת Lookup term
# ---------------------------------------------------------
def resolve_lookup_term(pdf_path, excel_output_path, term):
    """
    מבצעת פירוש של הערכים בעמודת Lookup term בהתאם לסוגם:
    Constant, Code Scalar, Input Variable וכו'.
    """
    term = str(term).strip()

    # 1️⃣ Constant: <*code_variable*>
    if term.startswith("Constant: <") and ">" in term:
        variable_name = re.search(r"<(.*?)>", term).group(1)
        print(f"🔍 Constant variable detected: {variable_name}")
        return variable_name

    # 2️⃣ Constant: "some text"
    elif term.startswith('Constant: "'):
        match = re.search(r'Constant:\s*"(.*?)"', term)
        if match:
            text_value = match.group(1)
            print(f"🔍 Constant text detected: {text_value}")
            return text_value

    # 3️⃣ Code Scalar: <variable> : <model>
    elif term.startswith("Code Scalar:"):
        match = re.search(r"<(.*?)>\s*:\s*<(.*?)>", term)
        if match:
            var_name, model_name = match.groups()
            print(f"🔍 Code Scalar detected → var={var_name}, model={model_name}")
            sheet_name = f"{model_name}_cflow_Scalars"
            try:
                df = pd.read_excel(excel_output_path, sheet_name=sheet_name)
                if var_name in df.columns:
                    value = df[var_name].iloc[0]
                    print(f"   Value from {sheet_name}.{var_name}: {value}")
                    return value
            except Exception as e:
                print(f"⚠️ Failed to load sheet {sheet_name}: {e}")

    # 4️⃣ Input Variable: <namemodel>_Data: input variable
    elif term.startswith("Input Variable:"):
        match = re.search(r"Input Variable:\s*<(.*?)>_Data: input variable", term)
        if match:
            model_name = match.group(1)
            print(f"🔍 Input Variable detected: model={model_name}")

            variable_name = f"{model_name}_Data"
            X = search_variable_Associated_Code_in_pdf(pdf_path, variable_name)
            if not X:
                return None

            sheet_name = f"{model_name}_Data"
            try:
                df = pd.read_excel(excel_output_path, sheet_name=sheet_name)
                found_row = df[df.iloc[:, 0] == X]
                if not found_row.empty and "value" in df.columns:
                    value = found_row["value"].iloc[0]
                    print(f"   Found value for {X}: {value}")
                    return value
            except Exception as e:
                print(f"⚠️ Failed to read sheet {sheet_name}: {e}")

    # ברירת מחדל – החזר כמו שהוא
    return term

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


def find_description_after_anchor(pdf_path, page_num, bbox):
    """
    מחפשת את הטקסט שמופיע אחרי 'Description:' באותה שורה (או בקרבת מקום)
    """
    doc = fitz.open(pdf_path)
    page = doc[page_num]
    blocks = page.get_text("blocks")

    for (x0, y0, x1, y1, text, *_ ) in blocks:
        if "Description:" in text:
            # נחלץ את מה שאחרי המילה Description:
            match = re.search(r"Description:\s*(.*)", text)
            if match:
                return match.group(1).strip()
    return None

def extract_description_same_line(pdf_path, start_page):
    """
    מחפש את המחרוזת שמופיעה אחרי 'Description:' באותה שורה בלבד.
    מחזיר את הטקסט שאחרי Description או None אם לא נמצא.
    """

    doc = fitz.open(pdf_path)
    for page_num in range(start_page, len(doc)):
        page = doc[page_num]
        blocks = page.get_text("blocks")

        for block in blocks:
            text = block[4]
            if "Description:" in text:
                # נחפש את הטקסט שמופיע אחרי Description:
                match = re.search(r"Description:\s*(.*)", text)
                if match:
                    return match.group(1).strip()
                else:
                    # לא נמצא טקסט אחרי הנקודותיים באותה שורה
                    return None
    return None


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

        print(table_data)

        # שלב 5: העמוד באקסל ממנו יחולץ ערך assumption
        print("Step 5:Description")

        description_value = extract_description_same_line(pdf_path, start_page=page_num + 1)
        print(f"📘 Description found: {description_value}")

        excel_input_path = r"C:\Users\user\Downloads\Main assumptions - variable - blank.xlsx"
        excel_output_path = r"C:\Users\user\Downloads\output (1).xlsx"
        result_value = process_lookup_logic(
            pdf_path,
            excel_input_path,
            excel_output_path,
            description_value,
            table_data
        )
        print(result_value)

    except Exception as e:
        print(f"An unexpected error occurred: {e}")
        import traceback
        traceback.print_exc()
        print(None)

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
        search_variable_Input_Manager_in_pdf(pdf_path, "takeup_age")


if __name__ == "__main__":
    main()
    # Call the lookup settings function if needed
