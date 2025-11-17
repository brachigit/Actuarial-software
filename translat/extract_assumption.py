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


# הפונקציה עורכת חיפוש בתוכן עניינים האם קיימת המחרוזתstart_str
# ומחזירה את מספר העמוד בו נמצאה המחרוזת
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
                if match:
                    page_number = int(match.group(1))
                    return page_number  # מחזיר את הראשון שנמצא
    print("No start_str found in the range of pages 1 to 5")
    return None


# חיפוש המשתנה מסוג External Source  שהתקבל +הפונקציה שולחת לפונקציה extract_first_input_page  כדי לדעת באיזה תווך לערוך את החיפוש
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
                                        return Search_variable_Lookup_Settings(pdf_path, updated_value, variable_name,
                                                                               depth + 1)

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

        header_text = " ".join((cell[0] if isinstance(cell, list) else str(cell)) for cell in table[0]).lower()
        sample_text = " ".join([
            " ".join(
                cell if isinstance(cell, str) else " ".join(cell) if isinstance(cell, list) else str(cell)
                for cell in r)
            for r in table[:2] ]).lower()

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

def detect_header_by_reverse_scan(df):
    """
    מוצאת את שורת הכותרת כך:
    1. מוצאים את השורה האחרונה שמלאה בנתונים.
    2. עולים כלפי מעלה עד שמוצאים שורה ריקה.
    3. השורה מתחת לשורה הריקה היא כותרת.
    """

    # פונקציה שמזהה האם שורה "מלאה" או "ריקה"
    def is_row_empty(row):
        # ריקה = כל התאים ריקים / NaN / *
        for cell in row:
            if isinstance(cell, str):
                if cell.strip() not in ["", "*"]:  # * נחשב ריק אצלך
                    return False
            elif pd.notna(cell):
                return False
        return True

    last_data_row = None

    # שלב 1: מוצאים את השורה האחרונה המלאה בנתונים
    for i in reversed(range(len(df))):
        if not is_row_empty(df.iloc[i]):
            last_data_row = i
            break

    if last_data_row is None:
        return 0  # fallback

    # שלב 2: עולים כלפי מעלה עד שפוגשים שורה ריקה
    for r in reversed(range(0, last_data_row)):
        if is_row_empty(df.iloc[r]):
            # השורה שמתחתיה היא הכותרת
            return r + 1

    return 0  # fallback אם הכול מלא


def process_lookup_logic(pdf_path, excel_input_path, excel_output_path, description_value, tables,input_variable=None):
    """
    🧩 פונקציה זו מממשת את לוגיקת העבודה המלאה על פי האפיון.
    היא משלבת בין טבלאות ה-Row וה-Column שנשלפו מה-PDF,
    ומבצעת חיפוש, השוואות, ושליפת ערכים מקבצי האקסל.
    """
    print(f"🔍 Constant variable detected: {input_variable}")
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

    header_row = detect_header_by_reverse_scan(df_input)
    print(f"📌 Detected REAL header row at index: {header_row}")

    # הגדרה מחדש של ה־header + reset index
    df_input.columns = df_input.iloc[header_row]
    df_input = df_input.iloc[header_row + 1:].reset_index(drop=True)

    print(f"🔄 Sheet rebuilt using detected header row. New shape: {df_input.shape}")

    # ---------------------------------------------------------
    # שלב 2 – זיהוי טבלאות Row ו-Column
    # ---------------------------------------------------------
    # שלב 2 – זיהוי טבלאות Row ו-Column לפי הסדר
    # ---------------------------------------------------------
    print("\n🔍 Debug: Checking extracted lookup tables:")
    for idx, tbl in enumerate(tables, start=1):
        if not tbl:
            continue
        header = tbl[0]
        print(f"Table {idx} header: {header[:5]} ...")

    if len(tables) < 2:
        print("❌ Expected at least two tables (Column and Row).")
        return None

    # הנחה: הטבלה הראשונה היא Column, השנייה היא Row
    column_table_raw = tables[0]
    row_table_raw = tables[1]

    # Flatten nested lists if needed
    if len(column_table_raw) == 1 and isinstance(column_table_raw[0], list):
        column_table_raw = column_table_raw[0]
    if len(row_table_raw) == 1 and isinstance(row_table_raw[0], list):
        row_table_raw = row_table_raw[0]

    try:
        print("📄 Raw Column Table:", column_table_raw)
        # אם יש רמה מיותרת ברשימה, נחלץ אותה

        row_table = pd.DataFrame(row_table_raw[1:], columns=row_table_raw[0])
        # יצירת DataFrame מהטבלה הראשונה
        column_table = pd.DataFrame(column_table_raw[1:], columns=column_table_raw[0])

        if row_table.empty:
            print("❌ Row table is empty – cannot extract row lookup term.")
            return None
        print("📊 Debug: row_table table content:")
        print(row_table)
        print("📊 Debug: Column table content:")
        print(column_table)

        # בדיקה האם יש שורות בטבלה
        if column_table.empty:
            print("❌ Column table is empty – cannot extract lookup term.")
            return None

        # במקום לגשת ישירות עם loc[0]
        try:
            col_lookup_term = column_table.iloc[0]["Lookup term"]
            print(f"✅ Lookup term from Column table: {col_lookup_term}")
        except Exception as e:
            print(f"❌ Failed to extract Lookup term from Column table: {e}")
            return None

    except Exception as e:
        print(f"❌ Failed to convert tables to DataFrames: {e}")
        return None

    # נוודא שאינדקסים מתחילים מ-0 כדי למנוע KeyError
    column_table.reset_index(drop=True, inplace=True)
    row_table.reset_index(drop=True, inplace=True)

    print(f"✅ Column table shape: {column_table.shape}")
    print(f"✅ Row table shape: {row_table.shape}")

    # ---------------------------------------------------------
    # שלב 3 – קריאת טבלת Column (עמודה אחת בלבד)
    # ---------------------------------------------------------
    print("\n📋 Debug: Columns in Column table:", list(column_table.columns))
    print(column_table.head())
    print("🔹 Processing Column table")

    col_target_columns = column_table["Lookup term"].tolist()
    col_target_rows = row_table["Lookup term"].tolist()

    print("🔹 Lookup terms (Column table):", col_target_columns)
    print("🔹 Lookup terms (Row table):", col_target_rows)

    resolved_columns = []

    print("🔹 Lookup terms (Column table):", col_target_columns)

    for term in col_target_columns:
        try:
            resolved_value = resolve_lookup_term(pdf_path, excel_output_path, term, input_variable)
            resolved_columns.append(resolved_value)
        except Exception as e:
            print(f"⚠️ Failed to resolve column term '{term}': {e}")
            resolved_columns.append(None)

    print("✅ Resolved column terms:",  resolved_columns)

    # יצירת רשימת תנאי השורה לאחר Resolve אמיתי
    resolved_row_terms = []
    for term in col_target_rows:
        try:
            resolved_term = resolve_lookup_term(pdf_path, excel_output_path, term, input_variable)
            resolved_row_terms.append(resolved_term)
        except Exception as e:
            print(f"⚠️ Failed to resolve row term '{term}': {e}")
            resolved_row_terms.append(None)

    print("✅ Final resolved row terms:", resolved_row_terms)

    # ---------------------------------------------------------
    # שלב 4 – חיפוש שורה מתאימה לפי resolved_row_terms
    # ---------------------------------------------------------
    matched_row_idx = None
    start_col_idx = 0  # נתחיל מהעמודה הראשונה

    for idx in range(len(df_input)):  # מעבר על שורות האקסל
        match = True
        current_col_idx = start_col_idx

        for term in resolved_row_terms:
            # אם הערך לא קיים – אין התאמה
            if term is None or term == "":
                match = False
                start_col_idx += 1  # דילוג לעמודה הבאה
                break

            # בדיקה שלא יצאנו מגבול העמודות
            if current_col_idx >= len(df_input.columns):
                match = False
                break

            cell_value = str(df_input.iloc[idx, current_col_idx]).strip()
            if str(cell_value) != str(term):
                match = False
                break

            current_col_idx += 1  # התקדמות לעמודה הבאה

        if match:
            matched_row_idx = idx
            print(f"✅ Found matching row at index {idx}")
            break

    if matched_row_idx is None:
        print("❌ No matching row found after applying lookup logic.")
        return None

    # ---------------------------------------------------------
    # שלב 5 – שליפת הערך לפי העמודה שנפתרה
    # ---------------------------------------------------------

    print("Columns in df_input:", list(df_input.columns))

    target_column = resolved_columns[0]

    if target_column not in df_input.columns:
        print(f"❌ Column '{target_column}' not found in Excel sheet.")
        return None

    final_value = df_input.loc[matched_row_idx, target_column]
    print(f"🎯 Final extracted value: {final_value}")

    return final_value






# ---------------------------------------------------------
# פונקציה פנימית – פירוש הערך שבעמודת Lookup term
# ---------------------------------------------------------
def resolve_lookup_term(pdf_path, excel_output_path, term, input_variable):
    """
    מבצעת פירוש של הערכים בעמודת Lookup term בהתאם לסוגם:
    Constant, Code Scalar, Input Variable וכו'.
    """
    print(f"🔍 Constant variable detected: {input_variable}")


    # ✳️ שלב ראשון — סינון לפני כל המרה או עיבוד
    if not isinstance(term, str) or not any(
        str(term).startswith(prefix) for prefix in ["Constant", "Code Scalar", "Input Variable"]
    ):
        print(f"⚠️ term לא רלוונטי או לא מחרוזת: {term}")
        return term

    # ✳️ רק עכשיו ננקה את המחרוזת
    term = term.replace("\n", " ").strip()


    # 1️⃣ Constant: <*code_variable*>
    if term.startswith("Constant: <") and ">" in term:
        return input_variable

    # 2️⃣ Constant: "some text"
    elif term.startswith('Constant: "'):
        match = re.search(r'Constant:\s*"(.*?)"', term)
        if match:
            text_value = match.group(1)
            print(f"🔍 Constant text detected: {text_value}")
            return text_value
            # 2️⃣ Constant: Assumption (ללא גרשיים)
    elif re.match(r'Constant:\s*\S', term) and "<" not in term and ">" not in term:
         text_value = re.sub(r'^Constant:\s*', '', term).strip()
         print(f"🔍 Constant text (no quotes) detected: {text_value}")
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
        # חילוץ פרטים מהמחרוזת לפי מבנה: Input Variable: life_Data: life: paid_up
        parts = term.split(":")
        if len(parts) >= 4:
            sheet_name = parts[1].strip()  # life_Data
            model_name = parts[2].strip()  # life
            variable_name = parts[3].strip()  # paid_up

            print(f"🔍 Input Variable detected: model={model_name}, variable={variable_name}")

            X = search_variable_Associated_Code_in_pdf(pdf_path, variable_name)
            if not X:
                return None
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
    """
    מחפש את שם המשתנה החל מעמוד ו/או מיקום נתון (bbox)
    ומחזיר את ה־bbox ואת מספר העמוד שבו נמצא.
    אם לא סופק bbox, החיפוש יתחיל מתחילת הדף.
    """
    doc = fitz.open(pdf_path)

    for page_num in range(start_page - 1, len(doc)):
        page = doc[page_num]


        matches = [r for r in page.search_for(variable_name)]

        if matches:
            match = matches[0]
            print(f"✅ Found '{variable_name}' on page {page_num + 1}")
            print(f"   BBox: {match}")
            return match, page_num

    print(f"❌ Variable '{variable_name}' not found in any page starting from {start_page}.")
    return None, None




def find_next_variable_anchor(pdf_path, variable_name, start_page=1, current_bbox=None):
    """
    מחפשת את המשתנה הבא אחרי המשתנה הנוכחי — מאותה שורה ומטה, ואם לא נמצא, בעמודים הבאים
    """
    doc = fitz.open(pdf_path)
    current_y = current_bbox.y0 if current_bbox else 0

    for page_num in range(start_page-1, len(doc)):
        page = doc[page_num]

        # חיפוש מדויק לפי שם המשתנה
        matches = [r for r in page.search_for(variable_name)]

        if matches:
            # אם אנחנו בעמוד ההתחלה – לדלג על מיקומים מעל current_y
            if page_num == start_page - 1:
                matches = [m for m in matches if m.y1 > current_y]

            if matches:
                bbox = matches[0]
                print(f"✅ Found variable '{variable_name}' on page {page_num + 1}")
                print(f"   BBox: {bbox}")
                return bbox, page_num

    # אם לא נמצא
    print(f"⚠️ Variable '{variable_name}' not found. Search ended at page {page_num + 1}.")
    return None, None

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

            """
            # ❌ אם כבר הופיעה בעבר אחת מהמילים — לא נמשיך לחלץ טבלאות
           elif found_column_marker or found_row_marker:
                print(f"⏭️ Skipping page {page_num + 1} – tables ignored after first 'Column'/'Row'")
                continue
                """
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

                        # נזהה אם מדובר בטבלה הראשונה (Column → Row)
                        if found_column_marker and not found_row_marker:
                            # טבלה ראשונה
                            table_group = "column"
                        elif found_row_marker:
                            # טבלה שנייה
                            table_group = "row"
                        else:
                            table_group = "unknown"

                        # שמירת טבלה לפי סוגה
                        if table_group == "column":
                            current_tables = all_tables
                        elif table_group == "row":
                            current_tables = all_tables
                        else:
                            current_tables = all_tables

                        # איחוד טבלאות בעלות אותה כותרת
                        if not current_tables:
                            current_tables.append(table)
                        else:
                            header_existing = current_tables[-1][0]
                            header_new = table[0]
                            ratio = SequenceMatcher(None, ",".join(header_existing), ",".join(header_new)).ratio()

                            if ratio > 0.9:
                                print(f"🔗 Merging continuation of {table_group} table (ratio={ratio:.2f})")
                                current_tables[-1].extend(table[1:])
                            else:
                                print(f"➕ Starting new {table_group} table (ratio={ratio:.2f})")
                                current_tables.append(table)



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
    bboxRow, page_Row = find_next_variable_anchor(pdf_path, "Row Lookup Details:", start_page + 1,main_bbox)
    print(f"📘📘 page_Column_index={page_Row}, start_page_index={start_page}")
    print(f"📘📘 page_Column_display={page_Row + 1}, start_page_display={start_page + 1}")

    # 2️⃣ חילוץ שתי טבלאות לפי גבולות העוגנים
    print(f"Extracting Column Lookup Details section...start_page{start_page+1} end_page,{ end_page+1}")
    table_Column = extract_tables_from_bbox(
        pdf_path,
        start_page=start_page,
        end_page=end_page,
        bbox=main_bbox,
        next_bbox=bboxRow
    )

    table_Row = extract_tables_from_bbox(
        pdf_path,
        start_page=page_Row,
        end_page=end_page,
        bbox=bboxRow,
        next_bbox=next_bbox
    )


    return table_Column, table_Row


def find_description_after_bbox(pdf_path, page_num, bbox):
    doc = fitz.open(pdf_path)
    page = doc[page_num]
    blocks = page.get_text("blocks")

    start_y = bbox[1]  # נתחיל רק מהגובה של ה־bbox
    for (x0, y0, x1, y1, text, *_) in blocks:
        if y0 >= start_y and "Description:" in text:
            match = re.search(r"Description:\s*(.*)", text)
            if match:
                return match.group(1).strip()
            else:
                desc_y = y0
                # נחפש בלוקים באותו y
                for (bx0, by0, bx1, by1, btext, *_) in blocks:
                    if abs(by0 - desc_y) < 2 and bx0 > x1:
                        return btext.strip()
    return None


def Search_variable_Lookup_Settings(pdf_path, variable_Lookup_Settings, input_variable=None, depth=0, max_depth=3):
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

        next_bbox, next_page = find_next_variable_anchor(pdf_path, "Types of Annuity Prop", page_num + 1, bbox)

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

        description_value = find_description_after_bbox(pdf_path,page_num,bbox)
        print(f"📘 Description found: {description_value}")

        excel_input_path = r"C:\Users\user\Downloads\Main assumptions - variable - blank.xlsx"
        excel_output_path = r"C:\Users\user\Downloads\output (1).xlsx"
        result_value = process_lookup_logic(
            pdf_path,
            excel_input_path,
            excel_output_path,
            description_value,
            table_data,
            input_variable
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
        pdf_path = os.path.join(base_path, files[0])  # בצורה הזו שם הקובץ אינדוודואל


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