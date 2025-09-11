import fitz  # PyMuPDF
import re
import os


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
def search_variable_Input_Manager_in_pdf(pdf_path, variable_name):
    """
    סורקת את ה-PDF החל מעמוד ה-Input Page ומחפשת את המשתנה הנתון.

    אופן הפעולה:
    1. מתחילה את החיפוש מעמוד ה-Input Page
    2. בכל עמוד, מחפשת את כל הטווחים בין "Associated Code Variables:" ל-"Modified On:"
    3. בכל טווח כזה, מחפשת את המשתנה הנתון
    4. אם נמצא, מחזירה את הערך המעודכן ומדפיסה את מספר העמוד
    5. מפסיקה את החיפוש אם מגיעה ל-Data Page או אם נמצא המשתנה

    Args:
        pdf_path: נתיב לקובץ ה-PDF
        variable_name: שם המשתנה לחיפוש

    Returns:
        str: הערך המעודכן של המשתנה אם נמצא, אחרת None
    """
    start_page = extract_first_str_in_table_of_contents(pdf_path, "Input Page")
    if start_page is None:
        print("Input Page not found in the PDF.")
        return None

    print(f"Starting search from Input Manager at page {start_page}")

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
                    for item in current_range:
                        # חלוקת הערכים לפי פסיקים
                        values = [v.strip() for v in item.split(',')]
                        if variable_name in values:
                            # חיפוש הערך המעודכן אחרי "Input Variable:"
                            for l in lines:
                                if "Input Variable:" in l:
                                    updated_value = l.split("Input Variable:", 1)[1].strip()
                                    print(f"Found matching variable '{variable_name}' on page {page_num + 1}")
                                    return updated_value

                    # איפוס הטווח הנוכחי
                    current_range = []

                # אם אנחנו בתוך טווח, מוסיפים את השורה הנוכחית
                elif in_range and line:  # מוסיפים רק שורות לא ריקות
                    current_range.append(line)

    print(f"Variable '{variable_name}' was not found in the PDF.")
    return None


def Search_variable_Lookup_Settings(pdf_path, variable_Lookup_Settings):
    """
    מוצאת את מספר העמוד של 'Lookup Settings' בקובץ ה-PDF ומחפשת את המשתנה הנתון.

    Args:
        pdf_path: נתיב לקובץ ה-PDF
        variable_Lookup_Settings: שם המשתנה לחיפוש

    Returns:
        int: מספר העמוד בו נמצא המשתנה או None אם לא נמצא
    """
    # חיפוש מספר העמוד של 'Lookup Settings'
    lookup_settings_page = extract_first_str_in_table_of_contents(pdf_path, "Lookup Settings")
    if lookup_settings_page is None:
        print("Lookup Settings page not found in the PDF.")
        return None

    print(f"Lookup Settings found on page: {lookup_settings_page}")

    with fitz.open(pdf_path) as doc:
        # מתחילים את החיפוש מעמוד ה-Lookup Settings
        for page_num in range(lookup_settings_page - 1, doc.page_count):
            page = doc[page_num]
            text = page.get_text()
            lines = text.split('\n')

            # סריקת כל שורה בעמוד
            for i in range(len(lines)):
                line = lines[i].strip()

                # בדיקה אם השורה הבאה מכילה "Description"
                if i < len(lines) - 1 and "Description" in lines[i + 1]:
                    # בדיקה אם השורה הנוכחית תואמת למשתנה המבוקש
                    if line == variable_Lookup_Settings:
                        print(f"Found matching variable '{variable_Lookup_Settings}' on page {page_num + 1}")
                        return page_num + 1

    print(f"Variable '{variable_Lookup_Settings}' was not found in the Lookup Settings section.")
    return None

    return lookup_settings_page

def return_variable_Lookup_Settings():

    base_path = r"uploads"
    # רשימת הקבצים בתיקייה
    files = os.listdir(base_path)
    # נניח שאת רוצה את הראשון
    if files:
        pdf_path = os.path.join(base_path, files[0])  #בצורה הזו שם הקובץ אינדוודואל
        pdf_num=search_variable_Input_Manager_in_pdf(pdf_path, "ann_inv_rate_m")
        print(pdf_num)

        Search_variable_Lookup_Settings(pdf_path, "inv_rates")

return_variable_Lookup_Settings()