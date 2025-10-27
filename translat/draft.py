import fitz  # pip install PyMuPDF


def print_pdf_page(pdf_path, page_number):
    """
    pdf_path: נתיב הקובץ
    page_number: מספר העמוד להתחיל מ-1
    """
    doc = fitz.open(pdf_path)

    # בדיקה אם העמוד קיים
    if page_number < 1 or page_number > len(doc):
        print(f"עמוד {page_number} לא קיים בקובץ. הקובץ מכיל {len(doc)} עמודים.")
        return

    # אחזור העמוד (index מתחיל מ-0)
    page = doc[page_number - 1]

    # חילוץ הטקסט
    text = page.get_text()

    print(f"--- תוכן עמוד {page_number} ---")
    print(text)
    print(f"--- סוף עמוד {page_number} ---")


# דוגמה לקריאה לפונקציה
print_pdf_page("C:\\Users\\user\\Downloads\\AuditReport With Lookups.pdf", 1416)


