import fitz  # PyMuPDF
import re
import pandas as pd
import os

pdf_path = "C:\\Users\\user\\Downloads\\AuditReport.pdf"
doc = fitz.open(pdf_path)

start_pattern = r"2\.1\.16\s*Data Page:\s*Ann_Data"
end_pattern = r"2\.1\.19\s*External Sources"

from_page = 72
to_page = 127

extracting = False
collected_text = ""

# שלב 1 – איסוף הטקסט הרלוונטי מהטווח
for page_num in range(from_page, to_page + 1):
    page = doc[page_num]
    text = page.get_text()

    if not extracting:
        start_match = re.search(start_pattern, text)
        if start_match:
            extracting = True
            text = text[start_match.start():]

    if extracting:
        end_match = re.search(end_pattern, text)
        if end_match:
            collected_text += text[:end_match.start()]
            break
        else:
            collected_text += text

doc.close()

# שלב 2 – עיבוד המידע והכנסה לעמודות Excel
lines = [line.strip() for line in collected_text.strip().split('\n') if line.strip()]

code_vars = []
input_values = []
values = []

current_code_var = ""
current_input_value = ""
current_value = ""

i = 0
while i < len(lines):
    line = lines[i]

    if "Input Variable:" in line:
        # שלב 1: חילוץ המחרוזת לאחר Input Variable:
        current_code_var = line.split("Input Variable:")[1].strip()
        code_vars.append(current_code_var)
        # הכנה לשתי העמודות האחרות
        input_values.append("")
        values.append("")
        i += 1
        continue

    if "Value:" in line:
        # בדיקה שהשורה הבאה קיימת
        if i + 1 < len(lines):
            next_line = lines[i + 1].strip()
            # הכנס לתא המתאים (אותו אינדקס של Code Variable האחרון)
            if len(input_values) > 0:
                input_values[-1] = next_line
            i += 2
            continue

    i += 1

# שלב 3 – בניית קובץ Excel
df = pd.DataFrame({
    "Code Variable": code_vars,
    "Input Value": input_values,
    "Value": values  # העמודה תישאר ריקה כרגע כי לא סיפקת מקור לחלץ אותה
})

output_path = "extracted_structured_output.xlsx"
df.to_excel(output_path, index=False)

print(f"הטבלה נשמרה כקובץ Excel בשם {output_path}")
os.startfile(output_path)
