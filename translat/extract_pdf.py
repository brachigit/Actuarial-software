# translat/extract_pdf.py

import fitz  # PyMuPDF
import re
import pandas as pd
import os
from datetime import datetime

def extract_pdf(pdf_path):
    doc = fitz.open(pdf_path)

    start_pattern = r"2\.1\.16\s*Data Page:\s*Ann_Data"
    end_pattern = r"2\.1\.19\s*External Sources"

    from_page = 72
    to_page = 127

    extracting = False
    collected_text = ""

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

    lines = [line.strip() for line in collected_text.strip().split('\n') if line.strip()]

    code_vars = []
    input_values = []
    values = []

    i = 0
    while i < len(lines):
        line = lines[i]

        if "Input Variable:" in line:
            current_code_var = line.split("Input Variable:")[1].strip()
            code_vars.append(current_code_var)
            input_values.append("")
            values.append("")
            i += 1
            continue

        if "Value:" in line:
            if i + 1 < len(lines):
                next_line = lines[i + 1].strip()
                if len(input_values) > 0:
                    input_values[-1] = next_line
                i += 2
                continue

        i += 1

    df = pd.DataFrame({
        "Code Variable": code_vars,
        "Input Value": input_values,
        "Value": values
    })

    # יצירת תיקיית outputs
    outputs_dir = "outputs"
    os.makedirs(outputs_dir, exist_ok=True)
    
    # יצירת שם קובץ ייחודי עם תאריך ושעה
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = os.path.splitext(os.path.basename(pdf_path))[0]
    output_path = os.path.join(outputs_dir, f"{filename}_extracted_{timestamp}.xlsx")
    
    df.to_excel(output_path, index=False)

    print(f"הטבלה נשמרה כקובץ Excel בשם {output_path}")
    
    return output_path

if __name__ == "__main__":
    pdf_path = "C:\\Users\\user\\Downloads\\AuditReport With Lookups.pdf"  # <-- תעדכני לנתיב הנכון
    extract_pdf(pdf_path)