import openpyxl
from docx import Document
import os
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn
import datetime

# Step 1: Read data from Excel file
def read_excel_data(excel_file_path):
    wb = openpyxl.load_workbook(excel_file_path)
    ws = wb.active
    data = []

    # Read header
    headers = [cell.value for cell in ws[1]]
    data.append(headers)

    # Read data rows
    for row in ws.iter_rows(min_row=2, values_only=True):
        formatted_row = []
        for cell in row:
            if isinstance(cell, (datetime.datetime, datetime.date)):
                formatted_row.append(cell.strftime('%Y-%m-%d'))  # Convert date to string
            else:
                formatted_row.append(str(cell))
        data.append(formatted_row)

    return data


# Step 2: Replace placeholders in Word document
def replace_placeholders(template_path, output_path, replacements):
    doc = Document(template_path)

    def replace_text_in_paragraph(paragraph, replacements):
        for placeholder, replacement in replacements.items():
            if placeholder in paragraph.text:
                paragraph.text = paragraph.text.replace(placeholder, replacement)
                

    def replace_text_in_table(table, replacements):
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    replace_text_in_paragraph(paragraph, replacements)

    # Replace in paragraphs
    for paragraph in doc.paragraphs:
        replace_text_in_paragraph(paragraph, replacements)

    # Replace in tables x`
    for table in doc.tables:
        replace_text_in_table(table, replacements)

    doc.save(output_path)


# Step 3: Main function to read from Excel and generate Word documents
def fill_word_template_from_excel(excel_file_path, template_doc_path, output_dir):
    data = read_excel_data(excel_file_path)
    headers = data[0]
    rows = data[1:]

    for row in rows:
        replacements = {
            "{{Name and Surname}}": row[0],
            "{{Passport}}": row[3],
            "{{DoB}}": row[1],
            "{{Citizenship}}": row[2],
            "{{Uni}}": row[4]
        }

        output_file_name = f"{row[0].replace(' ', '_')}.docx"
        output_file_path = os.path.join(output_dir, output_file_name)
        replace_placeholders(template_doc_path, output_file_path, replacements)
        print(f"Created document: {output_file_name}")

if __name__ == "__main__":

    excel_file_path = "path/to/excel"  # Path to your Excel file
    template_doc_path = "path/to/template"  # Path to your Word template
    output_dir ="path/to/output/folder"  # Directory to save the generated documents

    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    fill_word_template_from_excel(excel_file_path, template_doc_path, output_dir)