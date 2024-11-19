from PyPDF2 import PdfReader
from docx import Document
import pdfplumber
from openpyxl import Workbook
import tkinter as tk
import os


def pdf_to_docx(pdf_path, docx_path):
    reader = PdfReader(pdf_path)
    doc = Document()
    for page in reader.pages:
        doc.add_paragraph(page.extract_text())
    doc.save(docx_path)

def pdf_to_xlsx(pdf_path, xlsx_path, header_font_size=10):
    wb = Workbook()
    ws = wb.active
    # make sure the xlsx file at the filepath is closed before editing
    try:
        # Try to open the file in exclusive mode
        with open(xlsx_path, 'r+'):
            pass
    except IOError:
        # print error statement in red
        print(f"\033[91mError: The file {xlsx_path} is open. Please close it and try again.\033[0m")
        return

    with pdfplumber.open(pdf_path) as pdf:
        table_continued = False
        for page in pdf.pages:
            text = page.extract_text().split('\n')
            tables = page.extract_tables()
            # Extract headings based on font size, type, and indentation
            headings = []
            current_heading = ""
            for char in page.chars:
                if char['size'] > header_font_size:  # Font size
                    current_heading += char['text']
                elif current_heading:
                    headings.append(current_heading.strip())
                    current_heading = ""
            if current_heading:
                headings.append(current_heading.strip())
            # if headings:
            #     ws.append([" ".join(headings)])
            # print(headings)
            for i, table in enumerate(tables):
                # Append headings before each table
                if i < len(headings):
                    ws.append([headings[i]])
                # Append table rows
                for row in table:
                    ws.append(row)
                # Add a blank row after each table
                ws.append([])
                # Check if the table continues on the next page
                if i == len(tables) - 1 and len(table) > 0 and table[-1] == ['']:
                    table_continued = True
                else:
                    table_continued = False
    wb.save(xlsx_path)

#pdf_to_docx('HDS10M.pdf', 'HDS10M.docx')
pdf_to_xlsx('HDS10M.pdf', 'HDS10M.xlsx')

# print("\033[31mThis is red text\033[0m")
# print("\033[32mThis is green text\033[0m")
# print("\033[33mThis is yellow text\033[0m")
# print("\033[34mThis is blue text\033[0m")
# print("\033[35mThis is magenta text\033[0m")
# print("\033[36mThis is cyan text\033[0m")
# print("\033[1mThis is bold text\033[0m")
# print("\033[4mThis is underlined text\033[0m")
# print("\033[7mThis is reversed text\033[0m")
# print("\033[1;31mThis is bold red text\033[0m")
