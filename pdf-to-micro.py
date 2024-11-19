from PyPDF2 import PdfReader
from docx import Document
import pdfplumber
from openpyxl import Workbook
import tkinter as tk


def pdf_to_docx(pdf_path, docx_path):
    reader = PdfReader(pdf_path)
    doc = Document()
    for page in reader.pages:
        doc.add_paragraph(page.extract_text())
    doc.save(docx_path)

def pdf_to_xlsx(pdf_path, xlsx_path):
    wb = Workbook()
    ws = wb.active
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            for table in page.extract_tables():
                for row in table:
                    ws.append(row)
    wb.save(xlsx_path)

pdf_to_docx('HDS10M.pdf', 'HDS10M.docx')
pdf_to_xlsx('HDS10M.pdf', 'HDS10M.xlsx')
