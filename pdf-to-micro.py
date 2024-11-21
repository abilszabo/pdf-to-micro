import pdfplumber
import tkinter as tk
import os
import re  

from openpyxl import Workbook
from openpyxl.styles import Font
# from docx import Document
# from PyPDF2 import PdfReader

# ================== SUB-FUNCTIONS ================== #
# XLSX FIND HEADER
def find_header(page, table_bbox, table_settings):
    # Define the bounding box for the header
    h_top = max(table_bbox[1] - 35, 0)  # Top boundary of header bbox
    h_x0 = max(table_bbox[0] - 15, 0)  # Left boundary of header bbox
    header_bbox = (h_x0, h_top, table_bbox[2], table_bbox[1])
    x_tol = table_settings["text_x_tolerance"]
    y_tol = table_settings["text_y_tolerance"]
    # Extract words within the bounding box
    words = [
        word["text"] for word in page.extract_words(x_tolerance=x_tol, y_tolerance=y_tol)
        if word["x0"] >= header_bbox[0] and
           word["x1"] <= header_bbox[2] and
           word["top"] >= header_bbox[1] and
           word["bottom"] <= header_bbox[3]
    ]

    # Combine words into a single string
    header = " ".join(words)

    # Return the extracted header text
    return header.strip()



# ================== MAIN FUNCTIONS ================== #

# PDF TO XLSX
def pdf_to_xlsx(pdf_path, xlsx_path):
    wb = Workbook()
    ws = wb.active

    # if make sure the xlsx file at the filepath exists, make sure its closed before editing
    if os.path.exists(xlsx_path):
        try:
            # Try to open the file in exclusive mode
            with open(xlsx_path, 'r+'):
                pass
        except IOError:
            # print error statement in red
            print(f"\033[91mError: The file {xlsx_path} is open. Please close it and try again.\033[0m")
            return
    
    # Ensure the debugging directory exists
    os.makedirs('debugging', exist_ok=True)
    # get the debugging directory path
    debugging_dir = os.path.abspath('debugging')

    # pull the filename from the path without the extension
    filename = os.path.splitext(os.path.basename(pdf_path))[0]

    # set init. conditions and flags
    table_continued = False
    tables_sum = 0

    # open the pdf file
    with pdfplumber.open(pdf_path) as pdf:
        # loop through each page in the pdf
        ws.append([filename])
        for cell in ws[ws.max_row]:
            cell.font = Font(bold=True, size=18)
        ws.append([])

        for page in pdf.pages:
            # width = page.width
            # height = page.height

            # extract table settings
            table_settings = {
                "vertical_strategy": "lines",
                "horizontal_strategy": "lines",
                "snap_y_tolerance": 5,
                "intersection_x_tolerance": 15,
                "text_tolerance": 1,
                "text_x_tolerance": 1,
                "text_y_tolerance": 1
            }   

            # # create image for visual debugging    
            # im = page.to_image(300)
            # im.debug_tablefinder(table_settings)
            # # save to the debugging directory
            # im.save(f"{debugging_dir}/page_{page.page_number}.png")

            
            # find all tables in the page
            tables = page.find_tables(table_settings)
            legitimate_tables = []
            # filter out tables with only 1 row or 1 column
            for table in tables:
                if len(table.columns) > 1 and len(table.rows) > 1:
                    legitimate_tables.append(table)
            
            # loop through each legitimate table
            for i, table in enumerate(legitimate_tables):
                # extract the table's data
                table_valid = False
                data = table.extract(x_tolerance=table_settings["text_x_tolerance"], y_tolerance=table_settings["text_y_tolerance"])
                # check the table isn't empty
                for row in data:
                    for cell in row:
                        if cell is not None and cell != "":
                            table_valid = True

                # for valid, non-null tables with more than 1 row & 1 column
                if table_valid:
                    if not table_continued:
                        tables_sum += 1

                        # pull it's header
                        header = find_header(page, table.bbox, table_settings)
                        # add the header to the worksheet
                        ws.append([header])
                        for cell in ws[ws.max_row]:
                            cell.font = Font(bold=True)
                        ws.append([f"Table {tables_sum}, Page {page.page_number}"])
                        for cell in ws[ws.max_row]:
                            cell.font = Font(italic=True)

                    # add each row to the worksheet
                    for i, row in enumerate(data):
                        ws.append(row)
                        if i == 0 and not table_continued:
                            for cell in ws[ws.max_row]:
                                cell.font = Font(bold=True)

                # check if the table continues on the next page
                if i == len(legitimate_tables) - 1 and len(data) > 0 and data[-1] == ['']:
                    table_continued = True
                else:
                    table_continued = False
                    ws.append([])
                    ws.append([])
      
    wb.save(xlsx_path)
    print(f"\033[45mExcel file saved to {xlsx_path}\033[0m")


# PDF TO DOCX
# def pdf_to_docx(pdf_path, docx_path):
#     reader = PdfReader(pdf_path)
#     doc = Document()
#     for page in reader.pages:
#         doc.add_paragraph(page.extract_text())
#     doc.save(docx_path)



# ================== RUN TESTS ================== #

# get path to this python file
# path = os.path.dirname(os.path.abspath(__file__))

pdf_to_xlsx('example_pdfs/FUTURA-System-Manual.pdf', 'output/FUTURA-System-Manual.xlsx')
# pdf_to_xlsx('example_pdfs/HDS10M.pdf', 'output/HDS10M.xlsx')

# open the xlsx file
os.system('start excel.exe output/FUTURA-System-Manual.xlsx')


