# pdf-to-micro
 python desktop app to convert pdfs to microsoft compatible file formats (.docx and .xlsx)

## install dependencies
``` bash
pip install PyPDF2 python-docx openpyxl pdfplumber PyMuPDF fitz
```

## build application
``` bash
pyinstaller --onefile --noconsole app.py
```
