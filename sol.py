import fitz
from openpyxl import Workbook

def extract_data_from_pdf(pdf_path,page_no):
    doc=fitz.open(pdf_path)
    page=doc.load_page(page_no)
    blocks = page.get_text("blocks")
    
    sorted_blocks = sorted(blocks, key=lambda b: (b[1], b[0]))
    
    column=[]
    current_column=[sorted_blocks[0]]
    for block in sorted_blocks[1:]:
        if abs(block[0] - current_column[-1][0])<10:
            current_column.append(block)
        else:
            column.append(current_column)
            current_column=[block]
    column.append(current_column)
    
    
    paragraphs = []
    for column in column:
        column.sort(key=lambda b: b[1])
        paragraph = ""
        for block in column:
            paragraph += block[4] + " "
        paragraphs.append(paragraph.strip())

    return paragraphs

def export_to_excel(paragraphs, excel_path):
    wb = Workbook()
    ws = wb.active

    for i, block in enumerate(paragraphs, start=1):
        ws.cell(row=i, column=1, value=block)

    wb.save(excel_path)
    
    



pdf_path="D:\python project\walnut ai assignement 1\keppel-corporation-limited-annual-report-2018.pdf"
page_no=12
excel_path = "D:\python project\walnut ai assignement 1\document.xlsx"

paragraphs=extract_data_from_pdf(pdf_path,page_no)
export_to_excel(paragraphs, excel_path)