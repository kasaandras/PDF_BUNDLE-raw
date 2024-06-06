import os
from PyPDF2 import PdfMerger, PdfReader, PdfWriter
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas

def pdf_bundle(pdf_files, output_pdf):
    print("Starting pdf_bundle")

    merger = PdfMerger()

    try:
        for pdf_file in pdf_files:
            merger.append(pdf_file)
        merger.write(output_pdf)
        merger.close()

    except Exception as e:
        print("An error occurred during PDF bundling:", str(e))

    print("Finished pdf_bundle")

def create_page_pdf(num, tmp):
    c = canvas.Canvas(tmp, pagesize=letter)
    for i in range(1, num + 1):
        c.setFont('Courier', 9)
        c.drawCentredString(396, 20, f"Page {i} of {num}")
        c.showPage()
    c.save()

def add_page_numbers(pdf_path):
    tmp = "__tmp.pdf"
    writer = PdfWriter()
    print("Started paging!")
    with open(pdf_path, "rb") as f:
        reader = PdfReader(f)
        n = len(reader.pages)

        create_page_pdf(n, tmp)

        with open(tmp, "rb") as ftmp:
            number_pdf = PdfReader(ftmp)

            for p in range(n):
                page = reader.pages[p]
                number_layer = number_pdf.pages[p]
                page.merge_page(number_layer)
                writer.add_page(page)

            if len(writer.pages) > 0:
                with open(pdf_path, "wb") as f:
                    writer.write(f)
    print("Finished paging!")
    os.remove(tmp)
