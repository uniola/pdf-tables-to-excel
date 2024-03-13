import datetime
import sys
from pathlib import Path

import openpyxl
import pdfplumber


def convert_pdf_to_excel(pdf_path: str, output_path: str) -> None:
    """Convert a PDF file to an Excel file.

    Args:
        pdf_path (str): The path to the PDF file.
        output_path (str): The path to save the Excel file.
    """
    # Read the PDF file
    with pdfplumber.open(pdf_path) as pdf:
        workbook = openpyxl.Workbook()
        for i, page in enumerate(pdf.pages):
            # Create a new sheet
            sheet = (
                workbook.active
                if i == 0
                else workbook.create_sheet(title=f"Sheet{i+1}")
            )
            # Write table to sheet
            table = page.extract_table()
            for row in table:
                sheet.append(row)
    # Save the Excel file
    workbook.save(output_path)


def main() -> int:
    # Get the PDF file path from standard input
    if len(sys.argv) < 2:
        print("Usage: pdf-tables-to-excel [PDF_FILE]")
        sys.exit(1)

    pdf_path = Path(sys.argv[1])

    # Exist?
    if not pdf_path.exists():
        print("File does not exist.")
        sys.exit(1)

    # PDF?
    if pdf_path.suffix.lower() != ".pdf":
        print("Not a PDF file.")
        sys.exit(1)

    # Output path
    timestamp = datetime.datetime.now().strftime("%Y%m%d%H%M%S")
    output_path = pdf_path.with_name(f"{pdf_path.stem}_{timestamp}.xlsx")

    # Convert the PDF to Excel
    convert_pdf_to_excel(pdf_path, output_path)
    return 0
