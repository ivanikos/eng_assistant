import PyPDF2
import pdfplumber


file_path = r"C:\Users\ivaign\OneDrive - United Conveyor Corp\Documents\Python_Projects\Testing Area\SD-54880-13-1141_1_.pdf"
# Open the PDF file
with open(file_path, 'rb') as file:
    # Initialize a PDF reader object
    pdf_reader = PyPDF2.PdfReader(file)

    # Iterate through each page
    for page_number in range(len(pdf_reader.pages)):
        # Get a page object
        page = pdf_reader.pages[page_number]

        # Extract text from the page
        text = page.extract_text()
        print(f"--- Page {page_number + 1} ---")
        print(text)


with pdfplumber.open(file_path) as pdf:
    # Iterate through each page
    for page_number, page in enumerate(pdf.pages):
        # Extract tables from the page
        tables = page.extract_tables()

        # If tables are found, iterate through them
        for table_number, table in enumerate(tables):
            print(f"--- Page {page_number + 1}, Table {table_number + 1} ---")
            for row in table:
                print(row)  # Each 'row' is a list of cell data


