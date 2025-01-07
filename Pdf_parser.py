import tkinter as tk
from tkinter import filedialog, simpledialog
import fitz  # PyMuPDF
import re

def select_pdf_file():
    root = tk.Tk()
    root.withdraw()
    pdf_path = filedialog.askopenfilename(
        title="Select PDF File",
        filetypes=[("PDF Files", "*.pdf")]
    )
    return pdf_path

def read_pdf(pdf_path):
    try:
        document = fitz.open(pdf_path)
        full_text = []
        for page_num in range(len(document)):
            page = document.load_page(page_num)
            page_text = page.get_text("text")
            full_text.append(page_text)
        return "\n".join(full_text)
    except Exception as e:
        print(f"Failed to open and read the file: {e}")
        return ""

def print_with_line_numbers(text):
    lines = text.split('\n')
    for line_number, line in enumerate(lines, start=1):
        print(f"{line_number}: {line}")

def extract_data(text):
    data = {
        "Vendor Name": "Cypress Creek Pest Control",
        "Invoice Number": None,
        "Billing Date": None,
        "Amount Due": None,
        "Check Number": None
    }

    # Split the text into lines
    lines = text.split('\n')

    # Search for the invoice number within a range of lines
    invoice_pattern = re.compile(r'Invoice\s+(\d+)', re.IGNORECASE)
    for i in range(25, min(35, len(lines))):
        match = invoice_pattern.search(lines[i])
        if match:
            data["Invoice Number"] = match.group(1)
            break

    # Define pattern for Billing Date (first occurrence of mm/dd/yyyy)
    date_pattern = re.compile(r'(\d{2}/\d{2}/\d{4})')
    date_match = date_pattern.search(text)
    if date_match:
        data["Billing Date"] = date_match.group(1)

    # Search for Amount Due
    amount_pattern = re.compile(r'Amount Due', re.IGNORECASE)
    monetary_pattern = re.compile(r'\$[\d,]+\.\d{2}')

    for i, line in enumerate(lines):
        if amount_pattern.search(line):
            # Check the same line for a monetary value
            same_line_match = monetary_pattern.search(line)
            if same_line_match:
                data["Amount Due"] = same_line_match.group(0)
                break
            # Check the next line for a monetary value
            elif i + 1 < len(lines):
                next_line_match = monetary_pattern.search(lines[i + 1])
                if next_line_match:
                    data["Amount Due"] = next_line_match.group(0)
                    break

    return data

def main():
    pdf_file_path = select_pdf_file()
    if pdf_file_path:
        text = read_pdf(pdf_file_path)
        
        # Print the initial document content with line numbers
        print("\n--- Initial Document Content ---")
        print_with_line_numbers(text)
        
        extracted_data = extract_data(text)
        
        # User input for Check Number
        root = tk.Tk()
        root.withdraw()
        check_number = simpledialog.askstring("Input", "Enter Check Number:")
        extracted_data["Check Number"] = check_number
        
        # Print extracted data
        print("\n--- Extracted Data ---")
        for key, value in extracted_data.items():
            print(f"{key}: {value}")
    else:
        print("No PDF file selected.")

if __name__ == "__main__":
    main()
