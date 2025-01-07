import docx
import re
import os

def extract_text_with_line_numbers(docx_path):
    try:
        doc = docx.Document(docx_path)
        full_text = []
        for line_number, para in enumerate(doc.paragraphs):
            line_with_number = f"{line_number + 1}: {para.text}"
            full_text.append(line_with_number)
        return full_text
    except Exception as e:
        print(f"Failed to open and read the file: {e}")
        return []

def extract_specific_data(lines):
    data = {
        'Vendor Name': 'Cypress Creek Pest Control of Texas',
        'Invoice Number': None,
        'Invoice Date': None,
        'Amount': None,
        'Check Number': None
    }

    for i in range(10, 14): 
        if 'Invoice' in lines[i]:
            invoice_match = re.search(r'Invoice (\d+)', lines[i].split(": ")[1]) 
            if invoice_match:
                data['Invoice Number'] = invoice_match.group(1) 
                break 

    for i in range(11, 15): 
        elements = lines[i].split(": ")[1].split() 
        if len(elements) == 3 and re.match(r'\d{6}', elements[0]) and re.match(r'\d{2}/\d{2}/\d{4}', elements[1]) and re.match(r'\$\d+\.\d{2}', elements[2]): 
            data['Invoice Date'] = elements[1] 
            data['Amount'] = elements[2] 
            break

    return data

def process_document(file_path, check_number):
    if not os.path.exists(file_path):
        print(f"File not found: {file_path}")
        return None

    extracted_lines = extract_text_with_line_numbers(file_path)
    if not extracted_lines:
        print("No lines extracted. Please check the file content.")
        return None

    data = extract_specific_data(extracted_lines)
    
    # Add Check Number if provided
    data['Check Number'] = check_number

    return data

if __name__ == "__main__":
    processed_data = process_document("sample.docx", "12345")
    if processed_data:
        print("\n--- Processed Document Data ---")
        for key, value in processed_data.items():
            print(f"{key}: {value}")
