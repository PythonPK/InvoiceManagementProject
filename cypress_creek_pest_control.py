import docx
import re
# from tkinter import Tk
# from tkinter.filedialog import askopenfilename
import os

# def select_file():
#     Tk().withdraw()
#     filename = askopenfilename(
#         title="Select file",
#         filetypes=[("Word Documents", "*.docx")]
#     )
#     return filename

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
        #'Vendor Address': 'Po Box 690548, Houston, TX 77269',
        #'Vendor Phone': '281-469-2679, Fax:281-469-4720',
        #'Vendor Email': 'billing@cycreekpestcontrol.com'
    }

    data['Invoice Number'] = None

    for i in range(10, 14): 
        if 'Invoice' in lines[i]:
            invoice_match = re.search(r'Invoice (\d+)', lines[i].split(": ")[1]) 
            if invoice_match:
                data['Invoice Number'] = invoice_match.group(1) 
                break 

    # 'Account Number': '212736', 
    data['Invoice Date'] = None 
    data['Amount'] = None 
    for i in range(11, 15): 
        elements = lines[i].split(": ")[1].split() 
        if len(elements) == 3 and re.match(r'\d{6}', elements[0]) and re.match(r'\d{2}/\d{2}/\d{4}', elements[1]) and re.match(r'\$\d+\.\d{2}', elements[2]): 
        # data['Account Number'] = elements[0] 
            data['Invoice Date'] = elements[1] 
            data['Amount'] = elements[2] 
            break

    # Commented out Billed to and Billing Address
    # for i in range(5, 12):
    #     if 'Julie Rivers Office Condominiums' in lines[i].split(": ")[1]:
    #         data['Billed To'] = 'Julie Rivers Office Condominiums'
    #         data['Billing Address'] = 'P.o. Box 90669, Houston, TX 77298'
    #         break
    # else:
    #     data['Billed To'] = lines[7].split(": ")[1]
    #     data['Billing Address'] = f"{lines[8].split(': ')[1]} {lines[9].split(': ')[1]}"

    return data

# def get_user_input():
#     # service_location = input("Enter the location where service provided: ")
#     # service_date = input("Enter Service Date (MM/DD/YYYY): ")
#     # description = input("Enter Description: ")
#     # quantity = input("Enter Quantity: ")
#     # amount = float(input("Enter Amount ($0.00): "))
#     check_number = input("Enter Check Number: ")
#     user_data = {
#         'Check Number': check_number
#     }
#     return user_data
    
    # default_tax = round(amount * 0.0825, 2)
    # default_discount = 0.00
    # default_adjustment = 0.00
    # default_total = round(amount + default_tax - default_discount - default_adjustment, 2)

    # tax = input(f"Enter Tax ($0.00, default ${default_tax}): ") or default_tax
    # discount = input(f"Enter Discount ($0.00, default ${default_discount}): ") or default_discount
    # adjustment = input(f"Enter Adjustment ($0.00, default ${default_adjustment}): ") or default_adjustment
    # total = input(f"Enter Total ($0.00, default ${default_total}): ") or default_total

    # user_data = {
    #     'Service Location': service_location,
    #     'Service Date': service_date,
    #     'Description': description,
    #     'Quantity': quantity,
    #     'Amount': f"${amount:.2f}",
    #     'Tax': f"${float(tax):.2f}",
    #     'Discount': f"${float(discount):.2f}",
    #     'Adjustment': f"${float(adjustment):.2f}",
    #     'Total': f"${float(total):.2f}"
    # }

    # return user_data

def process_document(file_path, check_number):
    if not os.path.exists(file_path):
        print(f"File not found: {file_path}")
        return None

    extracted_lines = extract_text_with_line_numbers(file_path)
    if not extracted_lines:
        print("No lines extracted. Please check the file content.")
        return None

    data = extract_specific_data(extracted_lines)
    # Removed user_data in module to be handled from main program.
    # user_data = get_user_input()

    data.update({'Check Number': check_number})
    if check_number: # Only add check number if it is provided
        data.update({'Check Number': check_number})
    return data

if __name__ == "__main__":
    processed_data = process_document()
    if processed_data:
        print("\n--- Processed Document Data ---")
        for key, value in processed_data.items():
            print(f"{key}: {value}")
