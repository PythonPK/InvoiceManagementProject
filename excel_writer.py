import openpyxl
import tkinter as tk
from tkinter import messagebox

def open_workbook(file_path):
    try:
        workbook = openpyxl.load_workbook(file_path)
        print(f"Workbook '{file_path}' loaded successfully.")
    except FileNotFoundError:
        workbook = openpyxl.Workbook()
        print(f"Workbook '{file_path}' not found. Created a new one.")
    return workbook

def append_data_to_sheet(workbook, sheet_name, data):
    if sheet_name in workbook.sheetnames:
        sheet = workbook[sheet_name]
    else:
        sheet = workbook.create_sheet(sheet_name)

    # Convert data list to string for easier comparison
    data_str = "|".join([str(cell) for cell in data if cell])  # Handle blank values

    # Check for duplicate entries
    for row in sheet.iter_rows(values_only=True):
        row_str = "|".join([str(cell) for cell in row if cell])  # Handle blank values
        if row_str == data_str:
            # Prompt user for action
            if not messagebox.askyesno("Duplicate Entry Found", "An exact match was found in the Excel file. Do you want to add it anyway?"):
                return False

    # Append the data as a new row
    sheet.append(data)
    return True

def save_workbook(workbook, file_path):
    workbook.save(file_path)
    print(f"Workbook '{file_path}' saved successfully.")

def append_processed_data(file_path, sheet_name, data):
    workbook = open_workbook(file_path)
    if append_data_to_sheet(workbook, sheet_name, data):
        save_workbook(workbook, file_path)
        print("Data appended successfully.")
    else:
        print("Data append aborted due to duplicate entry.")
