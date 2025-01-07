import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import cypress_creek_pest_control as ccpc
import excel_writer as ew
import os
from datetime import datetime
import openpyxl  # Make sure to import openpyxl
import json

class VendorProcessingApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Invoice Processing App")

        # Vendor selection dropdown
        self.vendor_label = ttk.Label(root, text="Select Vendor:")
        self.vendor_label.pack(pady=5)

        self.vendor_var = tk.StringVar()
        self.vendor_dropdown = ttk.Combobox(root, textvariable=self.vendor_var)
        self.vendor_dropdown['values'] = ("Cypress Creek Pest Control",)  # Add other vendors here
        self.vendor_dropdown.pack(pady=5)

        # Check number entry
        self.check_number_label = ttk.Label(root, text="Enter Check Number:")
        self.check_number_label.pack(pady=5)
        self.check_number_entry = ttk.Entry(root)
        self.check_number_entry.pack(pady=5)

        # Button frame to hold Process, Transfer, and Create .xlsx buttons inline
        self.button_frame = ttk.Frame(root)
        self.button_frame.pack(pady=10)

        # Process button
        self.process_button = ttk.Button(self.button_frame, text="Process", command=self.process_vendor)
        self.process_button.grid(row=0, column=0, padx=5)

        # Transfer button
        self.transfer_button = ttk.Button(self.button_frame, text="Transfer", command=self.transfer_to_excel)
        self.transfer_button.grid(row=0, column=1, padx=5)

        # Create .xlsx button
        self.create_button = ttk.Button(self.button_frame, text="Create .xlsx", command=self.create_excel_file)
        self.create_button.grid(row=0, column=2, padx=5)

        # Output text box
        self.output_text = tk.Text(root, height=15, width=70)
        self.output_text.pack(pady=10)

    def process_vendor(self):
        vendor = self.vendor_var.get()
        check_number = self.check_number_entry.get()

        if vendor == "Cypress Creek Pest Control":
            file_path = filedialog.askopenfilename(title="Select file", filetypes=[("Word Documents", "*.docx")])
            if file_path:
                processed_data = ccpc.process_document(file_path, check_number)
                self.output_text.delete(1.0, tk.END)
                for key, value in processed_data.items():
                    self.output_text.insert(tk.END, f"{key}: {value}\n")

                # Ask user for confirmation after 2 seconds
                self.root.after(2000, self.ask_data_correct, processed_data)

    def ask_data_correct(self, processed_data):
        if messagebox.askyesno("Data Confirmation", "Did the data process correctly?"):
            return # If yes, do nothing and close the prompt
        else: 
            # If no, open file save dialog for error log
            file_path = filedialog.asksaveasfilename(
                title="Save Error Log", 
                filetypes=[("Text Files", "*.txt")],
                defaultextension=".txt"
            )
            if file_path: 
                log_entry = { 
                    "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"), 
                    "data": processed_data 
                    } 
                if os.path.exists(file_path): 
                    with open(file_path, 'a') as file: 
                        file.write("\n" + json.dumps(log_entry, indent=4)) 
                else: 
                    with open(file_path, 'w') as file: 
                        file.write(json.dumps(log_entry, indent=4)) 
                messagebox.showinfo("Error Log Saved", f"Error log saved to: {file_path}") 
            else: 
                messagebox.showwarning("No file selected", "No file was selected for saving the error log.")
                
                
    def transfer_to_excel(self):
        # Open a file dialog to select the .xlsx file
        file_path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel Files", "*.xlsx")],
            defaultextension=".xlsx"
        )

        if not file_path:
            messagebox.showwarning("No File Selected", "Please select an Excel file to transfer the data. Please close it and try again.")
            return

        try:
            workbook = openpyxl.load_workbook(file_path)
        except PermissionError:
            messagebox.showerror("File Open", "The file is currently open in Excel")
            return

        sheet_name = "Invoice Data"

        # Read data from text box
        data = self.output_text.get(1.0, tk.END).strip().split("\n")
        data_dict = {}
        for line in data:
            parts = line.split(": ")
            if len(parts) > 1 and (parts[0] == "Check Number" or parts[1]):  # Allow empty second part only for "Check Number"
                data_dict[parts[0]] = parts[1]
        data_list = list(data_dict.values())

        ew.append_processed_data(file_path, sheet_name, data_list)

        # Confirmation message box
        if messagebox.askyesno("Confirmation", "Did the data transfer correctly?"):
            self.show_success_message()
        else:
            messagebox.showinfo("Error", "Data transfer failed. Please correct the data.")

    def create_excel_file(self):
        # Open a file dialog to save the .xlsx file
        file_path = filedialog.asksaveasfilename(
            title="Create Excel File",
            filetypes=[("Excel Files", "*.xlsx")],
            defaultextension=".xlsx"
        )

        if not file_path:
            messagebox.showwarning("No file selected", "Please select a location to create the Excel file.")
            return

        # Create the file if it doesn't exist
        if not os.path.exists(file_path):
            workbook = openpyxl.Workbook()
            workbook.save(file_path)
            messagebox.showinfo("File Created", f"New Excel file created at: {file_path}")
        else:
            messagebox.showinfo("File Exists", "The file already exists. Please choose a different name or location.")

    def show_success_message(self):
        success_window = tk.Toplevel(self.root)
        success_window.title("Success")
        success_label = tk.Label(success_window, text="Data transferred successfully.", pady=20, padx=20)
        success_label.pack()
        success_window.after(1500, success_window.destroy)

if __name__ == "__main__":
    root = tk.Tk()
    app = VendorProcessingApp(root)
    root.mainloop()
