import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import cypress_creek_pest_control as ccpc
import excel_writer as ew

class VendorProcessingApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Vendor Processing App")

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

        # Button frame to hold Process and Transfer buttons inline
        self.button_frame = ttk.Frame(root)
        self.button_frame.pack(pady=10)

        # Process button
        self.process_button = ttk.Button(self.button_frame, text="Process", command=self.process_vendor)
        self.process_button.grid(row=0, column=0, padx=5)

        # Transfer button
        self.transfer_button = ttk.Button(self.button_frame, text="Transfer", command=self.transfer_to_excel)
        self.transfer_button.grid(row=0, column=1, padx=5)

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

    def transfer_to_excel(self):
        file_path = "processed_data.xlsx"
        sheet_name = "Vendor Data"

        # Read data from text box
        data = self.output_text.get(1.0, tk.END).strip().split("\n")
        data_dict = {line.split(": ")[0]: line.split(": ")[1] for line in data if line.split(": ")[1]}
        data_list = list(data_dict.values())

        ew.append_processed_data(file_path, sheet_name, data_list)

        # Confirmation message box
        if messagebox.askyesno("Confirmation", "Did the data transfer correctly?"):
            messagebox.showinfo("Success", "Data transferred successfully.")
        else:
            messagebox.showinfo("Error", "Data transfer failed. Please correct the data.")

if __name__ == "__main__":
    root = tk.Tk()
    app = VendorProcessingApp(root)
    root.mainloop()
