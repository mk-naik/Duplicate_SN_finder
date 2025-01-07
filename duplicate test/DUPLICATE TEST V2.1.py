import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd

# Global variables for file paths
file1_path = ""
file2_path = ""

def browse_file1():
    global file1_path
    file1_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    if file1_path:
        file1_label.config(text=f"Selected: {file1_path.split('/')[-1]}")
        update_sheets(file1_path, sheet1_dropdown)

def browse_file2():
    global file2_path
    file2_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    if file2_path:
        file2_label.config(text=f"Selected: {file2_path.split('/')[-1]}")
        update_sheets(file2_path, sheet2_dropdown)

def update_sheets(file_path, dropdown):
    try:
        sheets = pd.ExcelFile(file_path).sheet_names
        dropdown['values'] = sheets
        dropdown.current(0)
    except Exception as e:
        messagebox.showerror("Error", f"Failed to load sheets: {e}")

def find_duplicates_combined():
    try:
        # Get selected sheet names
        sheet1 = sheet1_dropdown.get()
        sheet2 = sheet2_dropdown.get()

        if not all([file1_path, file2_path, sheet1, sheet2]):
            messagebox.showerror("Error", "Please select all files and sheets before starting.")
            return

        # Load the selected sheets
        file1_data = pd.read_excel(file1_path, sheet_name=sheet1)
        file2_data = pd.read_excel(file2_path, sheet_name=sheet2)

        # Combine all BARCODE-like columns
        barcode_columns_file1 = [col for col in file1_data.columns if "BARCODE" in col.upper()]
        barcode_columns_file2 = [col for col in file2_data.columns if "BARCODE" in col.upper()]

        if not barcode_columns_file1 or not barcode_columns_file2:
            messagebox.showerror("Error", "No BARCODE columns found in one or both files.")
            return

        combined_barcodes_file1 = pd.concat([file1_data[col] for col in barcode_columns_file1], ignore_index=True)
        combined_barcodes_file2 = pd.concat([file2_data[col] for col in barcode_columns_file2], ignore_index=True)

        # Find duplicates
        duplicates = combined_barcodes_file1[combined_barcodes_file1.isin(combined_barcodes_file2)].dropna()

        if not duplicates.empty:
            duplicates.to_excel("Duplicates_Between_Files.xlsx", index=False, header=["DUPLICATES"])
            messagebox.showinfo("Result", "Duplicates found and saved to 'Duplicates_Between_Files.xlsx'.")
        else:
            messagebox.showinfo("Result", "No Duplicates Found.")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")

# Create the GUI
root = tk.Tk()
root.title("Duplicate Finder")
root.geometry("500x400")

# File selection
file1_button = tk.Button(root, text="Select File 1", command=browse_file1)
file1_button.pack(pady=5)

file1_label = tk.Label(root, text="No file selected")
file1_label.pack()

file2_button = tk.Button(root, text="Select File 2", command=browse_file2)
file2_button.pack(pady=5)

file2_label = tk.Label(root, text="No file selected")
file2_label.pack()

# Sheet selection
sheet1_label = tk.Label(root, text="Select Sheet from File 1:")
sheet1_label.pack()

sheet1_dropdown = ttk.Combobox(root, state="readonly")
sheet1_dropdown.pack(pady=5)

sheet2_label = tk.Label(root, text="Select Sheet from File 2:")
sheet2_label.pack()

sheet2_dropdown = ttk.Combobox(root, state="readonly")
sheet2_dropdown.pack(pady=5)

# Start button
start_button = tk.Button(root, text="Start", command=find_duplicates_combined)
start_button.pack(pady=20)

# Run the GUI
root.mainloop()
