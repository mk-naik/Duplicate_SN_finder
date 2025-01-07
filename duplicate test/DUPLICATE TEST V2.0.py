import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# Initialize global variables
file1_path = ""
file2_path = ""
sheet_names_file1 = []
sheet_names_file2 = []

def select_file1():
    global file1_path, sheet_names_file1
    file1_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
    if file1_path:
        file1_label.config(text=f"File 1: {file1_path.split('/')[-1]}")
        try:
            sheet_names_file1 = pd.ExcelFile(file1_path).sheet_names
            sheet1_dropdown['values'] = sheet_names_file1
        except Exception as e:
            messagebox.showerror("Error", f"Failed to read sheets from File 1: {e}")

def select_file2():
    global file2_path, sheet_names_file2
    file2_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
    if file2_path:
        file2_label.config(text=f"File 2: {file2_path.split('/')[-1]}")
        try:
            sheet_names_file2 = pd.ExcelFile(file2_path).sheet_names
            sheet2_dropdown['values'] = sheet_names_file2
        except Exception as e:
            messagebox.showerror("Error", f"Failed to read sheets from File 2: {e}")

def update_headers(file_path, sheet_name, dropdown):
    try:
        if file_path and sheet_name:
            headers = pd.read_excel(file_path, sheet_name=sheet_name, nrows=0).columns.tolist()
            dropdown['values'] = headers
    except Exception as e:
        messagebox.showerror("Error", f"Failed to read headers: {e}")

def find_duplicates():
    try:
        # Get selected sheet names and headers
        sheet1 = sheet1_dropdown.get()
        sheet2 = sheet2_dropdown.get()
        header1 = header1_dropdown.get()
        header2 = header2_dropdown.get()

        if not all([file1_path, file2_path, sheet1, sheet2, header1, header2]):
            messagebox.showerror("Error", "Please select all files, sheets, and headers before starting.")
            return

        # Load the selected columns
        file1 = pd.read_excel(file1_path, sheet_name=sheet1, usecols=[header1])
        file2 = pd.read_excel(file2_path, sheet_name=sheet2, usecols=[header2])

        # Find duplicates
        duplicates = file1[file1[header1].isin(file2[header2])]

        if not duplicates.empty:
            duplicates.to_excel("Duplicates_Between_Files.xlsx", index=False)
            messagebox.showinfo("Result", "Duplicates found and saved to 'Duplicates_Between_Files.xlsx'.")
        else:
            messagebox.showinfo("Result", "No Duplicates Found.")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")

# Tkinter GUI
root = tk.Tk()
root.title("Enhanced Duplicate Finder")

# File 1 selection
file1_label = tk.Label(root, text="No File 1 Selected", width=50, anchor="w")
file1_label.grid(row=0, column=0, padx=10, pady=5)
file1_button = tk.Button(root, text="Select File 1", command=select_file1)
file1_button.grid(row=0, column=1, padx=10, pady=5)

# Sheet 1 selection
tk.Label(root, text="Select Sheet from File 1:").grid(row=1, column=0, padx=10, pady=5)
sheet1_dropdown = ttk.Combobox(root, state="readonly", width=47)
sheet1_dropdown.grid(row=1, column=1, padx=10, pady=5)

# Header 1 selection
tk.Label(root, text="Select Header from File 1:").grid(row=2, column=0, padx=10, pady=5)
header1_dropdown = ttk.Combobox(root, state="readonly", width=47)
header1_dropdown.grid(row=2, column=1, padx=10, pady=5)

# File 2 selection
file2_label = tk.Label(root, text="No File 2 Selected", width=50, anchor="w")
file2_label.grid(row=3, column=0, padx=10, pady=5)
file2_button = tk.Button(root, text="Select File 2", command=select_file2)
file2_button.grid(row=3, column=1, padx=10, pady=5)

# Sheet 2 selection
tk.Label(root, text="Select Sheet from File 2:").grid(row=4, column=0, padx=10, pady=5)
sheet2_dropdown = ttk.Combobox(root, state="readonly", width=47)
sheet2_dropdown.grid(row=4, column=1, padx=10, pady=5)

# Header 2 selection
tk.Label(root, text="Select Header from File 2:").grid(row=5, column=0, padx=10, pady=5)
header2_dropdown = ttk.Combobox(root, state="readonly", width=47)
header2_dropdown.grid(row=5, column=1, padx=10, pady=5)

# Bind dropdowns to update headers
sheet1_dropdown.bind("<<ComboboxSelected>>", lambda e: update_headers(file1_path, sheet1_dropdown.get(), header1_dropdown))
sheet2_dropdown.bind("<<ComboboxSelected>>", lambda e: update_headers(file2_path, sheet2_dropdown.get(), header2_dropdown))

# Start button
start_button = tk.Button(root, text="Start", command=find_duplicates)
start_button.grid(row=6, column=0, columnspan=2, pady=10)

# Run the application
root.mainloop()
