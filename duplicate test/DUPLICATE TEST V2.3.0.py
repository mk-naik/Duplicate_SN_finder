import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter.ttk import Combobox
import os
import datetime

class DuplicateFinderApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Multi-File Barcode Duplicate Finder")
        self.root.geometry("480x360")

        self.selected_files = []
        self.sheet_selection_comboboxes = []
        self.sheet_headers = {}  # Store sheet names for each file

        # File selection
        self.file_button = tk.Button(root, text="Select Files", command=self.select_files)
        self.file_button.pack(pady=10)

        self.file_label = tk.Label(root, text="No files selected")
        self.file_label.pack()

        # Dynamic area for file-specific selections
        self.selection_frame = tk.Frame(root)
        self.selection_frame.pack(pady=10, fill=tk.BOTH, expand=True)

        # Start button
        self.start_button = tk.Button(root, text="Start Duplicate Check", command=self.find_duplicates)
        self.start_button.pack(pady=20)

    def select_files(self):
        files = filedialog.askopenfilenames(
            title="Select Excel Files",
            filetypes=[("Excel Files", "*.xlsx *.xls")]
        )
        if files:
            self.selected_files = files
            self.file_label.config(text=f"{len(files)} file(s) selected")
            self.display_file_selection()
        else:
            self.file_label.config(text="No files selected")

    def display_file_selection(self):
        # Clear previous selections
        for widget in self.selection_frame.winfo_children():
            widget.destroy()
        self.sheet_selection_comboboxes = []
        self.sheet_headers = {}

        for idx, file in enumerate(self.selected_files):
            try:
                # Get sheet names
                sheet_names = pd.ExcelFile(file).sheet_names
                self.sheet_headers[file] = sheet_names

                # Display sheet selection
                sheet_label = tk.Label(self.selection_frame, text=f"File {idx + 1}: {file.split('/')[-1]}")
                sheet_label.pack(anchor="w")

                sheet_combobox = Combobox(self.selection_frame, values=sheet_names, state="readonly")
                sheet_combobox.pack(pady=5)
                sheet_combobox.current(0)  # Default to first sheet
                self.sheet_selection_comboboxes.append(sheet_combobox)
            except Exception as e:
                messagebox.showerror("Error", f"Error reading file {file}: {str(e)}")

    def find_duplicates(self):
        if not self.selected_files:
            messagebox.showwarning("Warning", "Please select files first.")
            return

        barcode_locations = []

        try:
            for idx, file in enumerate(self.selected_files):
                selected_sheet = self.sheet_selection_comboboxes[idx].get()

                if not selected_sheet:
                    messagebox.showwarning("Warning", f"Please select a sheet for File {idx + 1}.")
                    return
                
                file_name = os.path.basename(file)

                # Read the selected sheet
                df = pd.read_excel(file, sheet_name=selected_sheet, header=1)

                # Automatically detect barcode-like columns
                barcode_columns = [col for col in df.columns if "barcode" in col.lower()]

                if not barcode_columns:
                    messagebox.showwarning("Warning", f"No 'Barcode' columns found in {file} ({selected_sheet}).")
                    continue

                # Aggregate all barcode data with location
                for col in barcode_columns:
                    for idx, value in enumerate(df[col].dropna()):
                        barcode_locations.append((value, str(file_name)))


            # Find duplicates
            barcode_df = pd.DataFrame(barcode_locations, columns=["BARCODE", "FILE_NAME"])
            duplicate_barcodes = barcode_df[barcode_df.duplicated("BARCODE", keep=False)]

            if not duplicate_barcodes.empty:
                # Count unique duplicate barcodes
                unique_duplicates_count = duplicate_barcodes["BARCODE"].nunique()
                
                grouped_duplicates = []

                for barcode, group in duplicate_barcodes.groupby("BARCODE"):
                    row_data = [barcode]

                    # Add all file names associated with the barcode
                    file_names = group["FILE_NAME"].tolist()
                    row_data.extend(file_names)
    
                    grouped_duplicates.append(row_data)

                # Define column headers dynamically
                max_files = max(len(row) - 1 for row in grouped_duplicates)  # Calculate max number of file entries
                headers = ["DUPLICATE_BARCODES"] + [f"FILE_NAME{i + 1}" for i in range(max_files)]

                # Align rows to match header length
                aligned_duplicates = [
                    row + [""] * (len(headers) - len(row)) for row in grouped_duplicates
                ]

                duplicates_df = pd.DataFrame(aligned_duplicates, columns=headers)

                # Save results to a new Excel file
                timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                output_filename = f"All_Duplicates_{timestamp}.xlsx"
                duplicates_df.to_excel(output_filename, index=False)
                messagebox.showinfo("Success", f"{unique_duplicates_count} Duplicates found and saved to '{output_filename}'.")
            else:
                messagebox.showinfo("No Duplicates", "No duplicates found across selected files.")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {str(e)}")


# Run the application
root = tk.Tk()
app = DuplicateFinderApp(root)
root.mainloop()
