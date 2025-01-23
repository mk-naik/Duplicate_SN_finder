import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter.ttk import Combobox

class DuplicateFinderApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Multi-File Barcode Duplicate Finder")
        self.root.geometry("600x500")

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

        all_barcodes = []

        try:
            for idx, file in enumerate(self.selected_files):
                selected_sheet = self.sheet_selection_comboboxes[idx].get()

                if not selected_sheet:
                    messagebox.showwarning("Warning", f"Please select a sheet for File {idx + 1}.")
                    return

                # Read the selected sheet
                df = pd.read_excel(file, sheet_name=selected_sheet)

                # Automatically detect barcode-like columns
                barcode_columns = [col for col in df.columns if "barcode" in col.lower()]

                if not barcode_columns:
                    messagebox.showwarning("Warning", f"No 'Barcode' columns found in {file} ({selected_sheet}).")
                    continue

                # Aggregate all barcode data
                for col in barcode_columns:
                    all_barcodes.extend(df[col].dropna().tolist())

            # Find duplicates
            barcode_series = pd.Series(all_barcodes)
            duplicates = barcode_series[barcode_series.duplicated()].unique()

            if len(duplicates) > 0:
                # Save results to a new Excel file
                duplicates_df = pd.DataFrame(duplicates, columns=["DUPLICATE_BARCODES"])
                duplicates_df.to_excel("All_Duplicates.xlsx", index=False)
                messagebox.showinfo("Success", "Duplicates found and saved to 'All_Duplicates.xlsx'.")
            else:
                messagebox.showinfo("No Duplicates", "No duplicates found across selected files.")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {str(e)}")


# Run the application
root = tk.Tk()
app = DuplicateFinderApp(root)
root.mainloop()
