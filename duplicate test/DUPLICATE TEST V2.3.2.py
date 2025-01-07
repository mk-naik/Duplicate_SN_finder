import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter.ttk import Combobox, Progressbar
import pandas as pd
import os
import datetime


class DuplicateFinderApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Multi-File Barcode Duplicate Finder")
        self.root.geometry("600x400")

        self.selected_files = []
        self.sheet_selection_comboboxes = []
        self.sheet_headers = {}  # Store sheet names for each file

        # File selection
        self.file_button = tk.Button(root, text="Select Files", command=self.select_files)
        self.file_button.pack(pady=10)

        self.file_label = tk.Label(root, text="No files selected")
        self.file_label.pack()

        # Create a scrollable canvas for file and sheet selection
        self.canvas_frame = tk.Frame(root)
        self.canvas_frame.pack(fill="both", expand=True)

        self.canvas = tk.Canvas(self.canvas_frame)
        self.scrollbar = tk.Scrollbar(self.canvas_frame, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = tk.Frame(self.canvas)

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )
        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        # Bind mouse wheel and keyboard events
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)
        self.canvas.bind_all("<Up>", lambda e: self.canvas.yview_scroll(-1, "units"))
        self.canvas.bind_all("<Down>", lambda e: self.canvas.yview_scroll(1, "units"))
        self.canvas.bind_all("<Prior>", lambda e: self.canvas.yview_scroll(-1, "pages"))  # Page Up
        self.canvas.bind_all("<Next>", lambda e: self.canvas.yview_scroll(1, "pages"))  # Page Down

        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")

        # Progress bar
        self.progress = Progressbar(root, orient="horizontal", mode="determinate", length=300)
        self.progress.pack(pady=10)

        # Start button
        self.start_button = tk.Button(root, text="Start Duplicate Check", command=self.find_duplicates)
        self.start_button.pack(pady=10)

    def _on_mousewheel(self, event):
        # Scroll on mouse wheel (Windows/Linux and macOS handling)
        if event.num == 5 or event.delta < 0:  # Down
            self.canvas.yview_scroll(1, "units")
        elif event.num == 4 or event.delta > 0:  # Up
            self.canvas.yview_scroll(-1, "units")

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
        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()
        self.sheet_selection_comboboxes = []
        self.sheet_headers = {}

        for idx, file in enumerate(self.selected_files):
            try:
                # Get sheet names
                sheet_names = pd.ExcelFile(file).sheet_names
                self.sheet_headers[file] = sheet_names

                 # Create a frame for each file and dropdown pair
                file_frame = tk.Frame(self.scrollable_frame)
                file_frame.pack(fill="x", pady=5)

                # Display the file label with increased width
                sheet_label = tk.Label(file_frame, text=f"File {idx + 1}: {os.path.basename(file)}", width=50, anchor="w")
                sheet_label.pack(side="left", padx=10)

                # Create and display the dropdown with reduced width
                sheet_combobox = Combobox(file_frame, values=sheet_names, state="readonly", width=7)
                sheet_combobox.pack(side="left", padx=10)
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
            self.progress["maximum"] = len(self.selected_files)  # Set max value for progress bar
            self.progress["value"] = 0  # Reset progress bar

            for idx, file in enumerate(self.selected_files):
                selected_sheet = self.sheet_selection_comboboxes[idx].get()

                if not selected_sheet:
                    messagebox.showwarning("Warning", f"Please select a sheet for File {idx + 1}.")
                    return

                file_name = os.path.basename(file)

                # Read the selected sheet, skipping the first row and using the second row as headers
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

                # Update progress bar
                self.progress["value"] += 1
                self.root.update_idletasks()

            # Find duplicates
            barcode_df = pd.DataFrame(barcode_locations, columns=["BARCODE", "FILE_NAME"])
            duplicate_barcodes = barcode_df[barcode_df.duplicated("BARCODE", keep=False)]

            if not duplicate_barcodes.empty:
                # Count unique duplicate barcodes
                unique_duplicates_count = duplicate_barcodes["BARCODE"].nunique()

                grouped_duplicates = []

                for barcode, group in duplicate_barcodes.groupby("BARCODE"):
                    rows = group.reset_index(drop=True)
                    row_data = [barcode]

                    for _, row in rows.iterrows():
                        row_data.append(row["FILE_NAME"])

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

                # Show success message with the count
                messagebox.showinfo("Success", f"{unique_duplicates_count} Duplicates found and saved to '{output_filename}'.")
            else:
                messagebox.showinfo("No Duplicates", "No duplicates found across selected files.")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {str(e)}")


# Run the application
root = tk.Tk()
app = DuplicateFinderApp(root)
root.mainloop()
