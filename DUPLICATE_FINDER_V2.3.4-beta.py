import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter.ttk import Combobox, Progressbar
import pandas as pd
import os
import datetime
import threading
from queue import Queue
import re


class DuplicateFinderApp:
    def __init__(self, root):
        self.root = root
        self.root.title("ICON Barcode Duplicate Finder")
        self.root.geometry("600x500")

        # Custom ICON barcode patterns
        self.barcode_patterns = {
            'ICON-17': r'^ICON\d{13}$',  # ICON followed by 13 digits
            'ICON-18': r'^ICON\d{3}[A-Z]\d{10}$'  # ICON + 3 digits + 1 letter + 10 digits
        }

        self.selected_files = []
        self.sheet_selection_comboboxes = []
        self.sheet_headers = {}
        
        # Queue for thread communication
        self.queue = Queue()
        
        # Create GUI elements
        self.create_gui()

    def create_gui(self):
        # Create a frame for the buttons
        button_frame = tk.Frame(self.root)
        button_frame.pack(pady=10)

        # File selection button
        self.file_button = tk.Button(button_frame, text="Select Files", command=self.select_files)
        self.file_button.pack(side="left", padx=5)

        # Folder selection button
        self.folder_button = tk.Button(button_frame, text="Select Folder", command=self.select_folder)
        self.folder_button.pack(side="left", padx=5)

        self.file_label = tk.Label(self.root, text="No files selected")
        self.file_label.pack()
        
        # Create a scrollable canvas
        self.canvas_frame = tk.Frame(self.root)
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

        # Bind scrolling events
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)
        self.canvas.bind_all("<Up>", lambda e: self.canvas.yview_scroll(-1, "units"))
        self.canvas.bind_all("<Down>", lambda e: self.canvas.yview_scroll(1, "units"))
        self.canvas.bind_all("<Prior>", lambda e: self.canvas.yview_scroll(-1, "pages"))
        self.canvas.bind_all("<Next>", lambda e: self.canvas.yview_scroll(1, "pages"))

        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")

        # Progress frame
        self.progress_frame = tk.Frame(self.root)
        self.progress_frame.pack(fill="x", pady=10)

        self.progress = Progressbar(
            self.progress_frame,
            orient="horizontal",
            mode="determinate",
            length=300
        )
        self.progress.pack(pady=5)

        self.status_var = tk.StringVar()
        self.status_label = tk.Label(
            self.progress_frame,
            textvariable=self.status_var,
            wraplength=500,
            height=2,
            justify=tk.CENTER
        )
        self.status_label.pack(fill="x", padx=10)

        # Start button
        self.start_button = tk.Button(
            self.root,
            text="Start Duplicate Check",
            command=self.start_processing
        )
        self.start_button.pack(pady=10)

            # Reset button (add this after the Start button)
        self.reset_button = tk.Button(
            self.root,
            text="Reset Selection",
            command=self.reset_selection
        )
        self.reset_button.pack(pady=5)

    def detect_barcodes(self, value):
        """
        Detect if a value matches ICON barcode pattern
        Returns (is_barcode, barcode_type)
        """
        if pd.isna(value):
            return False, None
        
        # Convert to string and remove any whitespace
        str_value = str(value).strip()
        
        # Skip empty strings
        if not str_value:
            return False, None

        # Check for exact length (17 or 18 characters)
        if len(str_value) not in [17, 18]:
            return False, None

        # Check if starts with 'ICON'
        if not str_value.startswith('ICON'):
            return False, None

        # For 17-character format
        if len(str_value) == 17:
            if re.match(self.barcode_patterns['ICON-17'], str_value):
                return True, 'ICON-17'
                
        # For 18-character format
        elif len(str_value) == 18:
            if re.match(self.barcode_patterns['ICON-18'], str_value):
                return True, 'ICON-18'

        return False, None

    def find_barcodes_in_dataframe(self, df):
        """Find all ICON barcode values in a DataFrame."""
        barcodes = []
        
        for column in df.columns:
            for idx, value in enumerate(df[column]):
                is_barcode, barcode_type = self.detect_barcodes(value)
                if is_barcode:
                    barcodes.append({
                        'value': str(value).strip(),
                        'column': column,
                        'row': idx + 2,  # Adding 2 for Excel row number (1-based + header)
                        'type': barcode_type
                    })
        
        return barcodes

    def _on_mousewheel(self, event):
        if event.num == 5 or event.delta < 0:
            self.canvas.yview_scroll(1, "units")
        elif event.num == 4 or event.delta > 0:
            self.canvas.yview_scroll(-1, "units")

    def select_files(self):
        files = filedialog.askopenfilenames(
            title="Select Excel Files",
            filetypes=[("Excel Files", "*.xlsx *.xls")]
        )
        if files:
            current_files = list(self.selected_files)
            self.selected_files = list(set(current_files + list(files)))
            self.file_label.config(text=f"{len(files)} file(s) selected")
            self.display_file_selection()
        else:
            self.file_label.config(text="No files selected")

    def select_folder(self):
        folder_selected = filedialog.askdirectory(title="Select Folder")
        
        if not folder_selected:
            return
            
        excel_files = []
        # Walk through all subdirectories and files
        for root, dirs, files in os.walk(folder_selected):
            for file in files:
                # Check if the file is an Excel file
                if file.endswith(('.xlsx', '.xls')):
                    # Construct the full file path
                    excel_files.append(os.path.join(root, file))
        
        if excel_files:
            # Add new files to existing selection
            current_files = list(self.selected_files)
            self.selected_files = list(set(current_files + excel_files))
            self.file_label.config(text=f"{len(self.selected_files)} file(s) selected")
            self.display_file_selection()
        else:
            messagebox.showinfo("Information", "No Excel files found in the selected folder and its subfolders.")
    
    def display_file_selection(self):
        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()
        self.sheet_selection_comboboxes = []
        self.sheet_headers = {}

        for idx, file in enumerate(self.selected_files):
            try:
                sheet_names = pd.ExcelFile(file).sheet_names
                self.sheet_headers[file] = sheet_names

                file_frame = tk.Frame(self.scrollable_frame)
                file_frame.pack(fill="x", pady=5)

                sheet_label = tk.Label(
                    file_frame,
                    text=f"File {idx + 1}: {os.path.basename(file)}",
                    width=50,
                    anchor="w"
                )
                sheet_label.pack(side="left", padx=10)

                sheet_combobox = Combobox(
                    file_frame,
                    values=sheet_names,
                    state="readonly",
                    width=7
                )
                sheet_combobox.pack(side="left", padx=10)
                sheet_combobox.current(0)
                self.sheet_selection_comboboxes.append(sheet_combobox)
            except Exception as e:
                messagebox.showerror("Error", f"Error reading file {file}: {str(e)}")

    def process_files(self):
        try:
            all_barcodes = []
            total_steps = len(self.selected_files) + 2

            for idx, file in enumerate(self.selected_files):
                selected_sheet = self.sheet_selection_comboboxes[idx].get()
                file_name = os.path.basename(file)
                
                self.update_status(idx * 100 / total_steps, f"Reading {file_name}...")
                
                try:
                    # Read all rows as strings to preserve leading zeros
                    df = pd.read_excel(file, sheet_name=selected_sheet, dtype=str)
                    self.update_status(idx * 100 / total_steps, f"Scanning for ICON barcodes in {file_name}...")
                    
                    barcodes = self.find_barcodes_in_dataframe(df)
                    
                    if not barcodes:
                        self.update_status(idx * 100 / total_steps, 
                                         f"No ICON barcodes found in {file_name}, continuing...")
                        continue

                    for barcode in barcodes:
                        all_barcodes.append({
                            'BARCODE': barcode['value'],
                            'FILE_NAME': file_name,
                            'FORMAT': barcode['type']
                        })
                        
                except Exception as e:
                    self.queue.put(("complete", False, f"Error processing {file_name}: {str(e)}"))
                    return

            if not all_barcodes:
                self.queue.put(("complete", False, "No ICON barcodes found in any of the selected files."))
                return

            self.update_status(90, "Processing duplicates...")
        
            # Convert to DataFrame and find duplicates
            barcode_df = pd.DataFrame(all_barcodes, columns=["BARCODE", "FILE_NAME", "FORMAT"])
            duplicate_barcodes = barcode_df[barcode_df.duplicated("BARCODE", keep=False)]

            if not duplicate_barcodes.empty:
                self.update_status(95, "Preparing detailed report...")

                # Create grouped duplicates with the format you want
                grouped_duplicates = []
                for barcode, group in duplicate_barcodes.groupby("BARCODE"):
                    # Count occurrences of each barcode
                    copies = len(group)
                
                    # Create row with barcode and its file locations
                    row_data = [barcode, copies]  # Start with barcode and copies count
                    file_names = group["FILE_NAME"].tolist()
                    row_data.extend(file_names)
                    grouped_duplicates.append(row_data)

                # Create headers with maximum number of files
                max_files = max(len(row) - 2 for row in grouped_duplicates)  # -2 for BARCODE and COPIES columns
                headers = ["DUPLICATE_BARCODES", "COPIES"] + [f"FILE_NAME{i + 1}" for i in range(max_files)]
            
                # Pad rows with empty strings to match header length
                aligned_duplicates = [row + [""] * (len(headers) - len(row)) for row in grouped_duplicates]
            
                # Create final DataFrame
                duplicates_df = pd.DataFrame(aligned_duplicates, columns=headers)

                # Save to Excel
                timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                destination_path = os.path.expanduser("~")
                folder_path = os.path.join(destination_path, "Desktop", "DUPLICATE_BARCODES")
                os.makedirs(folder_path, exist_ok=True)
                output_filename = os.path.join(folder_path, f"ICON_Duplicates_{timestamp}.xlsx")

                with pd.ExcelWriter(output_filename) as writer:
                    # Detailed duplicates sheet with new format
                    duplicates_df.to_excel(writer, sheet_name='Detailed_Report', index=False)
                
                    # Summary sheet
                    summary_data = {
                        'Metric': [
                            'Total Files Processed',
                            'Total ICON Barcodes Found',
                            'Unique Barcodes',
                            'Duplicate Barcodes',
                            '17-Character Barcodes',
                            '18-Character Barcodes'
                        ],
                        'Value': [
                            len(self.selected_files),
                            len(barcode_df),
                            len(barcode_df['BARCODE'].unique()),
                            len(duplicate_barcodes['BARCODE'].unique()),
                            len(barcode_df[barcode_df['FORMAT'] == 'ICON-17']),
                            len(barcode_df[barcode_df['FORMAT'] == 'ICON-18'])
                        ]
                    }
                    pd.DataFrame(summary_data).to_excel(writer, sheet_name='Summary', index=False)

                self.update_status(100, "")
                self.queue.put(("complete", True, 
                    f"Found {len(duplicate_barcodes['BARCODE'].unique())} duplicate ICON barcodes. "
                    f"Report saved to '{output_filename}'"))
                self.update_status(0, "")
            else:
                self.update_status(100, "")
                self.queue.put(("complete", True, "No duplicate ICON barcodes found."))
                self.update_status(0, "")

        except Exception as e:
            self.queue.put(("complete", False, f"An error occurred: {str(e)}"))
    
    def update_status(self, progress, status):
        self.queue.put(("status", progress, status))

    def check_queue(self):
        while not self.queue.empty():
            msg = self.queue.get()
            if msg[0] == "status":
                _, progress, status = msg
                self.progress["value"] = progress
                self.status_var.set(status)
                self.root.update_idletasks()
            elif msg[0] == "complete":
                _, success, message = msg
                self.enable_controls()
                if success:
                    messagebox.showinfo("Success", message)
                else:
                    messagebox.showerror("Error", message)
        
        self.root.after(100, self.check_queue)

    def start_processing(self):
        if not self.selected_files:
            messagebox.showwarning("Warning", "Please select files first.")
            return

        self.disable_controls()
        thread = threading.Thread(target=self.process_files, daemon=True)
        thread.start()
        self.check_queue()

    def reset_selection(self):
        self.selected_files = []
        self.file_label.config(text="No files selected")
        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()
        self.sheet_selection_comboboxes = []
        self.sheet_headers = {}

    def disable_controls(self):
        self.file_button.config(state="disabled")
        self.folder_button.config(state="disabled")
        self.start_button.config(state="disabled")
        self.reset_button.config(state="disabled")
        for combobox in self.sheet_selection_comboboxes:
            combobox.config(state="disabled")

    def enable_controls(self):
        self.file_button.config(state="normal")
        self.folder_button.config(state="normal")
        self.start_button.config(state="normal")
        self.reset_button.config(state="normal")
        for combobox in self.sheet_selection_comboboxes:
            combobox.config(state="readonly")


if __name__ == "__main__":
    root = tk.Tk()
    app = DuplicateFinderApp(root)
    root.mainloop()
