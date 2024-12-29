import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter.ttk import Combobox, Progressbar
import pandas as pd
import os
import datetime
import threading
from queue import Queue


class DuplicateFinderApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Multi-File Barcode Duplicate Finder")
        self.root.geometry("600x500")

        self.selected_files = []
        self.sheet_selection_comboboxes = []
        self.sheet_headers = {}
        
        # Queue for thread communication
        self.queue = Queue()
        
        # Create GUI elements
        self.create_gui()

    def create_gui(self):
        # File selection
        self.file_button = tk.Button(self.root, text="Select Files", command=self.select_files)
        self.file_button.pack(pady=10)

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
            self.selected_files = files
            self.file_label.config(text=f"{len(files)} file(s) selected")
            self.display_file_selection()
        else:
            self.file_label.config(text="No files selected")

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

    def disable_controls(self):
        self.file_button.config(state="disabled")
        self.start_button.config(state="disabled")
        for combobox in self.sheet_selection_comboboxes:
            combobox.config(state="disabled")

    def enable_controls(self):
        self.file_button.config(state="normal")
        self.start_button.config(state="normal")
        for combobox in self.sheet_selection_comboboxes:
            combobox.config(state="readonly")

    def process_files(self):
        try:
            barcode_locations = []
            total_steps = len(self.selected_files) + 2

            for idx, file in enumerate(self.selected_files):
                selected_sheet = self.sheet_selection_comboboxes[idx].get()
                file_name = os.path.basename(file)
                
                self.update_status(idx * 100 / total_steps, f"Reading {file_name}...")
                
                df = pd.read_excel(file, sheet_name=selected_sheet, header=1)
                barcode_columns = [col for col in df.columns if "barcode" in col.lower()]

                if not barcode_columns:
                    self.queue.put(("complete", False, f"No 'Barcode' columns found in {file_name}"))
                    return

                for col in barcode_columns:
                    for idx, value in enumerate(df[col].dropna()):
                        barcode_locations.append((value, str(file_name)))

            self.update_status(90, "Processing duplicates...")
            
            barcode_df = pd.DataFrame(barcode_locations, columns=["BARCODE", "FILE_NAME"])
            duplicate_barcodes = barcode_df[barcode_df.duplicated("BARCODE", keep=False)]

            if not duplicate_barcodes.empty:
                unique_duplicates_count = duplicate_barcodes["BARCODE"].nunique()
                grouped_duplicates = []

                for barcode, group in duplicate_barcodes.groupby("BARCODE"):
                    row_data = [barcode]
                    for _, row in group.iterrows():
                        row_data.append(row["FILE_NAME"])
                    grouped_duplicates.append(row_data)

                max_files = max(len(row) - 1 for row in grouped_duplicates)
                headers = ["DUPLICATE_BARCODES"] + [f"FILE_NAME{i + 1}" for i in range(max_files)]
                aligned_duplicates = [row + [""] * (len(headers) - len(row)) for row in grouped_duplicates]
                duplicates_df = pd.DataFrame(aligned_duplicates, columns=headers)

                self.update_status(95, "Saving results...")

                timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                output_filename = f"All_Duplicates_{timestamp}.xlsx"
                duplicates_df.to_excel(output_filename, index=False)

                self.update_status(100, "")
                self.queue.put(("complete", True, 
                    f"{unique_duplicates_count} Duplicates found and saved to '{output_filename}'"))
                self.update_status(0, "")
            else:
                self.update_status(100, "")
                self.queue.put(("complete", True, "No duplicates found across selected files."))
                self.update_status(0, "")

        except Exception as e:
            self.queue.put(("complete", False, f"An error occurred: {str(e)}"))


if __name__ == "__main__":
    root = tk.Tk()
    app = DuplicateFinderApp(root)
    root.mainloop()
