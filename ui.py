import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import pandas as pd


class DocumentGeneratorUI:
    def __init__(self, root, generator):
        self.root = root
        self.root.title("Document Generator - Excel to Docx Template")
        self.root.configure(bg="#2E3B4E")
        self.root.iconbitmap("icon.ico")
        self.generator = generator
        self.excel_path_var = tk.StringVar()
        self.template_doc_path_var = tk.StringVar()
        self.destination_folder_var = tk.StringVar()
        self.file_name_var = tk.StringVar()

        self.column_combobox = None
        self.concatenation_label = None
        self.main_frame = None

        self.create_widgets()

    def create_widgets(self):
        style = ttk.Style()
        style.configure("TButton", background="#4CAF50", foreground="white")

        self.root.grid_rowconfigure(0, weight=1)
        self.root.grid_columnconfigure(0, weight=1)

        style = ttk.Style()
        style.configure("TFrameStyle.TFrame", background="#2E3B4E")

        self.main_frame = ttk.Frame(
            self.root, padding=(10, 10, 10, 10), style="TFrameStyle.TFrame"
        )
        self.main_frame.grid(row=0, column=0, sticky="nsew")

        tk.Label(self.main_frame, text="Excel File:", fg="white", bg="#2E3B4E").grid(
            row=0, column=0, sticky="w", pady=5
        )
        tk.Entry(self.main_frame, textvariable=self.excel_path_var, width=40).grid(
            row=0, column=1, columnspan=2, pady=5
        )
        tk.Button(self.main_frame, text="Document", command=self.browse_excel).grid(
            row=0, column=3, pady=5
        )

        tk.Label(
            self.main_frame,
            text="Template Document:",
            fg="white",
            bg="#2E3B4E",
        ).grid(row=1, column=0, sticky="w", pady=5)
        tk.Entry(
            self.main_frame, textvariable=self.template_doc_path_var, width=40
        ).grid(row=1, column=1, columnspan=2, pady=5, padx=5)
        tk.Button(
            self.main_frame, text="Document", command=self.browse_template_doc
        ).grid(row=1, column=3, pady=5)

        tk.Label(
            self.main_frame, text="Destination Folder:", fg="white", bg="#2E3B4E"
        ).grid(row=2, column=0, sticky="w", pady=5)
        tk.Entry(
            self.main_frame, textvariable=self.destination_folder_var, width=40
        ).grid(row=2, column=1, columnspan=2, pady=5)
        tk.Button(
            self.main_frame, text="Document", command=self.browse_destination_folder
        ).grid(row=2, column=3, pady=5)

        tk.Label(self.main_frame, text="File Name:", fg="white", bg="#2E3B4E").grid(
            row=3, column=0, sticky="w", pady=5
        )
        tk.Entry(self.main_frame, textvariable=self.file_name_var).grid(
            row=3, column=1, columnspan=2, pady=5
        )

        tk.Label(self.main_frame, text="Select Column:", fg="white", bg="#2E3B4E").grid(
            row=4, column=0, sticky="w", pady=5
        )

        self.column_combobox = ttk.Combobox(self.main_frame, state="readonly")
        self.column_combobox.grid(row=4, column=1, columnspan=2, pady=5)

        self.concatenation_label = tk.Label(
            self.main_frame, text="", fg="white", bg="#2E3B4E"
        )
        self.concatenation_label.grid(row=5, column=0, columnspan=4, pady=5)

        tk.Button(
            self.main_frame, text="Generate Documents", command=self.generate_documents
        ).grid(row=6, column=0, columnspan=4, pady=10)

        self.file_name_var.trace_add("write", self.update_combobox_and_label)
        self.column_combobox.bind_all(
            "<<ComboboxSelected>>", lambda event=None: self.update_combobox_and_label()
        )

    def browse_excel(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if file_path:
            self.excel_path_var.set(file_path)

            excel_data = pd.read_excel(file_path, sheet_name="Sheet")
            column_names = excel_data.columns.tolist()

            if self.column_combobox:
                self.column_combobox.destroy()

            self.column_combobox = ttk.Combobox(
                self.main_frame, values=column_names, state="readonly"
            )
            self.column_combobox.grid(row=4, column=1, columnspan=2, pady=5)
            self.column_combobox.set(column_names[0])

    def browse_template_doc(self):
        file_path = filedialog.askopenfilename(filetypes=[("Word Files", "*.docx")])
        if file_path:
            self.template_doc_path_var.set(file_path)

    def browse_destination_folder(self):
        folder_path = filedialog.askdirectory()
        if folder_path:
            self.destination_folder_var.set(folder_path)

    def generate_documents(self):
        excel_path = self.excel_path_var.get()
        template_doc_path = self.template_doc_path_var.get()
        destination_folder = self.destination_folder_var.get()
        file_name = self.file_name_var.get()
        selected_column = self.column_combobox.get()

        if not os.path.exists(excel_path) or not os.path.exists(template_doc_path):
            messagebox.showerror(
                "Error",
                "Excel file or template document not found.",
            )
            return
        if not destination_folder:
            messagebox.showerror(
                "Error",
                "Destination folder not specified, it is where the generated documents will be saved.",
            )
            return

        try:
            excel_data = pd.read_excel(excel_path, sheet_name="Sheet")
        except Exception as e:
            messagebox.showerror("Error", f"Error loading Excel file: {e}")
            return

        if not os.path.exists(destination_folder):
            os.makedirs(destination_folder)

        field_mapping = {}
        for col_name in excel_data.columns:
            field_mapping[f"[{col_name.lower()}]"] = col_name.lower()

        column_like_doc_name = f"[{selected_column.lower()}]"

        for _, row in excel_data.iterrows():
            self.generator.generate_document_for_row(
                row,
                template_doc_path,
                destination_folder,
                field_mapping,
                column_like_doc_name,
                file_name,
            )

        self.update_combobox_value()

        messagebox.showinfo(
            "Success",
            "Documents generated successfully.",
        )

    def update_combobox_and_label(self, *args):
        self.update_combobox_value()
        self.update_concatenation_label()

    def update_combobox_value(self):
        selected_column = self.column_combobox.get()
        if selected_column:
            self.concatenation_label.config(
                text=f"Concatenation: {self.file_name_var.get()}_{selected_column.lower()}"
            )

    def update_concatenation_label(self, *args):
        selected_column = self.column_combobox.get()
        if selected_column:
            self.concatenation_label.config(
                text=f"File Name:  {self.file_name_var.get()}_{selected_column.lower()}"
            )
