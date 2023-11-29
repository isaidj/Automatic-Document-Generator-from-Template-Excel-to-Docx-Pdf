import TKinterModernThemes as TKMT
from TKinterModernThemes.WidgetFrame import Widget
from tkinter import ttk, filedialog, messagebox
import tkinter as tk
import pandas as pd
import os


def buttonCMD():
    print("Button clicked!")


class App(TKMT.ThemedTKinterFrame):
    def __init__(self, generator):
        super().__init__("Document Generator - Excel to Docx Template", "park", "dark")
        # self.root = root
        # Icon

        self.root.iconbitmap("icon.ico")
        # variables
        self.generator = generator
        self.excel_path_var = tk.StringVar()
        self.template_doc_path_var = tk.StringVar()
        self.destination_folder_var = tk.StringVar()
        self.file_name_var = tk.StringVar()

        self.column_combobox = None
        self.concatenation_label = None
        self.main_frame = None
        self.is_pdf = tk.BooleanVar()

        # ----------------------------Widgets-----------------------------------

        self.Label("Generate Documents", col=0, row=0, sticky="nsew")
        # ----------------------------Documents Frame-------------------------------------------
        self.documents_frame = self.addLabelFrame(
            "Documents routes"
        )  # placed at row 1, col 0

        # -------------------------------------------------------------------------------
        excel_label = ttk.Label(self.documents_frame.master, text="Excel File:")
        excel_label.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")
        excel_entry = ttk.Entry(
            self.documents_frame.master, textvariable=self.excel_path_var
        )
        excel_entry.grid(row=0, column=1, padx=10, pady=10, sticky="nsew", columnspan=2)
        excel_button = ttk.Button(
            self.documents_frame.master,
            text="...",
            command=self.browse_excel,
            cursor="hand2",
        )
        excel_button.grid(row=0, column=4, padx=10, pady=10, sticky="nsew")

        # ---------------------------------------------------------------------------
        template_doc_label = ttk.Label(
            self.documents_frame.master, text="Template Docx:"
        )
        template_doc_label.grid(row=1, column=0, padx=10, pady=10, sticky="nsew")
        template_doc_entry = ttk.Entry(
            self.documents_frame.master, textvariable=self.template_doc_path_var
        )
        template_doc_entry.grid(
            row=1, column=1, padx=10, pady=10, sticky="nsew", columnspan=2
        )
        template_doc_button = ttk.Button(
            self.documents_frame.master,
            text="...",
            command=self.browse_template_doc,
            width=2,
            cursor="hand2",
        )
        template_doc_button.grid(row=1, column=4, padx=10, pady=10, sticky="nsew")
        # ----------------------------Destination Folder-----------------------------------
        destination_folder_label = ttk.Label(
            self.documents_frame.master, text="Destination Folder:"
        )
        destination_folder_label.grid(row=2, column=0, padx=10, pady=10, sticky="nsew")
        destination_folder_entry = ttk.Entry(
            self.documents_frame.master, textvariable=self.destination_folder_var
        )
        destination_folder_entry.grid(
            row=2, column=1, padx=10, pady=10, sticky="nsew", columnspan=2
        )
        destination_folder_button = ttk.Button(
            self.documents_frame.master,
            text="...",
            command=self.browse_destination_folder,
            cursor="hand2",
        )
        destination_folder_button.grid(row=2, column=4, padx=10, pady=10, sticky="nsew")
        # ----------------------------//Documents Frame--------------------------------------------------------
        # -----------------------------Configuration Frame--------------------------------------------------------
        self.configuration_frame = self.addLabelFrame("Configuration")

        # ----------------------------File Name-----------------------------------
        file_name_label = ttk.Label(self.configuration_frame.master, text="File Name:")
        file_name_label.grid(row=3, column=0, padx=10, pady=10, sticky="nsew")
        file_name_entry = ttk.Entry(
            self.configuration_frame.master, textvariable=self.file_name_var
        )
        file_name_entry.grid(row=3, column=1, padx=10, pady=10, sticky="nsew")
        # ----------------------------Select Column-----------------------------------
        select_column_label = ttk.Label(
            self.configuration_frame.master, text="Select Column:"
        )
        select_column_label.grid(row=4, column=0, padx=10, pady=10, sticky="nsew")
        self.column_combobox = ttk.Combobox(
            self.configuration_frame.master, state="readonly"
        )
        self.column_combobox.grid(row=4, column=1)
        # ----------------------------Concatenation Label-----------------------------------
        self.concatenation_label = ttk.Label(
            self.configuration_frame.master,
            text="",
        )
        self.concatenation_label.grid(
            row=5, column=0, columnspan=4, padx=10, pady=10, sticky="nsew"
        )
        # ----------------------------Checkbox PDF-----------------------------------
        checkbox_pdf = ttk.Checkbutton(
            self.configuration_frame.master,
            text="PDF?",
            variable=self.is_pdf,
            cursor="hand2",
        )
        checkbox_pdf.grid(
            row=6, column=0, columnspan=4, padx=10, pady=10, sticky="nsew"
        )
        # ----------------------------Generate Button-----------------------------------
        generate_button = ttk.Button(
            self.master,
            text="Generate  Documents",
            command=self.generate_documents,
            cursor="hand2",
        )
        generate_button.grid(
            row=7, column=0, columnspan=4, padx=40, pady=20, sticky="nsew"
        )

        # -----------------------------//Configuration Frame--------------------------------------------------------
        # -----------------------------Events-------------------------------------------------------
        self.file_name_var.trace_add("write", self.update_combobox_and_label)
        self.column_combobox.bind_all(
            "<<ComboboxSelected>>", lambda event=None: self.update_combobox_and_label()
        )

        self.run()

    def browse_excel(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if file_path:
            self.excel_path_var.set(file_path)

            excel_data = pd.read_excel(file_path, sheet_name="Sheet")
            column_names = excel_data.columns.tolist()

            if self.column_combobox:
                self.column_combobox.destroy()

            self.column_combobox = ttk.Combobox(
                self.configuration_frame.master, values=column_names, state="readonly"
            )
            self.column_combobox.grid(row=4, column=1)
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
                pdf=self.is_pdf.get(),
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
