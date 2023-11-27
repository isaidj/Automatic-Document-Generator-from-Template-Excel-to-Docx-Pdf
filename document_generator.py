import os
import pandas as pd
from docx import Document


class DocumentGenerator:
    def __init__(self):
        pass

    def replace_fields(self, doc, row, field_mapping):
        for paragraph in doc.paragraphs:
            for field, col_name in field_mapping.items():
                for run in paragraph.runs:
                    run.text = run.text.replace(field, str(row[col_name]))

    def generate_document_for_row(
        self,
        row,
        template_doc,
        destination_folder,
        field_mapping,
        column_like_doc_name,
        file_name,
    ):
        doc = Document(template_doc)
        self.replace_fields(doc, row, field_mapping)
        document_name = f"{file_name}_{str(row[field_mapping[column_like_doc_name]]).replace(' ', '_')}"
        new_document_path = os.path.join(destination_folder, f"{document_name}.docx")
        doc.save(new_document_path)
        print(
            f"Document generated for {document_name}. File saved at: {new_document_path}"
        )
