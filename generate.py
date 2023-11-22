import pandas as pd
from docx import Document
import os


def replace_fields(doc, row, field_mapping):
    """
    Replaces the specified fields in the document.

    Parameters:
        - doc: python-docx Document.
        - row: Pandas DataFrame row containing the data.
        - field_mapping: Dictionary where keys are the placeholders in the document
          and values are the corresponding column names in the Excel file.
    """
    for paragraph in doc.paragraphs:
        for field, col_name in field_mapping.items():
            for run in paragraph.runs:
                run.text = run.text.replace(field, str(row[col_name]))


def generate_document_for_row(
    row, template_doc, destination_folder, field_mapping, column_like_doc_name
):
    """
    Generates a document for a specific row in the Excel file.

    Parameters:
        - row: Pandas DataFrame row containing the data.
        - template_doc: Path to the template document.
        - destination_folder: Folder where the generated documents will be saved.
        - field_mapping: Dictionary where keys are the placeholders in the document
          and values are the corresponding column names in the Excel file.
    """
    # Create a new document or load an existing one
    doc = Document(template_doc)

    # Call the function to replace the fields
    replace_fields(doc, row, field_mapping)

    # Save the document with replaced fields
    document_name = str(row[field_mapping[column_like_doc_name]]).replace(" ", "_")
    new_document_path = os.path.join(destination_folder, f"{document_name}.docx")
    doc.save(new_document_path)

    print(f"Document generated for {document_name}. File saved at: {new_document_path}")


if __name__ == "__main__":
    # Example of usage
    example_excel_path = os.path.join(os.getcwd(), "example_table.xlsx")
    example_template_doc_path = os.path.join(os.getcwd(), "example_doc.docx")
    example_destination_folder = os.path.join(os.getcwd(), "generated-documents")

    # Define the field mapping (placeholder in document: corresponding column name in Excel)
    example_field_mapping = {
        # --Marker--:--Excel Column----#
        "[nombre]": "NOMBRE",
        "[cedula]": "CEDULA",
        "[fecha]": "FECHA",
    }
    # Define the column name that will be used to name the generated documents
    example_column_like_doc_name = "[nombre]"

    # Verify the existence of files and folders
    if not os.path.exists(example_excel_path) or not os.path.exists(
        example_template_doc_path
    ):
        print("Error: Excel file or template document not found.")
    else:
        # Get data from the Excel file
        try:
            excel_data = pd.read_excel(example_excel_path, sheet_name="Sheet")
        except Exception as e:
            print(f"Error loading Excel file: {e}")
        else:
            # Create the destination folder if it doesn't exist
            if not os.path.exists(example_destination_folder):
                os.makedirs(example_destination_folder)

            for _, row in excel_data.iterrows():
                # Generate document for each row in the Excel file
                generate_document_for_row(
                    row,
                    example_template_doc_path,
                    example_destination_folder,
                    example_field_mapping,
                    example_column_like_doc_name,
                )
