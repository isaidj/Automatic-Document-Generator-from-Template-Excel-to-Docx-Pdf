import pandas as pd
from docx import Document
import os


def reemplazar_campos(doc, campos):
    """
    Reemplaza los campos especificados en el documento.

    Parameters:
        - doc: Documento de python-docx.
        - campos: Diccionario donde las claves son los campos a reemplazar y los valores son los nuevos textos.
    """
    for parrafo in doc.paragraphs:
        for campo, nuevo_texto in campos.items():
            for run in parrafo.runs:
                run.text = run.text.replace(campo, nuevo_texto)


def generar_documento_final(ruta_excel, ruta_documento_word, ruta_nuevo_documento):
    # Obtener los valores de los campos desde el archivo Excel
    archivo_excel = pd.read_excel(ruta_excel, sheet_name="Sheet")
    columnas = ["NOMBRE", "CEDULA", "FECHA"]
    valores_columnas = archivo_excel[columnas].iloc[0].tolist()

    # Crear un diccionario con los nombres de los campos y sus valores
    campos_a_reemplazar = {
        "[nombre]": valores_columnas[0],
        "[cedula]": str(
            valores_columnas[1]
        ),  # Asegurarse de que la cédula sea un string
        "[fecha]": str(valores_columnas[2]),  # Asegurarse de que la fecha sea un string
    }

    # Crear un nuevo documento o cargar uno existente
    doc = Document(ruta_documento_word)

    # Llamar a la función para reemplazar los campos
    reemplazar_campos(doc, campos_a_reemplazar)

    # Guardar el documento con los campos reemplazados
    doc.save(ruta_nuevo_documento)

    print(
        f"Campos reemplazados exitosamente. Puedes encontrar el archivo en: {ruta_nuevo_documento}"
    )


# Ejemplo de uso
ruta_excel_ejemplo = os.path.join(os.getcwd(), "base-de-datos-trabajadores.xlsx")
ruta_documento_word_ejemplo = os.path.join(
    os.path.expanduser("~"),
    "desktop",
    "FORMATO-SUSTITUCIÓN PATRONAL  SYVER SAS ORIOL INTENACIONAL SAS.docx",
)
ruta_nuevo_documento_ejemplo = os.path.join(
    os.path.expanduser("~"), "desktop", "documento_final.docx"
)

generar_documento_final(
    ruta_excel_ejemplo, ruta_documento_word_ejemplo, ruta_nuevo_documento_ejemplo
)
