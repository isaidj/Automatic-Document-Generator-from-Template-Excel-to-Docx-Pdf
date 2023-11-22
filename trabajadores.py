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


def generar_documentos(ruta_excel, ruta_documento_word, carpeta_destino):
    # Crear la carpeta de destino si no existe
    if not os.path.exists(carpeta_destino):
        os.makedirs(carpeta_destino)

    # Obtener los valores de los campos desde el archivo Excel
    archivo_excel = pd.rea(ruta_excel, sheet_name="Sheet")
    columnas = ["NOMBRE", "CEDULA", "FECHA"]

    for index, fila in archivo_excel.iterrows():
        # Crear un diccionario con los nombres de los campos y sus valores
        campos_a_reemplazar = {
            "[nombre]": fila["NOMBRE"],
            "[cedula]": str(
                fila["CEDULA"]
            ),  # Asegurarse de que la cédula sea un string
            "[fecha]": str(fila["FECHA"]),  # Asegurarse de que la fecha sea un string
        }

        # Crear un nuevo documento o cargar uno existente
        doc = Document(ruta_documento_word)

        # Llamar a la función para reemplazar los campos
        reemplazar_campos(doc, campos_a_reemplazar)

        # Guardar el documento con los campos reemplazados
        nombre_persona = campos_a_reemplazar["[nombre]"].replace(
            " ", "_"
        )  # Reemplazar espacios en blanco con guiones bajos
        ruta_nuevo_documento = os.path.join(carpeta_destino, f"{nombre_persona}.docx")
        doc.save(ruta_nuevo_documento)

        print(
            f"Documento generado para {nombre_persona}. Puedes encontrar el archivo en: {ruta_nuevo_documento}"
        )


# Ejemplo de uso
ruta_excel_ejemplo = os.path.join(os.getcwd(), "base-de-datos-trabajadores.xlsx")
ruta_documento_word_ejemplo = os.path.join(
    os.getcwd(), "FORMATO-SUSTITUCIÓN PATRONAL  SYVER SAS ORIOL INTENACIONAL SAS.docx"
)
carpeta_destino_ejemplo = os.path.join(os.getcwd(), "documentos-generados")
generar_documentos(
    ruta_excel_ejemplo, ruta_documento_word_ejemplo, carpeta_destino_ejemplo
)
