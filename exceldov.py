import pandas as pd
import os

ruta_actual = os.getcwd()
ruta_documento = os.path.join(ruta_actual, "base-de-datos-trabajadores.xlsx")
# Reemplaza 'nombre_del_archivo.xlsx' y 'nombre_de_la_columna' con tus valores reales
archivo_excel = pd.read_excel(ruta_documento, sheet_name="Sheet")
columnas = ["NOMBRE", "CEDULA", "FECHA"]
valores_columnas = archivo_excel[columnas].values.tolist()


print(valores_columnas[0])
