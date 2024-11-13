import pandas as pd

# Función para obtener los nombres de las columnas
def obtener_nombres_columnas(archivo, hoja):
    df = pd.read_excel(archivo, sheet_name=hoja)
    return df.columns.tolist()

# Función para obtener matrículas con más de 3 faltas como enteros
def obtener_matriculas_con_faltas(archivo, hoja, umbral_faltas=3):
    df = pd.read_excel(archivo, sheet_name=hoja)
    matriculas = df.loc[df['Faltas'] > umbral_faltas, 'Matricula'].astype(int).tolist()
    return matriculas

# # Función para obtener matrículas con más de 3 faltas
# def obtener_matriculas_con_faltas(archivo, hoja, umbral_faltas=3):
#     df = pd.read_excel(archivo, sheet_name=hoja)
#     matriculas = df.loc[df['Faltas'] > umbral_faltas, 'Matricula'].tolist()
#     return matriculas

# Especifica el archivo y la hoja
archivo = "Peggy3erPiso.xlsx"  # Cambia el nombre del archivo por el que necesitas
hoja = "1°A"

# Llamar a las funciones
# nombres_columnas = obtener_nombres_columnas(archivo, hoja)
matriculas = obtener_matriculas_con_faltas(archivo, hoja)

# Imprimir los resultados
# print("Nombres de las columnas:", nombres_columnas)
print("Matrículas con más de 3 faltas:", matriculas)
