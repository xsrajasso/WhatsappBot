import pandas as pd

def buscar_valor_columna_a(ruta_archivo, nombre_hoja, valor_busqueda):
    """
    Busca un valor en la columna A de una hoja de Excel y devuelve el número de fila.

    :param ruta_archivo: Ruta al archivo Excel.
    :param nombre_hoja: Nombre de la hoja donde buscar.
    :param valor_busqueda: Valor a buscar en la columna A.
    :return: Número de fila si se encuentra, o None si no se encuentra.
    """
    try:
        # Cargar la hoja específica del archivo Excel
        df = pd.read_excel(ruta_archivo, sheet_name=nombre_hoja, engine='openpyxl')
    except FileNotFoundError:
        print(f"Error: El archivo '{ruta_archivo}' no se encontró.")
        return None
    except Exception as e:
        print(f"Error al cargar el archivo: {e}")
        return None

    # Verificar si la columna A existe
    if 'A' in df.columns:
        columna_a = df['A']
    else:
        # Alternativamente, si la columna A no tiene nombre, usar el índice
        columna_a = df.iloc[:, 0]

    # Buscar el valor en la columna A
    filas_encontradas = df.index[df.iloc[:, 0] == valor_busqueda].tolist()

    if not filas_encontradas:
        print(f"El valor '{valor_busqueda}' no se encontró en la columna A.")
        return None
    else:
        # `pandas` usa índices basados en 0, mientras que Excel comienza en 1
        # Además, si hay encabezados, el índice 0 corresponde a la primera fila de datos (fila 2 en Excel)
        # Ajustaremos el número de fila en consecuencia
        # Supongamos que la primera fila es el encabezado
        numero_fila_excel = filas_encontradas[0] + 2  # +1 para index 0 y +1 para encabezado
        print(f"El valor '{valor_busqueda}' se encuentra en la fila {numero_fila_excel} de Excel.")
        return numero_fila_excel

# Uso de la función
ruta = "Directorio.xlsx"
hoja = "1A"
valor = 13186

fila_encontrada = buscar_valor_columna_a(ruta, hoja, valor)
