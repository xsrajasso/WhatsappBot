from datetime import datetime

# Obtener la fecha actual
fecha_actual = datetime.now()

# Obtener el día de la semana en número (lunes=0, martes=1, ..., domingo=6)
dia_semana_numero = fecha_actual.weekday()

# Opcional: convertir el número en el nombre del día de la semana
dias_semana = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"]
dia_semana_nombre = dias_semana[dia_semana_numero]

# Imprimir el resultado
print("Hoy es:", dia_semana_nombre)
