import pandas as pd

# Cargar el archivo Excel
file_path = 'C:\\Users\\nmercado\\Documents\\Scripts\\solar_data4.xlsx'  # Cambia la ruta a donde tienes tu archivo
df = pd.read_excel(file_path)

# Convertir 'Fecha' y 'Hora' a un solo campo de datetime
df['Fecha_Hora'] = pd.to_datetime(df['Fecha'].astype(str) + ' ' + df['Hora'].astype(str))

# Ordenar los datos por fecha y hora en caso de que no estén ordenados
df = df.sort_values(by='Fecha_Hora')

# Establecer umbral de potencia y tiempo consecutivo (en minutos)
threshold = 150
time_interval = 15  # minutos
interval_rows = time_interval // 5  # cada fila es un intervalo de 5 minutos

# Identificar los periodos donde la potencia supera el umbral (usando la columna "Potencia (KW)")
df['Above_Threshold'] = df['Potencia (KW)'] > threshold

# Agrupar valores consecutivos que están por encima del umbral
df['Consecutive'] = df['Above_Threshold'].groupby(
    (df['Above_Threshold'] != df['Above_Threshold'].shift()).cumsum()
).transform('size') * df['Above_Threshold']

# Filtrar los periodos donde hay 3 o más filas consecutivas con valores superiores al umbral
filtered_df = df[df['Consecutive'] >= interval_rows]


# Crear un archivo Excel con las recomendaciones en diferentes hojas
with pd.ExcelWriter('C:\\Users\\nmercado\\Documents\\Scripts\\filtered_solar_data_with_recommendations4.xlsx', engine='xlsxwriter') as writer:
    # Guardar los datos filtrados
    filtered_df[['Fecha_Hora', 'Potencia (KW)']].to_excel(writer, sheet_name='Filtrado', index=False)
    