import pyodbc
import pandas as pd

# Configuración de la conexión a la base de datos
dsn = 'AltieriGestion'
username = 'sa'  # Si tienes un usuario y contraseña configurados, agrégalos aquí
password = 'Axoft1988'

# Conexión a la base de datos
connection_string = f'DSN={dsn};UID={username};PWD={password}'
connection = pyodbc.connect(connection_string)

# Consulta a la base de datos
query = "SELECT * FROM dbo_MANT_OTM"
df = pd.read_sql(query, connection)

# Cierre de la conexión
connection.close()

# Guardar los datos en un archivo Excel
excel_file = 'output.xlsx'
df.to_excel(excel_file, index=False)

print(f"Datos exportados a {excel_file}")
