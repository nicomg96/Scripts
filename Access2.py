import pyodbc
import pandas as pd
import time
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# Configuración de la conexión a la base de datos
dsn = 'AltieriGestion'
username = 'sa'
password = 'Axoft1988'
mant_otm_table = 'dbo.MANT_OTM'
prod_maquinas_table = 'dbo.PROD_MAQUINAS'

# Columnas que queremos extraer de MANT_OTM
mant_otm_columns = [
    'ID_OTM', 'FECHA_CARGA', 'ESTADO', 'DESCRIPCION_CORTA', 'ID_MAQUINA'
]

# Mapeo de estados
estado_mapping = {
    1: 'PENDIENTE',
    5: 'EN CURSO',
    3: 'TERMINADA',
    4: 'EN PAUSA',
    2: 'ANULADA'
}

# Configuración de email
sender_email = "mantenimiento2@altieri.com.ar"
receiver_email = ["ingenieria@altieri.com.ar", "nicolasmercadog@gmail.com", "mantenimiento@altieri.com.ar", "operaciones@altieri.com.ar", "rsarmiento@altieri.com.ar"]
smtp_server = "smtp.office365.com"
smtp_port = 587
email_password = "Fal11053"  # Cambia esto a la contraseña correcta

# Función para enviar un correo electrónico
def send_email(changes):
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = ", ".join(receiver_email)  # Convertir la lista de destinatarios a una cadena
    msg['Subject'] = "Nuevo orden o modficiación de orden"

    body = f"Se detectaron los siguientes cambios:\n\n{changes}"
    msg.attach(MIMEText(body, 'plain'))

    with smtplib.SMTP(smtp_server, smtp_port) as server:
        server.starttls()
        server.login(sender_email, email_password)
        server.sendmail(sender_email, receiver_email, msg.as_string())

# Función para formatear los cambios detectados
def format_changes(diff_df):
    formatted_changes = ""
    for _, row in diff_df.iterrows():
        formatted_changes += f"ID_OTM: {row['ID_OTM']}\n"
        formatted_changes += f"ID_MAQUINA: {row['ID_MAQUINA']}\n"
        if pd.notna(row['FECHA_CARGA']):
            formatted_changes += f"FECHA_CARGA: {row['FECHA_CARGA']}\n"
        if pd.notna(row['ESTADO']):
            formatted_changes += f"ESTADO: {row['ESTADO']}\n"
        if pd.notna(row['DESCRIPCION_CORTA']):
            formatted_changes += f"DESCRIPCION_CORTA: {row['DESCRIPCION_CORTA']}\n"
        formatted_changes += "\n"
    return formatted_changes

# Función para conectarse a la base de datos y extraer los datos filtrados
def fetch_data():
    # Crear conexión
    conn = pyodbc.connect(f'DSN={dsn};UID={username};PWD={password}')
    
    # Query SQL para obtener los datos de MANT_OTM filtrados
    query = f"""
    SELECT {', '.join(mant_otm_columns)}
    FROM {mant_otm_table}
    WHERE TIPO_OTM = 'CORRECTIVO'
    """
    
    # Leer los datos de MANT_OTM en un DataFrame
    mant_otm_df = pd.read_sql(query, conn)
    
    # Query SQL para obtener los datos de PROD_MAQUINAS
    query_maquinas = f"""
    SELECT ID_MAQUINA, NOMBRE_MAQUINA
    FROM {prod_maquinas_table}
    """
    
    # Leer los datos de PROD_MAQUINAS en un DataFrame
    maquinas_df = pd.read_sql(query_maquinas, conn)
    
    # Cerrar la conexión
    conn.close()
    
    # Reemplazar ID_MAQUINA por NOMBRE_MAQUINA
    merged_df = pd.merge(mant_otm_df, maquinas_df, on='ID_MAQUINA', how='left')
    merged_df['ID_MAQUINA'] = merged_df['NOMBRE_MAQUINA']
    merged_df = merged_df.drop(columns=['NOMBRE_MAQUINA'])  # Eliminar la columna temporal
    
    # Reemplazar valores en la columna ESTADO
    merged_df['ESTADO'] = merged_df['ESTADO'].map(estado_mapping)
    
    return merged_df

# Función para comparar DataFrames y detectar cambios
def detect_changes(df_old, df_new):
    changes = ""
    diff = pd.concat([df_old, df_new]).drop_duplicates(keep=False)
    if not diff.empty:
        changes = format_changes(diff)
    return changes

# Ruta del archivo Excel
excel_file = 'data.xlsx'

# Bucle infinito para actualizar el archivo cada 1 minuto y enviar correos en caso de cambios
previous_df = pd.DataFrame()

while True:
    # Extraer los datos
    current_df = fetch_data()
    
    # Comparar con el DataFrame anterior para detectar cambios
    changes = detect_changes(previous_df, current_df)
    
    if changes:
        send_email(changes)
        print(f'Se detectaron cambios:\n{changes}')
    
    # Actualizar el archivo Excel
    current_df.to_excel(excel_file, index=False)
    print(f'Datos actualizados y escritos en {excel_file}')
    
    # Actualizar el DataFrame anterior
    previous_df = current_df.copy()
    
    # Esperar 60 segundos antes de la siguiente actualización
    time.sleep(60)
