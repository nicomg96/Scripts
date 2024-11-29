import os
import pyodbc
import pandas as pd
import time
import requests
import logging
from dotenv import load_dotenv

# Configuración de logging
logging.basicConfig(
    filename='app.log',
    level=logging.DEBUG,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

# Cargar variables de entorno desde un archivo .env
load_dotenv()

telegram_bot_token = os.getenv("TELEGRAM_BOT_TOKEN")
telegram_chat_id = os.getenv("TELEGRAM_CHAT_ID")

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

def send_test_message():
    url = f"https://api.telegram.org/bot{telegram_bot_token}/sendMessage"
    payload = {
        "chat_id": telegram_chat_id,
        "text": "Este es un mensaje de prueba."
    }
    try:
        response = requests.post(url, data=payload)
        response.raise_for_status()
        logging.info(f"Mensaje de prueba enviado: {response.json()}")
    except requests.exceptions.RequestException as e:
        logging.error(f"Error al enviar el mensaje de prueba: {e}")

# Llamar a la función de prueba
send_test_message()


# Función para enviar un mensaje de Telegram
def send_telegram(changes):
    message = f"Se detectaron los siguientes cambios:\n\n{changes}"
    url = f"https://api.telegram.org/bot{telegram_bot_token}/sendMessage"
    payload = {
        "chat_id": telegram_chat_id,
        "text": message
    }
    try:
        response = requests.post(url, data=payload)
        response.raise_for_status()
        logging.info(f"Mensaje enviado: {response.json()}")
    except requests.exceptions.RequestException as e:
        logging.error(f"Error al enviar el mensaje: {e}")

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
    try:
        logging.info("Intentando conectar a la base de datos.")
        conn = pyodbc.connect(f'DSN={dsn};UID={username};PWD={password}')
        logging.info("Conexión exitosa.")

        query = f"""
        SELECT {', '.join(mant_otm_columns)}
        FROM {mant_otm_table}
        WHERE TIPO_OTM = 'CORRECTIVO'
        """
        mant_otm_df = pd.read_sql(query, conn)

        query_maquinas = f"""
        SELECT ID_MAQUINA, NOMBRE_MAQUINA
        FROM {prod_maquinas_table}
        """
        maquinas_df = pd.read_sql(query_maquinas, conn)
        conn.close()
        logging.info("Datos extraídos exitosamente.")
    except Exception as e:
        logging.error(f"Error al conectarse a la base de datos: {e}")
        return pd.DataFrame()

    try:
        merged_df = pd.merge(mant_otm_df, maquinas_df, on='ID_MAQUINA', how='left')
        merged_df['ID_MAQUINA'] = merged_df['ID_MAQUINA'].fillna('DESCONOCIDA')
        merged_df['ID_MAQUINA'] = merged_df['NOMBRE_MAQUINA']
        merged_df = merged_df.drop(columns=['NOMBRE_MAQUINA'])
        merged_df['ESTADO'] = merged_df['ESTADO'].map(estado_mapping).fillna('DESCONOCIDO')
        logging.info("Datos procesados correctamente.")
    except Exception as e:
        logging.error(f"Error al procesar los datos: {e}")
        return pd.DataFrame()

    return merged_df

# Función para comparar DataFrames y detectar cambios
def detect_changes(df_old, df_new):
    try:
        changes = ""
        diff = pd.concat([df_old, df_new]).drop_duplicates(keep=False)
        if not diff.empty:
            changes = format_changes(diff)
        logging.info("Cambios detectados exitosamente.")
        return changes
    except Exception as e:
        logging.error(f"Error al detectar cambios: {e}")
        return ""

# Ruta del archivo Excel
excel_file = 'data.xlsx'

# Bucle infinito para actualizar el archivo cada 1 minuto y enviar notificaciones en caso de cambios
previous_df = pd.DataFrame()

while True:
    current_df = fetch_data()
    
    if current_df.empty:
        logging.warning("No se pudo obtener datos de la base de datos, reintentando en 60 segundos...")
    else:
        changes = detect_changes(previous_df, current_df)
        
        if changes:
            send_telegram(changes)
            logging.info(f'Se detectaron cambios:\n{changes}')
        
        try:
            current_df.to_excel(excel_file, index=False)
            logging.info(f'Datos actualizados y escritos en {excel_file}')
        except Exception as e:
            logging.error(f"Error al escribir datos en Excel: {e}")
        
        previous_df = current_df.copy()
    
    time.sleep(60)
