import os
from sqlalchemy import create_engine
import pandas as pd
import time
import logging
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from dotenv import load_dotenv

from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

# Configuración de logging
logging.basicConfig(
    filename='app.log',
    level=logging.DEBUG,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

# Cargar variables de entorno desde un archivo .env
load_dotenv()

# Configuración de la conexión a la base de datos usando SQLAlchemy
dsn = 'AltieriGestion'
username = 'sa'
password = 'Axoft1988'
mant_otm_table = 'dbo.MANT_OTM'
prod_maquinas_table = 'dbo.PROD_MAQUINAS'

# Crear la cadena de conexión para SQLAlchemy
connection_string = f'mssql+pyodbc://{username}:{password}@{dsn}?driver=SQL+Server'

# Crear el motor de conexión
engine = create_engine(connection_string)

# Columnas que queremos extraer de MANT_OTM
mant_otm_columns = [
    'ID_OTM', 'FECHA_CARGA', 'ESTADO', 'DESCRIPCION_CORTA', 'ID_MAQUINA', 'ID_OPERARIO'
]

# Mapeo de estados
estado_mapping = {
    1: 'PENDIENTE',
    5: 'EN CURSO',
    3: 'TERMINADA',
    4: 'EN PAUSA',
    2: 'ANULADA'
}

# Crear una lista de diccionarios con los datos proporcionados
operarios = [
    {'ID_OPERARIO': 1, 'NOMBRE CORTO OPER': 'CC', 'APELLIDO': 'CARPIO', 'NOMBRE': 'CRISTIAN'},
    {'ID_OPERARIO': 2, 'NOMBRE CORTO OPER': 'CG', 'APELLIDO': 'CARPIO', 'NOMBRE': 'GUSTAVO'},
    {'ID_OPERARIO': 4, 'NOMBRE CORTO OPER': 'FA', 'APELLIDO': 'ALFIERI', 'NOMBRE': 'FRANCO'},
    {'ID_OPERARIO': 5, 'NOMBRE CORTO OPER': 'GJ', 'APELLIDO': 'GONZALEZ', 'NOMBRE': 'JOSE'},
    {'ID_OPERARIO': 6, 'NOMBRE CORTO OPER': 'LD', 'APELLIDO': 'LUNA', 'NOMBRE': 'DANIEL'},
    {'ID_OPERARIO': 7, 'NOMBRE CORTO OPER': 'CA', 'APELLIDO': 'CORNEJO', 'NOMBRE': 'ALEJANDRO'},
    {'ID_OPERARIO': 9, 'NOMBRE CORTO OPER': 'SM', 'APELLIDO': 'SUAREZ', 'NOMBRE': 'MARCELO'},
    {'ID_OPERARIO': 12, 'NOMBRE CORTO OPER': 'MG', 'APELLIDO': 'MARCOS', 'NOMBRE': 'GUSTAVO'},
    {'ID_OPERARIO': 13, 'NOMBRE CORTO OPER': 'MA', 'APELLIDO': 'MENESES', 'NOMBRE': 'ADRIAN'},
    {'ID_OPERARIO': 14, 'NOMBRE CORTO OPER': 'ZF', 'APELLIDO': 'ZARATE', 'NOMBRE': 'FERNANDO'},
    {'ID_OPERARIO': 15, 'NOMBRE CORTO OPER': 'PM', 'APELLIDO': 'PANELLA', 'NOMBRE': 'MARIANO NICOLAS'},
    {'ID_OPERARIO': 16, 'NOMBRE CORTO OPER': 'GG', 'APELLIDO': 'GUARÑOLO', 'NOMBRE': 'GUSTAVO'},
    {'ID_OPERARIO': 18, 'NOMBRE CORTO OPER': 'RJ', 'APELLIDO': 'REINOSO', 'NOMBRE': 'JAVIER'},
    {'ID_OPERARIO': 19, 'NOMBRE CORTO OPER': 'FC', 'APELLIDO': 'CALDERÓN', 'NOMBRE': 'FACUNDO'},
    {'ID_OPERARIO': 20, 'NOMBRE CORTO OPER': 'GZ', 'APELLIDO': 'ZANON', 'NOMBRE': 'GUSTAVO'},
    {'ID_OPERARIO': 23, 'NOMBRE CORTO OPER': 'JG', 'APELLIDO': 'GUTIERREZ', 'NOMBRE': 'JORGE LUIS'},
    {'ID_OPERARIO': 24, 'NOMBRE CORTO OPER': 'AE', 'APELLIDO': 'ALVAREZ', 'NOMBRE': 'MARTIN EMANUEL'},
    {'ID_OPERARIO': 25, 'NOMBRE CORTO OPER': 'AL', 'APELLIDO': 'ALFIERI', 'NOMBRE': 'LEANDRO CRISTIAN'},
    {'ID_OPERARIO': 27, 'NOMBRE CORTO OPER': 'CH', 'APELLIDO': 'HERRERO', 'NOMBRE': 'CRISTIAN'},
    {'ID_OPERARIO': 29, 'NOMBRE CORTO OPER': 'AC', 'APELLIDO': 'CRESPIN', 'NOMBRE': 'ALBERTO'},
    {'ID_OPERARIO': 30, 'NOMBRE CORTO OPER': 'TE', 'APELLIDO': 'TORO', 'NOMBRE': 'EDUARDO CARLOS'},
    {'ID_OPERARIO': 32, 'NOMBRE CORTO OPER': 'JO', 'APELLIDO': 'ORTIZ', 'NOMBRE': 'JAVIER'},
    {'ID_OPERARIO': 33, 'NOMBRE CORTO OPER': 'PG', 'APELLIDO': 'PACHECO', 'NOMBRE': 'GUILLERMO'},
    {'ID_OPERARIO': 34, 'NOMBRE CORTO OPER': 'BP', 'APELLIDO': 'BUTTINI', 'NOMBRE': 'PABLO'},
    {'ID_OPERARIO': 35, 'NOMBRE CORTO OPER': 'FP', 'APELLIDO': 'FRANCO', 'NOMBRE': 'PABLO'},
    {'ID_OPERARIO': 36, 'NOMBRE CORTO OPER': 'VR', 'APELLIDO': 'VIDELA', 'NOMBRE': 'ROBERTO'},
    {'ID_OPERARIO': 40, 'NOMBRE CORTO OPER': 'MP', 'APELLIDO': 'PELETAY', 'NOMBRE': 'MARCELO'},
    {'ID_OPERARIO': 42, 'NOMBRE CORTO OPER': 'RM', 'APELLIDO': 'ROIZ', 'NOMBRE': 'MARCELO'},
    {'ID_OPERARIO': 43, 'NOMBRE CORTO OPER': 'ZJ', 'APELLIDO': 'ROIZ', 'NOMBRE': 'JORGE'},
    {'ID_OPERARIO': 44, 'NOMBRE CORTO OPER': 'AS', 'APELLIDO': 'AGOSTINI', 'NOMBRE': 'SERGIO'},
    {'ID_OPERARIO': 46, 'NOMBRE CORTO OPER': 'GO', 'APELLIDO': 'GONZALEZ', 'NOMBRE': 'OMAR'},
    {'ID_OPERARIO': 47, 'NOMBRE CORTO OPER': 'HC', 'APELLIDO': 'HERRERO', 'NOMBRE': 'CRISTIAN'},
    {'ID_OPERARIO': 48, 'NOMBRE CORTO OPER': 'GH', 'APELLIDO': 'GUAQUINCHAY', 'NOMBRE': 'HUGO'},
    {'ID_OPERARIO': 51, 'NOMBRE CORTO OPER': 'AD', 'APELLIDO': 'AMPUERO', 'NOMBRE': 'DANIEL'},
    {'ID_OPERARIO': 54, 'NOMBRE CORTO OPER': 'AH', 'APELLIDO': 'CECCHIN', 'NOMBRE': 'ALEJANDRO'},
    {'ID_OPERARIO': 55, 'NOMBRE CORTO OPER': 'MM', 'APELLIDO': 'MENESES', 'NOMBRE': 'MAXIMILIANO'},
    {'ID_OPERARIO': 59, 'NOMBRE CORTO OPER': 'GM', 'APELLIDO': 'MONTENEGRO', 'NOMBRE': 'GERMAN'},
    {'ID_OPERARIO': 62, 'NOMBRE CORTO OPER': 'AG', 'APELLIDO': 'GUTIERREZ', 'NOMBRE': 'ALEXANDER'},
    {'ID_OPERARIO': 63, 'NOMBRE CORTO OPER': 'VM', 'APELLIDO': 'MASON', 'NOMBRE': 'MARIO'},
    {'ID_OPERARIO': 64, 'NOMBRE CORTO OPER': 'OL', 'APELLIDO': 'LUCERO', 'NOMBRE': 'GABRIEL'},
    {'ID_OPERARIO': 65, 'NOMBRE CORTO OPER': 'PD', 'APELLIDO': 'DIAZ', 'NOMBRE': 'PABLO'},
    {'ID_OPERARIO': 66, 'NOMBRE CORTO OPER': '66', 'APELLIDO': 'CAMBIAR', 'NOMBRE': 'CAMBIAR'},
    {'ID_OPERARIO': 67, 'NOMBRE CORTO OPER': 'MN', 'APELLIDO': 'MERCADO', 'NOMBRE': 'NICOLAS'},
    {'ID_OPERARIO': 70, 'NOMBRE CORTO OPER': 'GR', 'APELLIDO': 'REINOSO', 'NOMBRE': 'GONZALO'},
    {'ID_OPERARIO': 71, 'NOMBRE CORTO OPER': 'JV', 'APELLIDO': 'VALERA', 'NOMBRE': 'JOSE'},
    {'ID_OPERARIO': 72, 'NOMBRE CORTO OPER': 'FB', 'APELLIDO': 'BONACORSO', 'NOMBRE': 'FRANCO'},
    {'ID_OPERARIO': 73, 'NOMBRE CORTO OPER': 'GC', 'APELLIDO': 'GASTILLO', 'NOMBRE': 'GUSTAVO'},
    {'ID_OPERARIO': 74, 'NOMBRE CORTO OPER': 'BN', 'APELLIDO': 'BAEZ', 'NOMBRE': 'NICOLAS'},
    {'ID_OPERARIO': 75, 'NOMBRE CORTO OPER': 'EM', 'APELLIDO': 'MURUA', 'NOMBRE': 'EXEQUIEL'},
    {'ID_OPERARIO': 76, 'NOMBRE CORTO OPER': 'JD', 'APELLIDO': 'DOLCEMASCOLO', 'NOMBRE': 'JOEL'},
    {'ID_OPERARIO': 78, 'NOMBRE CORTO OPER': 'GA', 'APELLIDO': 'GONZALEZ', 'NOMBRE': 'ABEL'},
    {'ID_OPERARIO': 79, 'NOMBRE CORTO OPER': 'EG', 'APELLIDO': 'GIMENEZ', 'NOMBRE': 'EBER'},
    {'ID_OPERARIO': 80, 'NOMBRE CORTO OPER': 'G1', 'APELLIDO': 'GIMENEZ', 'NOMBRE': 'CRISTIAN'},
    {'ID_OPERARIO': 81, 'NOMBRE CORTO OPER': 'GT', 'APELLIDO': 'GUTIERREZ', 'NOMBRE': 'MAURO'},
    {'ID_OPERARIO': 82, 'NOMBRE CORTO OPER': 'NV', 'APELLIDO': 'VALENZUELA', 'NOMBRE': 'NAHUEL'}
]

operario_nombre_completo = {op['ID_OPERARIO']: f"{op['APELLIDO']} {op['NOMBRE']}" for op in operarios}


# Configuración de Selenium y WhatsApp Web
def setup_whatsapp():
    # Usar ChromeDriver con WebDriver Manager
    options = webdriver.ChromeOptions()
    options.add_argument("--user-data-dir=C:\\Users\\nmercado\\Documents\\Scripts\\chrome_data2")  # Para mantener la sesión iniciada
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    driver.get("https://web.whatsapp.com")
    input("Escanea el código QR de WhatsApp y presiona Enter cuando estés listo.")
    time.sleep(10)
    return driver

# Función para enviar un mensaje a un grupo de WhatsApp
def send_whatsapp_message(driver, group_name, message):
    try:
        # Esperar a que el cuadro de búsqueda esté presente
        search_box = WebDriverWait(driver, 50).until(
            EC.presence_of_element_located((By.XPATH, '//div[@contenteditable="true"][@data-tab="3"]'))
        )
        search_box.click()

        # Limpiar el cuadro de búsqueda antes de escribir
        search_box.clear()
        search_box.send_keys(Keys.CONTROL + "a")
        search_box.send_keys(Keys.DELETE)

        # Escribir el nombre del grupo
        search_box.send_keys(group_name)
        search_box.send_keys(Keys.ENTER)

        # Verificar si el grupo ha sido seleccionado
        group = WebDriverWait(driver, 50).until(
            EC.presence_of_element_located((By.XPATH, f'//span[@title="{group_name}"]'))
        )

        # Usar JavaScript para hacer clic en el grupo
        driver.execute_script("arguments[0].click();", group)

        # Esperar a que el cuadro de mensaje esté presente con data-tab="10"
        message_box = WebDriverWait(driver, 50).until(
            EC.presence_of_element_located((By.XPATH, '//div[@contenteditable="true"][@data-tab="10"]'))
        )
        message_box.click()

        # Escribir y enviar el mensaje
        for part in message.split("\n"):  # Enviar mensajes de varias líneas
            message_box.send_keys(part)
            message_box.send_keys(Keys.SHIFT + Keys.ENTER)  # Para crear una nueva línea
        message_box.send_keys(Keys.ENTER)  # Enviar el mensaje final
        
        logging.info(f"Mensaje enviado al grupo {group_name}.")
    except TimeoutException:
        logging.error(f"No se encontró el cuadro de búsqueda o el cuadro de mensajes a tiempo.")
    except Exception as e:
        logging.error(f"Error al enviar mensaje por WhatsApp: {e}")




# Función para formatear los cambios detectados
def format_changes(diff_df, df_old):
    formatted_changes = ""
    for _, row in diff_df.iterrows():
        # Verificar si la OTM ya existía en el df_old
        otm_id = row['ID_OTM']
        old_row = df_old[df_old['ID_OTM'] == otm_id]

        if not old_row.empty:  # Si la OTM ya existía, mostrar solo los cambios más recientes
            # Comparar el estado
            old_estado = old_row.iloc[0]['ESTADO']
            new_estado = row['ESTADO']
            
            if old_estado != new_estado:  # Solo mostrar el cambio si hay una diferencia
                formatted_changes += "➡️            *CAMBIO EN OTM*            ⬅️\n"
                formatted_changes += f"N° OTM: *{row['ID_OTM']}*\n"
                formatted_changes += f"N° MAQUINA: {row['ID_MAQUINA']}\n"
                formatted_changes += f"FECHA CARGA: {pd.to_datetime(row['FECHA_CARGA']).strftime('%Y-%m-%d %H:%M')}\n"
                formatted_changes += f"ESTADO: ~{old_estado}~ --> *{new_estado}*\n"
                formatted_changes += f"DESCRIPCION: {row['DESCRIPCION_CORTA']}\n"
                formatted_changes += f"OPERARIO: {row['ID_OPERARIO']}\n"
        else:  # Si la OTM es nueva, mostrarla como nueva
            formatted_changes += "*❗                    Nueva OTM                      *\n"
            formatted_changes += f"N° OTM: {row['ID_OTM']}\n"
            formatted_changes += f"N° MAQUINA: {row['ID_MAQUINA']}\n"
            formatted_changes += f"FECHA CARGA: {pd.to_datetime(row['FECHA_CARGA']).strftime('%Y-%m-%d %H:%M')}\n"
            formatted_changes += f"ESTADO: {row['ESTADO']}\n"
            formatted_changes += f"DESCRIPCION_CORTA: {row['DESCRIPCION_CORTA']}\n"
            formatted_changes += f"OPERARIO: {row['ID_OPERARIO']}\n"

        formatted_changes += "\n"  # Separador entre OTMs
    return formatted_changes



# Función para conectarse a la base de datos y extraer los datos filtrados
def fetch_data():
    try:
        logging.info("Intentando conectar a la base de datos.")
        # Incluir la columna ID_OPERARIO en la consulta de MANT_OTM
        query = f"""
        SELECT {', '.join(mant_otm_columns)}
        FROM {mant_otm_table}
        WHERE TIPO_OTM = 'CORRECTIVO'
        """
        mant_otm_df = pd.read_sql(query, engine)

        # Consulta para obtener datos de PROD_MAQUINAS
        query_maquinas = f"""
        SELECT ID_MAQUINA, NOMBRE_MAQUINA
        FROM {prod_maquinas_table}
        """
        maquinas_df = pd.read_sql(query_maquinas, engine)
        logging.info("Datos extraídos exitosamente.")
    except Exception as e:
        logging.error(f"Error al conectarse a la base de datos: {e}")
        return pd.DataFrame()

    try:
        # Realizar merge entre MANT_OTM y PROD_MAQUINAS
        merged_df = pd.merge(mant_otm_df, maquinas_df, on='ID_MAQUINA', how='left')
        merged_df['ID_MAQUINA'] = merged_df['NOMBRE_MAQUINA'].fillna('DESCONOCIDA')
        merged_df = merged_df.drop(columns=['NOMBRE_MAQUINA'])

        # Mapear el ID_OPERARIO al nombre completo
        # Verifica si 'ID_OPERARIO' tiene valores válidos antes de mapear
        if 'ID_OPERARIO' in merged_df.columns:
            merged_df['ID_OPERARIO'] = merged_df['ID_OPERARIO'].map(operario_nombre_completo).fillna('DESCONOCIDO')
        else:
            logging.warning("ID_OPERARIO no encontrado en los datos.")

        # Mapear el estado
        merged_df['ESTADO'] = merged_df['ESTADO'].map(estado_mapping).fillna('DESCONOCIDO')

        logging.info("Datos procesados correctamente.")
    except Exception as e:
        logging.error(f"Error al procesar los datos: {e}")
        return pd.DataFrame()

    return merged_df

def detect_changes(df_old, df_new):
    try:
        changes = ""
        diff = pd.concat([df_old, df_new]).drop_duplicates(keep=False)
        if not diff.empty:
            changes = format_changes(diff, df_old)  # Pasa df_old para comparar
        logging.info("Cambios detectados exitosamente.")
        return changes
    except Exception as e:
        logging.error(f"Error al detectar cambios: {e}")
        return ""

# Ruta del archivo Excel
excel_file = 'data.xlsx'

# Inicializar Selenium y WhatsApp
driver = setup_whatsapp()

# Bucle infinito para actualizar el archivo cada 1 minuto y enviar notificaciones en caso de cambios
previous_df = pd.DataFrame()

while True:
    current_df = fetch_data()
    
    if current_df.empty:
        logging.warning("No se pudo obtener datos de la base de datos, reintentando en 60 segundos...")
    else:
        changes = detect_changes(previous_df, current_df)
        
        if changes:
            send_whatsapp_message(driver, "Mantenimiento-Notificaciones", changes)
            logging.info(f'Se detectaron cambios:\n{changes}')
        
        try:
            current_df.to_excel(excel_file, index=False)
            logging.info(f'Datos actualizados y escritos en {excel_file}')
        except Exception as e:
            logging.error(f"Error al escribir datos en Excel: {e}")
        
        previous_df = current_df.copy()
    
    time.sleep(5)
