import pandas as pd
from sqlalchemy import create_engine, text
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import logging
import datetime
from datetime import datetime, timedelta


# Configurar logging
logging.basicConfig(filename='stock_monitor.log', level=logging.DEBUG,
                    format='%(asctime)s:%(levelname)s:%(message)s')

# Cargar el motor de conexión con la base de datos
dsn = 'AltieriGestion'
username = 'sa'
password = 'Axoft1988'
connection_string = f'mssql+pyodbc://{username}:{password}@{dsn}?driver=SQL+Server'
engine = create_engine(connection_string)

# Configuración inicial de Selenium y WhatsApp Web
def setup_whatsapp():
    options = webdriver.ChromeOptions()
    options.add_argument("--user-data-dir=C:\\Users\\nmercado\\Documents\\Scripts\\chrome_data")
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    driver.get("https://web.whatsapp.com")
    input("Escanea el código QR de WhatsApp y presiona Enter cuando estés listo.")
    return driver

# Función para enviar un mensaje a un grupo de WhatsApp
def send_whatsapp_message(driver, group_name, message):
    try:
        search_box = WebDriverWait(driver, 50).until(
            EC.presence_of_element_located((By.XPATH, '//div[@contenteditable="true"][@data-tab="3"]'))
        )
        search_box.send_keys(group_name + Keys.ENTER)

        message_box = WebDriverWait(driver, 50).until(
            EC.presence_of_element_located((By.XPATH, '//div[@contenteditable="true"][@data-tab="10"]'))
        )
        for part in message.split("\n"):
            message_box.send_keys(part)
            message_box.send_keys(Keys.SHIFT + Keys.ENTER)  # Para crear una nueva línea
        message_box.send_keys(Keys.ENTER)  # Enviar el mensaje final
        logging.info(f"Mensaje enviado al grupo {group_name}.")
    except Exception as e:
        logging.error(f"Error al enviar mensaje por WhatsApp: {e}")


def format_full_list_message(stock_list):
    full_list_msg = "*LISTA DE STOCK BAJO:*\n"
    for index, row in stock_list.iterrows():
        full_list_msg += f"- {row['DESCRIPCION']} (Stock: {row['StockActual']})\n"
    return full_list_msg

# Consulta SQL para obtener el stock actual
sql_query = text("""
WITH StockTotal AS (
    SELECT 
        m.COD_ART, 
        m.ID_UBICACION,
        SUM(m.CANTIDAD) AS StockActual
    FROM dbo.STOCK_REGISTRO_DE_MOVIMIENTOS m
    WHERE m.ID_UBICACION IN ('01alm1')
    GROUP BY m.COD_ART, m.ID_UBICACION
)
SELECT 
    s.COD_ART, 
    a.DESCRIPCION, 
    s.ID_UBICACION,
    f.FAMILIA,
    a.STK_MIN,
    COALESCE(s.StockActual, 0) AS StockActual, 
    a.STK_MIN - COALESCE(s.StockActual, 0) AS Diferencia
FROM dbo.STOCK_ARTICULOS a
LEFT JOIN StockTotal s ON a.COD_ART = s.COD_ART
LEFT JOIN dbo.STOCK_FAMILIAS f ON a.FAM_ART = f.ID_TIP_ART
WHERE (a.STK_MIN - COALESCE(s.StockActual, 0)) > 0
AND f.FAMILIA NOT IN ('INSUMOS DE MANTENIMIENTO', 'REP.MAQUINAS', 'BARNICES', 'CAJAS', 'INSUMOS P/BARNIZ', 'MAT.PRIMA TAPONES', 'PACKAGING', 'PINTURA PREPARADAS', 'PINTURAS', 'SEPARADORES', 'STAMPING')
AND s.COD_ART IS NOT NULL AND s.COD_ART != ''
ORDER BY f.FAMILIA, s.COD_ART, s.ID_UBICACION;
""")

driver = setup_whatsapp()
group_name = "Stock-Notificaciones"

previous_df = pd.DataFrame()
last_full_send_time = datetime.now() - timedelta(minutes=5)

while True:
    current_time = datetime.now()
    try:
        with engine.connect() as connection:
            current_df = pd.read_sql(sql_query, connection)
            logging.info("Consulta SQL ejecutada con éxito.")
    except Exception as e:
        logging.error(f"Error al realizar la consulta SQL: {e}")
        continue

    # Enviar la lista completa cada 5 minutos
    if (current_time - last_full_send_time).total_seconds() >= 3000:
        full_list_msg = format_full_list_message(current_df)
        send_whatsapp_message(driver, group_name, full_list_msg)
        logging.info("Lista completa de stock bajo enviada.")
        last_full_send_time = current_time

    time.sleep(60)  # Esperar 1 minuto antes de la próxima verificación