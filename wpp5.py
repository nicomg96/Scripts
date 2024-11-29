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
from datetime import datetime, timedelta

# Configuración del logging
logging.basicConfig(filename='combined_monitor.log', level=logging.DEBUG,
                    format='%(asctime)s:%(levelname)s:%(message)s')

# Datos para la conexión con la base de datos
dsn = 'AltieriGestion'
username = 'sa'
password = 'Axoft1988'
connection_string = f'mssql+pyodbc://{username}:{password}@{dsn}?driver=SQL+Server'
engine = create_engine(connection_string)

# Iniciar Selenium y WhatsApp Web
def setup_whatsapp():
    options = webdriver.ChromeOptions()
    options.add_argument("--user-data-dir=C:\\Users\\nmercado\\Documents\\Scripts\\chrome_data")
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    driver.get("https://web.whatsapp.com")
    input("Escanea el código QR de WhatsApp y presiona Enter cuando estés listo.")
    return driver

# Función para enviar mensajes a WhatsApp
def send_whatsapp_message(driver, group_name, message):
    try:
        search_box = WebDriverWait(driver, 50).until(
            EC.presence_of_element_located((By.XPATH, '//div[@contenteditable="true"][@data-tab="3"]'))
        )
        search_box.click()
        search_box.clear()
        search_box.send_keys(group_name + Keys.ENTER)

        message_box = WebDriverWait(driver, 50).until(
            EC.presence_of_element_located((By.XPATH, '//div[@contenteditable="true"][@data-tab="10"]'))
        )
        for part in message.split("\n"):
            message_box.send_keys(part)
            message_box.send_keys(Keys.SHIFT + Keys.ENTER)
        message_box.send_keys(Keys.ENTER)
        logging.info(f"Mensaje enviado al grupo {group_name}.")
    except Exception as e:
        logging.error(f"Error al enviar mensaje por WhatsApp: {e}")

# Función para formatear y enviar el mensaje del stock
def monitor_stock():
    # Consulta SQL para el stock
    sql_query = """
    SELECT 
        COD_ART, DESCRIPCION, StockActual, STK_MIN
    FROM 
        dbo.STOCK_TABLE
    WHERE 
        StockActual < STK_MIN
    """
    try:
        stock_data = pd.read_sql(sql_query, engine)
        if not stock_data.empty:
            message = "*Alerta de Stock Bajo:*\n" + "\n".join(
                f"{row['DESCRIPCION']}: Actual: {row['StockActual']}, Mínimo: {row['STK_MIN']}"
                for index, row in stock_data.iterrows()
            )
            send_whatsapp_message(driver, "Stock-Notificaciones", message)
    except Exception as e:
        logging.error("Error al monitorizar el stock: " + str(e))

# Función para formatear y enviar el mensaje de mantenimiento
def monitor_maintenance():
    # Consulta SQL para el mantenimiento
    maintenance_query = """
    SELECT 
        ID_OTM, FECHA_CARGA, ESTADO, DESCRIPCION_CORTA, ID_MAQUINA, ID_OPERARIO
    FROM 
        dbo.MANT_OTM
    WHERE 
        ESTADO = 'PENDIENTE'
    """
    try:
        maintenance_data = pd.read_sql(maintenance_query, engine)
        if not maintenance_data.empty:
            message = "*Alerta de Mantenimiento Pendiente:*\n" + "\n".join(
                f"OTM: {row['ID_OTM']}, Máquina: {row['ID_MAQUINA']}, Estado: {row['ESTADO']}"
                for index, row in maintenance_data.iterrows()
            )
            send_whatsapp_message(driver, "Mantenimiento-Notificaciones", message)
    except Exception as e:
        logging.error("Error al monitorizar el mantenimiento: " + str(e))

# Inicialización de WebDriver
driver = setup_whatsapp()

# Bucle infinito para la actualización de mensajes
while True:
    monitor_stock()
    monitor_maintenance()
    time.sleep(30)  # Espera 5 minutos antes de la próxima verificación
