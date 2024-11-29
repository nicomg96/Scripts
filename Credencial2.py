from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import json
from datetime import datetime, timedelta, timezone
import pandas as pd
import pytz  # Asegúrate de tener instalado pytz

options = webdriver.ChromeOptions()
profile_directory = "Default"
profile_path = "C:\\Users\\nmercado\\AppData\\Local\\Google\\Chrome\\User Data"
options.add_argument(f"user-data-dir={profile_path}")
options.add_argument(f"profile-directory={profile_directory}")

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

try:
    start_date = datetime(2023, 11, 1)
    end_date = datetime(2024, 9, 10)
    excel_file = 'solar_data4.xlsx'
    all_data = []
    tz = pytz.timezone('Etc/GMT+3')  # Estableciendo la zona horaria GMT-03:00

    current_date = start_date
    while current_date <= end_date:
        year = current_date.year
        month = current_date.month
        day = current_date.day
        url = f"https://www.solarweb.com/Chart/GetChartNew?pvSystemId=169f4d3e-e6f5-43db-979e-1d6aeebc32e1&year={year}&month={month}&day={day}&interval=day&view=consumption"
        driver.get(url)
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.TAG_NAME, "pre")))
        json_text = driver.find_element(By.TAG_NAME, "pre").text
        data = json.loads(json_text)

        # Verificamos si 'settings' y 'series' existen
        if 'settings' in data and 'series' in data['settings']:
            series_data = data['settings']['series']
            for item in series_data:
                if item.get('type') == 'areaspline' and item.get('name') == 'Potencia de la red':
                    consumo_data = item.get('data', [])
                    for timestamp, power in consumo_data:
                        utc_time = datetime.fromtimestamp(timestamp / 1000, tz=timezone.utc)
                        local_time = utc_time.astimezone(tz)  # Convertir de UTC a GMT-03:00
                        formatted_date = local_time.strftime('%Y-%m-%d')
                        hour = local_time.strftime('%H:%M')
                        power_kw = power / 1000
                        all_data.append([formatted_date, hour, power_kw])

        current_date += timedelta(days=1)

    if all_data:
        df = pd.DataFrame(all_data, columns=['Fecha', 'Hora', 'Potencia (KW)'])
        df.to_excel(excel_file, index=False)
        print(f"Datos guardados en {excel_file}")
    else:
        print("No data to save.")

finally:
    driver.quit()
