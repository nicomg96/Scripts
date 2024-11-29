import requests

def fetch_and_update_excel(url, excel_file):
    # Realizar la solicitud GET
    response = requests.get(url)
    
    # Imprimir el contenido de la respuesta para inspección
    print(f"Contenido de la respuesta: {response.text[:500]}")  # Muestra los primeros 500 caracteres

    try:
        # Intentar parsear la respuesta JSON
        data = response.json()
        print(data)  # Esto te mostrará el JSON completo en la consola
    except requests.exceptions.JSONDecodeError as e:
        print(f"Error al decodificar el JSON: {e}")
        print("Parece que la respuesta no es un JSON válido. Aquí está el texto crudo:")
        print(response.text)  # Imprimir el contenido crudo de la respuesta

# URL de la solicitud
url = 'https://www.solarweb.com/Chart/GetChartNew?pvSystemId=4391ad5e-69be-49a7-ba8c-a12c4468f5f7&year=2024&month=9&day=2&interval=day&view=production&_t=172527213377'

# Ruta al archivo de Excel (aquí es donde trabajarías con el archivo Excel)
excel_file = 'ruta_a_tu_archivo_excel.xlsx'

# Llamada a la función para realizar la solicitud y manejar la respuesta
fetch_and_update_excel(url, excel_file)
