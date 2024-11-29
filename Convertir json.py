import json
import pandas as pd
from datetime import datetime

# Función para convertir de Watts a Kilowatts
def watts_to_kw(watts):
    return watts / 1000

# Cargar datos del archivo JSON
def load_data(json_file):
    with open(json_file, 'r') as file:
        data = json.load(file)
    return data

# Procesar datos y convertirlos en un DataFrame
def process_data(data):
    # Extraer datos de consumo
    consumo_data = data['settings']['series'][0]['data']
    
    # Crear listas para timestamps, valores en Watts y en Kilowatts
    timestamps = []
    watts_values = []
    kw_values = []

    for entry in consumo_data:
        timestamp = datetime.fromtimestamp(entry[0] / 1000)  # Convertir a datetime
        watts = entry[1]
        kw = watts_to_kw(watts)
        
        timestamps.append(timestamp)
        watts_values.append(watts)
        kw_values.append(kw)

    # Crear un DataFrame de Pandas
    df = pd.DataFrame({
        'Timestamp': timestamps,
        'Watts': watts_values,
        'Kilowatts': kw_values
    })
    
    return df

# Guardar DataFrame en un archivo CSV
def save_to_csv(df, output_file):
    df.to_csv(output_file, index=False)
    print(f"Datos guardados en {output_file}")

# Función principal
def main():
    json_file = '1prueba.json'  # Reemplaza con la ruta correcta a tu archivo JSON
    output_file = 'output.csv'  # Nombre del archivo CSV de salida
    
    # Cargar, procesar y guardar datos
    data = load_data(json_file)
    df = process_data(data)
    save_to_csv(df, output_file)

if __name__ == '__main__':
    main()
