import requests
import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import LineChart, Reference
from openpyxl.utils.dataframe import dataframe_to_rows  # Importar la función necesaria

def obtener_precio_dolar_blue():
    url = 'https://api.bluelytics.com.ar/v2/evolution.json'  # URL actualizada
    try:
        response = requests.get(url)
        response.raise_for_status()  # Esto lanzará un error si la respuesta no es 200

        data = response.json()  # Intenta convertir la respuesta a JSON
        print("Datos recibidos de la API:", data)  # Imprime los datos para ver qué se está recibiendo
        
        # Comprobar que hay datos
        if len(data) > 0:
            # Crear una lista para almacenar los registros
            registros = []
            
            for registro in data:
                fecha = registro['date']  # Fecha
                precio_dolar_blue = registro['value_sell']  # Precio de venta
                registros.append({'Fecha': fecha, 'Precio Dólar Blue': precio_dolar_blue})

            # Crear un DataFrame a partir de los registros
            df = pd.DataFrame(registros)

            # Guardar el DataFrame en un archivo Excel
            archivo_excel = 'historial_precio_dolar_blue.xlsx'
            df.to_excel(archivo_excel, index=False)  # Guardar en Excel
            print(f"Historial guardado en '{archivo_excel}'")

            # Crear un libro de trabajo de Excel y seleccionar la hoja activa
            wb = Workbook()
            ws = wb.active

            # Escribir los datos en la hoja de Excel
            for r in dataframe_to_rows(df, index=False, header=True):
                ws.append(r)

            # Crear un gráfico de líneas
            chart = LineChart()
            chart.title = "Historial del Precio del Dólar Blue"
            chart.x_axis.title = "Fecha"
            chart.y_axis.title = "Precio (en $)"

            # Referencia de datos para el gráfico
            data_ref = Reference(ws, min_col=2, min_row=1, max_col=2, max_row=len(df) + 1)
            labels_ref = Reference(ws, min_col=1, min_row=2, max_row=len(df) + 1)  # Fechas para el eje x

            chart.add_data(data_ref, titles_from_data=True)
            chart.set_categories(labels_ref)

            # Añadir el gráfico a la hoja
            ws.add_chart(chart, "E5")  # Ajusta la posición del gráfico como desees

            # Guardar el libro de trabajo con el gráfico
            wb.save(archivo_excel)
            print("Gráfico añadido y archivo guardado.")
        else:
            print("No se encontraron datos en la respuesta.")
    except requests.exceptions.RequestException as e:
        print(f"Ocurrió un error en la conexión: {e}")
    except ValueError:
        print("Error al convertir la respuesta a JSON.")
    except KeyError as e:
        print(f"Error al acceder a la clave: {e}")

obtener_precio_dolar_blue()
