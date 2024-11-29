import pandas as pd
from datetime import datetime

start_date = datetime(2018, 1, 1)
end_date = datetime.now()

date_range = pd.date_range(start_date, end_date)

calendar_df = pd.DataFrame(date_range, columns=['Fecha'])

calendar_df['Semana'] = calendar_df['Fecha'].dt.isocalendar().week
calendar_df['Mes'] = calendar_df['Fecha'].dt.strftime('%B')
calendar_df['Año'] = calendar_df['Fecha'].dt.year

file_path = "C:/Users/nmercado/Documents/calendario.xlsx"
calendar_df.to_excel(file_path, index=False)

