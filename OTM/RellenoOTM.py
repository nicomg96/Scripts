import sqlalchemy as sa
from sqlalchemy.exc import SQLAlchemyError
from datetime import datetime
import calendar

# Configuración de la conexión a la base de datos
dsn = 'AltieriGestion'
username = 'sa'
password = 'Axoft1988'
mant_otm_table = 'dbo.MANT_OTM'
mant_tareas_table = 'dbo.MANT_TAREAS'
prod_maquinas_table = 'dbo.PROD_MAQUINAS'

# Crear el motor de la base de datos
engine = sa.create_engine(f'mssql+pyodbc://{username}:{password}@{dsn}')

# Obtener lista de sectores disponibles
def obtener_sectores():
    query = f"SELECT DISTINCT SECTOR FROM {mant_otm_table}"
    with engine.connect() as conn:
        result = conn.execute(sa.text(query)).fetchall()
    return [row[0] for row in result]

# Obtener las OTM filtradas
def obtener_otms_filtradas(sector, anio, mes, semana):
    query = f"""
    SELECT ID_OTM, TIPO_OTM, FECHA_PROGRAMADA, SECTOR, ID_MAQUINA, ID_TAREA
    FROM {mant_otm_table}
    WHERE SECTOR = :sector AND YEAR(FECHA_PROGRAMADA) = :anio AND MONTH(FECHA_PROGRAMADA) = :mes
          AND DATEPART(WEEK, FECHA_PROGRAMADA) = :semana AND TIPO_OTM = 'PREDICTIVO'
    """
    with engine.connect() as conn:
        result = conn.execute(sa.text(query), {
            'sector': sector,
            'anio': anio,
            'mes': mes,
            'semana': semana
        }).fetchall()
    return result

# Obtener el nombre de la máquina
def obtener_nombre_maquina(id_maquina):
    query = f"""
    SELECT NOMBRE_MAQUINA
    FROM {prod_maquinas_table}
    WHERE ID_MAQUINA = :id_maquina
    """
    with engine.connect() as conn:
        result = conn.execute(sa.text(query), {'id_maquina': id_maquina}).fetchone()
    return result[0] if result else " "

# Obtener el nombre de la tarea
def obtener_nombre_tarea(id_tarea):
    query = f"""
    SELECT NOMBRE_TAREA
    FROM {mant_tareas_table}
    WHERE ID_TAREA = :id_tarea
    """
    with engine.connect() as conn:
        result = conn.execute(sa.text(query), {'id_tarea': id_tarea}).fetchone()
    return result[0] if result else " "

# Actualizar la OTM en la base de datos
def actualizar_otm(numero_otm, id_mecanico1, cant_horas_mec_1, id_mecanico2, cant_horas_mec_2, fecha_cierre, estado):
    update_query = f"UPDATE {mant_otm_table} SET "
    update_fields = []
    params = {'numero_otm': numero_otm}

    # Solo añadir campos si se ingresaron valores
    if id_mecanico1:
        update_fields.append("ID_MECANICO1 = :id_mecanico1")
        params['id_mecanico1'] = id_mecanico1
    if cant_horas_mec_1:
        update_fields.append("CANT_HORAS_MEC_1 = :cant_horas_mec_1")
        params['cant_horas_mec_1'] = cant_horas_mec_1
    if id_mecanico2:
        update_fields.append("ID_MECANICO2 = :id_mecanico2")
        params['id_mecanico2'] = id_mecanico2
    if cant_horas_mec_2:
        update_fields.append("CANT_HORAS_MEC_2 = :cant_horas_mec_2")
        params['cant_horas_mec_2'] = cant_horas_mec_2
    if fecha_cierre:
        update_fields.append("FECHA_CIERRE = :fecha_cierre")
        params['fecha_cierre'] = fecha_cierre
    if estado:
        update_fields.append("ESTADO = :estado")
        params['estado'] = estado

    # Si no hay campos para actualizar, terminar la función
    if not update_fields:
        print(f"No se realizaron cambios en la OTM {numero_otm}.")
        return

    # Construir la consulta de actualización final
    update_query += ", ".join(update_fields)
    update_query += " WHERE ID_OTM = :numero_otm"

    try:
        with engine.begin() as conn:
            conn.execute(sa.text(update_query), params)
        print(f"OTM {numero_otm} actualizada correctamente.")
    except SQLAlchemyError as e:
        print(f"Error al actualizar la OTM {numero_otm}: {e}")

def main():
    # Obtener sectores disponibles
    sectores = obtener_sectores()
    print("Sectores disponibles:")
    for idx, sector in enumerate(sectores, 1):
        print(f"{idx}. {sector}")

    # Pedir el sector, año, mes, y semana
    sector_idx = int(input("Seleccione el número del sector: ")) - 1
    sector = sectores[sector_idx]

    anio = int(input("Ingrese el año (YYYY): "))
    mes = int(input("Ingrese el mes (1-12): "))

    # Obtener la semana actual
    semana_actual = datetime.now().isocalendar()[1]
    semana = input(f"Ingrese la semana del año (1-52, por defecto es {semana_actual}): ")
    semana = int(semana) if semana else semana_actual

    # Obtener OTM filtradas
    otms = obtener_otms_filtradas(sector, anio, mes, semana)

    if not otms:
        print(f"No se encontraron OTM para el sector {sector}, año {anio}, mes {mes}, semana {semana}.")
        return

    # Mostrar las OTM filtradas y procesar una por una
    for datos_otm in otms:
        id_otm = datos_otm[0]
        id_maquina = datos_otm[4]
        id_tarea = datos_otm[5]

        # Obtener nombres de la máquina y la tarea
        nombre_maquina = obtener_nombre_maquina(id_maquina)
        nombre_tarea = obtener_nombre_tarea(id_tarea)

        print(f"\nOTM {id_otm}:")
        print(f"Máquina: {nombre_maquina}")
        print(f"Tarea: {nombre_tarea}")

        # Pedir los nuevos valores para los campos
        id_mecanico1 = input("Introduce el ID del primer mecánico (dejar vacío si no aplica): ")
        cant_horas_mec_1 = input("Introduce la cantidad de horas del primer mecánico (dejar vacío si no aplica): ")
        id_mecanico2 = input("Introduce el ID del segundo mecánico (dejar vacío si no aplica): ")
        cant_horas_mec_2 = input("Introduce la cantidad de horas del segundo mecánico (dejar vacío si no aplica): ")
        fecha_cierre = input("Introduce la fecha de cierre (YYYY-MM-DD, dejar vacío si no aplica): ")
        estado = input("Introduce el nuevo estado (dejar vacío si no aplica): ")

        # Convertir la fecha de cierre a formato datetime si se proporcionó
        if fecha_cierre:
            try:
                fecha_cierre = datetime.strptime(fecha_cierre, "%Y-%m-%d")
            except ValueError:
                print("Formato de fecha inválido. Debe ser YYYY-MM-DD.")
                continue

        # Actualizar la OTM en la base de datos
        actualizar_otm(id_otm, id_mecanico1, cant_horas_mec_1, id_mecanico2, cant_horas_mec_2, fecha_cierre, estado)

if __name__ == "__main__":
    main()
