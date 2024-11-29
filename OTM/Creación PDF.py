import sqlalchemy as sa
import fitz  # PyMuPDF
import pandas as pd


# Configuración de la conexión a la base de datos
dsn = 'AltieriGestion'
username = 'sa'
password = 'Axoft1988'
mant_otm_table = 'dbo.MANT_OTM'
mant_tareas_table = 'dbo.MANT_TAREAS'
prod_maquinas_table = 'dbo.PROD_MAQUINAS'
mant_mecanicos_table = 'dbo.MANT_MECANICOS'
sistema_usuarios_table = 'dbo.SISTEMA_USUARIOS'

engine = sa.create_engine(f'mssql+pyodbc://{username}:{password}@{dsn}')

def obtener_nombre_usuario(usu_sis):
    query = f"""
    SELECT USU_NOM_UNO, USU_APE 
    FROM {sistema_usuarios_table}
    WHERE USU_SIS = :usu_sis
    """
    with engine.connect() as conn:
        result = conn.execute(sa.text(query), {'usu_sis': usu_sis}).fetchone()
    if result:
        return f"{result[0]} {result[1]}"
    return " "


def obtener_tareas_excel(id_tarea):
    excel_path = r"C:\Users\nmercado\Documents\Scripts\OTM\Tareas.xlsx"
    sheet_name = 'Base_TODAS'
    
    # Leer el archivo y filtrar las filas que coinciden con el ID_TAREA
    df_tareas = pd.read_excel(excel_path, sheet_name=sheet_name)
    tareas_filtradas = df_tareas[df_tareas['Id_Tarea'] == id_tarea]
    
    # Obtener las columnas necesarias
    tareas = []
    for _, row in tareas_filtradas.iterrows():
        tarea = {
            "Id": row["Id_ST"],
            "Descripción": row["Descripción_ST"],
            "Ptos": row["Puntos"],
            "INS": row["INS"],
            "Descripción2": row["Descripción_INS"]
        }
        tareas.append(tarea)
    
    print(f"Tareas encontradas para ID_TAREA {id_tarea}: {tareas}")  # Depuración
    return tareas

# Diccionario con los nombres de operarios
operarios = {
    1: "CARPIO CRISTIAN",
    2: "CARPIO GUSTAVO",
    4: "ALFIERI FRANCO",
    5: "GONZALEZ JOSE",
    6: "LUNA DANIEL",
    7: "CORNEJO ALEJANDRO",
    9: "SUAREZ MARCELO",
    12: "MARCOS GUSTAVO",
    13: "MENESES ADRIAN",
    14: "ZARATE FERNANDO",
    15: "PANELLA MARIANO NICOLAS",
    16: "GUARÑOLO GUSTAVO",
    18: "REINOSO JAVIER",
    19: "CALDERÓN FACUNDO",
    20: "ZANON GUSTAVO",
    23: "GUTIERREZ JORGE LUIS",
    24: "ALVAREZ MARTIN EMANUEL",
    25: "ALFIERI LEANDRO CRISTIAN",
    27: "HERRERO CRISTIAN",
    29: "CRESPIN ALBERTO",
    30: "TORO EDUARDO CARLOS",
    32: "ORTIZ JAVIER",
    33: "PACHECO GUILLERMO",
    34: "BUTTINI PABLO",
    35: "FRANCO PABLO",
    36: "VIDELA ROBERTO",
    40: "PELETAY MARCELO",
    42: "ROIZ MARCELO",
    43: "ROIZ JORGE",
    44: "AGOSTINI SERGIO",
    46: "GONZALEZ OMAR",
    51: "AMPUERO DANIEL",
    54: "CECCHIN ALEJANDRO",
    55: "MENESES MAXIMILIANO",
    59: "MONTENEGRO GERMAN",
    62: "GUTIERREZ ALEXANDER",
    63: "MASON MARIO",
    64: "LUCERO GABRIEL",
    65: "DIAZ PABLO",
    67: "MERCADO NICOLAS",
    70: "REINOSO GONZALO",
    71: "VALERA JOSE",
    73: "GASTILLO GUSTAVO",
    74: "BAEZ NICOLAS",
    75: "MURUA EXEQUIEL",
    76: "DOLCEMASCOLO JOEL",
    78: "GONZALEZ ABEL",
    79: "GIMENEZ EBER",
    80: "GIMENEZ CRISTIAN",
    81: "GUTIERREZ MAURO",
    82: "VALENZUELA NAHUEL",
    83: "MOLINA MATIAS",
    84: "HIDALGO ALEJANDRO",
    85: "JUAREZ EXEQUIEL"
}

# Diccionario de estados
estados = {
    1: "PENDIENTE",
    2: "ANULADO",
    3: "TERMINADA",
    4: "CUMPLIDA_PARCIAL",
    5: "CURSADA"
}

# Diccionario de tipos de falla
tipos_falla = {
    1: "MECÁNICA",
    2: "ELÉCTRICA / ELECTROMECÁNICA",
    3: "NEUMÁTICA",
    4: "HIDRÁULICA"
}

def obtener_datos_otm(numero_otm):
    query = f"""
    SELECT ID_OTM, TIPO_OTM, FECHA_PROGRAMADA, SECTOR, USU_SIS, ID_MAQUINA, ID_TAREA, 
           ID_OPERARIO, DESCRIPCION_CORTA, ESTADO, COMENTARIOS_FINALES, ID_TIPO_FALLA, 
           CANT_HORAS_MEC_1, FECHA_CIERRE, ID_MECANICO1, ID_MECANICO2, CANT_HORAS_MEC_2
    FROM {mant_otm_table} 
    WHERE ID_OTM = :numero_otm
    """
    with engine.connect() as conn:
        result = conn.execute(sa.text(query), {'numero_otm': numero_otm}).fetchone()
    return result

def obtener_nombre_tarea(id_tarea):
    query = f"""
    SELECT NOMBRE_TAREA 
    FROM {mant_tareas_table} 
    WHERE ID_TAREA = :id_tarea
    """
    with engine.connect() as conn:
        result = conn.execute(sa.text(query), {'id_tarea': id_tarea}).fetchone()
    return result[0] if result else " "

def obtener_nombre_maquina(id_maquina):
    query = f"""
    SELECT NOMBRE_MAQUINA 
    FROM {prod_maquinas_table} 
    WHERE ID_MAQUINA = :id_maquina
    """
    with engine.connect() as conn:
        result = conn.execute(sa.text(query), {'id_maquina': id_maquina}).fetchone()
    return result[0] if result else " "

def obtener_nombre_mecanico(id_mecanico):
    if id_mecanico is None:
        return " "
    query = f"""
    SELECT NOMBRE, APELLIDO 
    FROM {mant_mecanicos_table} 
    WHERE ID_MECANICO = :id_mecanico
    """
    with engine.connect() as conn:
        result = conn.execute(sa.text(query), {'id_mecanico': id_mecanico}).fetchone()
    if result:
        nombre_completo = f"{result[0]} {" "} {result[1]}"  # Formato: "APELLIDO NOMBRE"
        return nombre_completo
    return " "

def rellenar_pdf(datos_otm, pdf_template, pdf_salida):
    # Abre el PDF plantilla y define la página
    doc = fitz.open(pdf_template)
    page = doc[0]  # Supongamos que estamos trabajando en la primera página

    # Obtener el nombre del operario usando su ID
    id_operario = datos_otm[7]
    nombre_operario = operarios.get(id_operario, " ")

    # Obtener el nombre de la tarea usando el ID_TAREA
    id_tarea = datos_otm[6]
    nombre_tarea = obtener_nombre_tarea(id_tarea)

    # Obtener el nombre de la máquina usando el ID_MAQUINA
    id_maquina = datos_otm[5]
    nombre_maquina = obtener_nombre_maquina(id_maquina)

    # Obtener el nombre del estado usando el ID_ESTADO
    id_estado = datos_otm[9]
    nombre_estado = estados.get(id_estado, " ")

    # Obtener el nombre completo de los mecánicos usando ID_MECANICO1 y ID_MECANICO2
    id_mecanico1 = datos_otm[14]
    nombre_mecanico1 = obtener_nombre_mecanico(id_mecanico1)

    id_mecanico2 = datos_otm[15]  # Puede estar vacío
    nombre_mecanico2 = obtener_nombre_mecanico(id_mecanico2)

    # Obtener el tipo de falla usando el ID_TIPO_FALLA
    id_tipo_falla = datos_otm[11]
    nombre_tipo_falla = tipos_falla.get(id_tipo_falla, " ")

        # Obtener el nombre completo del usuario usando USU_SIS
    usu_sis = datos_otm[4]
    nombre_usuario = obtener_nombre_usuario(usu_sis)

      # Obtener las filas de tareas relacionadas con el ID_TAREA del archivo de Excel
    data_rows = obtener_tareas_excel(id_tarea)
 # Definir encabezado de la tabla
    encabezado = {"Id": "Id", "Descripción": "Descripción", "Ptos": "Ptos", "INS": "INS", "Descripción2": "Descripción"}

    # Posiciones de inicio de las columnas
    start_x_id = 30  # Coordenada X para la columna "Id"
    start_x_desc = 200  # Coordenada X para la columna "Descripción"
    start_x_ptos = 450  # Coordenada X para la columna "Ptos"
    start_x_ins = 547  # Coordenada X para la columna "INS"
    start_x_desc2 = 650  # Coordenada X para la segunda "Descripción"

    # Coordenada Y inicial para el encabezado
    encabezado_y = 140  # Ajusta esta posición a la altura inicial del encabezado

    # Imprimir el encabezado
    page.insert_text((start_x_id, encabezado_y), encabezado["Id"], fontsize=9, color=(0, 0, 0))
    page.insert_text((start_x_desc, encabezado_y), encabezado["Descripción"], fontsize=9, color=(0, 0, 0))
    page.insert_text((start_x_ptos, encabezado_y), encabezado["Ptos"], fontsize=9, color=(0, 0, 0))
    page.insert_text((start_x_ins, encabezado_y), encabezado["INS"], fontsize=9, color=(0, 0, 0))
    page.insert_text((start_x_desc2, encabezado_y), encabezado["Descripción2"], fontsize=9, color=(0, 0, 0))

    # Coordenada Y inicial para las filas de datos, debajo del encabezado
    start_y = encabezado_y + 20  # Desplazar las filas de datos por debajo del encabezado
    line_height = 12  # Altura entre filas

    # Iterar sobre las filas de datos y añadirlas al PDF
    for i, row in enumerate(data_rows):
        y = start_y + i * line_height
        page.insert_text((start_x_id-5, y), str(row["Id"]), fontsize=8, color=(0, 0, 0))
        page.insert_text((start_x_desc-150, y), str(row["Descripción"]), fontsize=8, color=(0, 0, 0))
        page.insert_text((start_x_ptos+3, y), str(row["Ptos"]), fontsize=8, color=(0, 0, 0))
        page.insert_text((start_x_ins-22, y), str(row["INS"]), fontsize=8, color=(0, 0, 0))
        descripcion2_truncada = (row["Descripción2"][:40] + '...') if len(row["Descripción2"]) > 30 else row["Descripción2"]
        page.insert_text((start_x_desc2-60, y), descripcion2_truncada, fontsize=9, color=(0, 0, 0))

          # Dibujar un cuadro de tick al lado del valor de "Ptos"
        tick_x = start_x_ptos + 15  # Ajusta la posición del cuadro según sea necesario
        tick_y = y - 10  # Ajusta la posición vertical si es necesario
        tick_size = 12  # Tamaño del cuadro de selección (más pequeño para ajustarse mejor)

    # Dibuja un cuadrado para tick
        page.draw_rect(fitz.Rect(tick_x, tick_y, tick_x + tick_size + 10, tick_y + tick_size), color=(0, 0, 0), width=0.5)

    # Agregar el símbolo de tick dentro del cuadro
        page.insert_text((tick_x + 7, tick_y +9), "X", fontsize=10, color=(0, 0, 0))


    # Definir posiciones de los campos (x, y) y valores para otros datos en el PDF
    campos = {
        "ID_OTM": (363, 50.5, datos_otm[0]),
        "TIPO_OTM": (425, 50.5, datos_otm[1]),
        "FECHA_PROGRAMADA": (530, 50.5, datos_otm[2].date() if datos_otm[2] else None),
        "SECTOR": (630, 50.5, datos_otm[3]),
        "USU_SIS": (145, 76, nombre_usuario),  # Reemplazado con nombre completo del usuario
        "ID_MAQUINA": (502, 76, id_maquina),
        "NOMBRE_MAQUINA": (547, 76, nombre_maquina),
        "ID_TAREA": (155, 96, id_tarea),
        "NOMBRE_TAREA": (190, 96, nombre_tarea),
        "NOMBRE_OPERARIO": (560, 96, nombre_operario),
        "DESCRIPCION_CORTA": (145, 117, datos_otm[8]),
        "ID_ESTADO": (590, 117, id_estado),
        "NOMBRE_ESTADO": (615, 117, nombre_estado),
        "COMENTARIOS_FINALES": (120, 365, datos_otm[10]),
        "ID_TIPO_FALLA": (100, 477, id_tipo_falla),
        "NOMBRE_TIPO_FALLA": (110, 477, nombre_tipo_falla),
        "NOMBRE_MECANICO1": (20, 535, nombre_mecanico1),
        "NOMBRE_MECANICO2": (20, 558, nombre_mecanico2),
        "CANT_HORAS_MEC_1": (360, 535, datos_otm[12]),
        "CANT_HORAS_MEC_2": (360, 558, datos_otm[15]),
        "FECHA_CIERRE": (695, 506, datos_otm[13].date() if datos_otm[13] else None),
    }

    # Escribe los datos en el PDF
    for _, (x, y, valor) in campos.items():
        if valor is not None:  # Verifica que el valor no sea None
            page.insert_text((x, y), str(valor), fontsize=9, color=(0, 0, 0))

    # Guarda el PDF completado
    doc.save(pdf_salida)
    doc.close()


def main():
    numero_otm = input("Introduce el N° de OTM: ")
    datos_otm = obtener_datos_otm(numero_otm)
    if datos_otm:
        nombre_pdf_salida = f"{numero_otm}.pdf"  # Generar el nombre del archivo usando el número de OTM
        rellenar_pdf(
            datos_otm, 
            r"C:\Users\nmercado\Documents\Scripts\OTM\Base OTM.pdf",
            nombre_pdf_salida
        )
        print(f"PDF completado guardado como '{nombre_pdf_salida}'")
    else:
        print("No se encontraron datos para el N° de OTM proporcionado.")

if __name__ == "__main__":
    main()