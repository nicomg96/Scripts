import sqlalchemy as sa
from sqlalchemy.exc import SQLAlchemyError
from datetime import datetime
import calendar
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
from tkcalendar import DateEntry
from PIL import Image, ImageTk
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter

# Configuración de la conexión a la base de datos
dsn = 'AltieriGestion'
username = 'sa'
password = 'Axoft1988'
mant_otm_table = 'dbo.MANT_OTM'
mant_tareas_table = 'dbo.MANT_TAREAS'
prod_maquinas_table = 'dbo.PROD_MAQUINAS'
mant_mecanicos_table = 'dbo.MANT_MECANICOS'
mant_subtareas_table = 'dbo.MANT_SUBTAREAS'

# Crear el motor de la base de datos
engine = sa.create_engine(f'mssql+pyodbc://{username}:{password}@{dsn}')

# Obtener lista de sectores disponibles
def obtener_sectores():
    query = f"SELECT DISTINCT SECTOR FROM {mant_otm_table}"
    with engine.connect() as conn:
        result = conn.execute(sa.text(query)).fetchall()
    return [row[0] for row in result]

# Obtener lista de mecánicos disponibles
def obtener_mecanicos():
    query = f"SELECT ID_MECANICO, NOMBRE, APELLIDO FROM {mant_mecanicos_table} WHERE INHABILITADO = 'NO'"
    with engine.connect() as conn:
        result = conn.execute(sa.text(query)).fetchall()
    return {row[0]: f"{row[1]} {row[2]}" for row in result}

# Obtener las OTM filtradas
def obtener_otms_filtradas(sector, fecha_seleccionada, id_mecanico1):
    semana = fecha_seleccionada.isocalendar()[1]
    anio = fecha_seleccionada.year
    mes = fecha_seleccionada.month
    query = f"""
    SELECT ID_OTM, TIPO_OTM, FECHA_PROGRAMADA, SECTOR, ID_MAQUINA, ID_TAREA
    FROM {mant_otm_table}
    WHERE SECTOR = :sector AND YEAR(FECHA_PROGRAMADA) = :anio AND MONTH(FECHA_PROGRAMADA) = :mes
          AND DATEPART(WEEK, FECHA_PROGRAMADA) = :semana AND TIPO_OTM = 'PREDICTIVO'
    """
    params = {'sector': sector, 'anio': anio, 'mes': mes, 'semana': semana}
    if id_mecanico1:
        query += " AND ID_MECANICO1 = :id_mecanico1"
        params['id_mecanico1'] = id_mecanico1

    with engine.connect() as conn:
        result = conn.execute(sa.text(query), params).fetchall()
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

# Obtener subtareas de una tarea específica
def obtener_subtareas(id_tarea):
    query = f"""
    SELECT SUBTAREA_NOMBRE
    FROM {mant_subtareas_table}
    WHERE ID_TAREA = :id_tarea
    """
    with engine.connect() as conn:
        result = conn.execute(sa.text(query), {'id_tarea': id_tarea}).fetchall()
    return [row[0] for row in result]

# Actualizar la OTM en la base de datos
# Actualizar la OTM en la base de datos y generar PDF con firma
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
        messagebox.showinfo("Info", f"No se realizaron cambios en la OTM {numero_otm}.")
        return

    # Construir la consulta de actualización final
    update_query += ", ".join(update_fields)
    update_query += " WHERE ID_OTM = :numero_otm"

    try:
        with engine.begin() as conn:
            conn.execute(sa.text(update_query), params)

        # Generar el PDF con la firma digital
        generar_pdf_con_firma(numero_otm, id_mecanico1, fecha_cierre)

        messagebox.showinfo("Éxito", f"OTM {numero_otm} actualizada correctamente.")
    except SQLAlchemyError as e:
        messagebox.showerror("Error", f"Error al actualizar la OTM {numero_otm}: {e}")

def generar_pdf_con_firma(numero_otm, id_mecanico1, fecha_cierre):
    output_pdf = f"OTM_{numero_otm}.pdf"
    firma_imagen = "firma_digital.png"  # Asegúrate de tener este archivo en el mismo directorio

    try:
        # Crear el PDF
        c = canvas.Canvas(output_pdf, pagesize=letter)
        width, height = letter

        # Añadir información al PDF
        c.setFont("Helvetica", 12)
        c.drawString(100, height - 100, f"OTM: {numero_otm}")
        c.drawString(100, height - 120, f"Mecánico Responsable: {id_mecanico1}")
        c.drawString(100, height - 140, f"Fecha de Cierre: {fecha_cierre}")

        # Añadir la firma digital
        firma_x = 100
        firma_y = height - 200
        c.drawImage(firma_imagen, firma_x, firma_y, width=200, height=50)

        # Finalizar y guardar el PDF
        c.showPage()
        c.save()

        messagebox.showinfo("Éxito", f"PDF generado: {output_pdf}")
    except Exception as e:
        messagebox.showerror("Error", f"Error al generar el PDF: {e}")



# Ventana principal
def main():
    root = tk.Tk()
    root.title("Actualización de OTM")

    # Widgets para seleccionar los filtros
    ttk.Label(root, text="Sector:").grid(row=0, column=0, padx=10, pady=5)
    sector_combobox = ttk.Combobox(root, values=obtener_sectores())
    sector_combobox.grid(row=0, column=1, padx=10, pady=5)

    ttk.Label(root, text="Fecha Seleccionada:").grid(row=1, column=0, padx=10, pady=5)
    fecha_seleccionada_entry = DateEntry(root, date_pattern='yyyy-mm-dd')
    fecha_seleccionada_entry.grid(row=1, column=1, padx=10, pady=5)

    ttk.Label(root, text="ID Mecánico 1:").grid(row=2, column=0, padx=10, pady=5)
    mecanicos = obtener_mecanicos()
    id_mecanico1_combobox = ttk.Combobox(root, values=[f"{value}" for value in mecanicos.values()])
    id_mecanico1_combobox.grid(row=2, column=1, padx=10, pady=5)

    # Botón para buscar OTM
    def buscar_otm():
        sector = sector_combobox.get()
        fecha_seleccionada = fecha_seleccionada_entry.get_date()
        id_mecanico1 = [key for key, value in mecanicos.items() if value == id_mecanico1_combobox.get()]
        id_mecanico1 = id_mecanico1[0] if id_mecanico1 else None

        otms = obtener_otms_filtradas(sector, fecha_seleccionada, id_mecanico1)
        for item in tree.get_children():
            tree.delete(item)
        for otm in otms:
            id_otm, tipo_otm, fecha_programada, sector, id_maquina, id_tarea = otm
            nombre_maquina = obtener_nombre_maquina(id_maquina)
            nombre_tarea = obtener_nombre_tarea(id_tarea)
            tree.insert("", "end", values=(id_otm, fecha_programada, nombre_maquina, id_tarea))

    buscar_button = ttk.Button(root, text="Buscar OTM", command=buscar_otm)
    buscar_button.grid(row=3, column=1, pady=10)

    # Tabla para mostrar las OTM filtradas
    tree = ttk.Treeview(root, columns=("ID OTM", "Fecha Programada", "Máquina", "Tarea"), show='headings')
    tree.heading("ID OTM", text="ID OTM")
    tree.heading("Fecha Programada", text="Fecha Programada")
    tree.heading("Máquina", text="Máquina")
    tree.heading("Tarea", text="Tarea")
    tree.grid(row=4, column=0, columnspan=4, padx=10, pady=10)

    # Tabla para mostrar las subtareas de la tarea seleccionada
    ttk.Label(root, text="Subtareas:").grid(row=5, column=2, padx=10, pady=5)
    subtareas_listbox = tk.Listbox(root, width=80, height=6)
    subtareas_listbox.grid(row=6, column=2, columnspan=4, padx=10, pady=5)

    def mostrar_subtareas(event):
        selected_item = tree.selection()
        if selected_item:
            item = tree.item(selected_item)
            id_tarea = item['values'][3]  # Obtener ID_TAREA de la fila seleccionada
            subtareas = obtener_subtareas(id_tarea)
            subtareas_listbox.delete(0, tk.END)
            for subtarea in subtareas:
                subtareas_listbox.insert(tk.END, subtarea)

    tree.bind("<<TreeviewSelect>>", mostrar_subtareas)

    # Campos para actualizar la OTM
    ttk.Label(root, text="ID Mecánico 1:").grid(row=5, column=0, padx=10, pady=5)
    id_mecanico1_update_entry = ttk.Entry(root)
    id_mecanico1_update_entry.grid(row=5, column=1, padx=10, pady=5)

    ttk.Label(root, text="Horas Mecánico 1:").grid(row=6, column=0, padx=10, pady=5)
    horas_mec_1_entry = ttk.Entry(root)
    horas_mec_1_entry.grid(row=6, column=1, padx=10, pady=5)

    ttk.Label(root, text="ID Mecánico 2:").grid(row=7, column=0, padx=10, pady=5)
    id_mecanico2_entry = ttk.Entry(root)
    id_mecanico2_entry.grid(row=7, column=1, padx=10, pady=5)

    ttk.Label(root, text="Horas Mecánico 2:").grid(row=8, column=0, padx=10, pady=5)
    horas_mec_2_entry = ttk.Entry(root)
    horas_mec_2_entry.grid(row=8, column=1, padx=10, pady=5)

    ttk.Label(root, text="Fecha de Cierre (YYYY-MM-DD):").grid(row=9, column=0, padx=10, pady=5)
    fecha_cierre_entry = DateEntry(root, date_pattern='yyyy-mm-dd')
    fecha_cierre_entry.grid(row=9, column=1, padx=10, pady=5)

    ttk.Label(root, text="Estado:").grid(row=10, column=0, padx=10, pady=5)
    estado_entry = ttk.Entry(root)
    estado_entry.grid(row=10, column=1, padx=10, pady=5)

    # Botón para actualizar la OTM
    actualizar_button = ttk.Button(root, text="Actualizar OTM")
    actualizar_button.grid(row=11, column=1, pady=10)

    root.mainloop()

if __name__ == "__main__":
    main()
