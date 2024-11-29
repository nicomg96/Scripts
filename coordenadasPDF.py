import fitz  # PyMuPDF

pdf_path = 'Base OTM.pdf'

def explorar_coordenadas(pdf_path):
    doc = fitz.open(pdf_path)
    page = doc[0]  # Cambia el índice de página si necesitas otra página

    print("Haz clic en la página para obtener coordenadas de posiciones:")

    def imprimir_posicion(event):
        x, y = event.x, event.y
        print(f"Posición: x={x}, y={y}")

    page.insert_text((100, 750), "Prueba", fontsize=10, color=(0, 0, 0))
    doc.save('Prueba_Posiciones.pdf')
    doc.close()

explorar_coordenadas(pdf_path)