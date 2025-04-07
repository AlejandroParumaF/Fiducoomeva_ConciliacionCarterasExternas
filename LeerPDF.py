import easyocr
import fitz  # PyMuPDF
import os
import io
from PIL import Image # Necesario si EasyOCR no acepta bytes directamente (aunque suele hacerlo)
import numpy as np   # Necesario para convertir desde PIL Image

def leer_pdf_scaneado_linea_por_linea(ruta_pdf, idiomas=['es'], usar_gpu=False, dpi=300):
    """
    Lee un documento PDF escaneado página por página usando EasyOCR
    e imprime la información detectada línea por línea (por bloque detectado).

    Args:
        ruta_pdf (str): La ruta al archivo PDF escaneado.
        idiomas (list): Lista de códigos de idioma para EasyOCR (ej: ['es', 'en']).
                        Ver documentación de EasyOCR para códigos disponibles.
        usar_gpu (bool): Si es True, intentará usar la GPU (requiere configuración).
                         Por defecto es False (usa CPU).
        dpi (int): Resolución en puntos por pulgada para renderizar las páginas
                   del PDF a imágenes. Mayor DPI puede mejorar la precisión de OCR
                   pero consume más memoria y tiempo. 300 es un buen valor.
    """
    # 1. Verificar si el archivo PDF existe
    if not os.path.exists(ruta_pdf):
        print(f"Error: El archivo PDF '{ruta_pdf}' no fue encontrado.")
        return

    # 2. Inicializar el lector de EasyOCR
    print(f"Inicializando EasyOCR para idioma(s): {idiomas} (GPU: {usar_gpu})...")
    try:
        reader = easyocr.Reader(idiomas, gpu=usar_gpu)
        print("EasyOCR inicializado correctamente.")
    except Exception as e:
        print(f"Error al inicializar EasyOCR: {e}")
        print("Asegúrate de que PyTorch está instalado correctamente y los modelos de lenguaje descargados.")
        return

    # 3. Abrir el documento PDF con PyMuPDF
    print(f"Abriendo PDF: {ruta_pdf}...")
    try:
        doc_pdf = fitz.open(ruta_pdf)
    except Exception as e:
        print(f"Error al abrir el archivo PDF '{ruta_pdf}': {e}")
        return

    print(f"Procesando {doc_pdf.page_count} página(s) con DPI={dpi}...\n")

    # 4. Iterar sobre cada página del PDF
    for num_pagina, pagina in enumerate(doc_pdf):
        print(f"--- Procesando Página {num_pagina + 1} ---")

        try:
            # a. Renderizar la página como una imagen (pixmap) con el DPI especificado
            # Calculamos el factor de zoom necesario para alcanzar el DPI deseado
            zoom = dpi / 72  # El DPI base de PDF suele ser 72
            matriz = fitz.Matrix(zoom, zoom)
            pixmap = pagina.get_pixmap(matrix=matriz, alpha=False) # alpha=False para ignorar transparencia

            # b. Convertir el pixmap a un formato que EasyOCR pueda leer (bytes o NumPy array)

            # Opción 1: Intentar con bytes directamente (más eficiente si funciona)
            img_bytes = pixmap.tobytes("png") # Convertir a bytes en formato PNG

            # Opción 2: Convertir a NumPy array (más compatible si los bytes fallan)
            # img = Image.open(io.BytesIO(img_bytes))
            # img_np = np.array(img)

            # c. Usar EasyOCR para extraer texto de la imagen de la página
            # Pasar los bytes directamente o el array NumPy
            # detail=0 devuelve solo la lista de textos detectados
            # paragraph=False devuelve bloques de texto individuales
            resultados = reader.readtext(img_bytes, detail=0, paragraph=False)
            # Si usas NumPy array: resultados = reader.readtext(img_np, detail=0, paragraph=False)


            # d. Imprimir los resultados línea por línea (cada detección en una línea nueva)
            if resultados:
                print(f"Texto encontrado en la página {num_pagina + 1}:")
                for linea_detectada in resultados:
                    print(linea_detectada)
            else:
                print(f"No se detectó texto en la página {num_pagina + 1}.")

        except Exception as e:
            print(f"Error al procesar la página {num_pagina + 1}: {e}")
            # Continuar con la siguiente página en caso de error en una

        print("-" * (len(f"--- Procesando Página {num_pagina + 1} ---") + 1)) # Separador visual
        print() # Línea en blanco entre páginas

    # 5. Cerrar el documento PDF
    doc_pdf.close()
    print("--- Proceso de OCR completado ---")

# --- Ejemplo de cómo usar la función ---

# Reemplaza 'ruta/a/tu/documento_escaneado.pdf' con la ruta real de tu archivo
ruta_archivo_pdf = "C:/Users/alpf0069/Desktop/Path Proyectos/BOT CONCILIACION CARTERA EXTERNA/NO LEIDOS/71316 Avista Dic.pdf"

# Llama a la función para procesar el PDF en español usando CPU
leer_pdf_scaneado_linea_por_linea(ruta_archivo_pdf, idiomas=['es'], usar_gpu=False, dpi=2000)

