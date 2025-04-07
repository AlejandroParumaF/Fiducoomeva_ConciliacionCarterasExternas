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
    pagina = doc_pdf[-1]

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
            periodo = 'ERROR'
            aporte = 'ERROR'
            retiro = 'ERROR'
            rend = 'ERROR'
            retefuente = 'ERROR'
            saldo_final = 'ERROR'

            for x, line in enumerate(resultados):
                if 'PERIODO' in line.upper() and periodo == 'ERROR':
                    periodo = resultados[x+1]
                        
                elif ('APORTES' in line.upper() or 'APORLES' in line.upper()) and aporte == 'ERROR':
                    aporte = resultados[x+1]                      

                elif ('RETIROS' in line.upper() or 'RELIROS' in line.upper()) and retiro == 'ERROR':
                    retiro = resultados[x+1]                      

                elif ('RETENCIÓN EN LA FUENTE' in line.upper() or 'RETENCION EN LA FVENTE' in line.upper() or 'RETENCION EN LA FUENTE' in line.upper() or 'RELENCION EN LA FUENTE' in line.upper() or 'RELENCIÓN EN LA FUENTE' in line.upper() or 'RETENCIÓN EN LA FVENTE' in line.upper()) and retefuente == 'ERROR':
                    retefuente = resultados[x+1]
                     
                elif ('RENDIMIENTOS NETOS' in line.upper() or 'RENDIMIENLOS NELOS' in line.upper() or 'RENDIMIENLOS NETOS' in line.upper() or 'RENDIMIENTOS NELOS' in line.upper()) and rend == 'ERROR':
                    rend = resultados[x+1]                   

                elif 'SALDO FINAL' in line.upper() and saldo_final == 'ERROR':
                    saldo_final = resultados[x+1]
      
            print(f"PERIODO: {periodo}")
            print(f"APORTE: {aporte}")  
            print(f"RETIRO: {retiro}")
            print(f"RETEFUENTE: {retefuente}")
            print(f"RENIMIENTOS: {rend}")
            print(f"SALDO FINAL: {saldo_final}")
                
        else:
            print(f"No se detectó texto en la página.")

    except Exception as e:
        print(f"Error al procesar la página: {e}")
        # Continuar con la siguiente página en caso de error en una

   
    # 5. Cerrar el documento PDF
    doc_pdf.close()
    print("--- Proceso de OCR completado ---")

# --- Ejemplo de cómo usar la función ---

# Reemplaza 'ruta/a/tu/documento_escaneado.pdf' con la ruta real de tu archivo
ruta_archivo_pdf = "C:/Users/alpf0069/Desktop/Path Proyectos/BOT CONCILIACION CARTERA EXTERNA/NO LEIDOS/71316 Avista Dic.pdf"

# Llama a la función para procesar el PDF en español usando CPU
leer_pdf_scaneado_linea_por_linea(ruta_archivo_pdf, idiomas=['es'], usar_gpu=False, dpi=2000)

