from pathlib import Path
import os # Necesario si quieres imprimir solo el nombre del archivo

ruta_carpeta_str = r"C:\Users\alpf0069\Documents\Proyectos\Area Contabilidad\Proyectos Python\Conciliación Carteras Externas\BOT CONCILIACION CARTERA EXTERNA\Lista Extractos"
lista_codigos = [119656, 102699, 93235, 102022, 71316, 98080, 116047, 106295, 106919] # Lista de códigos a buscar (como strings)
lista_codigos = [str(codigo) for codigo in lista_codigos] # Convertir a strings para la comparación
archivos_seleccionados = []
carpeta = Path(ruta_carpeta_str)

if not carpeta.is_dir():
    print(f"Error: La ruta '{ruta_carpeta_str}' no existe o no es un directorio.")
else:
    print(f"Buscando archivos PDF en: {carpeta}")
    print(f"Códigos a buscar en el nombre: {lista_codigos}")

    try:
        # Iterar sobre todos los elementos en la carpeta
        for item in carpeta.iterdir():
            # Verificar si es un archivo y si la extensión es .pdf (ignorando mayúsculas/minúsculas)
            if item.is_file() and item.suffix.lower() == ".pdf":
                # Obtener el nombre del archivo sin la extensión
                nombre_sin_extension = item.stem

                # Verificar si alguno de los códigos está contenido en el nombre del archivo
                # Usamos any() para detener la búsqueda tan pronto como se encuentre una coincidencia
                if any(codigo in nombre_sin_extension for codigo in lista_codigos):
                    archivos_seleccionados.append(item) # Añadir el objeto Path completo

    except PermissionError:
        print(f"Error: No tienes permisos para leer la carpeta '{ruta_carpeta_str}'.")
    except Exception as e:
        print(f"Ocurrió un error inesperado: {e}")

# --- Mostrar Resultados ---
if archivos_seleccionados:
    print("\n--- Archivos PDF encontrados que coinciden: ---")
    for archivo_path in archivos_seleccionados:
        # archivo_path es un objeto Path. Puedes obtener la ruta como string o solo el nombre:
        print(f"Ruta completa: {archivo_path}")
        # print(f"Solo nombre: {archivo_path.name}")
        # print(f"Ruta como string: {str(archivo_path)}")
else:
    if carpeta.is_dir(): # Solo mostrar si no hubo error de directorio
      print("\nNo se encontraron archivos PDF que cumplan los criterios.")