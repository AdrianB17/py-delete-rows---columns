import pandas as pd
import os

def eliminar_filas_columnas_vacias_multiples(archivos_entrada, carpeta_salida):
    """
    Procesa múltiples archivos Excel, elimina filas y columnas vacías, y aplica las reglas definidas.

    :param archivos_entrada: Lista de rutas de archivos Excel de entrada.
    :param carpeta_salida: Carpeta donde guardar los archivos procesados.
    """
    if not os.path.exists(carpeta_salida):
        os.makedirs(carpeta_salida)  # Crear la carpeta de salida si no existe

    for archivo_entrada in archivos_entrada:
        try:
            # Cargar el archivo Excel
            df = pd.read_excel(archivo_entrada, sheet_name=None)  # Cargar todas las hojas
        except Exception as e:
            print(f"Error al cargar el archivo {archivo_entrada}: {e}")
            continue

        hojas_procesadas = {}
        for hoja, datos in df.items():
            try:
                # Asignar encabezados desde la fila 4 si existen
                datos.columns = datos.iloc[2]  # La fila 4 (índice 3) como encabezado
                #datos = datos[3:]  # Eliminar las filas usadas como encabezado

                # Eliminar siempre las 3 primeras filas y la quinta fila
                datos = datos.drop(index=[0, 1, 2, 3], errors='ignore')

                # Eliminar las 3 últimas filas
                datos = datos.iloc[:-3, :]  # Seleccionar todo menos las últimas 3 filas

                # Eliminar filas y columnas completamente vacías
                datos = datos.dropna(how='all', axis=0)  # Filas vacías
                datos = datos.dropna(how='all', axis=1)  # Columnas vacías

                hojas_procesadas[hoja] = datos
            except Exception as e:
                print(f"Error al procesar la hoja {hoja} del archivo {archivo_entrada}: {e}")
                continue

        # Guardar el archivo procesado
        archivo_nombre = os.path.basename(archivo_entrada).replace(".xlsx", "_procesado.xlsx")
        archivo_salida = os.path.join(carpeta_salida, archivo_nombre)

        try:
            with pd.ExcelWriter(archivo_salida, engine='openpyxl') as writer:
                for hoja, datos in hojas_procesadas.items():
                    datos.to_excel(writer, sheet_name=hoja, index=False)
            print(f"Archivo procesado guardado en: {archivo_salida}")
        except Exception as e:
            print(f"Error al guardar el archivo {archivo_salida}: {e}")

# Lista de archivos Excel de entrada
archivos_entrada = [
    "./carpeta_entrada/data01.xlsx",
    "./carpeta_entrada/data02.xlsx",
]

# Carpeta de salida
carpeta_salida = "carpeta_salida"

# Procesar los archivos
eliminar_filas_columnas_vacias_multiples(archivos_entrada, carpeta_salida)
