import pandas as pd

def eliminar_filas_columnas_vacias(archivo_entrada, archivo_salida):
    try:
        # Cargar el archivo Excel
        df = pd.read_excel(archivo_entrada, sheet_name=None)  # `header=None` para procesar encabezados personalizados
    except Exception as e:
        print(f"Error al cargar el archivo Excel: {e}")
        return

    hojas_procesadas = {}
    for hoja, datos in df.items():
        # Asignar la primera fila con valores como encabezados si existen
        datos.columns = datos.iloc[2]  # La fila 4 (índice 3) se considera encabezado
        #datos = datos[4:]  # Eliminar las filas usadas como encabezado
        
        # Eliminar siempre las 3 primeras filas y la quinta fila
        datos = datos.drop(index=[0, 1, 2, 3], errors='ignore')  # `errors='ignore'` evita errores si no existen estas filas

        # Eliminar filas y columnas completamente vacías
        datos = datos.dropna(how='all', axis=0)  # Elimina filas vacías
        datos = datos.dropna(how='all', axis=1)  # Elimina columnas vacías

        hojas_procesadas[hoja] = datos

    # Guardar el resultado en un nuevo archivo
    try:
        with pd.ExcelWriter(archivo_salida, engine='openpyxl') as writer:
            for hoja, datos in hojas_procesadas.items():
                datos.to_excel(writer, sheet_name=hoja, index=False)
        print(f"Archivo procesado guardado en: {archivo_salida}")
    except Exception as e:
        print(f"Error al guardar el archivo Excel: {e}")

# Ruta de tu archivo Excel de entrada y salida
archivo_entrada = "MARA.xlsx"  # Cambia por tu archivo de entrada
archivo_salida = "data_salida18.xlsx"    # Cambia por el archivo de salida deseado

eliminar_filas_columnas_vacias(archivo_entrada, archivo_salida)