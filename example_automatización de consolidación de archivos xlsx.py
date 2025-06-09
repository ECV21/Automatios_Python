import pandas as pd
import os

def concatenar_archivos_mensuales(nombre_carpeta_mes):
    """
    Concatena todos los archivos .xlsx diarios dentro de una carpeta mensual
    y guarda el resultado en un único archivo Excel con el nombre del mes.

    Args:
        nombre_carpeta_mes (str): El nombre de la carpeta del mes (ej. "Mes_1", "Mes_2").
    """

    # Definir el directorio base donde se encuentran las carpetas de los meses.
    # AJUSTA ESTA RUTA A LA UBICACIÓN REAL DE TUS CARPETAS (Mes_1, Mes_2, etc.)
    # Por ejemplo, si tus carpetas Mes_1, Mes_2, etc. están en
    # "C:\Users\tu_usuario\Documentos\MisDatos\Mes_1",
    # entonces tu_directorio_base sería "C:\Users\tu_usuario\Documentos\MisDatos"
    
    # Asumiendo que las carpetas Mes_X están en el mismo nivel que el script,
    # o puedes especificar una ruta absoluta aquí.
    # Por ejemplo:
    directorio_base = "C:/Users/eduar/OneDrive/Escritorio/Ejercicios Python/Files"
    # O, si las carpetas de los meses están justo donde ejecutas el script:
    #directorio_base = os.getcwd() # Obtiene el directorio de trabajo actual

    # Lista de nombres de meses para mapear Mes_X a Enero, Febrero, etc.
    nombres_meses = {
        'Mes_1': 'Enero', 'Mes_2': 'Febrero', 'Mes_3': 'Marzo',
        'Mes_4': 'Abril', 'Mes_5': 'Mayo', 'Mes_6': 'Junio',
        'Mes_7': 'Julio', 'Mes_8': 'Agosto', 'Mes_9': 'Septiembre',
        'Mes_10': 'Octubre', 'Mes_11': 'Noviembre', 'Mes_12': 'Diciembre'
    }

    # Construir la ruta completa a la carpeta del mes
    ruta_carpeta_mes = os.path.join(directorio_base, nombre_carpeta_mes)

    # Validar si la carpeta existe
    if not os.path.isdir(ruta_carpeta_mes):
        print(f"Error: La carpeta '{ruta_carpeta_mes}' no existe.")
        return

    # Lista para almacenar los DataFrames de cada archivo diario
    lista_df_diarios = []

    print(f"Procesando carpeta: '{nombre_carpeta_mes}'...")

    # Recorrer todos los archivos en la carpeta del mes
    for archivo in os.listdir(ruta_carpeta_mes):
        if archivo.endswith('.xlsx'): # Verificar si el archivo es un Excel
            ruta_completa_archivo = os.path.join(ruta_carpeta_mes, archivo)
            try:
                # Leer el archivo Excel y añadirlo a la lista
                df_diario = pd.read_excel(ruta_completa_archivo)
                lista_df_diarios.append(df_diario)
                print(f"  - Añadido: {archivo}")
            except Exception as e:
                print(f"  - Error al leer el archivo {archivo}: {e}")

    # Verificar si se encontraron archivos para concatenar
    if not lista_df_diarios:
        print(f"No se encontraron archivos .xlsx en la carpeta '{nombre_carpeta_mes}'.")
        return

    # Concatenar todos los DataFrames de la lista
    df_concatenado = pd.concat(lista_df_diarios, ignore_index=True)
    print(f"\nSe han concatenado {len(lista_df_diarios)} archivos.")

    # Crear la carpeta 'FILES' si no existe en el directorio base
    ruta_carpeta_files = os.path.join(directorio_base, 'FILES')
    os.makedirs(ruta_carpeta_files, exist_ok=True) # exist_ok=True evita errores si ya existe

    # Obtener el nombre del mes en español
    nombre_mes_final = nombres_meses.get(nombre_carpeta_mes, f'Mes_Desconocido_{nombre_carpeta_mes.split("_")[-1]}')

    # Definir la ruta y el nombre del archivo Excel de salida
    nombre_archivo_salida = f"{nombre_mes_final}.xlsx"
    ruta_archivo_salida = os.path.join(ruta_carpeta_files, nombre_archivo_salida)

    # Guardar el DataFrame concatenado en el nuevo archivo Excel
    try:
        df_concatenado.to_excel(ruta_archivo_salida, index=False)
        print(f"Archivo consolidado guardado exitosamente en: '{ruta_archivo_salida}'")
    except Exception as e:
        print(f"Error al guardar el archivo consolidado: {e}")
        
        
        
# Llama a la función para cada mes
concatenar_archivos_mensuales('Mes_1')        
        
        