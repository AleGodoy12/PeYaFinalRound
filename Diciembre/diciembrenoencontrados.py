import pandas as pd
import os

# Ruta de la carpeta con los archivos
ruta_datos = r"C:\Users\eyleen.godoy_arbusta\Desktop\Consolidar datos\Diciembre\DatosNOencontradosDiciembre"
ruta_salida = r"C:\Users\eyleen.godoy_arbusta\Desktop\Consolidar datos\Diciembre"

def extraer_nombre_analista(nombre_archivo):
    """Extrae el nombre del analista del nombre del archivo"""
    try:
        if nombre_archivo.startswith("Reporte_QA_"):
            # Para archivos que empiezan con "Reporte_QA_"
            nombre = nombre_archivo.replace("Reporte_QA_", "").replace(".xlsx", "").strip()
        else:
            # Para archivos que empiezan con "QA busqueda"
            nombre = nombre_archivo.replace("QA busqueda ", "").replace(".xlsx", "").strip()
        return nombre
    except Exception as e:
        print(f"Error al extraer nombre del archivo {nombre_archivo}: {str(e)}")
        return nombre_archivo

def unificar_archivos():
    print("Iniciando unificación de archivos no encontrados...")
    
    # Lista para almacenar todos los DataFrames
    dfs = []
    
    # Obtener lista de archivos Excel
    archivos = [f for f in os.listdir(ruta_datos) if f.endswith('.xlsx')]
    
    if not archivos:
        print("No se encontraron archivos Excel en la carpeta.")
        return
    
    print(f"Se encontraron {len(archivos)} archivos para procesar.")
    
    # Procesar cada archivo
    for archivo in archivos:
        try:
            print(f"\nProcesando {archivo}...")
            ruta_completa = os.path.join(ruta_datos, archivo)
            
            # Leer el archivo Excel
            df = pd.read_excel(ruta_completa)
            
            # Extraer nombre del analista
            nombre_analista = extraer_nombre_analista(archivo)
            
            # Agregar columna del analista
            df['analista'] = nombre_analista
            
            # Agregar a la lista de DataFrames
            dfs.append(df)
            
            print(f"Registros procesados: {len(df)}")
            
        except Exception as e:
            print(f"Error al procesar {archivo}: {str(e)}")
            continue
    
    if not dfs:
        print("No se pudo procesar ningún archivo correctamente.")
        return
    
    # Unificar todos los DataFrames
    print("\nUnificando datos...")
    df_final = pd.concat(dfs, ignore_index=True)
    
    # Guardar resultado
    archivo_salida = os.path.join(ruta_salida, "NoEncontradosDiciembre.xlsx")
    df_final.to_excel(archivo_salida, index=False)
    
    print("\nResumen final:")
    print(f"Total de archivos procesados: {len(archivos)}")
    print(f"Total de registros en archivo final: {len(df_final)}")
    print(f"Archivo guardado en: {archivo_salida}")

if __name__ == "__main__":
    unificar_archivos()
