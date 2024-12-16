import pandas as pd
import os

# Rutas de los archivos
ruta_noviembre = r"C:\Users\eyleen.godoy_arbusta\Desktop\Consolidar datos\Noviembre"
ruta_comparativo = r"C:\Users\eyleen.godoy_arbusta\Desktop\Consolidar datos\comparativobusquedas"
ruta_principal = r"C:\Users\eyleen.godoy_arbusta\Desktop\Consolidar datos"

# Archivos a procesar
archivo_principal = os.path.join(ruta_principal, "Priorización Productos PedidosYa-Arbusta.xlsx")
archivos_a_revisar = [
    (os.path.join(ruta_noviembre, "entregablenoviembre.xlsx"), "master_code"),
    (os.path.join(ruta_comparativo, "Datos encontrados.xlsx"), "master_code"),
    (os.path.join(ruta_comparativo, "Datos_encontrados_backup_20241207_044921.xlsx"), "master_code")
]

def cargar_archivo(archivo, nombre_columna):
    """Función para cargar un archivo y obtener sus master_codes"""
    try:
        df = pd.read_excel(archivo)
        if nombre_columna not in df.columns:
            print(f"Advertencia: No se encontró la columna {nombre_columna} en {os.path.basename(archivo)}")
            print("Columnas disponibles:", df.columns.tolist())
            return set()
        return set(df[nombre_columna].astype(str))
    except Exception as e:
        print(f"Error al cargar {os.path.basename(archivo)}: {str(e)}")
        return set()

def comparar_archivos():
    print("Leyendo archivo principal...")
    try:
        # Leer específicamente la hoja "Resultado" y obtener los MASTERCODE
        df_principal = pd.read_excel(archivo_principal, sheet_name="Resultado")
        
        if 'MASTERCODE' not in df_principal.columns:
            print("No se encontró la columna MASTERCODE en el archivo principal.")
            print("Columnas disponibles:", df_principal.columns.tolist())
            return
        
        # Obtener los 42000 MASTERCODE
        codigos_principal = df_principal['MASTERCODE'].astype(str).tolist()
        print(f"Total de MASTERCODE a procesar: {len(codigos_principal)}")
        
    except Exception as e:
        print(f"Error al leer el archivo principal: {str(e)}")
        return
    
    # Diccionario para almacenar resultados
    resultados = {codigo: "NUEVO" for codigo in codigos_principal}
    codigos_pendientes = set(codigos_principal)
    
    # Procesar cada archivo en orden
    for archivo, nombre_columna in archivos_a_revisar:
        if not codigos_pendientes:  # Si ya no hay códigos pendientes, terminar
            break
            
        print(f"\nProcesando {os.path.basename(archivo)}...")
        codigos_archivo = cargar_archivo(archivo, nombre_columna)
        
        # Encontrar coincidencias
        encontrados = codigos_pendientes.intersection(codigos_archivo)
        
        # Actualizar resultados
        for codigo in encontrados:
            resultados[codigo] = os.path.basename(archivo)
        
        # Actualizar pendientes
        codigos_pendientes -= encontrados
        print(f"Se encontraron {len(encontrados)} códigos en {os.path.basename(archivo)}")
    
    # Crear DataFrame de resultados
    print("\nGenerando archivo de resultados...")
    df_resultado = pd.DataFrame({
        'MASTERCODE': list(resultados.keys()),
        'UBICACION': list(resultados.values())
    })
    
    # Guardar resultados
    archivo_salida = os.path.join(ruta_principal, "resultado_comparacion.xlsx")
    df_resultado.to_excel(archivo_salida, index=False)
    print(f"\nArchivo guardado en: {archivo_salida}")
    
    # Resumen final
    total_nuevos = sum(1 for v in resultados.values() if v == "NUEVO")
    print(f"\nResumen:")
    print(f"Total de códigos procesados: {len(codigos_principal)}")
    print(f"Códigos encontrados: {len(codigos_principal) - total_nuevos}")
    print(f"Códigos nuevos: {total_nuevos}")

if __name__ == "__main__":
    comparar_archivos()
