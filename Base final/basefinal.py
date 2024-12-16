import pandas as pd
import os

# Rutas de los archivos
ruta_principal = r"C:\Users\eyleen.godoy_arbusta\Desktop\Consolidar datos"
ruta_base_final = os.path.join(ruta_principal, "Base final")

# Archivos a procesar
archivo_priorizacion = os.path.join(ruta_principal, "Priorización Productos PedidosYa-Arbusta.xlsx")
archivo_bbdd = os.path.join(ruta_principal, "BBDD Formato operativo - Con Barcode.xlsx")

def crear_base_final():
    print("Iniciando creación de base final...")
    
    try:
        # Cargar archivo de priorización
        print("Cargando archivo de priorización...")
        df_priorizacion = pd.read_excel(archivo_priorizacion, sheet_name="Resultado")
        
        if 'MASTERCODE' not in df_priorizacion.columns:
            print("No se encontró la columna MASTERCODE en el archivo de priorización")
            return
            
        # Obtener lista de MASTERCODE
        mastercode_list = set(df_priorizacion['MASTERCODE'].astype(str))
        print(f"Total de MASTERCODE a buscar: {len(mastercode_list)}")
        
        # Cargar BBDD
        print("\nCargando BBDD Formato operativo...")
        df_bbdd = pd.read_excel(archivo_bbdd)
        
        # Asegurar que la columna de master_code en BBDD sea string
        columna_master_bbdd = df_bbdd.columns[0]  # Asumimos que es la primera columna
        df_bbdd[columna_master_bbdd] = df_bbdd[columna_master_bbdd].astype(str)
        
        # Filtrar registros que coinciden con los MASTERCODE de priorización
        print("Filtrando registros...")
        df_final = df_bbdd[df_bbdd[columna_master_bbdd].isin(mastercode_list)]
        
        # Crear carpeta Base final si no existe
        if not os.path.exists(ruta_base_final):
            os.makedirs(ruta_base_final)
        
        # Guardar resultado
        archivo_salida = os.path.join(ruta_base_final, "basefinal.xlsx")
        df_final.to_excel(archivo_salida, index=False)
        
        # Mostrar resumen
        print("\nResumen:")
        print(f"Total MASTERCODE buscados: {len(mastercode_list)}")
        print(f"Total registros encontrados: {len(df_final)}")
        print(f"Archivo guardado en: {archivo_salida}")
        
        # Verificar si hay diferencias
        encontrados = set(df_final[columna_master_bbdd])
        no_encontrados = mastercode_list - encontrados
        
        if no_encontrados:
            print(f"\nHay {len(no_encontrados)} MASTERCODE que no se encontraron en la BBDD")
            
            # Guardar lista de no encontrados
            df_no_encontrados = pd.DataFrame(list(no_encontrados), columns=['MASTERCODE'])
            archivo_no_encontrados = os.path.join(ruta_base_final, "mastercode_no_encontrados.xlsx")
            df_no_encontrados.to_excel(archivo_no_encontrados, index=False)
            print(f"Lista de MASTERCODE no encontrados guardada en: {archivo_no_encontrados}")
        
    except Exception as e:
        print(f"Error en el proceso: {str(e)}")

if __name__ == "__main__":
    crear_base_final()
