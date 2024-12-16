import pandas as pd
import os

# Rutas de los archivos
ruta_diciembre = r"C:\Users\eyleen.godoy_arbusta\Desktop\Consolidar datos\Diciembre"
ruta_principal = r"C:\Users\eyleen.godoy_arbusta\Desktop\Consolidar datos"

# Archivos a procesar
archivo_priorizacion = os.path.join(ruta_principal, "Priorización Productos PedidosYa-Arbusta.xlsx")
archivos_resultado = [
    "Resultado_NoEncontradosDiciembre.xlsx",
    "Resultado_DatosEncontradosDiciembreUnificados.xlsx"
]

def cargar_archivos_resultado():
    """Carga los archivos resultado y crea un diccionario de master_codes"""
    master_codes_dict = {}
    
    for archivo in archivos_resultado:
        try:
            ruta_completa = os.path.join(ruta_diciembre, archivo)
            df = pd.read_excel(ruta_completa)
            columna_master = df.columns[0]  # Primera columna (master_code)
            
            # Guardar cada master_code y su archivo de origen
            for codigo in df[columna_master].astype(str):
                master_codes_dict[codigo] = archivo
                
            print(f"Procesados {len(df)} registros de {archivo}")
            
        except Exception as e:
            print(f"Error al procesar {archivo}: {str(e)}")
    
    return master_codes_dict

def comparar_con_priorizacion():
    print("Iniciando comparación...")
    
    # Cargar master_codes de los archivos resultado
    master_codes_dict = cargar_archivos_resultado()
    
    try:
        # Cargar archivo de priorización
        df_priorizacion = pd.read_excel(archivo_priorizacion, sheet_name="Resultado")
        
        if 'MASTERCODE' not in df_priorizacion.columns:
            print("No se encontró la columna MASTERCODE en el archivo de priorización")
            return
        
        # Crear DataFrame de resultados
        resultados = []
        
        # Comparar cada MASTERCODE
        for codigo in df_priorizacion['MASTERCODE'].astype(str):
            ubicacion = master_codes_dict.get(codigo, "No encontrado")
            resultados.append({
                'MASTERCODE': codigo,
                'UBICACION': ubicacion
            })
        
        # Crear DataFrame con los resultados
        df_resultado = pd.DataFrame(resultados)
        
        # Guardar resultado
        archivo_salida = os.path.join(ruta_diciembre, "Resultado_Priorizacion.xlsx")
        df_resultado.to_excel(archivo_salida, index=False)
        
        # Mostrar resumen
        encontrados = len([r for r in resultados if r['UBICACION'] != "No encontrado"])
        total = len(resultados)
        
        print("\nResumen:")
        print(f"Total MASTERCODE procesados: {total}")
        print(f"Encontrados en archivos resultado: {encontrados}")
        print(f"No encontrados: {total - encontrados}")
        print(f"\nResultado guardado en: Resultado_Priorizacion.xlsx")
        
    except Exception as e:
        print(f"Error al procesar archivo de priorización: {str(e)}")

if __name__ == "__main__":
    comparar_con_priorizacion()
