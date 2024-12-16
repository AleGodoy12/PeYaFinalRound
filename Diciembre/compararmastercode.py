import pandas as pd
import os

# Rutas de los archivos
ruta_diciembre = r"C:\Users\eyleen.godoy_arbusta\Desktop\Consolidar datos\Diciembre"
ruta_comparativa = r"C:\Users\eyleen.godoy_arbusta\Desktop\Consolidar datos\comparativobusquedas"

# Archivos a procesar de diciembre
archivos_diciembre = [
    "NoEncontradosDiciembre.xlsx",
    "DatosEncontradosDiciembreUnificados.xlsx"
]

def cargar_archivo_diciembre(archivo):
    """Carga un archivo de diciembre y obtiene sus master_codes"""
    ruta_completa = os.path.join(ruta_diciembre, archivo)
    df = pd.read_excel(ruta_completa)
    
    # Asumimos que master_code está en la primera columna
    columna_master = df.columns[0]
    # Copiamos el DataFrame original y agregamos una columna para el resultado
    df['archivo_encontrado'] = 'No encontrado'
    return df, columna_master

def buscar_en_comparativa(df, columna_master):
    """Busca los master_codes en los archivos de la carpeta comparativa"""
    # Obtener lista de archivos Excel en la carpeta comparativa
    archivos_comparativa = [f for f in os.listdir(ruta_comparativa) if f.endswith('.xlsx')]
    
    # Para cada archivo en la carpeta comparativa
    for archivo in archivos_comparativa:
        print(f"Buscando en {archivo}...")
        ruta_completa = os.path.join(ruta_comparativa, archivo)
        
        try:
            df_comp = pd.read_excel(ruta_completa)
            # Asumimos que master_code está en la primera columna del archivo comparativo
            columna_comp = df_comp.columns[0]
            
            # Convertir a strings para comparación
            codigos_comp = set(df_comp[columna_comp].astype(str))
            
            # Actualizar 'archivo_encontrado' solo para los que aún no se han encontrado
            mascara = (df['archivo_encontrado'] == 'No encontrado') & (df[columna_master].astype(str).isin(codigos_comp))
            df.loc[mascara, 'archivo_encontrado'] = archivo
            
        except Exception as e:
            print(f"Error al procesar {archivo}: {str(e)}")
            continue

def procesar_archivos():
    for archivo_dic in archivos_diciembre:
        print(f"\nProcesando {archivo_dic}...")
        try:
            # Cargar archivo de diciembre
            df, columna_master = cargar_archivo_diciembre(archivo_dic)
            print(f"Registros a procesar: {len(df)}")
            
            # Buscar cada master_code en los archivos comparativos
            buscar_en_comparativa(df, columna_master)
            
            # Guardar resultado
            nombre_salida = f"Resultado_{archivo_dic}"
            ruta_salida = os.path.join(ruta_diciembre, nombre_salida)
            df.to_excel(ruta_salida, index=False)
            
            # Mostrar resumen
            encontrados = len(df[df['archivo_encontrado'] != 'No encontrado'])
            print(f"\nResumen para {archivo_dic}:")
            print(f"Total registros procesados: {len(df)}")
            print(f"Registros encontrados: {encontrados}")
            print(f"Registros no encontrados: {len(df) - encontrados}")
            print(f"Resultado guardado en: {nombre_salida}")
            
        except Exception as e:
            print(f"Error al procesar {archivo_dic}: {str(e)}")
            continue

if __name__ == "__main__":
    procesar_archivos()
