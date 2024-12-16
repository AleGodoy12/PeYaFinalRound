import pandas as pd
import os

# Rutas de los archivos
ruta_principal = r"C:\Users\eyleen.godoy_arbusta\Desktop\Consolidar datos"
ruta_base_final = os.path.join(ruta_principal, "Base final")

# Archivo a procesar
archivo_base = os.path.join(ruta_base_final, "basefinal.xlsx")

def transformar_base():
    print("Iniciando transformación de base final...")
    
    try:
        # Cargar archivo base
        print("Cargando base final...")
        df_base = pd.read_excel(archivo_base)
        
        # Lista de campos que queremos convertir a filas, con orden específico
        campos_busqueda = {
            'title': 1,
            'image': 2,
            'brand': 3,
            'content_value': 4,
            'content_unit': 5,
            'units_per_pack': 6,
            'additional_image': 7,
            'description': 8,
            'storage_type': 9,
            'packaging': 10,
            'info_nutricional': 11
        }
        
        # Lista para almacenar las nuevas filas
        filas_nuevas = []
        
        # Procesar cada registro
        total_registros = len(df_base)
        print(f"Procesando {total_registros} registros...")
        
        for idx, row in df_base.iterrows():
            mastercode = row[df_base.columns[0]]  # Primera columna (mastercode)
            barcode = row.get('barcode', '')      # Columna barcode
            image = row.get('image', '')          # Columna image
            pais = row.get('pais', '')            # Columna pais
            
            # Crear una fila para cada campo de búsqueda
            for campo, orden in campos_busqueda.items():
                filas_nuevas.append({
                    'mastercode': mastercode,
                    'barcode': barcode,
                    'image': image,
                    'pais': pais,
                    'campo_busqueda': campo,
                    'valor_base': row.get(campo, ''),
                    'orden': orden  # Campo temporal para ordenar
                })
        
        # Crear nuevo DataFrame con la estructura transformada
        df_transformado = pd.DataFrame(filas_nuevas)
        
        # Ordenar primero por mastercode y luego por el orden de los campos
        df_transformado = df_transformado.sort_values(['mastercode', 'orden'])
        
        # Eliminar la columna de orden que ya no necesitamos
        df_transformado = df_transformado.drop('orden', axis=1)
        
        # Guardar resultado
        archivo_salida = os.path.join(ruta_base_final, "base_transformada.xlsx")
        df_transformado.to_excel(archivo_salida, index=False)
        
        # Mostrar resumen
        print("\nResumen:")
        print(f"Total registros procesados: {total_registros}")
        print(f"Total filas generadas: {len(df_transformado)}")
        print(f"Archivo guardado en: {archivo_salida}")
        
    except Exception as e:
        print(f"Error en el proceso: {str(e)}")

if __name__ == "__main__":
    transformar_base() 