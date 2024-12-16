import os
import pandas as pd

# Especifica la ruta a la carpeta
ruta_carpeta = r"C:\Users\eyleen.godoy_arbusta\Desktop\Consolidar datos\ConBarcode-Consolidado"

# Lista para almacenar los DataFrames
dataframes = []

# Itera sobre cada archivo en la carpeta
for nombre_archivo in os.listdir(ruta_carpeta):
    if nombre_archivo.endswith('.xlsx'):
        # Obtiene el nombre del analista del nombre del archivo
        nombre_analista = os.path.splitext(nombre_archivo)[0]
        
        # Lee el archivo Excel
        ruta_archivo = os.path.join(ruta_carpeta, nombre_archivo)
        df = pd.read_excel(ruta_archivo)
        
        # Agrega una columna con el nombre del analista
        df['Analista'] = nombre_analista
        
        # Añade el DataFrame a la lista
        dataframes.append(df)

# Concatena todos los DataFrames en uno solo
df_consolidado = pd.concat(dataframes, ignore_index=True)

# Guarda el DataFrame consolidado en un nuevo archivo Excel
ruta_salida = r"C:\Users\eyleen.godoy_arbusta\Desktop\Consolidar datos\Consolidado.xlsx"
df_consolidado.to_excel(ruta_salida, index=False)

print(f"Datos consolidados guardados en: {ruta_salida}")
