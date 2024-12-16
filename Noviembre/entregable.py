import pandas as pd

# Paso 1: Leer los archivos
buscados_df = pd.read_excel('buscados_consolidados.xlsx')
consolidado_df = pd.read_excel('Consolidado.xlsx')
consolidado_encontrados_df = pd.read_excel('consolidado_encontrados.xlsx')

# Paso 2: Limpiar y renombrar columnas
# Eliminar la columna 'Unnamed: 8' del primer DataFrame
buscados_df.drop(columns=['Unnamed: 8'], inplace=True)

# Renombrar columnas en consolidado_encontrados_df para unificar
consolidado_encontrados_df.rename(columns={'país1': 'país', 'Valor scrapper': 'Valor encontrado'}, inplace=True)

# Asegurar que ambas columnas de códigos de barras estén en la misma forma
# Renombramos 'barcode' a 'Barcode' en el Consolidado para unificar
consolidado_df.rename(columns={'barcode': 'Barcode'}, inplace=True)

# Paso 3: Concatenar los tres DataFrames
resultado_df = pd.concat([buscados_df, consolidado_df, consolidado_encontrados_df], ignore_index=True)

# Paso 4: Identificar y eliminar duplicados
duplicados_df = resultado_df[resultado_df.duplicated(keep=False)]  # Mantener todos los duplicados
resultado_df_unique = resultado_df.drop_duplicates(keep='first')   # Mantener solo el primero

# Paso 5: Guardar los resultados en archivos Excel
resultado_df_unique.to_excel('entregablenoviembre.xlsx', index=False)
duplicados_df.to_excel('duplicados_noviembre.xlsx', index=False)

print("Eliminación de duplicados completada. Archivos guardados:")
print("- entregablenoviembre.xlsx: Contiene los datos únicos.")
print("- duplicados_noviembre.xlsx: Contiene los datos duplicados.")
