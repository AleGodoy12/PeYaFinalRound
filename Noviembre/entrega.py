import pandas as pd

# Leer el archivo
resultado_df = pd.read_excel('entregablenoviembre.xlsx')

# Contar las ocurrencias de cada barcode
barcode_counts = resultado_df['Barcode'].value_counts()

# Filtrar los barcodes que no aparecen exactamente 11 veces
duplicados = barcode_counts[barcode_counts != 11]

# Mostrar el resultado
total_duplicados = duplicados.count()
print(f"Total de barcodes que no cumplen con el criterio de 11 apariciones: {total_duplicados}")
print("Barcodes duplicados y sus respectivas cuentas:")
print(duplicados)
