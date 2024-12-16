import pandas as pd
import os

# Ruta de la carpeta que contiene los archivos
ruta_datos = r"C:\Users\eyleen.godoy_arbusta\Desktop\Consolidar datos\DatosBuscados"
# Filtrar los archivos para evitar archivos temporales de Excel
archivos = [f for f in os.listdir(ruta_datos) if f.endswith('.xlsx') and not f.startswith('~$')]

# Crear un DataFrame vacío para almacenar los resultados
resultados = pd.DataFrame()

# Contador de verdaderos por analista
contador_verdaderos = {}

# Función para procesar archivos y contar verdaderos
def procesar_archivo(archivo, analista):
    global resultados, contador_verdaderos
    df = pd.read_excel(os.path.join(ruta_datos, archivo))
    
    # Filtrar los datos donde 'Check envio' es True
    df_verdaderos = df[df['Check envio'] == True]
    
    # Agregar la columna 'Analista'
    df_verdaderos['Analista'] = analista
    
    # Eliminar la columna 'Check envio'
    df_verdaderos = df_verdaderos.drop(columns=['Check envio'])
    
    # Concatenar los resultados
    resultados = pd.concat([resultados, df_verdaderos], ignore_index=True)
    
    # Contar los verdaderos por analista
    contador_verdaderos[analista] = contador_verdaderos.get(analista, 0) + df_verdaderos.shape[0]

# Procesar archivos de la carpeta DatosBuscados
for archivo in archivos:
    # Obtener el nombre completo del analista del archivo
    if archivo.startswith('Reporte_QA_') and archivo.endswith('.xlsx'):
        analista = archivo.replace('Reporte_QA_', '').replace('.xlsx', '').strip()  # Extraer el nombre del analista
    else:
        continue  # Saltar archivos que no cumplen con el formato
    procesar_archivo(archivo, analista)

# Guardar el DataFrame resultante en un nuevo archivo Excel
ruta_salida = r"C:\Users\eyleen.godoy_arbusta\Desktop\Consolidar datos\buscados_consolidados.xlsx"
resultados.to_excel(ruta_salida, index=False)

# Imprimir el mensaje en consola
suma_total = 0
for analista, cantidad in contador_verdaderos.items():
    print(f"Analista: {analista}, Cantidad: {cantidad}")
    suma_total += cantidad

# Imprimir la suma total
print(f"Suma total de 'VERDADERO': {suma_total}")
