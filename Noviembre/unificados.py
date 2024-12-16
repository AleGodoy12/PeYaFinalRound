import pandas as pd
import re

# Rutas corregidas (usa rutas absolutas o relativas según sea necesario)
archivos = {
    "buscados_consolidados": r"C:\Users\eyleen.godoy_arbusta\Desktop\Consolidar datos\buscados_consolidados.xlsx",
    "consolidado_encontrados": r"C:\Users\eyleen.godoy_arbusta\Desktop\Consolidar datos\consolidado_encontrados.xlsx",
    "Consolidado": r"C:\Users\eyleen.godoy_arbusta\Desktop\Consolidar datos\Consolidado.xlsx"
}

# Función para normalizar nombres
def normalizar_nombre(nombre):
    equivalencias = {
        "alecci aylen": "aylen alecci",
        "axel esteban": "aguilera gonzález axel esteban",
        "danna eberle": "eberle danna",
        "enzo ezequiel": "juarez enzo",
        "maria de los angeles castello": "maría de los ángeles castello",
        "pucheta pucheta sasha antonella": "pucheta sasha antonella",
        "sasha antonella": "pucheta sasha antonella",
        "douglas cacharani": "douglas axel cacharani soliz",
        "franco guzman": "guzmán franco",
        "maia perovich": "perovich maia"
    }
    nombre = nombre.lower()
    nombre = re.sub(r"[^\w\s]", "", nombre)  # Elimina caracteres especiales
    nombre = " ".join(nombre.split())  # Elimina espacios extra
    return equivalencias.get(nombre, nombre)

# Diccionario para almacenar los DataFrames
dataframes = {}

# Lee cada archivo y almacena el DataFrame
for nombre, ruta in archivos.items():
    try:
        df = pd.read_excel(ruta)
        if 'Analista' in df.columns:
            df['Analista'] = df['Analista'].apply(normalizar_nombre)
        dataframes[nombre] = df
    except FileNotFoundError:
        print(f"Archivo no encontrado: {ruta}")
    except Exception as e:
        print(f"Error al procesar {ruta}: {e}")

# Combina todos los DataFrames en uno solo
df_consolidado = pd.concat(dataframes.values(), ignore_index=True)

# Cantidad total gestionada por analista
cantidad_por_analista = df_consolidado.groupby('Analista').size()
print("Cantidad total gestionada por analista:")
print(cantidad_por_analista)

# Suma total gestionada por todo el equipo
total_gestionado = cantidad_por_analista.sum()
print("\nSuma total gestionada por todo el equipo:", total_gestionado)
