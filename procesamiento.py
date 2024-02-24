import pandas as pd
import openpyxl
import requests
from datetime import date
import sqlite3

# Leyendo el archivo de Excel
df = pd.read_excel('reactiva.xlsx', sheet_name='TRANSFERENCIAS 2020')
df.head(0)

def limpiar_nombres_columnas(df):
    df.columns = df.columns.str.lower()
    df.columns = df.columns.str.replace(' ', '_')
    df.columns = df.columns.str.normalize('NFKD').str.encode('ascii', errors='ignore').str.decode('utf-8')
    return df

df_limpio_nombres = limpiar_nombres_columnas(df)
df_limpio_nombres.head(1)

# Eliminando columnas no deseadas
df_sin_columnas = df_limpio_nombres.drop(['id', 'tipo_moneda.1'], axis=1)
df_sin_columnas.head(1)

# Reemplazando valores en la columna 'dispositivo_legal'
df_con_dispositivo_legal = df_sin_columnas
df_con_dispositivo_legal['dispositivo_legal'] = df_con_dispositivo_legal['dispositivo_legal'].replace({'0m': ''}, regex=True)
df_con_dispositivo_legal.head(1)

# Obteniendo el tipo de cambio actual del dólar
def obtener_tipo_cambio_sunat(fecha):
    try:
        url = f"https://api.apis.net.pe/v1/tipo-cambio-sunat?fecha={fecha}"
        response = requests.get(url)
        response.raise_for_status()
        return response.json()['compra']
    except requests.RequestException as e:
        print("Error al obtener el tipo de cambio:", e)
        return None

fecha_actual = date.today().strftime('%Y-%m-%d')
tipo_cambio_usd = obtener_tipo_cambio_sunat(fecha_actual)

# Dolarizando los montos de inversión y transferencia
df_dolarizado = df_con_dispositivo_legal
df_dolarizado['monto_inversion_dol'] = (df_dolarizado['monto_de_inversion'] / tipo_cambio_usd).round(2)
df_dolarizado['monto_transferencia2020_dol'] = (df_dolarizado['monto_de_transferencia_2020'] / tipo_cambio_usd).round(2)
if 'monto_dolares' in df_dolarizado.columns:
    df_dolarizado = df_dolarizado.drop('monto_dolares', axis=1)
df_dolarizado.head(3)

# Mapeando valores en la columna 'estado'
df_con_valores_estado = df_dolarizado
df_con_valores_estado['estado'] = df_con_valores_estado['estado'].replace('En Ejecución', 'Ejecución')
df_con_valores_estado['estado'] = df_con_valores_estado['estado'].replace('Convenio y/o Contrato Resuelto', 'Resuelto')
df_con_valores_estado.estado.unique()

def asignar_puntuacion(estado):
    valor = estado
    if valor == 'Resuelto':
        puntuacion = 0
    elif valor == 'Actos Previos':
        puntuacion = 1
    elif valor == 'Ejecución':
        puntuacion = 2
    elif valor == 'Concluido':
        puntuacion = 3
    else:
        puntuacion = None
    return puntuacion

# Creando una nueva columna con la puntuación del estado
df_con_puntuacion_estado = df_con_valores_estado
df_con_puntuacion_estado['puntuacion'] = df_con_puntuacion_estado['estado'].apply(asignar_puntuacion)
df_con_puntuacion_estado.head(2)

# Conectando a la base de datos SQLite y almacenando datos únicos
conexion = sqlite3.connect('ubicaciones_reactiva.db')
datos_ubigeo = df_con_puntuacion_estado[['ubigeo', 'region', 'provincia', 'distrito']].drop_duplicates()
datos_ubigeo.to_sql('ubigeo', conexion, if_exists='replace', index=False)
conexion.commit()
conexion.close()
print("Tabla de ubigeos almacenada en la base de datos.")

# Filtrando por tipo 'Urbano' y estados 1, 2, 3
condicion_tipo_estado = (df_con_puntuacion_estado['tipologia'] == 'Equipamiento Urbano') & (df_con_puntuacion_estado['puntuacion'].between(1, 3))
df_filtrado_por_tipo_estado = df_con_puntuacion_estado[condicion_tipo_estado]

# Obteniendo lista de regiones únicas
lista_regiones_unicas = df_filtrado_por_tipo_estado['region'].unique()

for region_unica in lista_regiones_unicas:
    condicion_region = df_filtrado_por_tipo_estado['region'] == region_unica
    df_por_region = df_filtrado_por_tipo_estado[condicion_region]
    
    if not df_por_region.empty:
        df_ordenado_por_region = df_por_region.sort_values(by='monto_de_inversion', ascending=False)
        top_5_obras_por_region = df_ordenado_por_region.head(5)
        
        # Guardando el resultado en un archivo Excel
        nombre_archivo_excel = f"top5_inversion_{region_unica.replace(' ', '_')}.xlsx"
        top_5_obras_por_region.to_excel(nombre_archivo_excel, index=False)
        print(f"Se generó el reporte para la región {region_unica} en: {nombre_archivo_excel}")
    else:
        print(f"No hay datos para la región {region_unica}. No se generará el reporte.")
