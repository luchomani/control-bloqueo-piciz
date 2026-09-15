import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime

# --- 1. FUNCIÓN PARA LIMPIAR LOS NÚMEROS DE DOCUMENTO Y TRÁNSITO ---
def limpiar_numeros(valor):
    if pd.isna(valor) or valor == '':
        return ""
    try:
        # Convierte a float, luego a int para quitar el .000000, y finalmente a string
        return str(int(float(valor)))
    except ValueError:
        return str(valor)

# Supongamos que tu DataFrame inicial se llama 'df'
# Aplicamos la limpieza a las columnas problemáticas
df['Número Tipo Documento'] = df['Número Tipo Documento'].apply(limpiar_numeros)
df['Transito'] = df['Transito'].apply(limpiar_numeros)

# --- 2. ELIMINAR DUPLICADOS AGRUPANDO PLACAS ---
# Definimos qué columnas deben ser idénticas para considerarse un duplicado
columnas_base = [
    'Fecha Registro', 'Compañía usuaria', 'Número Tipo Documento', 
    'Transito', 'Fecha de Báscula', 'Limite'
]

# Agrupamos y unimos las placas con una barra (ej: SPW433 / TAV719)
df = df.groupby(columnas_base, dropna=False, as_index=False).agg({
    'Placa': lambda x: ' / '.join(x.dropna().unique())
})

# Reorganizamos para que la columna 'Placa' vuelva a estar al principio
columnas_orden = ['Placa'] + columnas_base
df = df[columnas_orden]

# --- 3. CÁLCULO DE FECHAS Y DÍAS HÁBILES (Emulando DIA.LAB de Excel) ---
# Opcional: Si tienes una lista de fechas de festivos (como tu hoja 'BLOQUEO PICIZ WEB Y DATOS'), 
# puedes agregarla aquí en formato 'YYYY-MM-DD'.
festivos_colombia = [
    # '2026-01-01', '2026-01-12', etc...
]

# Aseguramos que la columna sea tipo fecha
df['Fecha de Báscula'] = pd.to_datetime(df['Fecha de Báscula'], errors='coerce')

def calcular_vencimiento(fecha_inicio, dias_limite):
    if pd.isna(fecha_inicio):
        return pd.NaT
    # np.busday_offset suma días hábiles saltando fines de semana. 
    # Para incluir festivos, añade el parámetro: holidays=festivos_colombia
    return pd.to_datetime(np.busday_offset(fecha_inicio.date(), int(dias_limite), roll='forward'))

# Aplicamos la fórmula de Vencimiento
df['Vencimiento 5 DIAS HABILES'] = df.apply(
    lambda row: calcular_vencimiento(row['Fecha de Báscula'], row['Limite']), axis=1
)

# Calculamos los días restantes basados en la fecha actual real
hoy = pd.Timestamp.today().normalize() # Esto toma la fecha de hoy a las 00:00:00
df['Días restantes'] = (df['Vencimiento 5 DIAS HABILES'] - hoy).dt.days

# Formateamos las fechas a texto YYYY-MM-DD para que se vean limpias en la tabla
df['Fecha de Báscula'] = df['Fecha de Báscula'].dt.strftime('%Y-%m-%d')
df['Vencimiento 5 DIAS HABILES'] = df['Vencimiento 5 DIAS HABILES'].dt.strftime('%Y-%m-%d')

# --- 4. APLICACIÓN DE COLORES Y RENDERIZADO EN STREAMLIT ---
def resaltar_filas(row):
    dias = row['Días restantes']
    
    # Si no hay fecha, no aplicamos color
    if pd.isna(dias):
        return [''] * len(row)
        
    # Lógica de colores según tus indicadores
    if dias <= 0:
        # Vencidos o Vencen Hoy (Rojo)
        color = '#d32f2f' 
        texto = 'white'
    elif 1 <= dias <= 2:
        # Próximos a Vencer (Amarillo)
        color = '#fbc02d'
        texto = 'black'
    else:
        # A Tiempo >= 3 (Verde)
        color = '#388e3c'
        texto = 'white'
        
    return [f'background-color: {color}; color: {texto}'] * len(row)

# Mostramos el DataFrame estilizado en Streamlit
st.dataframe(
    df.style.apply(resaltar_filas, axis=1),
    use_container_width=True,
    hide_index=True
)
