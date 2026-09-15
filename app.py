import streamlit as st
import pandas as pd
import numpy as np
import datetime

# --- Configuración de la página ---
st.set_page_config(page_title="Control de Bloqueo | Zona Franca", page_icon="🚦", layout="wide")

st.markdown("""
    <style>
    .main-header { color: #1f4e3d; font-weight: bold; }
    .sub-header { color: #4a4a4a; }
    .stDataFrame { font-size: 14px; }
    </style>
    """, unsafe_allow_html=True)

st.markdown("<h1 class='main-header'>🚦 Módulo de Control de Bloqueos (Reporte PW)</h1>", unsafe_allow_html=True)
st.markdown("<p class='sub-header'>Carga el archivo Excel <b>ReportePW</b> descargado de PICIZ para automatizar el filtrado y visualizar las alertas de vencimiento.</p>", unsafe_allow_html=True)

# --- Funciones Auxiliares ---
def limpiar_numeros(valor):
    """Elimina decimales y ceros sobrantes (.000000) de los IDs"""
    if pd.isna(valor) or valor == '':
        return ""
    try:
        return str(int(float(valor)))
    except ValueError:
        return str(valor)

# --- Carga de Archivo ---
uploaded_file = st.file_uploader("Sube el archivo Excel (ReportePW.xlsx)", type=["xlsx", "xls"])

if uploaded_file is not None:
    with st.spinner("Procesando, unificando placas y calculando fechas..."):
        try:
            # 1. Buscar la fila correcta de encabezados
            df_raw = pd.read_excel(uploaded_file, header=None)
            header_row_index = -1
            for i, row in df_raw.iterrows():
                if "NOMBRE COMPANIA" in str(row.values):
                    header_row_index = i
                    break
            
            if header_row_index == -1:
                st.error("❌ No se encontró la fila de encabezados. Verifica el formato del Excel.")
                st.stop()
            
            # 2. Leer el Excel desde los encabezados
            df = pd.read_excel(uploaded_file, header=header_row_index)
            df.columns = df.columns.str.strip().str.replace(r'\r\n', '', regex=True)
            
            # 3. Filtrar columnas solicitadas
            columnas_esperadas = {
                'NOMBRE COMPANIA': 'Compañia usuaria',
                'PLACA': 'Placa',
                'FECHA REGISTRO': 'Fecha Registro',
                'NUM DEL DOC. ADUANERO': 'Número Tipo Documento',
                'TRANSITO': 'Transito',
                'FECHA BASCULA': 'Fecha de Báscula'
            }
            columnas_existentes = {k: v for k, v in columnas_esperadas.items() if k in df.columns}
            df_filtrado = df[list(columnas_existentes.keys())].rename(columns=columnas_existentes)
            
            # Limpiar filas vacías sin placas
            if 'Placa' in df_filtrado.columns:
                df_filtrado = df_filtrado.dropna(subset=['Placa'])

            # 4. Limpieza de números (Documentos y Tránsito)
            for col in ['Número Tipo Documento', 'Transito']:
                if col in df_filtrado.columns:
                    df_filtrado[col] = df_filtrado[col].apply(limpiar_numeros)

            # 5. Formateo de fechas a Date
            for col in ['Fecha Registro', 'Fecha de Báscula']:
                if col in df_filtrado.columns:
                    df_filtrado[col] = pd.to_datetime(df_filtrado[col], format='%d/%m/%Y', errors='coerce').dt.date

            # Límite constante
            df_filtrado['Limite'] = 5

            # 6. Agrupar duplicados (Unir Placas)
            # Agrupa por todas las columnas base para fusionar las placas de un mismo tránsito
            columnas_base = [col for col in ['Fecha Registro', 'Compañia usuaria', 'Número Tipo Documento', 'Transito', 'Fecha de Báscula', 'Limite'] if col in df_filtrado.columns]
            
            df_filtrado = df_filtrado.groupby(columnas_base, dropna=False, as_index=False).agg({
                'Placa': lambda x: ' / '.join(x.dropna().unique())
            })

            # 7. Cálculo de Fechas Equivalente a DIA.LAB de Excel
            def calcular_vencimiento(fecha_registro):
                if pd.isna(fecha_registro):
                    return None
                try:
                    # np.busday_offset suma los 5 días ignorando sábados y domingos
                    fecha_venc = np.busday_offset(np.datetime64(fecha_registro), 5, roll='forward')
                    return pd.to_datetime(fecha_venc).date()
                except:
                    return None
                    
            df_filtrado['Vencimiento 5 DIAS HABILES'] = df_filtrado['Fecha Registro'].apply(calcular_vencimiento)
            
            # Cálculo de días restantes contra la fecha actual
            hoy = datetime.date.today()
            
            def calcular_dias_restantes(fecha_venc):
                if pd.isna(fecha_venc):
                    return None
                return (fecha_venc - hoy).days
                
            df_filtrado['Días restantes'] = df_filtrado['Vencimiento 5 DIAS HABILES'].apply(calcular_dias_restantes)
            
            # 8. Reorganizar columnas para la vista final
            orden_columnas = ['Placa', 'Fecha Registro', 'Compañia usuaria', 'Número Tipo Documento', 'Transito', 'Fecha de Báscula', 'Limite', 'Vencimiento 5 DIAS HABILES', 'Días restantes']
            orden_columnas = [col for col in orden_columnas if col in df_filtrado.columns]
            df_final = df_filtrado[orden_columnas].fillna('')

            # 9. Aplicar Colores (Semaforización)
            def apply_row_colors(row):
                try:
                    dias = float(row['Días restantes'])
                    if dias <= 0:
                        return ['background-color: #d32f2f; color: white;'] * len(row) # Rojo
                    elif 1 <= dias <= 2:
                        return ['background-color: #fbc02d; color: black;'] * len(row)  # Amarillo
                    elif dias >= 3:
                        return ['background-color: #388e3c; color: white;'] * len(row)  # Verde
                except:
                    pass
                return [''] * len(row)

            # --- Interfaz de Resultados ---
            # Calcular KPIs
            vencidos = df_final[pd.to_numeric(df_final['Días restantes'], errors='coerce') <= 0].shape[0]
            riesgo = df_final[(pd.to_numeric(df_final['Días restantes'], errors='coerce') >= 1) & (pd.to_numeric(df_final['Días restantes'], errors='coerce') <= 2)].shape[0]
            a_tiempo = df_final[pd.to_numeric(df_final['Días restantes'], errors='coerce') >= 3].shape[0]
            
            st.markdown("### 📊 Resumen de Operaciones")
            col1, col2, col3 = st.columns(3)
            col1.metric("🔴 Vencidos o Vencen Hoy", vencidos)
            col2.metric("🟡 Próximos a Vencer (1-2 días)", riesgo)
            col3.metric("🟢 A Tiempo (>= 3 días)", a_tiempo)
            
            st.markdown("### 📋 Panel de Control de Ingresos")
            styled_df = df_final.style.apply(apply_row_colors, axis=1)
            st.dataframe(styled_df, use_container_width=True, height=500, hide_index=True)
            
            # 10. Exportar a Excel
            st.markdown("### 💾 Exportar Datos")
            def convert_df_to_excel(df_styled):
                from io import BytesIO
                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_styled.to_excel(writer, index=False, sheet_name='Control_Bloqueo')
                return output.getvalue()

            st.download_button(
                label="📥 Descargar Reporte Formateado (Excel)",
                data=convert_df_to_excel(styled_df),
                file_name=f"Control_Bloqueos_{datetime.date.today()}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
        except Exception as e:
            st.error(f"❌ Ocurrió un error al procesar el archivo: {e}")

else:
    st.info("Esperando archivo... Por favor, carga el `ReportePW.xlsx` para comenzar.")
