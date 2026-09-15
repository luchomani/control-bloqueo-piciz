import streamlit as st
import pandas as pd
import datetime
import numpy as np

# Configuración de la página
st.set_page_config(
    page_title="Control de Bloqueos | Zona Franca de Cúcuta",
    page_icon="🏢",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- ESTILOS CSS CORPORATIVOS E INGENIERÍA DE DISEÑO ---
st.markdown("""
    <style>
    /* Colores institucionales y tipografía */
    :root {
        --primary-green: #12402A;
        --secondary-green: #1F4E3D;
        --accent-olive: #7A8B2C;
        --bg-light: #F8F9FA;
        --border-color: #D1D5DB;
    }
    
    .main {
        background-color: #FFFFFF;
    }
    
    /* Encabezados principales */
    .executive-title {
        font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif;
        color: #12402A;
        font-size: 26px;
        font-weight: 700;
        margin-bottom: 0px;
        letter-spacing: -0.5px;
    }
    
    .executive-subtitle {
        font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif;
        color: #4B5563;
        font-size: 14px;
        margin-top: 5px;
        margin-bottom: 25px;
    }
    
    /* Contenedores y Tarjetas con bordes limpios */
    .metric-container {
        background-color: #F9FAFB;
        border: 1px solid #E5E7EB;
        border-radius: 6px;
        padding: 15px;
        text-align: center;
        box-shadow: 0 1px 2px rgba(0,0,0,0.02);
    }
    
    /* Sidebar ejecutivo */
    [data-testid="stSidebar"] {
        background-color: #F3F4F6;
        border-right: 1px solid #E5E7EB;
    }
    
    /* Tablas y DataFrames */
    .stDataFrame {
        border: 1px solid #E5E7EB;
        border-radius: 6px;
    }
    </style>
    """, unsafe_allow_html=True)

# --- GESTIÓN DE ESTADO DE SESIÓN ---
if 'file_processed' not in st.session_state:
    st.session_state.file_processed = False
if 'df_final' not in st.session_state:
    st.session_state.df_final = None
if 'fecha_ref' not in st.session_state:
    st.session_state.fecha_ref = datetime.date.today()

# --- PANEL LATERAL CORPORATIVO ---
with st.sidebar:
    # Mostrar el logo corporativo si está disponible en el directorio local
    try:
        st.image("LOGO ZFS-ZFC.jpeg", use_container_width=True)
    except:
        st.markdown("<h3 style='color: #12402A; text-align: center;'>ZONA FRANCA</h3>", unsafe_allow_html=True)
    
    st.markdown("<hr style='margin: 10px 0; border: 1px solid #D1D5DB;'>", unsafe_allow_html=True)
    st.markdown("### Configuración de Operación", help="Parámetros de control logístico")
    
    fecha_actual_input = st.date_input(
        "Fecha Actual de Referencia",
        value=st.session_state.fecha_ref,
        help="Equivalente a la celda de referencia temporal en el sistema de control."
    )
    st.session_state.fecha_ref = fecha_actual_input

    st.markdown("---")
    uploaded_file = st.file_uploader("Cargar Reporte PW (Excel)", type=["xlsx", "xls"])
    
    st.markdown("---")
    # Botón de reinicio / limpieza
    if st.button("Limpiar Datos y Sesión", use_container_width=True):
        st.session_state.file_processed = False
        st.session_state.df_final = None
        st.rerun()

# --- CONTENIDO PRINCIPAL ---
st.markdown("<p class='executive-title'>Módulo de Control de Bloqueos y Vencimientos</p>", unsafe_allow_html=True)
st.markdown("<p class='executive-subtitle'>Sistema de auditoría y seguimiento de ingresos para operaciones de comercio exterior.</p>", unsafe_allow_html=True)

# Función auxiliar para limpieza numérica
def limpiar_numeros(valor):
    if pd.isna(valor) or valor == '':
        return ""
    try:
        val_str = str(valor).strip()
        if val_str == '' or val_str.lower() in ['nan', 'none']:
            return ""
        if '.' in val_str:
            return str(int(float(val_str)))
        return val_str
    except ValueError:
        return str(valor).strip()

if uploaded_file is not None:
    try:
        with st.spinner("Procesando registros aduaneros..."):
            # 1. Cargar archivo crudo
            df_raw = pd.read_excel(uploaded_file, header=None, dtype=str)
            
            # 2. Localizar cabeceras
            header_row_index = -1
            for i, row in df_raw.iterrows():
                row_str = " ".join(str(val) for val in row.values)
                if "NOMBRE COMPANIA" in row_str or "PLACA" in row_str:
                    header_row_index = i
                    break
            
            if header_row_index == -1:
                st.error("No se localizó la fila de encabezados estándar en el archivo cargado.")
                st.stop()
            
            df = pd.read_excel(uploaded_file, header=header_row_index, dtype=str)
            df.columns = df.columns.str.strip().str.replace(r'\r\n', '', regex=True)
            
            columnas_esperadas = {
                'NOMBRE COMPANIA': 'Compañía Usuaria',
                'PLACA': 'Placa',
                'FECHA REGISTRO': 'Fecha Registro',
                'TIPO INGRESO': 'Tipo Ingreso', 
                'NUM DEL DOC. ADUANERO': 'Número Documento',
                'TRANSITO': 'Tránsito',
                'FECHA BASCULA': 'Fecha de Báscula'
            }
            
            columnas_existentes = {k: v for k, v in columnas_esperadas.items() if k in df.columns}
            df_filtrado = df[list(columnas_existentes.keys())].rename(columns=columnas_existentes)
            
            if 'Tipo Ingreso' in df_filtrado.columns:
                df_filtrado = df_filtrado[df_filtrado['Tipo Ingreso'].str.contains('FORMULARIO', na=False, case=False)]
                df_filtrado = df_filtrado.drop(columns=['Tipo Ingreso'])
            
            df_filtrado = df_filtrado.dropna(subset=['Placa'])
            df_filtrado['Placa'] = df_filtrado['Placa'].astype(str).str.strip()
            df_filtrado = df_filtrado[df_filtrado['Placa'] != '']

            for col in ['Número Documento', 'Tránsito']:
                if col in df_filtrado.columns:
                    df_filtrado[col] = df_filtrado[col].apply(limpiar_numeros)

            for col in ['Fecha Registro', 'Fecha de Báscula']:
                if col in df_filtrado.columns:
                    df_filtrado[col] = pd.to_datetime(df_filtrado[col], dayfirst=True, errors='coerce').dt.date

            columnas_dedup = [col for col in ['Placa', 'Fecha Registro', 'Compañía Usuaria', 'Número Documento', 'Tránsito', 'Fecha de Báscula'] if col in df_filtrado.columns]
            df_filtrado = df_filtrado.drop_duplicates(subset=columnas_dedup, keep='first')

            df_filtrado['Límite'] = 5
            
            def calcular_vencimiento(row):
                fecha_base = row['Fecha de Báscula'] if pd.notna(row['Fecha de Báscula']) and str(row['Fecha de Báscula']) != 'NaT' else row['Fecha Registro']
                if pd.isna(fecha_base) or str(fecha_base) == 'NaT':
                    return None
                try:
                    fecha_venc = np.busday_offset(np.datetime64(fecha_base), 5, roll='forward')
                    return pd.to_datetime(fecha_venc).date()
                except:
                    return None
                    
            df_filtrado['Vencimiento (5 Días Hábiles)'] = df_filtrado.apply(calcular_vencimiento, axis=1)
            
            def calcular_dias_restantes(fecha_venc):
                if pd.isna(fecha_venc) or str(fecha_venc) == 'NaT':
                    return None
                delta = fecha_venc - st.session_state.fecha_ref
                return delta.days
                
            df_filtrado['Días Restantes'] = df_filtrado['Vencimiento (5 Días Hábiles)'].apply(calcular_dias_restantes)
            
            orden_columnas = ['Placa', 'Fecha Registro', 'Compañía Usuaria', 'Número Documento', 'Tránsito', 'Fecha de Báscula', 'Límite', 'Vencimiento (5 Días Hábiles)', 'Días Restantes']
            orden_columnas = [col for col in orden_columnas if col in df_filtrado.columns]
            df_final = df_filtrado[orden_columnas].copy()
            
            df_final['Días Restantes'] = pd.to_numeric(df_final['Días Restantes'], errors='coerce').fillna(0).astype(int)
            st.session_state.df_final = df_final
            st.session_state.file_processed = True

    except Exception as e:
        st.error(f"Error en el procesamiento del archivo: {e}")

# --- VISUALIZACIÓN DE RESULTADOS ---
if st.session_state.file_processed and st.session_state.df_final is not None:
    df_res = st.session_state.df_final
    
    # KPIs Ejecutivos
    dias_series = df_res['Días Restantes']
    vencidos = (dias_series <= 0).sum()
    riesgo = ((dias_series >= 1) & (dias_series <= 2)).sum()
    a_tiempo = (dias_series >= 3).sum()
    
    st.markdown(f"<p style='font-size: 15px; font-weight: 600; color: #12402A;'>Resumen de Estado Operativo — Fecha de Corte: {st.session_state.fecha_ref.strftime('%d/%m/%Y')}</p>", unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("Vencidos o Vencen Hoy", int(vencidos))
    with col2:
        st.metric("Próximos a Vencer (1-2 días)", int(riesgo))
    with col3:
        st.metric("En Plazo (>= 3 días)", int(a_tiempo))
        
    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("<p style='font-size: 16px; font-weight: 600; color: #12402A;'>Detalle de Registros y Control de Plazos</p>", unsafe_allow_html=True)
    
    def apply_executive_colors(row):
        try:
            dias = int(row['Días Restantes'])
            if dias <= 0:
                return ['background-color: #FEE2E2; color: #991B1B; font-weight: 500;'] * len(row) # Rojo sobrio
            elif 1 <= dias <= 2:
                return ['background-color: #FEF3C7; color: #92400E; font-weight: 500;'] * len(row) # Amarillo sobrio
            elif dias >= 3:
                return ['background-color: #ECFDF5; color: #065F46;'] * len(row) # Verde sobrio
        except:
            pass
        return [''] * len(row)

    styled_df = df_res.style.apply(apply_executive_colors, axis=1).format({'Días Restantes': '{:d}'})
    st.dataframe(styled_df, use_container_width=True, height=450, hide_index=True)
    
    # Sección de Exportación
    st.markdown("<br>", unsafe_allow_html=True)
    col_exp1, col_exp2 = st.columns([3, 1])
    with col_exp2:
        def convert_df_to_excel(df_styled):
            from io import BytesIO
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_styled.to_excel(writer, index=False, sheet_name='Control_Bloqueos')
            return output.getvalue()

        excel_data = convert_df_to_excel(styled_df)
        st.download_button(
            label="Descargar Reporte",
            data=excel_data,
            file_name=f"Control_Bloqueos_{st.session_state.fecha_ref}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

else:
    st.info("Cargue un archivo en la barra lateral para iniciar el procesamiento de control aduanero.")
