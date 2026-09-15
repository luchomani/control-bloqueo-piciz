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
    """Elimina decimales y ceros sobrantes (.000000) de los IDs y fuerza a string limpio"""
    if pd.isna(valor) or valor == '':
        return ""
    try:
        val_str = str(valor).strip()
        if val_str == '' or val_str.lower() in ['nan', 'none']:
            return ""
        # Si es un número con decimales de Excel
        if '.' in val_str:
            return str(int(float(val_str)))
        return val_str
    except ValueError:
        return str(valor).strip()

# --- Carga de Archivo ---
uploaded_file = st.file_uploader("Sube el archivo Excel (ReportePW.xlsx)", type=["xlsx", "xls"])

if uploaded_file is not None:
    with st.spinner("Procesando datos, limpiando duplicados y calculando fechas..."):
        try:
            # 1. Buscar la fila correcta de encabezados leyendo todo como string
            df_raw = pd.read_excel(uploaded_file, header=None, dtype=str)
            header_row_index = -1
            for i, row in df_raw.iterrows():
                row_str = " ".join(str(val) for val in row.values)
                if "NOMBRE COMPANIA" in row_str or "PLACA" in row_str:
                    header_row_index = i
                    break
            
            if header_row_index == -1:
                st.error("❌ No se encontró la fila de encabezados en el Excel. Asegúrate de que contenga las columnas requeridas.")
                st.stop()
            
            # 2. Leer el Excel desde los encabezados detectados forzando tipo texto
            df = pd.read_excel(uploaded_file, header=header_row_index, dtype=str)
            df.columns = df.columns.str.strip().str.replace(r'\r\n', '', regex=True)
            
            # 3. Mapear y filtrar columnas solicitadas (incluyendo Conductor/Identificación)
            columnas_esperadas = {
                'NOMBRE COMPANIA': 'Compañía usuaria',
                'PLACA': 'Placa',
                'CONDUCTOR': 'Identificación / Conductor',
                'IDENTIFICACION': 'Identificación / Conductor',
                'CEDULA': 'Identificación / Conductor',
                'FECHA REGISTRO': 'Fecha Registro',
                'NUM DEL DOC. ADUANERO': 'Número Tipo Documento',
                'TRANSITO': 'Transito',
                'FECHA BASCULA': 'Fecha de Báscula'
            }
            
            # Buscar columnas existentes que coincidan con el mapa
            columnas_existentes = {}
            for col_excel in df.columns:
                col_upper = col_excel.upper()
                for key, target_name in columnas_esperadas.items():
                    if key in col_upper and target_name not in columnas_existentes.values():
                        columnas_existentes[col_excel] = target_name
            
            if not columnas_existentes:
                st.error("❌ El archivo no contiene las columnas esperadas de PICIZ.")
                st.stop()
                
            df_filtrado = df[list(columnas_existentes.keys())].rename(columns=columnas_existentes)
            
            # Asegurar que todas las columnas clave existan aunque no vengan en el Excel
            for target in ['Placa', 'Identificación / Conductor', 'Fecha Registro', 'Compañía usuaria', 'Número Tipo Documento', 'Transito', 'Fecha de Báscula']:
                if target not in df_filtrado.columns:
                    df_filtrado[target] = ""

            # Limpiar filas completamente vacías sin placas
            df_filtrado = df_filtrado.dropna(subset=['Placa'])
            df_filtrado['Placa'] = df_filtrado['Placa'].astype(str).str.strip()
            df_filtrado = df_filtrado[df_filtrado['Placa'] != '']

            # 4. Limpieza de números (Documentos, Tránsito e Identificación)
            for col in ['Número Tipo Documento', 'Transito', 'Identificación / Conductor']:
                if col in df_filtrado.columns:
                    df_filtrado[col] = df_filtrado[col].apply(limpiar_numeros)

            # 5. Formateo seguro de fechas
            for col in ['Fecha Registro', 'Fecha de Báscula']:
                if col in df_filtrado.columns:
                    df_filtrado[col] = pd.to_datetime(df_filtrado[col], errors='coerce').dt.date

            # Límite constante (5 días hábiles)
            df_filtrado['Limite'] = 5

            # 6. Eliminar registros duplicados idénticos (evitando mezcla confusa de placas)
            columnas_dedup = [col for col in ['Placa', 'Identificación / Conductor', 'Fecha Registro', 'Compañía usuaria', 'Número Tipo Documento', 'Transito', 'Fecha de Báscula'] if col in df_filtrado.columns]
            df_filtrado = df_filtrado.drop_duplicates(subset=columnas_dedup, keep='first')

            # 7. Cálculo de Fechas Equivalente a DIA.LAB de Excel
            def calcular_vencimiento(fecha_registro):
                if pd.isna(fecha_registro) or str(fecha_registro) == 'NaT':
                    return None
                try:
                    fecha_venc = np.busday_offset(np.datetime64(fecha_registro), 5, roll='forward')
                    return pd.to_datetime(fecha_venc).date()
                except:
                    return None
                    
            df_filtrado['Vencimiento 5 DIAS HABILES'] = df_filtrado['Fecha Registro'].apply(calcular_vencimiento)
            
            # Cálculo de días restantes contra la fecha actual real
            hoy = datetime.date.today()
            
            def calcular_dias_restantes(fecha_venc):
                if pd.isna(fecha_venc) or str(fecha_venc) == 'NaT':
                    return None
                return (fecha_venc - hoy).days
                
            df_filtrado['Días restantes'] = df_filtrado['Vencimiento 5 DIAS HABILES'].apply(calcular_dias_restantes)
            
            # 8. Reorganizar columnas para la vista final (Incluyendo Conductor)
            orden_columnas = [
                'Placa', 'Identificación / Conductor', 'Fecha Registro', 
                'Compañía usuaria', 'Número Tipo Documento', 'Transito', 
                'Fecha de Báscula', 'Limite', 'Vencimiento 5 DIAS HABILES', 'Días restantes'
            ]
            orden_columnas = [col for col in orden_columnas if col in df_filtrado.columns]
            df_final = df_filtrado[orden_columnas].fillna('')

            # 9. Aplicar Colores (Semaforización)
            def apply_row_colors(row):
                try:
                    dias_val = row['Días restantes']
                    if dias_val == '' or pd.isna(dias_val):
                        return [''] * len(row)
                    dias = float(dias_val)
                    if dias <= 0:
                        return ['background-color: #d32f2f; color: white;'] * len(row) # Rojo (Vencidos / Vencen hoy)
                    elif 1 <= dias <= 2:
                        return ['background-color: #fbc02d; color: black;'] * len(row)  # Amarillo (Próximos a vencer)
                    else:
                        return ['background-color: #388e3c; color: white;'] * len(row)  # Verde (A tiempo)
                except:
                    pass
                return [''] * len(row)

            # --- Interfaz de Resultados ---
            # Calcular KPIs de forma segura
            def parse_dias(val):
                try:
                    return float(val)
                except:
                    return 999.0

            dias_series = df_final['Días restantes'].apply(parse_dias)
            vencidos = (dias_series <= 0).sum()
            riesgo = ((dias_series >= 1) & (dias_series <= 2)).sum()
            a_tiempo = (dias_series >= 3).sum()
            
            st.markdown("### 📊 Resumen de Operaciones")
            col1, col2, col3 = st.columns(3)
            col1.metric("🔴 Vencidos o Vencen Hoy", int(vencidos))
            col2.metric("🟡 Próximos a Vencer (1-2 días)", int(riesgo))
            col3.metric("🟢 A Tiempo (>= 3 días)", int(a_tiempo))
            
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
