import streamlit as st
import pandas as pd
import datetime
import numpy as np

# Configuración de la página
st.set_page_config(page_title="Control de Bloqueo | Zona Franca", page_icon="🚦", layout="wide")

# Estilos CSS personalizados para simular el tema corporativo
st.markdown("""
    <style>
    .main-header {
        color: #1f4e3d; /* Verde corporativo oscuro */
        font-weight: bold;
    }
    .sub-header {
        color: #4a4a4a;
    }
    .stDataFrame {
        font-size: 14px;
    }
    </style>
    """, unsafe_allow_html=True)

st.markdown("<h1 class='main-header'>🚦 Módulo de Control de Bloqueos (Reporte PW)</h1>", unsafe_allow_html=True)
st.markdown("<p class='sub-header'>Carga el archivo Excel <b>ReportePW</b> descargado de PICIZ para automatizar el filtrado y visualizar las alertas de vencimiento.</p>", unsafe_allow_html=True)

# --- Función auxiliar para eliminar ceros decimales (.000000) ---
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

uploaded_file = st.file_uploader("Sube el archivo Excel (ReportePW.xlsx)", type=["xlsx", "xls"])

if uploaded_file is not None:
    with st.spinner("Procesando y limpiando datos..."):
        try:
            # 1. Cargar el archivo sin encabezados para encontrar la fila correcta
            df_raw = pd.read_excel(uploaded_file, header=None, dtype=str)
            
            # 2. Buscar la fila que contiene las cabeceras reales (ej. "NOMBRE COMPANIA")
            header_row_index = -1
            for i, row in df_raw.iterrows():
                row_str = " ".join(str(val) for val in row.values)
                if "NOMBRE COMPANIA" in row_str or "PLACA" in row_str:
                    header_row_index = i
                    break
            
            if header_row_index == -1:
                st.error("❌ No se encontró la fila de encabezados esperada (ej. 'NOMBRE COMPANIA'). Verifica el formato del archivo.")
                st.stop()
            
            # 3. Leer el archivo con la fila de encabezado correcta
            df = pd.read_excel(uploaded_file, header=header_row_index, dtype=str)
            
            # Limpiar nombres de columnas
            df.columns = df.columns.str.strip().str.replace(r'\r\n', '', regex=True)
            
            # 4. Seleccionar las columnas requeridas
            columnas_esperadas = {
                'NOMBRE COMPANIA': 'Compañia usuaria',
                'PLACA': 'Placa',
                'FECHA REGISTRO': 'Fecha Registro',
                'TIPO INGRESO': 'Tipo Ingreso', 
                'NUM DEL DOC. ADUANERO': 'Número Tipo Documento',
                'TRANSITO': 'Transito',
                'FECHA BASCULA': 'Fecha de Báscula'
            }
            
            columnas_faltantes = [col for col in columnas_esperadas.keys() if col not in df.columns]
            if columnas_faltantes:
                st.warning(f"⚠️ Faltan las siguientes columnas en el archivo: {', '.join(columnas_faltantes)}")
                columnas_existentes = {k: v for k, v in columnas_esperadas.items() if k in df.columns}
                df_filtrado = df[list(columnas_existentes.keys())].rename(columns=columnas_existentes)
            else:
                df_filtrado = df[list(columnas_esperadas.keys())].rename(columns=columnas_esperadas)
            
            # 5. Filtrar solo filas donde "Tipo Ingreso" sea FORMULARIO 
            if 'Tipo Ingreso' in df_filtrado.columns:
                df_filtrado = df_filtrado[df_filtrado['Tipo Ingreso'].str.contains('FORMULARIO', na=False, case=False)]
                df_filtrado = df_filtrado.drop(columns=['Tipo Ingreso'])
            
            # Limpiar Valores Nulos en Placa
            df_filtrado = df_filtrado.dropna(subset=['Placa'])
            df_filtrado['Placa'] = df_filtrado['Placa'].astype(str).str.strip()
            df_filtrado = df_filtrado[df_filtrado['Placa'] != '']

            # 6. Limpieza de ceros y decimales en Documento y Tránsito
            for col in ['Número Tipo Documento', 'Transito']:
                if col in df_filtrado.columns:
                    df_filtrado[col] = df_filtrado[col].apply(limpiar_numeros)

            # 7. Limpieza y formateo de fechas
            for col in ['Fecha Registro', 'Fecha de Báscula']:
                if col in df_filtrado.columns:
                    df_filtrado[col] = pd.to_datetime(df_filtrado[col], errors='coerce').dt.date

            # Eliminar duplicados idénticos
            columnas_dedup = [col for col in ['Placa', 'Fecha Registro', 'Compañia usuaria', 'Número Tipo Documento', 'Transito', 'Fecha de Báscula'] if col in df_filtrado.columns]
            df_filtrado = df_filtrado.drop_duplicates(subset=columnas_dedup, keep='first')

            # 8. Cálculos de Vencimiento y Días Restantes basados en Fecha de Báscula
            df_filtrado['Limite'] = 5
            
            def calcular_vencimiento(row):
                fecha_base = row['Fecha de Báscula'] if pd.notna(row['Fecha de Báscula']) and str(row['Fecha de Báscula']) != 'NaT' else row['Fecha Registro']
                if pd.isna(fecha_base) or str(fecha_base) == 'NaT':
                    return None
                try:
                    fecha_venc = np.busday_offset(np.datetime64(fecha_base), 5, roll='forward')
                    return pd.to_datetime(fecha_venc).date()
                except:
                    return None
                    
            df_filtrado['Vencimiento 5 DIAS HABILES'] = df_filtrado.apply(calcular_vencimiento, axis=1)
            
            hoy = datetime.date.today() 
            
            def calcular_dias_restantes(fecha_venc):
                if pd.isna(fecha_venc) or str(fecha_venc) == 'NaT':
                    return None
                delta = fecha_venc - hoy
                return delta.days
                
            df_filtrado['Días restantes'] = df_filtrado['Vencimiento 5 DIAS HABILES'].apply(calcular_dias_restantes)
            
            # Reorganizar columnas
            orden_columnas = ['Placa', 'Fecha Registro', 'Compañia usuaria', 'Número Tipo Documento', 'Transito', 'Fecha de Báscula', 'Limite', 'Vencimiento 5 DIAS HABILES', 'Días restantes']
            orden_columnas = [col for col in orden_columnas if col in df_filtrado.columns]
            df_final = df_filtrado[orden_columnas].copy()
            
            # Convertir Días restantes explícitamente a numérico para evitar errores de tipo
            df_final['Días restantes'] = pd.to_numeric(df_final['Días restantes'], errors='coerce')

            # --- 9. Lógica de Semaforización ---
            def apply_row_colors(row):
                try:
                    dias_val = row['Días restantes']
                    if pd.isna(dias_val):
                        return [''] * len(row)
                    dias = float(dias_val)
                    if dias <= 0:
                        return ['background-color: #d32f2f; color: white;'] * len(row) # Rojo (Vencidos / Hoy)
                    elif 1 <= dias <= 2:
                        return ['background-color: #fff59d; color: black;'] * len(row) # Amarillo (Riesgo)
                    elif dias >= 3:
                        return ['background-color: #a5d6a7; color: black;'] * len(row) # Verde (A tiempo)
                except:
                    pass
                return [''] * len(row)

            # KPIs precisos con datos numéricos limpios
            dias_series = df_final['Días restantes']
            vencidos = (dias_series <= 0).sum()
            riesgo = ((dias_series >= 1) & (dias_series <= 2)).sum()
            a_tiempo = (dias_series >= 3).sum()
            
            st.markdown("### 📊 Resumen de Operaciones")
            col1, col2, col3 = st.columns(3)
            col1.metric("🔴 Vencidos o Vencen Hoy", int(vencidos))
            col2.metric("🟡 Próximos a Vencer (1-2 días)", int(riesgo))
            col3.metric("🟢 A Tiempo (>= 3 días)", int(a_tiempo))
            
            st.markdown("### 📋 Panel de Control de Ingresos")
            
            # Aplicar estilo antes de rellenar vacíos con texto para preservar tipos numéricos
            styled_df = df_final.style.apply(apply_row_colors, axis=1)
            st.dataframe(styled_df, use_container_width=True, height=500, hide_index=True)
            
            # 10. Descarga
            st.markdown("### 💾 Exportar Datos")
            
            def convert_df_to_excel(df_styled):
                from io import BytesIO
                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_styled.to_excel(writer, index=False, sheet_name='Control_Bloqueo')
                return output.getvalue()

            excel_data = convert_df_to_excel(styled_df)
            
            st.download_button(
                label="📥 Descargar Reporte Formateado (Excel)",
                data=excel_data,
                file_name=f"Control_Bloqueos_{datetime.date.today()}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
        except Exception as e:
            st.error(f"❌ Ocurrió un error al procesar el archivo: {e}")

else:
    st.info("Esperando archivo... Por favor, carga el `ReportePW.xlsx` para comenzar.")
