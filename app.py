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
    /* Estilizar la tabla cuando se usa st.dataframe */
    .stDataFrame {
        font-size: 14px;
    }
    </style>
    """, unsafe_allow_html=True)

st.markdown("<h1 class='main-header'>🚦 Módulo de Control de Bloqueos (Reporte PW)</h1>", unsafe_allow_html=True)
st.markdown("<p class='sub-header'>Carga el archivo Excel <b>ReportePW</b> descargado de PICIZ para automatizar el filtrado y visualizar las alertas de vencimiento.</p>", unsafe_allow_html=True)

uploaded_file = st.file_uploader("Sube el archivo Excel (ReportePW.xlsx)", type=["xlsx", "xls"])

if uploaded_file is not None:
    with st.spinner("Procesando y limpiando datos..."):
        try:
            # 1. Cargar el archivo sin encabezados para encontrar la fila correcta
            df_raw = pd.read_excel(uploaded_file, header=None)
            
            # 2. Buscar la fila que contiene las cabeceras reales (ej. "NOMBRE COMPANIA")
            header_row_index = -1
            for i, row in df_raw.iterrows():
                if "NOMBRE COMPANIA" in row.values or "NOMBRE COMPANIA" in str(row.values):
                    header_row_index = i
                    break
            
            if header_row_index == -1:
                st.error("❌ No se encontró la fila de encabezados esperada (ej. 'NOMBRE COMPANIA'). Verifica el formato del archivo.")
                st.stop()
            
            # 3. Leer el archivo con la fila de encabezado correcta
            df = pd.read_excel(uploaded_file, header=header_row_index)
            
            # Limpiar nombres de columnas (quitar espacios extra y saltos de línea)
            df.columns = df.columns.str.strip().str.replace('\r\n', '', regex=True)
            
            # 4. Seleccionar las columnas requeridas (mapeo a lo solicitado)
            columnas_esperadas = {
                'NOMBRE COMPANIA': 'Compañia usuaria',
                'PLACA': 'Placa',
                'FECHA REGISTRO': 'Fecha Registro',
                'TIPO INGRESO': 'Tipo Ingreso', 
                'NUM DEL DOC. ADUANERO': 'Número Tipo Documento',
                'TRANSITO': 'Transito',
                'FECHA BASCULA': 'Fecha de Báscula'
            }
            
            # Verificar si las columnas esperadas existen
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
            
            # 6. Limpieza y formateo de fechas
            for col in ['Fecha Registro', 'Fecha de Báscula']:
                if col in df_filtrado.columns:
                    df_filtrado[col] = pd.to_datetime(df_filtrado[col], format='%d/%m/%Y', errors='coerce').dt.date

            # Limpiar Valores Nulos
            df_filtrado = df_filtrado.dropna(subset=['Placa'])
            
            # 7. Cálculos de Vencimiento y Días Restantes
            df_filtrado['Limite'] = 5
            
            def calcular_vencimiento(fecha_registro):
                if pd.isna(fecha_registro):
                    return None
                try:
                    # Sumar 5 días hábiles
                    fecha_venc = np.busday_offset(np.datetime64(fecha_registro), 5, roll='forward')
                    return pd.to_datetime(fecha_venc).date()
                except:
                    return None
                    
            df_filtrado['Vencimiento 5 DIAS HABILES'] = df_filtrado['Fecha Registro'].apply(calcular_vencimiento)
            
            # Usar fecha actual para el cálculo
            hoy = datetime.date.today() 
            
            def calcular_dias_restantes(fecha_venc):
                if pd.isna(fecha_venc):
                    return None
                delta = fecha_venc - hoy
                return delta.days
                
            df_filtrado['Días restantes'] = df_filtrado['Vencimiento 5 DIAS HABILES'].apply(calcular_dias_restantes)
            
            # Reorganizar columnas
            orden_columnas = ['Placa', 'Fecha Registro', 'Compañia usuaria', 'Número Tipo Documento', 'Transito', 'Fecha de Báscula', 'Limite', 'Vencimiento 5 DIAS HABILES', 'Días restantes']
            orden_columnas = [col for col in orden_columnas if col in df_filtrado.columns]
            df_final = df_filtrado[orden_columnas]
            
            # Reemplazar NaN por vacíos
            df_final = df_final.fillna('')
            
            # --- 8. Lógica de Semaforización ---
            def apply_row_colors(row):
                try:
                    dias = float(row['Días restantes'])
                    if dias <= 0:
                        return ['background-color: #d32f2f; color: white;'] * len(row)
                    elif 1 <= dias <= 2:
                        return ['background-color: #fff59d; color: black;'] * len(row) 
                    elif dias >= 3:
                        return ['background-color: #a5d6a7; color: black;'] * len(row) 
                except:
                    pass
                return [''] * len(row)

            # KPIs
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
            
            # 9. Descarga
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