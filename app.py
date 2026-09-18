# --- DENTRO DE: if uploaded_file is not None: ---

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
            df.columns = (
                df.columns.astype(str)
                .str.strip()
                .str.replace(r'[\r\n]+', ' ', regex=True)
                .str.replace(r'\s+', ' ', regex=True)
            )

            # 3. Mapeo de columnas (AHORA INCLUYE COMENTARIO Y TIPO INGRESO)
            columnas_esperadas = {
                'NOMBRE COMPANIA': 'Compañía Usuaria',
                'PLACA': 'Placa',
                'FECHA REGISTRO': 'Fecha Registro',
                'FECHA BASCULA': 'Fecha de Báscula',
                'TIPO INGRESO': 'Tipo Ingreso',
                'NUM DEL DOC. ADUANERO': 'Número Documento',
                'DOCUMENTO DE TRANSPORTE': 'Documento Transporte',
                'TRANSITO': 'Tránsito',
                'ESTADO PLANILLA': 'Estado Planilla',
                'COMENTARIO': 'Comentario'
            }

            columnas_existentes = {k: v for k, v in columnas_esperadas.items() if k in df.columns}
            df_filtrado = df[list(columnas_existentes.keys())].rename(columns=columnas_existentes)

            # 4. Filtro de tipos de ingreso: FORMULARIO + OTROS INGRESOS
            if 'Tipo Ingreso' in df_filtrado.columns:
                tipos_validos = ['FORMULARIO', 'OTROS INGRESOS']
                df_filtrado['Tipo Ingreso'] = (
                    df_filtrado['Tipo Ingreso'].astype(str).str.upper().str.strip()
                )
                df_filtrado = df_filtrado[df_filtrado['Tipo Ingreso'].isin(tipos_validos)]

            # 5. Limpieza básica
            df_filtrado = df_filtrado.dropna(subset=['Placa'])
            df_filtrado['Placa'] = df_filtrado['Placa'].astype(str).str.strip()
            df_filtrado = df_filtrado[df_filtrado['Placa'] != '']

            for col in ['Número Documento', 'Tránsito', 'Documento Transporte']:
                if col in df_filtrado.columns:
                    df_filtrado[col] = df_filtrado[col].apply(limpiar_numeros)

            for col in ['Fecha Registro', 'Fecha de Báscula']:
                if col in df_filtrado.columns:
                    df_filtrado[col] = pd.to_datetime(
                        df_filtrado[col], dayfirst=True, errors='coerce'
                    ).dt.date

            # 6. Deduplicación (incluye Tipo Ingreso para no perder OTROS INGRESOS)
            columnas_dedup = [
                col for col in [
                    'Placa', 'Fecha Registro', 'Compañía Usuaria',
                    'Número Documento', 'Tránsito', 'Fecha de Báscula',
                    'Tipo Ingreso'
                ] if col in df_filtrado.columns
            ]
            df_filtrado = df_filtrado.drop_duplicates(subset=columnas_dedup, keep='first')

            # 7. Cálculo de vencimiento
            df_filtrado['Límite'] = 5

            def calcular_vencimiento(row):
                fecha_base = (
                    row['Fecha de Báscula']
                    if pd.notna(row.get('Fecha de Báscula')) and str(row.get('Fecha de Báscula')) != 'NaT'
                    else row.get('Fecha Registro')
                )
                if pd.isna(fecha_base) or str(fecha_base) == 'NaT':
                    return None
                try:
                    fecha_venc = np.busday_offset(np.datetime64(fecha_base), 5, roll='forward')
                    return pd.to_datetime(fecha_venc).date()
                except Exception:
                    return None

            df_filtrado['Vencimiento (5 Días Hábiles)'] = df_filtrado.apply(calcular_vencimiento, axis=1)

            def calcular_dias_restantes(fecha_venc):
                if pd.isna(fecha_venc) or str(fecha_venc) == 'NaT':
                    return None
                return (fecha_venc - st.session_state.fecha_ref).days

            df_filtrado['Días Restantes'] = df_filtrado['Vencimiento (5 Días Hábiles)'].apply(calcular_dias_restantes)

            # 8. Orden final de columnas (AHORA INCLUYE Tipo Ingreso y Comentario)
            orden_columnas = [
                'Placa', 'Fecha Registro', 'Compañía Usuaria',
                'Tipo Ingreso', 'Número Documento', 'Documento Transporte',
                'Tránsito', 'Estado Planilla', 'Fecha de Báscula',
                'Límite', 'Vencimiento (5 Días Hábiles)', 'Días Restantes',
                'Comentario'
            ]
            orden_columnas = [col for col in orden_columnas if col in df_filtrado.columns]
            df_final = df_filtrado[orden_columnas].copy()

            df_final['Días Restantes'] = (
                pd.to_numeric(df_final['Días Restantes'], errors='coerce')
                .fillna(0).astype(int)
            )

            st.session_state.df_final = df_final
            st.session_state.file_processed = True
