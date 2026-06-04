import streamlit as st
import pandas as pd
import numpy as np
import re
import plotly.express as px

st.set_page_config(page_title="Analytics ENESEM", page_icon="📊", layout="wide")

st.title("📊 Consolidador y Visor Evolutivo ENESEM")
st.markdown("""
Esta herramienta consolida los tabulados de Manufactura y Minería, y permite analizar 
la evolución temporal de las variables por cada código CIIU de forma interactiva.
""")


# ==========================================
# Funciones de Procesamiento
# ==========================================
@st.cache_data
def procesar_mineria_df(df_raw):
    mask_ciiu = df_raw.apply(lambda x: x.astype(str).str.contains('CIIU', case=False, na=False)).any(axis=1)
    mask_ind = df_raw.apply(lambda x: x.astype(str).str.contains('INDUSTRIA', case=False, na=False)).any(axis=1)

    try:
        fila_anios_idx = df_raw[mask_ciiu & mask_ind].index[0]
        fila_vars_idx = df_raw[
            df_raw.apply(lambda x: x.astype(str).str.contains('NUMERO DE ESTABLECIMIENTOS', case=False, na=False)).any(
                axis=1)].index[0]
    except IndexError:
        return pd.DataFrame()

    macro_variables = df_raw.iloc[fila_vars_idx].replace('', np.nan).ffill()
    anios_raw = df_raw.iloc[fila_anios_idx]

    columnas_finales = []
    for macro, anio in zip(macro_variables, anios_raw):
        macro_str = str(macro).strip()
        anio_str = str(anio).strip().upper()

        if 'CIIU' in anio_str:
            columnas_finales.append('CIIU')
        elif 'INDUSTRIA' in anio_str:
            columnas_finales.append('INDUSTRIA')
        elif re.search(r'\d{4}', anio_str):
            anio_limpio = re.search(r'\d{4}', anio_str).group()
            columnas_finales.append(f"{macro_str}|{anio_limpio}")
        else:
            columnas_finales.append("ELIMINAR")

    df_datos = df_raw.iloc[fila_anios_idx + 1:].copy()
    df_datos.columns = columnas_finales

    df_datos = df_datos.loc[:, ~df_datos.columns.duplicated()]
    df_datos = df_datos.loc[:, df_datos.columns != "ELIMINAR"]

    if 'CIIU' not in df_datos.columns or 'INDUSTRIA' not in df_datos.columns:
        return pd.DataFrame()

    df_datos = df_datos.dropna(subset=['CIIU', 'INDUSTRIA'])

    df_melted = df_datos.melt(id_vars=['CIIU', 'INDUSTRIA'], var_name='Variable_Año', value_name='Valor')
    df_melted[['Variable', 'Año']] = df_melted['Variable_Año'].str.split('|', expand=True)
    df_melted.drop('Variable_Año', axis=1, inplace=True)
    df_melted['Sector'] = 'Minería'

    return df_melted


@st.cache_data
def procesar_manufactura_df(df_raw, nombre_pestana):
    mapa_variables = {
        'Cuad.02': 'NUMERO DE EMPRESAS',
        'Cuad.03': 'NUMERO DE PERSONAS OCUPADAS',
        'Cuad.04': 'NUMERO DE EMPLEADOS, ADMINISTRATIVOS Y OBREROS',
        'Cuad.05': 'SUELDOS Y SALARIOS PAGADOS',
        'Cuad.11': 'PRODUCCION (Output at basic prices)',
        'Cuad.17': 'VALOR AGREGADO (Value added at basic prices)',
        'Cuad.21': 'FORMACION BRUTA DE CAPITAL FIJO',
        'Cuad.31': 'NUMERO DE MUJERES EMPLEADAS, ADMINISTRATIVAS Y OBRERAS'
    }

    codigo_cuadro = re.search(r'Cuad\.\d{2}', nombre_pestana)
    if not codigo_cuadro:
        return pd.DataFrame()

    variable_actual = mapa_variables.get(codigo_cuadro.group(), 'Variable Desconocida')

    mask_ciiu = df_raw.apply(lambda x: x.astype(str).str.contains('CIIU', case=False, na=False)).any(axis=1)
    mask_ind = df_raw.apply(lambda x: x.astype(str).str.contains('INDUSTRIA', case=False, na=False)).any(axis=1)

    try:
        fila_encabezado = df_raw[mask_ciiu & mask_ind].index[0]
    except IndexError:
        return pd.DataFrame()

    df_datos = df_raw.iloc[fila_encabezado:].copy()
    df_datos.columns = [str(c).strip().replace('\n', ' ') for c in df_datos.iloc[0]]
    df_datos = df_datos[1:]
    df_datos = df_datos.loc[:, ~df_datos.columns.duplicated()]

    col_ciiu = next((c for c in df_datos.columns if 'CIIU' in str(c).upper()), None)
    col_industria = next((c for c in df_datos.columns if 'INDUSTRIA' in str(c).upper()), None)

    if not col_ciiu or not col_industria:
        return pd.DataFrame()

    cols_a_mantener = [col_ciiu, col_industria]
    cols_anios = [col for col in df_datos.columns if
                  str(col).strip().replace('.0', '').isdigit() and len(str(col).strip().replace('.0', '')) == 4]

    df_datos = df_datos[cols_a_mantener + cols_anios].dropna(subset=[col_ciiu])

    df_melted = df_datos.melt(id_vars=cols_a_mantener, value_vars=cols_anios, var_name='Año', value_name='Valor')

    df_melted.rename(columns={col_ciiu: 'CIIU', col_industria: 'INDUSTRIA'}, inplace=True)
    df_melted['Variable'] = variable_actual
    df_melted['Sector'] = 'Manufactura'
    df_melted['Año'] = df_melted['Año'].astype(str).str.replace('.0', '', regex=False)

    return df_melted


# ==========================================
# Carga de Datos y Consolidación
# ==========================================
st.sidebar.header("📁 Carga de Datos")
archivo_subido = st.sidebar.file_uploader("Sube el archivo Excel (.xls o .xlsx)", type=['xls', 'xlsx'])

if archivo_subido:
    with st.spinner("Procesando matriz estructural..."):
        try:
            diccionario_hojas = pd.read_excel(archivo_subido, sheet_name=None, header=None)

            df_mineria = pd.DataFrame()
            dfs_manufactura = []

            for nombre_pestana, df_raw in diccionario_hojas.items():
                if 'metadatos' in nombre_pestana.lower():
                    continue
                if 'Minería' in nombre_pestana:
                    df_mineria = procesar_mineria_df(df_raw)
                elif 'Cuad.' in nombre_pestana:
                    df_temp = procesar_manufactura_df(df_raw, nombre_pestana)
                    if not df_temp.empty:
                        dfs_manufactura.append(df_temp)

            df_manufactura = pd.concat(dfs_manufactura, ignore_index=True) if dfs_manufactura else pd.DataFrame()

            if not df_mineria.empty or not df_manufactura.empty:
                df_consolidado = pd.concat([df_mineria, df_manufactura], ignore_index=True)

                # Estandarización de tipos de datos
                df_consolidado['Valor'] = pd.to_numeric(
                    df_consolidado['Valor'].astype(str).str.replace(',', '').str.replace('*', ''), errors='coerce')
                df_consolidado['CIIU'] = df_consolidado['CIIU'].astype(str).str.replace('.0', '',
                                                                                        regex=False).str.strip()
                df_consolidado['Año'] = df_consolidado['Año'].astype(str).str.replace('.0', '', regex=False).str.strip()
                df_consolidado = df_consolidado.sort_values(by=['Año', 'CIIU'])

                st.sidebar.success("📊 ¡Base estructurada correctamente!")

                # ==========================================
                # Panel de Control de Filtros (Sidebar)
                # ==========================================
                st.sidebar.header("🎯 Filtros de Visualización")

                sector_sel = st.sidebar.selectbox("1. Selecciona el Sector", options=df_consolidado['Sector'].unique())
                df_filtrado_sec = df_consolidado[df_consolidado['Sector'] == sector_sel]

                variable_sel = st.sidebar.selectbox("2. Selecciona la Variable",
                                                    options=df_filtrado_sec['Variable'].unique())
                df_filtrado_var = df_filtrado_sec[df_filtrado_sec['Variable'] == variable_sel]

                # Crear etiqueta combinada 'CIIU - Nombre' para facilitar la selección al usuario
                df_filtrado_var['CIIU_Etiqueta'] = df_filtrado_var['CIIU'] + " - " + df_filtrado_var[
                    'INDUSTRIA'].astype(str)
                opciones_ciiu = sorted(df_filtrado_var['CIIU_Etiqueta'].unique())

                # Por defecto seleccionamos las primeras 3 para que el gráfico no nazca vacío ni saturado
                ciiu_sel = st.sidebar.multiselect(
                    "3. Selecciona Códigos CIIU a comparar",
                    options=opciones_ciiu,
                    default=opciones_ciiu[:3] if len(opciones_ciiu) >= 3 else opciones_ciiu
                )

                # Filtrado final para visualización
                df_final_ui = df_filtrado_var[df_filtrado_var['CIIU_Etiqueta'].isin(ciiu_sel)]

                # ==========================================
                # Pestañas de la Interfaz Principal
                # ==========================================
                tab1, tab2 = st.tabs(["📈 Evolución Temporal", "📋 Datos Consolidados"])

                with tab1:
                    st.subheader(f"Evolución de: {variable_sel}")

                    if not df_final_ui.empty:
                        # Generación del gráfico de líneas interactivo con Plotly
                        fig = px.line(
                            df_final_ui,
                            x='Año',
                            y='Valor',
                            color='CIIU_Etiqueta',
                            markers=True,
                            labels={'Valor': 'Cantidad / Valor Registrado', 'Año': 'Año de Referencia',
                                    'CIIU_Etiqueta': 'Código CIIU / Rama'},
                            title=f"Tendencia Temporal - Sector {sector_sel}"
                        )

                        # Mejoras visuales al gráfico
                        fig.update_layout(
                            hovermode="x unified",
                            legend=dict(orientation="h", yanchor="bottom", y=-0.5, xanchor="left", x=0),
                            xaxis=dict(type='category')  # Forza a que los años se traten de forma discreta
                        )

                        st.plotly_chart(fig, use_container_width=True)
                    else:
                        st.warning(
                            "Por favor, selecciona al menos un código CIIU en el panel izquierdo para generar el gráfico.")

                with tab2:
                    st.subheader("Registros Consolidados Filtrados")
                    st.dataframe(df_final_ui[['CIIU', 'INDUSTRIA', 'Año', 'Variable', 'Valor']],
                                 use_container_width=True)

                    # Descarga de la base de datos completa o filtrada
                    st.markdown("---")
                    csv = df_consolidado.to_csv(index=False).encode('utf-8')
                    st.download_button(
                        label="📥 Descargar Base Completa en CSV",
                        data=csv,
                        file_name="Base_Consolidada_ENESEM.csv",
                        mime="text/csv",
                    )

        except Exception as e:
            st.error(f"Ocurrió un error al procesar el Excel: {e}")
else:
    st.info("Por favor, sube el archivo Excel original en el panel izquierdo para comenzar.")
