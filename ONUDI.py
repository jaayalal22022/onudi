import streamlit as st
import pandas as pd
import numpy as np
import re
import plotly.express as px

st.set_page_config(page_title="ENESEM", page_icon="📊", layout="wide")

st.title("📊Visualizador ONUDI - ENESEM")
st.markdown("""
Esta herramienta consolida los tabulados de Manufactura y Minería, y permite analizar 
la evolución de las variables por cada código CIIU.
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
        'Cuad.11': 'PRODUCCION ',
        'Cuad.17': 'VALOR AGREGADO ',
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
archivo_subido = st.sidebar.file_uploader("Suba el archivo Excel (.xls o .xlsx)", type=['xls', 'xlsx'])

if archivo_subido:
    with st.spinner("Procesando matriz..."):
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
                
                # --- SOLUCIÓN AL ERROR: Forzar todo a string para evitar conflictos str vs float ---
                df_consolidado['Valor'] = pd.to_numeric(df_consolidado['Valor'].astype(str).str.replace(',', '').str.replace('*', ''), errors='coerce')
                df_consolidado['CIIU'] = df_consolidado['CIIU'].astype(str).str.replace('.0', '', regex=False).str.strip()
                df_consolidado['Año'] = df_consolidado['Año'].astype(str).str.replace('.0', '', regex=False).str.strip()
                df_consolidado['INDUSTRIA'] = df_consolidado['INDUSTRIA'].astype(str).str.strip()
                df_consolidado['Sector'] = df_consolidado['Sector'].astype(str).str.strip()
                df_consolidado['Variable'] = df_consolidado['Variable'].astype(str).str.strip()
                
                # Eliminar filas basura que pudieron generarse al final del Excel en la nube
                df_consolidado = df_consolidado[df_consolidado['CIIU'] != 'nan']
                
                df_consolidado = df_consolidado.sort_values(by=['Año', 'CIIU'])
                
                st.sidebar.success("📊Base cargada correctamente")
                
                # ==========================================
                # Panel de Control de Filtros (Sidebar)
                # ==========================================
                st.sidebar.header("🎯 Filtros de Visualización")
                
                # Asegurar que las listas desplegables se ordenen correctamente casteando a str
                opciones_sector = sorted([str(x) for x in df_consolidado['Sector'].unique()])
                sector_sel = st.sidebar.selectbox("1. Seleccione el Sector", options=opciones_sector)
                df_filtrado_sec = df_consolidado[df_consolidado['Sector'] == sector_sel]
                
                opciones_variable = sorted([str(x) for x in df_filtrado_sec['Variable'].unique()])
                variable_sel = st.sidebar.selectbox("2. Seleccione la Variable", options=opciones_variable)
                df_filtrado_var = df_filtrado_sec[df_filtrado_sec['Variable'] == variable_sel]
                
                # Crear etiqueta combinada
                df_filtrado_var['CIIU_Etiqueta'] = df_filtrado_var['CIIU'] + " - " + df_filtrado_var['INDUSTRIA']
                
                # --- SOLUCIÓN AL ERROR: Ordenar usando una comprensión de listas en formato string ---
                opciones_ciiu = sorted([str(x) for x in df_filtrado_var['CIIU_Etiqueta'].unique() if str(x) != 'nan - nan'])
                
                # Por defecto seleccionamos las primeras 3
                ciiu_sel = st.sidebar.multiselect(
                    "3. Seleccione Códigos CIIU a comparar", 
                    options=opciones_ciiu,
                    default=opciones_ciiu[:3] if len(opciones_ciiu) >= 3 else opciones_ciiu
                )
                
                df_final_ui = df_filtrado_var[df_filtrado_var['CIIU_Etiqueta'].isin(ciiu_sel)]
                
                # ==========================================
                # Pestañas de la Interfaz Principal
                # ==========================================
                tab1, tab2 = st.tabs(["📈 Evolución Anual", "📋 Datos Consolidados"])
                
                with tab1:
                    st.subheader(f"Evolución de: {variable_sel}")
                    
                    if not df_final_ui.empty:
                        # ==========================================
                        # 1. GRÁFICO PRINCIPAL (Los seleccionados)
                        # ==========================================
                        fig = px.line(
                            df_final_ui, 
                            x='Año', 
                            y='Valor', 
                            color='CIIU_Etiqueta',
                            markers=True,
                            labels={'Valor': 'Cantidad / Valor Registrado', 'Año': 'Año de Referencia', 'CIIU_Etiqueta': 'Código CIIU / Rama'},
                            title=f"Tendencias Principales - Sector {sector_sel}"
                        )
                        
                        fig.update_layout(
                            hovermode="x unified",
                            legend=dict(orientation="h", yanchor="bottom", y=-0.5, xanchor="left", x=0),
                            xaxis=dict(type='category')
                        )
                        
                        st.plotly_chart(fig, use_container_width=True)
                        
                        # ==========================================
                        # 2. DESPLEGABLE CON EL RESTO DE CIIU (No seleccionados)
                        # ==========================================
                        with st.expander("👀 Explorar esta variable en el RESTO de industrias (No seleccionadas)"):
                            st.markdown(f"Esta vista te permite analizar rápidamente la tendencia de **{variable_sel}** en las demás ramas industriales que no están en el gráfico principal.")
                            
                            # Filtramos los datos para excluir (~) los CIIU que el usuario ya seleccionó
                            df_otros_ciiu = df_filtrado_var[~df_filtrado_var['CIIU_Etiqueta'].isin(ciiu_sel)]
                            
                            # Obtenemos la lista única de esos CIIU sobrantes
                            otros_ciiu_lista = sorted([str(x) for x in df_otros_ciiu['CIIU_Etiqueta'].unique() if str(x) != 'nan - nan'])
                            
                            if not df_otros_ciiu.empty and len(otros_ciiu_lista) > 0:
                                # Creamos dos columnas
                                col1, col2 = st.columns(2)
                                
                                # Iteramos sobre cada CIIU que no fue seleccionado
                                for idx, ciiu_loop in enumerate(otros_ciiu_lista):
                                    df_mini = df_otros_ciiu[df_otros_ciiu['CIIU_Etiqueta'] == ciiu_loop]
                                    
                                    if not df_mini.empty:
                                        fig_mini = px.line(
                                            df_mini, 
                                            x='Año', 
                                            y='Valor', 
                                            markers=True,
                                            title=f"{ciiu_loop}"
                                        )
                                        
                                        # Ajustes visuales
                                        fig_mini.update_layout(
                                            hovermode="x unified",
                                            xaxis=dict(type='category'),
                                            margin=dict(l=10, r=10, t=40, b=10),
                                            title_font=dict(size=12) # Letra más pequeña para que quepa bien el texto de la industria
                                        )
                                        
                                        # Le damos un color neutro (gris) para diferenciarlos psicológicamente de los principales
                                        fig_mini.update_traces(line_color='#8C8C8C') 
                                        
                                        # Alternamos entre columna 1 y 2
                                        if idx % 2 == 0:
                                            col1.plotly_chart(fig_mini, use_container_width=True)
                                        else:
                                            col2.plotly_chart(fig_mini, use_container_width=True)
                            else:
                                st.info("Has seleccionado todos los CIIU disponibles para esta variable, no hay más industrias que mostrar.")
                    else:
                        st.warning("Por favor, selecciona al menos un código CIIU en el panel izquierdo para generar el gráfico.")


                        
                with tab2:
                    st.subheader("Registros Consolidados Filtrados")
                    st.dataframe(df_final_ui[['CIIU', 'INDUSTRIA', 'Año', 'Variable', 'Valor']], use_container_width=True)
                    
                    st.markdown("---")
                    csv = df_consolidado.to_csv(index=False).encode('utf-8')
                    st.download_button(
                        label="📥 Descargar Base en CSV",
                        data=csv,
                        file_name="Base_Consolidada_ENESEM.csv",
                        mime="text/csv",
                    )
                    
        except Exception as e:
            st.error(f"Ocurrió un error al procesar el Excel: {e}")
else:
    st.info("Por favor, suba el archivo Excel original en el panel izquierdo para comenzar.")

