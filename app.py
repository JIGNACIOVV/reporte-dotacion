import streamlit as st
import pandas as pd
import numpy as np
import io
import os

# ==========================================
# ⚙️ CONFIGURACIÓN DE LA PÁGINA
# ==========================================
st.set_page_config(page_title="Balance de Dotación", layout="wide")

st.title("📊 Tablero de Control de Dotación Diario")
st.markdown("Visualización actualizada del balance entre el personal requerido y el contratado.")

# ==========================================
# 🧠 LÓGICA
# ==========================================
def limpiar_texto(serie):
    return serie.astype(str).str.strip().str.upper()

def clasificar_jornada(texto):
    texto = str(texto).lower()
    if 'peak' in texto: return 'PEAK'
    if 'full' in texto or 'completa' in texto: return 'FT'
    if 'part' in texto or 'parcial' in texto or 'media' in texto: return 'PT'
    return 'FT'

def estilo_balance(val):
    if val < 0:
        return 'color: red; font-weight: bold'
    elif val > 0:
        return 'color: green'
    return ''

# ==========================================
# 🚀 PROCESAMIENTO AUTOMÁTICO
# ==========================================
# Verifica que los archivos existan en la carpeta de la nube
if os.path.exists("Meta.xlsx") and os.path.exists("Buk.xlsx"):
    try:
        # 1. Procesar Meta
        df_metas = pd.read_excel("Meta.xlsx")
        if 'PK' in df_metas.columns:
            df_metas = df_metas.rename(columns={'PK': 'PEAK'})
        
        for c in ['FT', 'PT', 'PEAK']:
            df_metas[c] = pd.to_numeric(df_metas[c], errors='coerce').fillna(0)
            
        df_metas['Terminal'] = limpiar_texto(df_metas['Terminal'])
        df_metas['Total_Req'] = df_metas['FT'] + df_metas['PT'] + df_metas['PEAK']

        # 2. Procesar BUK
        df_buk = pd.read_excel("Buk.xlsx", header=5)
        df_buk.columns = df_buk.columns.str.strip()
        
        COL_TERMINAL_BUK = "Nombre de Recintos" 
        COL_JORNADA_BUK = "Tipo_jornada"

        df_buk['Terminal_Norm'] = limpiar_texto(df_buk[COL_TERMINAL_BUK])
        df_buk['Clasificacion'] = df_buk[COL_JORNADA_BUK].apply(clasificar_jornada)

        resumen_real = df_buk.pivot_table(
            index='Terminal_Norm',
            columns='Clasificacion',
            aggfunc='size',
            fill_value=0
        ).reset_index()

        for col in ['FT', 'PT', 'PEAK']:
            if col not in resumen_real.columns: resumen_real[col] = 0
            
        resumen_real = resumen_real.rename(columns={
            'FT': 'R_FT', 'PT': 'R_PT', 'PEAK': 'R_PEAK', 'Terminal_Norm': 'Terminal'
        })

        # 3. Merge y Cálculos
        reporte = pd.merge(df_metas, resumen_real, on='Terminal', how='left').fillna(0)
        reporte['R_Total'] = reporte['R_FT'] + reporte['R_PT'] + reporte['R_PEAK']
        
        reporte['B_FT'] = reporte['R_FT'] - reporte['FT']
        reporte['B_PT'] = reporte['R_PT'] - reporte['PT']
        reporte['B_PEAK'] = reporte['R_PEAK'] - reporte['PEAK']
        reporte['B_Total'] = reporte['R_Total'] - reporte['Total_Req']

        cols_finales = [
            'Terminal', 
            'FT', 'PT', 'PEAK', 'Total_Req',       
            'R_FT', 'R_PT', 'R_PEAK', 'R_Total',   
            'B_FT', 'B_PT', 'B_PEAK', 'B_Total'    
        ]
        reporte = reporte[cols_finales]

        # ==========================================
        # 🖥️ VISUALIZACIÓN EN PANTALLA
        # ==========================================
        # Filtros opcionales
        filtro_terminal = st.multiselect("🔍 Filtrar por Terminal:", options=reporte['Terminal'].unique())
        
        if filtro_terminal:
            df_view = reporte[reporte['Terminal'].isin(filtro_terminal)]
        else:
            df_view = reporte

        cols_balance = ['B_FT', 'B_PT', 'B_PEAK', 'B_Total']
        st.dataframe(
            df_view.style.map(estilo_balance, subset=cols_balance)
                   .format("{:.0f}", subset=df_view.columns[1:]),
            use_container_width=True,
            height=600
        )

        # ==========================================
        # 📥 DESCARGA EXCEL (Para los compañeros que quieran el archivo)
        # ==========================================
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            reporte.replace(0, np.nan).to_excel(writer, sheet_name='Balance', index=False, startrow=2, header=False)
            workbook = writer.book
            worksheet = writer.sheets['Balance']

            fmt_terminal = workbook.add_format({'bg_color': '#F4B084', 'bold': True, 'border': 1, 'align': 'center'})
            fmt_req = workbook.add_format({'bg_color': '#BDD7EE', 'bold': True, 'border': 1, 'align': 'center'})
            fmt_real = workbook.add_format({'bg_color': '#FFE699', 'bold': True, 'border': 1, 'align': 'center'})
            fmt_bal = workbook.add_format({'bg_color': '#C6E0B4', 'bold': True, 'border': 1, 'align': 'center'})
            fmt_sub = workbook.add_format({'bold': True, 'align': 'center', 'border': 1, 'valign': 'vcenter'})
            fmt_neg = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'})

            worksheet.merge_range('A1:A2', 'Terminal', fmt_terminal)
            worksheet.merge_range('B1:E1', 'REQUERIDO', fmt_req)
            worksheet.merge_range('F1:I1', 'CONTRATADOS', fmt_real)
            worksheet.merge_range('J1:M1', 'BALANCE', fmt_bal)

            subs = ['FT', 'PT', 'PK', 'Total', 'FT', 'PT', 'PK', 'Total', 'FT', 'PT', 'PK', 'TOTAL']
            for i, t in enumerate(subs):
                worksheet.write(1, i + 1, t, fmt_sub)

            worksheet.set_column('A:A', 25)
            worksheet.set_column('B:M', 8)
            
            ultima_fila = len(reporte) + 2
            worksheet.conditional_format(2, 9, ultima_fila, 12, {
                'type': 'cell', 'criteria': '<', 'value': 0, 'format': fmt_neg
            })

        st.download_button(
            label="📥 Descargar Reporte en Excel",
            data=buffer,
            file_name="Reporte_Balance_Diario.xlsx",
            mime="application/vnd.ms-excel"
        )

    except Exception as e:
        st.error(f"Error procesando los datos: {e}. Revisa que el formato de Buk y Meta no haya cambiado.")
else:
    st.warning("⚠️ El administrador aún no ha cargado los datos de hoy. Faltan los archivos 'Meta.xlsx' o 'Buk.xlsx'.")
