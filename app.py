
import streamlit as st
import pandas as pd
import os
import io
import matplotlib.pyplot as plt

st.set_page_config(page_title="Análisis Técnico: Cajas vs Picking", layout="wide")
st.title("Análisis Técnico: Cajas vs Picking")

st.markdown("""
Cargue un archivo Excel de inspección para calcular cajas completas, unidades a picking y revisar técnicos activos.

- Detecta automáticamente técnicos desde la fila 16.
- Analiza datos desde la fila 21.
- Si un técnico no tiene actividad, se indica explícitamente.
- Incluye gráfica de torta + resumen numérico por técnico.
""")

archivo_inspeccion = st.file_uploader("Suba su archivo Excel de inspección:", type=[".xlsx"])

BASE_PATH = "Articulos_Filtrados_Completos.xlsx"

def cargar_base():
    if os.path.exists(BASE_PATH):
        base = pd.read_excel(BASE_PATH)
        base["Articulo, Nombre"] = base["Articulo, Nombre"].astype(str).str.strip().str.upper()
        base["Unidades/Caja"] = pd.to_numeric(base["Unidades/Caja"], errors='coerce').fillna(0).astype(int)
        return base
    return pd.DataFrame(columns=["Artículo", "Unidades/Caja", "Articulo, Nombre"])

def guardar_base(base):
    base.to_excel(BASE_PATH, index=False)

if archivo_inspeccion:
    df_crudo = pd.read_excel(archivo_inspeccion, sheet_name='LASER', header=None)
    df_datos = pd.read_excel(archivo_inspeccion, sheet_name='LASER', header=20)

    codigos = df_datos.iloc[:, 1].astype(str).str.strip().str.upper()
    unidades_caja_archivo = pd.to_numeric(df_datos.iloc[:, 3], errors='coerce')
    base = cargar_base()
    resultados = []
    tecnicos_sin_datos = []
    resumen_por_tecnico = {}

    encabezados = df_crudo.iloc[15]
    for col_index, encabezado in enumerate(encabezados):
        if isinstance(encabezado, str) and encabezado.strip().upper().startswith("TÉCNICO"):
            datos_col = pd.to_numeric(df_datos.iloc[:, col_index], errors='coerce').fillna(0)
            col_defectuosos = df_datos.columns[col_index + 1] if col_index + 1 < len(df_datos.columns) else None
            defectuosos = pd.to_numeric(df_datos[col_defectuosos], errors='coerce').fillna(0) if col_defectuosos else pd.Series([0]*len(datos_col))

            if datos_col.sum() == 0:
                tecnicos_sin_datos.append(encabezado.strip())
                continue

            tecnico = encabezado.strip()
            total_buenas = 0
            total_defectuosas = 0
            total_cajas = 0
            total_picking = 0

            for i, codigo in enumerate(codigos):
                cantidad = datos_col[i]
                defectuosa = defectuosos[i] if i < len(defectuosos) else 0
                if cantidad == 0:
                    continue

                unidades = unidades_caja_archivo[i]
                if pd.isna(unidades) or unidades == 0:
                    match = base[base["Articulo, Nombre"] == codigo]
                    if not match.empty:
                        unidades = match["Unidades/Caja"].values[0]
                else:
                    if codigo not in base["Articulo, Nombre"].values:
                        nuevo = pd.DataFrame({
                            "Artículo": [""],
                            "Unidades/Caja": [int(unidades)],
                            "Articulo, Nombre": [codigo]
                        })
                        base = pd.concat([base, nuevo], ignore_index=True)
                        guardar_base(base)

                if unidades > 0:
                    cajas = int(cantidad // unidades)
                    picking = int(cantidad % unidades)
                    total_buenas += cantidad
                    total_defectuosas += int(defectuosa)
                    total_cajas += cajas
                    total_picking += picking

                    resultados.append([tecnico, codigo, int(cantidad), unidades, cajas, picking])

            resumen_por_tecnico[tecnico] = {
                "Unidades Buenas": total_buenas,
                "Unidades Defectuosas": total_defectuosas,
                "Cajas Completas": total_cajas,
                "Unidades a Picking": total_picking
            }

    if resultados:
        df_resultados = pd.DataFrame(resultados, columns=["Técnico", "Código", "Revisado", "Uds/Caja", "Cajas", "Picking"])
        st.success("Análisis completado")
        st.dataframe(df_resultados, use_container_width=True)

        output = io.BytesIO()
        df_resultados.to_excel(output, index=False, engine='openpyxl')
        output.seek(0)

        st.download_button(
            label="📂 Descargar Excel de resultados",
            data=output,
            file_name="analisis_tecnico_resultado.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        # Mostrar gráfica de torta + resumen por técnico
        for tecnico, resumen in resumen_por_tecnico.items():
            st.markdown(f"""---  
            ### {tecnico}""")
            col1, col2 = st.columns([1, 2])
            with col1:
                st.metric(label="Unidades Buenas", value=resumen['Unidades Buenas'])
                st.metric(label="Unidades Defectuosas", value=resumen['Unidades Defectuosas'])
                st.metric(label="Cajas Completas", value=resumen['Cajas Completas'])
                st.metric(label="Unidades a Picking", value=resumen['Unidades a Picking'])
            with col2:
                fig, ax = plt.subplots()
                ax.pie(
                    [resumen['Cajas Completas'], resumen['Unidades a Picking']],
                    labels=["Cajas Completas", "Picking"],
                    autopct='%1.1f%%',
                    startangle=90
                )
                ax.axis('equal')
                st.pyplot(fig)

    else:
        st.warning("No se encontraron datos para analizar.")

    if tecnicos_sin_datos:
        for t in tecnicos_sin_datos:
            st.info(f"{t} no tiene actividad registrada.")
