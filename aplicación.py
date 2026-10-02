from io import BytesIO
import openpyxl
import pandas as pd
import streamlit as st

# Configuración de la página
st.set_page_config(
    page_title="Control Electoral - Seccional", page_icon="🗳️", layout="wide"
)

st.title("🗳️ Centro de Control Rápido y Monitoreo Electoral")
st.markdown(
    "Sistema optimizado para el control de asistencia en tiempo real, con"
    " semáforo de dirigentes y listado de pendientes para rescate."
)

# 1. Subir archivo Excel
uploaded_file = st.file_uploader(
    "Cargar planilla Excel de votantes (.xlsx)", type=["xlsx", "xls"]
)

if uploaded_file is not None:


  @st.cache_data
  def load_data(file):
    df = pd.read_excel(file)
    if "Estado_Voto" not in df.columns:
      df["Estado_Voto"] = False  # False = No ha votado, True = Ya votó

    # Normalizar columna de cédula a texto
    cedula_col = None
    for col in df.columns:
      if "cedula" in col.lower() or "dni" in col.lower() or "doc" in col.lower():
        cedula_col = col
        break
    if cedula_col:
      df[cedula_col] = df[cedula_col].astype(str).str.replace(".0", "", regex=False)
    return df, cedula_col


  df, cedula_col = load_data(uploaded_file)

  if not cedula_col:
    st.error(
        "❌ No se pudo detectar una columna de Cédula o Documento en el Excel."
    )
  else:
    # Detectar columnas clave opcionales
    dirigente_col = next(
        (
            col
            for col in df.columns
            if "dirigente" in col.lower()
            or "lider" in col.lower()
            or "responsable" in col.lower()
        ),
        None,
    )
    mesa_col = next(
        (
            col
            for col in df.columns
            if "mesa" in col.lower() or "seccion" in col.lower()
        ),
        None,
    )
    nombre_col = next(
        (
            col
            for col in df.columns
            if "nombre" in col.lower() or "apellido" in col.lower()
        ),
        None,
    )
    telefono_col = next(
        (
            col
            for col in df.columns
            if "telefono" in col.lower()
            or "celular" in col.lower()
            or "telf" in col.lower()
        ),
        None,
    )

    # --- 2. ESTADÍSTICAS EN VIVO (CONTROL PARALELO) ---
    st.markdown("---")
    col1, col2, col3, col4 = st.columns(4)

    total_padron = len(df)
    total_votaron = int(df["Estado_Voto"].sum())
    faltantes = total_padron - total_votaron
    porcentaje = (
        (total_votaron / total_padron) * 100 if total_padron > 0 else 0
    )

    col1.metric("Padrón Total", f"{total_padron:,}")
    col2.metric("Ya Votaron", f"{total_votaron:,}", delta=f"{porcentaje:.1f}%")
    col3.metric("Faltantes", f"{faltantes:,}")
    col4.metric("Participación General", f"{porcentaje:.1f}%")
    st.markdown("---")

    # Pestañas principales de navegación
    tab_consulta, tab_rescate, tab_dirigentes = st.tabs([
        "🔍 Consulta Rápida por Cédula",
        "🚗 Listado de Pendientes (Rescate)",
        "📊 Rendimiento y Semáforo",
    ])

    # --- PESTAÑA 1: CONSULTA RÁPIDA ---
    with tab_consulta:
      st.subheader("Búsqueda y Registro Instantáneo")
      busqueda = st.text_input(
          "Ingrese el número de cédula del elector:",
          placeholder="Ej: 1234567",
          key="input_cedula",
      )

      if busqueda:
        busqueda_limpia = busqueda.strip()
        resultado = df[df[cedula_col].str.contains(busqueda_limpia, na=False)]

        if len(resultado) == 0:
          st.warning(
              f"⚠️ No se encontró ningún elector con la cédula"
              f" '{busqueda_limpia}'."
          )
        else:
          st.success(
              f"✅ ¡Elector encontrado! ({len(resultado)} coincidencia(s))"
          )

          for idx, row in resultado.iterrows():
            estado_actual = row["Estado_Voto"]

            with st.container(border=True):
              cols = st.columns([2, 2, 2, 2])

              with cols[0]:
                st.markdown(f"**Cédula:** `{row[cedula_col]}`")
                if nombre_col:
                  st.markdown(f"**Nombre:** {row[nombre_col]}")

              with cols[1]:
                if mesa_col:
                  st.markdown(f"🏛️ **Mesa:** {row[mesa_col]}")
                if dirigente_col:
                  st.markdown(f"👤 **Dirigente:** {row[dirigente_col]}")

              with cols[2]:
                if estado_actual:
                  st.markdown("### 🟢 YA VOTÓ")
                else:
                  st.markdown("### 🔴 PENDIENTE")

              with cols[3]:
                if not estado_actual:
                  if st.button(
                      "✅ Marcar Como Votó",
                      key=f"btn_voto_{idx}",
                      type="primary",
                  ):
                    df.at[idx, "Estado_Voto"] = True
                    st.rerun()
                else:
                  if st.button(
                      "↩️ Desmarcar (Error)", key=f"btn_desvoto_{idx}"
                  ):
                    df.at[idx, "Estado_Voto"] = False
                    st.rerun()

    # --- PESTAÑA 2: LISTADO DE PENDIENTES (RESCATE) ---
    with tab_rescate:
      st.subheader("🚗 Operativo Rescate: Electores Faltantes por Dirigente")
      st.markdown(
          "Filtre por dirigente para ver exactamente quiénes de su listado"
          " todavía no votaron al finalizar la tarde."
      )

      if dirigente_col:
        lista_dirigentes = ["Todos"] + list(df[dirigente_col].dropna().unique())
        dirigente_elegido = st.selectbox(
            "Seleccionar Dirigente para ver pendientes", lista_dirigentes
        )

        df_pendientes = df[df["Estado_Voto"] == False]
        if dirigente_elegido != "Todos":
          df_pendientes = df_pendientes[
              df_pendientes[dirigente_col] == dirigente_elegido
          ]

        st.info(f"Total de electores pendientes en este filtro: {len(df_pendientes)}")

        # Mostrar tabla resumida de pendientes (restringiendo columnas innecesarias si existen)
        cols_mostrar = [cedula_col]
        if nombre_col:
          cols_mostrar.append(nombre_col)
        if mesa_col:
          cols_mostrar.append(mesa_col)
        if telefono_col:
          cols_mostrar.append(telefono_col)
        if dirigente_col:
          cols_mostrar.append(dirigente_col)

        st.dataframe(
            df_pendientes[cols_mostrar],
            use_container_width=True,
            hide_index=True,
        )
      else:
        st.warning(
            "No se detectó una columna de Dirigente en el Excel para usar este"
            " módulo."
        )

    # --- PESTAÑA 3: RENDIMIENTO Y SEMÁFORO ---
    with tab_dirigentes:
      st.subheader("📊 Monitoreo y Semáforo de Dirigentes")
      st.markdown(
          "🟢 **Verde:** >50% de participación | 🟡 **Amarillo:** 25% a 50% | 🔴"
          " **Rojo:** Menor al 25%"
      )

      if dirigente_col:
        resumen_dir = (
            df.groupby(dirigente_col)
            .agg(
                Total_Asignados=("Estado_Voto", "count"),
                Ya_Votaron=("Estado_Voto", "sum"),
            )
            .reset_index()
        )
        resumen_dir["Faltantes"] = (
            resumen_dir["Total_Asignados"] - resumen_dir["Ya_Votaron"]
        )
        resumen_dir["% Participación"] = (
            resumen_dir["Ya_Votaron"] / resumen_dir["Total_Asignados"] * 100
        ).round(1)

        # Asignar semáforo visual
        def asignar_semaforo(pct):
          if pct >= 50:
            return "🟢 Óptimo (>50%)"
          elif pct >= 25:
            return "🟡 Regular (25-50%)"
          else:
            return "🔴 Alerta (<25%)"

        resumen_dir["Estado Operativo"] = resumen_dir[
            "% Participación"
        ].apply(asignar_semaforo)

        st.dataframe(
            resumen_dir.sort_values(by="Ya_Votaron", ascending=False),
            use_container_width=True,
            hide_index=True,
        )
      else:
        st.warning(
            "No se detectó una columna de Dirigente para generar el resumen."
        )

    # --- 4. DESCARGA DEL EXCEL EJECUTIVO CON MULTI-PESTAÑAS ---
    st.markdown("---")
    st.subheader("📥 Generar Reporte Final y Descargar Excel Ejecutivo")


    def generar_excel_ejecutivo(dataframe, d_col, m_col):
      output = BytesIO()
      with pd.ExcelWriter(output, engine="openpyxl") as writer:
        dataframe.to_excel(writer, index=False, sheet_name="Padron_Actualizado")

        if d_col:
          resumen_dir = (
              dataframe.groupby(d_col)
              .agg(
                  Total_Asignados=("Estado_Voto", "count"),
                  Ya_Votaron=("Estado_Voto", "sum"),
              )
              .reset_index()
          )
          resumen_dir["Faltantes"] = (
              resumen_dir["Total_Asignados"] - resumen_dir["Ya_Votaron"]
          )
          resumen_dir["% Participacion"] = (
              resumen_dir["Ya_Votaron"] / resumen_dir["Total_Asignados"] * 100
          ).round(2)
          resumen_dir.to_excel(
              writer, index=False, sheet_name="Resumen_Dirigentes"
          )

        if m_col and d_col:
          cruce = (
              dataframe.groupby([m_col, d_col])
              .agg(
                  Total_Asignados=("Estado_Voto", "count"),
                  Votos_Confirmados=("Estado_Voto", "sum"),
              )
              .reset_index()
          )
          cruce["% Efectividad"] = (
              cruce["Votos_Confirmados"] / cruce["Total_Asignados"] * 100
          ).round(2)
          cruce.to_excel(writer, index=False, sheet_name="Cruce_Mesas")

        tot_padron = len(dataframe)
        tot_votos = int(dataframe["Estado_Voto"].sum())
        tot_falt = tot_padron - tot_votos
        pct_part = (tot_votos / tot_padron * 100) if tot_padron > 0 else 0

        dashboard_data = pd.DataFrame({
            "Indicador Electoral": [
                "Padrón Total",
                "Total Votos Emitidos",
                "Votos Faltantes",
                "Porcentaje de Participación General",
            ],
            "Valor": [tot_padron, tot_votos, tot_falt, f"{pct_part:.2f}%"],
        })
        dashboard_data.to_excel(writer, index=False, sheet_name="Dashboard")

      return output.getvalue()


    excel_bytes = generar_excel_ejecutivo(df, dirigente_col, mesa_col)

    st.download_button(
        label=(
            "📥 Descargar Reporte Ejecutivo Completo (Excel con Múltiples"
            " Pestañas)"
        ),
        data=excel_bytes,
        file_name="reporte_electoral_cierre.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        ),
    )

else:
  st.warning(
      "Por favor, cargue un archivo Excel con el padrón de la seccional para"
      " iniciar."
  )
