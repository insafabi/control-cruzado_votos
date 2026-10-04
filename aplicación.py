from datetime import datetime
from io import BytesIO
import openpyxl
import pandas as pd
import streamlit as st
from sqlalchemy import create_engine, text

# Configuración de la página
st.set_page_config(
    page_title="Control Electoral - Seccional", page_icon="🗳️", layout="wide"
)

# --- CONFIGURACIÓN DE PERSISTENCIA (BASE DE DATOS SQLITE) ---
DB_NAME = "control_electoral_persistente.db"
engine = create_engine(f"sqlite:///{DB_NAME}", echo=False)
TABLE_NAME = "padron_votantes"

st.title("🗳️️ Centro de Control Rápido y Mensajería Electoral")
st.markdown(
    "Sistema optimizado para control en tiempo real, semáforo y pantalla de"
    " auditoría (Optimizado para uso en Celular via Web)."
)

# Función para cargar datos persistidos si ya existen previamente en el disco
@st.cache_data
def verificar_tabla_existente():
  try:
    with engine.connect() as conn:
      query = f"SELECT name FROM sqlite_master WHERE type='table' AND name='{TABLE_NAME}';"
      res = conn.execute(text(query)).fetchone()
      if res:
        return pd.read_sql(TABLE_NAME, con=engine)
  except Exception:
    pass
  return None


# Inicializar el session_state buscando si hay una BD guardada previamente
if "df_estado" not in st.session_state:
  df_persitido = verificar_tabla_existente()
  if df_persitido is not None and not df_persitido.empty:
    st.session_state.df_estado = df_persitido
    st.session_state.file_name = "Cargado desde Base de Datos Local"
  else:
    st.session_state.df_estado = None
    st.session_state.file_name = None

# 1. Subir archivo Excel (Solo necesario la primera vez o para actualizar padrón)
uploaded_file = st.file_uploader(
    "Cargar tu planilla Excel de votantes (.xlsx) - (Si ya cargaste antes, los"
    " datos previos están guardados)",
    type=["xlsx", "xls"],
)

if uploaded_file is not None:
  if (
      "file_name_subido" not in st.session_state
      or st.session_state.file_name_subido != uploaded_file.name
  ):
    df_raw = pd.read_excel(uploaded_file)
    df_raw.columns = df_raw.columns.str.strip()

    doc_col_temp = next(
        (
            c
            for c in df_raw.columns
            if "documento" in c.lower() or "cedula" in c.lower()
        ),
        None,
    )
    if doc_col_temp:
      df_raw[doc_col_temp] = (
          df_raw[doc_col_temp].astype(str).str.replace(".0", "", regex=False)
      )

    if "Estado_Voto" not in df_raw.columns:
      if "VOTO" in df_raw.columns:
        df_raw["Estado_Voto"] = (
            df_raw["VOTO"]
            .apply(
                lambda x: True
                if str(x).strip().upper() in ["S", "SI", "1", "TRUE", "X"]
                else False
            )
            .astype(bool)
        )
      else:
        df_raw["Estado_Voto"] = False
    else:
      df_raw["Estado_Voto"] = df_raw["Estado_Voto"].astype(bool)

    if "VOTO" in df_raw.columns:
      df_raw["VOTO"] = df_raw["VOTO"].astype(object)

    # Guardar en base de datos permanente
    df_raw.to_sql(TABLE_NAME, con=engine, if_exists="replace", index=False)

    st.session_state.df_estado = df_raw
    st.session_state.file_name_subido = uploaded_file.name
    st.success("¡Planilla cargada y respaldada en la memoria persistente!")

# Verificar si tenemos un DataFrame activo para mostrar la aplicación
if st.session_state.get("df_estado") is not None:
  df = st.session_state.df_estado

  # Detectar columnas clave de tu planilla de forma inteligente
  doc_col = next(
      (
          c
          for c in df.columns
          if "documento" in c.lower() or "cedula" in c.lower()
      ),
      None,
  )
  dirigente_col = next(
      (c for c in df.columns if "dirigente" in c.lower()), None
  )
  nombre_col = next((c for c in df.columns if "nombres" in c.lower()), None)
  apellido_col = next((c for c in df.columns if "apellidos" in c.lower()), None)
  mesa_col = next((c for c in df.columns if "mesa" in c.lower()), None)
  orden_col = next((c for c in df.columns if "orden" in c.lower()), None)
  telefono_col = next(
      (
          c
          for c in df.columns
          if "telefono" in c.lower()
          or "celular" in c.lower()
          or "tel" in c.lower()
      ),
      None,
  )
  local_col = next(
      (
          c
          for c in df.columns
          if "local" in c.lower()
          or "institucion" in c.lower()
          or "escuela" in c.lower()
          or "colegio" in c.lower()
      ),
      None,
  )

  if not doc_col:
    st.error(
        "❌ No se encontró la columna 'Documento' o 'Cédula' en tu Excel."
    )
  else:

    def actualizar_bd_y_estado():
      st.session_state.df_estado.to_sql(
          TABLE_NAME, con=engine, if_exists="replace", index=False
      )
      st.cache_data.clear()

    # --- BARRA LATERAL / CONTROLES DE EMERGENCIA (RESET TOTAL) ---
    with st.sidebar:
      st.header("⚙️ Opciones y Herramientas")
      st.markdown("---")
      st.warning(
          "⚠️ **Zona de Peligro:** Si necesitas reiniciar absolutamente"
          " **todos** los votos y empezar desde cero:"
      )
      confirmar_reset = st.checkbox(
          "Confirmar reseteo total", key="chk_reset_seguro"
      )
      if st.button("🔄 Reiniciar Todos los Votos", type="primary"):
        if confirmar_reset:
          st.session_state.df_estado["Estado_Voto"] = False
          if "VOTO" in st.session_state.df_estado.columns:
            st.session_state.df_estado["VOTO"] = ""
          actualizar_bd_y_estado()
          st.success("¡Se han borrado todos los votos registrados!")
          st.rerun()
        else:
          st.error("Debes marcar la casilla de confirmación.")

    # --- 2. ESTADÍSTICAS EN VIVO ---
    st.markdown("---")

    total_padron = len(df)
    total_votaron = int(df["Estado_Voto"].sum())
    faltantes = total_padron - total_votaron
    porcentaje = (
        (total_votaron / total_padron) * 100 if total_padron > 0 else 0
    )

    if dirigente_col:
      votaron_con_dirigente = int(
          df[
              (df["Estado_Voto"] == True)
              & (df[dirigente_col].notna())
              & (df[dirigente_col].astype(str).str.strip() != "")
              & (df[dirigente_col].astype(str).str.lower() != "nan")
          ].shape[0]
      )
    else:
      votaron_con_dirigente = 0

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Padrón Total", f"{total_padron:,}")
    col2.metric(
        "Ya Votaron (General)", f"{total_votaron:,}", delta=f"{porcentaje:.1f}%"
    )
    col3.metric("Faltantes", f"{faltantes:,}")
    col4.metric("Participación General", f"{porcentaje:.1f}%")

    col_a, _ = st.columns([2, 2])
    with col_a:
      st.metric(
          "🗳️ Ya Votaron (Con Dirigente Identificado)",
          f"{votaron_con_dirigente:,}",
      )

    st.markdown("---")

    # Pestañas principales de navegación
    tab_consulta, tab_cortes, tab_dirigentes, tab_correccion = st.tabs([
        "🔍 Consulta Rápida por Cédula",
        "📱 Cortes Horarios y Envío WhatsApp",
        "📊 Rendimiento y Semáforo",
        "✏️ Corrección y Auditoría",
    ])

    # --- PESTAÑA 1: CONSULTA RÁPIDA ---
    with tab_consulta:
      st.subheader("Búsqueda y Registro Instantáneo")
      busqueda = st.text_input(
          "Ingrese el número de Documento / Cédula:",
          placeholder="Ej: 4610728",
          key="busqueda_cedula_t1",
          autocomplete="off",
      )

      if busqueda:
        busqueda_limpia = busqueda.strip()
        resultado = df[df[doc_col].str.contains(busqueda_limpia, na=False)]

        if len(resultado) == 0:
          st.warning(
              f"⚠ No se encontró ningún elector con el documento"
              f" '{busqueda_limpia}'."
          )
        else:
          st.success(
              f"✅ ¡Elector encontrado! ({len(resultado)} coincidencia(s))"
          )

          for idx, row in resultado.iterrows():
            estado_actual = row["Estado_Voto"]
            nom_completo = (
                f"{row.get(nombre_col, '')} {row.get(apellido_col, '')}"
                if nombre_col and apellido_col
                else ""
            )

            with st.container(border=True):
              cols = st.columns([2, 2, 2, 2])
              with cols[0]:
                st.markdown(f"**Documento:** `{row[doc_col]}`")
                if nom_completo:
                  st.markdown(f"**Nombre:** {nom_completo}")
              with cols[1]:
                if mesa_col:
                  st.markdown(f"🏛️ **Mesa:** {row[mesa_col]}")
                if orden_col:
                  st.markdown(f"🔢 **Orden:** {row[orden_col]}")
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
                      key=f"btn_voto_t1_{idx}",
                      type="primary",
                  ):
                    st.session_state.df_estado.loc[idx, "Estado_Voto"] = True
                    if "VOTO" in st.session_state.df_estado.columns:
                      st.session_state.df_estado.loc[idx, "VOTO"] = "S"
                    actualizar_bd_y_estado()
                    st.rerun()
                else:
                  if st.button(
                      "↩️ Desmarcar (Error)", key=f"btn_desvoto_t1_{idx}"
                  ):
                    st.session_state.df_estado.loc[idx, "Estado_Voto"] = False
                    if "VOTO" in st.session_state.df_estado.columns:
                      st.session_state.df_estado.loc[idx, "VOTO"] = ""
                    actualizar_bd_y_estado()
                    st.rerun()

    # --- PESTAÑA 2: CORTES HORARIOS Y ENVÍO WHATSAPP ---
    with tab_cortes:
      st.subheader("⏰ Cortes Horarios para Operativo de Mensajería")
      corte_seleccionado = st.selectbox(
          "Seleccionar Corte Horario:",
          ["09:00 AM", "11:00 AM", "13:00 PM", "14:00 PM", "15:00 PM", "16:00 PM"],
          key="corte_horario_sel",
      )

      if dirigente_col:
        dirigentes_lista = ["Todos los Dirigentes"] + list(
            df[dirigente_col].dropna().unique()
        )
        dirigente_filtro = st.selectbox(
            "Filtrar por Dirigente para este corte",
            dirigentes_lista,
            key="dirigente_filtro_corte",
        )
      else:
        dirigente_filtro = "Todos los Dirigentes"

      df_pendientes = df[df["Estado_Voto"] == False].copy()
      if dirigente_filtro != "Todos los Dirigentes" and dirigente_col:
        df_pendientes = df_pendientes[
            df_pendientes[dirigente_col] == dirigente_filtro
        ]

      st.info(
          f"📊 Electores pendientes para el corte de las {corte_seleccionado}:"
          f" **{len(df_pendientes):,} personas**"
      )

      mi_numero_contacto = st.text_input(
          "Tu número de celular (para recibir/coordinar avisos en WhatsApp si"
          " el elector no tiene número cargado):",
          value="595981000000",
          key="mi_numero_wpp_input",
          autocomplete="off",
      )

      if len(df_pendientes) > 0:
        st.markdown(
            "### 🔎 Búsqueda Rápida de Pendientes (Optimizado para evitar"
            " bloqueos)"
        )
        busq_pend = st.text_input(
            "Escribe el nombre o cédula de un pendiente específico para enviarle"
            " mensaje:",
            placeholder="Ej: Juan o Cédula",
            key="busq_pendientes_input",
        )

        if busq_pend:
          df_filtrado_wpp = df_pendientes[
              df_pendientes.astype(str)
              .apply(lambda x: x.str.contains(busq_pend, case=False, na=False))
              .any(axis=1)
          ]
        else:
          df_filtrado_wpp = df_pendientes.head(50)
          st.caption(
              "Mostrando los primeros 50 electores pendientes. Utiliza el"
              " buscador superior si buscas a alguien específico."
          )

        import urllib.parse

        for idx, row in df_filtrado_wpp.iterrows():
          nom = row.get(nombre_col, "Sin Nombre")
          ape = row.get(apellido_col, "")
          cedula = row.get(doc_col, "S/N")
          mesa = row.get(mesa_col, "S/N")
          orden = row.get(orden_col, "S/N")
          local = row.get(local_col, "S/N") if local_col else "S/N"
          dirigente = row.get(dirigente_col, "S/N") if dirigente_col else "S/N"
          telefono = (
              str(row.get(telefono_col, "")).strip() if telefono_col else ""
          )

          mensaje = (
              f"Hola {nom} {ape} (Cédula: {cedula}). Te recordamos pasar a"
              f" votar (Corte: {corte_seleccionado}). Datos de votación ->"
              f" Local: {local} | Mesa: {mesa} | Orden: {orden}."
          )

          telefono_limpio = (
              telefono.replace(".0", "")
              .replace(" ", "")
              .replace("+", "")
              .replace("-", "")
          )
          if (
              telefono_limpio
              and telefono_limpio.lower() != "nan"
              and len(telefono_limpio) > 6
          ):
            whatsapp_url = (
                f"https://wa.me/{telefono_limpio}?text="
                f"{urllib.parse.quote(mensaje)}"
            )
            etiqueta_btn = f"💬 Enviar a Elector ({telefono})"
            color_btn = "#25D366"
          else:
            mi_num_limpio = (
                mi_numero_contacto.replace("+", "").replace(" ", "").strip()
            )
            whatsapp_url = (
                f"https://wa.me/{mi_num_limpio}?text="
                f"{urllib.parse.quote('REPORTE PENDIENTE: ' + mensaje)}"
            )
            etiqueta_btn = "📤 Enviar a Mi WhatsApp (Sin cel del elector)"
            color_btn = "#007BFF"

          with st.container(border=True):
            col_info1, col_info2, col_btn = st.columns([3, 3, 2])
            with col_info1:
              st.markdown(
                  f"**👤 {nom} {ape}**<br>Cédula: `{cedula}`<br>Dirigente:"
                  f" `{dirigente}`",
                  unsafe_allow_html=True,
              )
            with col_info2:
              st.markdown(
                  f"🏛️ **Local:** {local}<br>🔢 **Mesa:** {mesa} | **Orden:**"
                  f" {orden}",
                  unsafe_allow_html=True,
              )
            with col_btn:
              st.markdown("<br>", unsafe_allow_html=True)
              st.markdown(
                  f'<a href="{whatsapp_url}" target="_blank"><button'
                  f' style="background-color:{color_btn}; color:white;'
                  ' border:none; padding:10px 14px; border-radius:5px;'
                  f' cursor:pointer; font-weight:bold; width:100%;">{etiqueta_btn}</button></a>',
                  unsafe_allow_html=True,
              )

    # --- PESTAÑA 3: RENDIMIENTO Y SEMÁFORO ---
    with tab_dirigentes:
      st.subheader("📊 Monitoreo y Semáforo de Dirigentes")
      st.markdown(
          "🟢 **Óptimo (>50%)** | 🟡 **Regular (25-50%)** | 🔴 **Alerta (<25%)**"
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
        st.warning("No se detectó una columna de Dirigente.")

    # --- PESTAÑA 4: CORRECCIÓN Y AUDITORÍA INDIVIDUAL ---
    with tab_correccion:
      st.subheader("✏️ Pantalla de Corrección y Auditoría por Cédula")
      st.markdown(
          "Utiliza este buscador para localizar rápidamente a una persona, ver"
          " si se registró su voto por error y **cambiar su condición de forma"
          " manual** (marcar como pendiente o como votó)."
      )

      busqueda_corr = st.text_input(
          "Buscar por Cédula o Documento para corregir:",
          placeholder="Ej: 4610728",
          key="busqueda_cedula_t4",
          autocomplete="off",
      )

      if busqueda_corr:
        busqueda_corr_limpia = busqueda_corr.strip()
        resultado_corr = df[
            df[doc_col].str.contains(busqueda_corr_limpia, na=False)
        ]

        if len(resultado_corr) == 0:
          st.warning(
              f"⚠️ No se encontró ningún registro con el documento"
              f" '{busqueda_corr_limpia}'."
          )
        else:
          st.success(
              f"✅ ¡Elector encontrado! ({len(resultado_corr)} coincidencia(s))"
          )

          for idx, row in resultado_corr.iterrows():
            estado_actual = row["Estado_Voto"]
            nom_completo = (
                f"{row.get(nombre_col, '')} {row.get(apellido_col, '')}"
                if nombre_col and apellido_col
                else ""
            )

            with st.container(border=True):
              cols = st.columns([2, 2, 2, 2])
              with cols[0]:
                st.markdown(f"**Documento:** `{row[doc_col]}`")
                if nom_completo:
                  st.markdown(f"**Nombre:** {nom_completo}")
              with cols[1]:
                if mesa_col:
                  st.markdown(f"🏛️ **Mesa:** {row[mesa_col]}")
                if orden_col:
                  st.markdown(f"🔢 **Orden:** {row[orden_col]}")
                if dirigente_col:
                  st.markdown(f"👤 **Dirigente:** {row[dirigente_col]}")
              with cols[2]:
                if estado_actual:
                  st.markdown("### 🟢 Estado: YA VOTÓ")
                else:
                  st.markdown("### 🔴 Estado: PENDIENTE")
              with cols[3]:
                st.markdown("**Modificar Condición:**")
                if estado_actual:
                  if st.button(
                      "🔄 Cambiar a Pendiente (Borrar Voto)",
                      key=f"btn_corr_pend_{idx}",
                      type="secondary",
                  ):
                    st.session_state.df_estado.loc[idx, "Estado_Voto"] = False
                    if "VOTO" in st.session_state.df_estado.columns:
                      st.session_state.df_estado.loc[idx, "VOTO"] = ""
                    actualizar_bd_y_estado()
                    st.success("¡Se ha revertido el voto a pendiente!")
                    st.rerun()
                else:
                  if st.button(
                      "🔄 Cambiar a Ya Votó",
                      key=f"btn_corr_voto_{idx}",
                      type="primary",
                  ):
                    st.session_state.df_estado.loc[idx, "Estado_Voto"] = True
                    if "VOTO" in st.session_state.df_estado.columns:
                      st.session_state.df_estado.loc[idx, "VOTO"] = "S"
                    actualizar_bd_y_estado()
                    st.success("¡Se ha marcado el voto correctamente!")
                    st.rerun()

    # --- 5. DESCARGA DEL EXCEL EJECUTIVO CON MULTI-PESTAÑAS ---
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
        file_name="reporte_electoral_actualizado.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        ),
    )
else:
  st.info(
      "📁 Por favor, carga tu planilla Excel inicial de la seccional para"
      " comenzar. Una vez cargada, los votos quedarán guardados"
      " automáticamente en la base de datos local temporal."
  )
