from datetime import datetime
from io import BytesIO
import openpyxl
import pandas as pd
import streamlit as st

# Configuración de la página
st.set_page_config(
    page_title="Control Electoral - Seccional", page_icon="🗳️", layout="wide"
)

st.title("🗳️ Centro de Control Rápido y Mensajería Electoral")
st.markdown(
    "Sistema optimizado para tu planilla: Control en tiempo real, semáforo de"
    " dirigentes, rescate y envíos de WhatsApp por cortes horarios (9h, 11h,"
    " 13h, 14h, 15h, 16h)."
)

# 1. Subir archivo Excel
uploaded_file = st.file_uploader(
    "Cargar tu planilla Excel de votantes (.xlsx)", type=["xlsx", "xls"]
)

if uploaded_file is not None:


  @st.cache_data
  def load_data(file):
    df = pd.read_excel(file)

    # Normalizar nombres de columnas por si tienen espacios extras
    df.columns = df.columns.str.strip()

    # Mapear o verificar columnas basándonos en tu estructura exacta
    # Esperadas: Documento, Dirigente, Local, Apellidos, Nombres, VOTO, Partido, Mesa, Orden, Telefono
    if "Estado_Voto" not in df.columns:
      # Si la columna VOTO original tiene 'S' o valores, podemos inicializar el estado
      if "VOTO" in df.columns:
        df["Estado_Voto"] = df["VOTO"].apply(
            lambda x: True
            if str(x).strip().upper() in ["S", "SI", "1", "TRUE", "X"]
            else False
        )
      else:
        df["Estado_Voto"] = False

    # Limpiar columna de cédula/documento a texto puro
    doc_col = next(
        (c for c in df.columns if "documento" in c.lower() or "cedula" in c.lower()),
        None,
    )
    if doc_col:
      df[doc_col] = df[doc_col].astype(str).str.replace(".0", "", regex=False)

    return df


  df = load_data(uploaded_file)

  # Detectar columnas clave de tu planilla
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
          if "telefono" in c.lower() or "celular" in c.lower()
      ),
      None,
  )

  if not doc_col:
    st.error(
        "❌ No se encontró la columna 'Documento' o 'Cédula' en tu Excel."
    )
  else:
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
    tab_consulta, tab_cortes, tab_dirigentes = st.tabs([
        "🔍 Consulta Rápida por Cédula",
        "📱 Cortes Horarios y Envío WhatsApp",
        "📊 Rendimiento y Semáforo",
    ])

    # --- PESTAÑA 1: CONSULTA RÁPIDA ---
    with tab_consulta:
      st.subheader("Búsqueda y Registro Instantáneo")
      busqueda = st.text_input(
          "Ingrese el número de Documento / Cédula:",
          placeholder="Ej: 4610728",
          key="input_cedula",
      )

      if busqueda:
        busqueda_limpia = busqueda.strip()
        resultado = df[df[doc_col].str.contains(busqueda_limpia, na=False)]

        if len(resultado) == 0:
          st.warning(
              f"⚠️ No se encontró ningún elector con el documento"
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
                      key=f"btn_voto_{idx}",
                      type="primary",
                  ):
                    df.at[idx, "Estado_Voto"] = True
                    df.at[idx, "VOTO"] = "S"
                    st.rerun()
                else:
                  if st.button(
                      "↩️ Desmarcar (Error)", key=f"btn_desvoto_{idx}"
                  ):
                    df.at[idx, "Estado_Voto"] = False
                    df.at[idx, "VOTO"] = ""
                    st.rerun()

    # --- PESTAÑA 2: CORTES HORARIOS Y ENVÍO WHATSAPP ---
    with tab_cortes:
      st.subheader("⏰ Cortes Horarios para Operativo de Mensajería")
      st.markdown(
          "Selecciona el corte horario actual para filtrar a los electores"
          " pendientes y disparar mensajes recordatorios desde tu celular."
      )

      # Selector de cortes horarios solicitados
      corte_seleccionado = st.selectbox(
          "Seleccionar Corte Horario:",
          ["09:00 AM", "11:00 AM", "13:00 PM", "14:00 PM", "15:00 PM", "16:00 PM"],
      )

      # Filtro opcional por dirigente dentro del corte
      if dirigente_col:
        dirigentes_lista = ["Todos los Dirigentes"] + list(
            df[dirigente_col].dropna().unique()
        )
        dirigente_filtro = st.selectbox(
            "Filtrar por Dirigente para este corte", dirigentes_lista
        )
      else:
        dirigente_filtro = "Todos los Dirigentes"

      # Filtrar solo los que NO han votado
      df_pendientes = df[df["Estado_Voto"] == False].copy()
      if dirigente_filtro != "Todos los Dirigentes" and dirigente_col:
        df_pendientes = df_pendientes[
            df_pendientes[dirigente_col] == dirigente_filtro
        ]

      st.info(
          f"📊 Electores pendientes para el corte de las {corte_seleccionado}:"
          f" **{len(df_pendientes)} personas**"
      )

      if len(df_pendientes) > 0 and telefono_col:
        st.markdown(
            "#### Listado interactivo con enlace directo de WhatsApp para cada"
            " elector pendiente:"
        )

        for idx, row in df_pendientes.head(50).iterrows():  # Muestra los primeros para agilizar
          telefono = str(row.get(telefono_col, "")).strip()
          nom = row.get(nombre_col, "")
          ape = row.get(apellido_col, "")
          mesa = row.get(mesa_col, "S/N")
          orden = row.get(orden_col, "S/N")
          dirigente = row.get(dirigente_col, "S/N")

          # Limpiar teléfono para link de whatsapp (ej: quitar .0, espacios)
          telefono_limpio = (
              telefono.replace(".0", "").replace(" ", "").replace("+", "")
          )

          # Mensaje personalizado para el corte
          mensaje = (
              f"¡Hola {nom}! Te saludamos desde el comando. Vemos que aún"
              f" no pudiste pasar a votar en este corte de las"
              f" {corte_seleccionado}. Tu mesa es {mesa} (Orden {orden})."
              f" ¡Contamos con tu presencia!"
          )
          import urllib.parse

          mensaje_encoded = urllib.parse.quote(mensaje)
          whatsapp_url = f"https://wa.me/{telefono_limpio}?text={mensaje_encoded}"

          col_a, col_b, col_c = st.columns([3, 2, 2])
          with col_a:
            st.text(f"{nom} {ape} (Tel: {telefono})")
          with col_b:
            st.text(f"Mesa: {mesa} | Dir: {dirigente}")
          with col_c:
            st.markdown(
                f'<a href="{whatsapp_url}" target="_blank"><button'
                ' style="background-color:#25D366; color:white; border:none;'
                ' padding:8px 12px; border-radius:5px; cursor:pointer; font-weight:bold;">💬'
                " Enviar WhatsApp</button></a>",
                unsafe_allow_html=True,
            )
        if len(df_pendientes) > 50:
          st.caption(
              "Mostrando los primeros 50 pendientes para optimizar la velocidad"
              " de visualización."
          )
      elif not telefono_col:
        st.warning(
            "No se detectó una columna de Teléfono o Celular en tu planilla"
            " para habilitar los enlaces de WhatsApp."
        )
      else:
        st.success("🎉 ¡Excelente! No hay electores pendientes en este filtro.")

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
  st.warning(
      "Por favor, cargue su archivo Excel de la seccional para comenzar."
  )
