"""
Reporte Mensual Upselling - GOUT Argentina
App Streamlit: dashboard interactivo + generador de PDF
"""

import io
import sys
import importlib

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from reportlab.pdfgen import canvas as pdf_canvas
from reportlab.lib.pagesizes import A4

# ─────────────────────────────────────────────────────────────────────────────
# CONFIG
# ─────────────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Reporte Upselling - GOUT",
    page_icon="📊",
    layout="wide",
)

MESES = [
    "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
    "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"
]

COLORES = {
    "dark":  "#1D3D2F",
    "med":   "#2D5A3D",
    "teal":  "#80CBC4",
    "light": "#C5E8C5",
    "pale":  "#EAF5EA",
    "pink":  "#F8BBD9",
}

# ─────────────────────────────────────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## ⚙️ Configuración")
    mes = st.selectbox("Mes", MESES, index=2)
    ano = st.text_input("Año", "2026")

    st.markdown("---")
    st.markdown("**IDs de Google Sheets**")
    sheet1_id = st.text_input(
        "Sheet 1 (Upselling)",
        "1umPdIIJ3v3CBBSeNKKqylZ4t2EKLtm0Ne0UBVR_gmRo",
        help="Hoja con DIARIO, DATOS, VTAS, 3HS, CONVXSKILL"
    )
    sheet2_id = st.text_input(
        "Sheet 2 (Dimensionamiento)",
        "1Lq6QeGJJvGM6XbdDATeps6CvFr4mye6bk6a_wtuiWuY",
        help="Hoja con ASISTENCIA"
    )

    st.markdown("---")
    cargar_btn = st.button("🔄 Cargar datos", use_container_width=True, type="primary")

# ─────────────────────────────────────────────────────────────────────────────
# CARGA DE DATOS
# ─────────────────────────────────────────────────────────────────────────────
def _write_temp_credentials():
    """
    En Streamlit Cloud las credenciales vienen de st.secrets.
    Localmente usa credentials.json si existe.
    Devuelve la ruta al archivo de credenciales a usar.
    """
    import os, json, tempfile
    if os.path.exists("credentials.json"):
        return "credentials.json"
    # Streamlit Cloud: reconstruir el JSON desde secrets
    creds_dict = dict(st.secrets["gcp_service_account"])
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".json", mode="w")
    json.dump(creds_dict, tmp)
    tmp.close()
    return tmp.name


def cargar_datos(mes, ano, sheet1_id, sheet2_id):
    """Carga generate_report con los parámetros del mes y retorna el dict de datos."""
    import generate_report as gr
    gr.MES              = mes
    gr.ANO              = ano
    gr.SHEET1_ID        = sheet1_id
    gr.SHEET2_ID        = sheet2_id
    gr.OUTPUT           = f"Reporte_Upselling_{mes}{ano}.pdf"
    gr.CREDENTIALS_FILE = _write_temp_credentials()
    return gr.load_data()


def datos_a_dataframes(d):
    """Convierte el dict de datos a DataFrames de pandas."""
    n = min(len(d['dates']), len(d['datos_daily']),
            len(d['ventas_daily']), len(d['final_daily']), len(d['efect_serie']))

    df_diario = pd.DataFrame({
        "Fecha":        d['dates'][:n],
        "Datos":        d['datos_daily'][:n],
        "Ventas":       d['ventas_daily'][:n],
        "Finalizadas":  d['final_daily'][:n],
        "Efectividad%": d['efect_serie'][:n],
    })

    asesores = {}
    for nombre, ventas in d['asesores_vtas'].items():
        asesores[nombre] = {
            "Asesor":  nombre,
            "Turno":   d['turno_map'].get(nombre, ""),
            "Ventas":  ventas,
            "Datos":   d['asesores_datos'].get(nombre, 0),
        }
    for nombre, datos in d['asesores_datos'].items():
        if nombre not in asesores:
            asesores[nombre] = {
                "Asesor": nombre,
                "Turno":  d['turno_map'].get(nombre, ""),
                "Ventas": 0,
                "Datos":  datos,
            }
    df_asesores = pd.DataFrame(list(asesores.values()))
    if not df_asesores.empty:
        df_asesores["Conv%"] = (
            df_asesores["Ventas"] / df_asesores["Datos"].replace(0, 1) * 100
        ).round(2)
        df_asesores = df_asesores.sort_values("Ventas", ascending=False).reset_index(drop=True)

    df_horas = pd.DataFrame([
        {"Hora": f"{h:02d}:00", "Total": t}
        for h, t in sorted(d['hora_totals'].items())
    ])

    df_cvx = pd.DataFrame(d['cvx_tm'] + d['cvx_tn'])
    if not df_cvx.empty:
        df_cvx = df_cvx.rename(columns={
            "nombre": "Asesor", "turno": "Turno",
            "datos": "Datos", "prom_datos": "Prom Datos",
            "horas": "Horas", "dias_lab": "Días Lab",
            "vts_brutas": "Vtas Brutas", "conv_brtas": "Conv% Brutas",
            "vts_activ": "Vtas Activ", "conv_activ": "Conv% Activ",
        })

    df_team = pd.DataFrame(d['team']).rename(columns={
        "grupo": "Grupo", "rol": "Rol",
        "nombre": "Nombre", "horas": "Hs/Turno", "dias_p": "Días P",
    }) if d['team'] else pd.DataFrame()

    return df_diario, df_asesores, df_horas, df_cvx, df_team


# ─────────────────────────────────────────────────────────────────────────────
# GENERADOR DE PDF (usa las funciones del script original)
# ─────────────────────────────────────────────────────────────────────────────
def generar_pdf(d, mes, ano) -> bytes:
    import generate_report as gr
    gr.MES = mes
    gr.ANO = ano

    buf = io.BytesIO()
    cv  = pdf_canvas.Canvas(buf, pagesize=A4)

    gr.page_cover(cv);                cv.showPage()
    gr.page_contents(cv);             cv.showPage()
    gr.page_results(cv, d);           cv.showPage()
    gr.page_additional(cv, d);        cv.showPage()
    gr.page_detail(cv, d);            cv.showPage()
    gr.page_attendance(cv, d);        cv.showPage()
    gr.page_recommendations(cv, d);   cv.showPage()
    gr.page_closing(cv);              cv.showPage()

    cv.save()
    buf.seek(0)
    return buf.read()


# ─────────────────────────────────────────────────────────────────────────────
# ESTADO DE SESION
# ─────────────────────────────────────────────────────────────────────────────
if cargar_btn:
    with st.spinner("Conectando a Google Sheets..."):
        try:
            d = cargar_datos(mes, ano, sheet1_id, sheet2_id)
            st.session_state["data"] = d
            st.session_state["mes"]  = mes
            st.session_state["ano"]  = ano
            st.success("Datos cargados correctamente.")
        except Exception as e:
            st.error(f"Error al cargar datos: {e}")
            st.stop()

# ─────────────────────────────────────────────────────────────────────────────
# DASHBOARD
# ─────────────────────────────────────────────────────────────────────────────
if "data" not in st.session_state:
    st.title("📊 Reporte Mensual Upselling")
    st.info("Configurá el mes y hacé click en **Cargar datos** para comenzar.")
    st.stop()

d          = st.session_state["data"]
mes_actual = st.session_state["mes"]
ano_actual = st.session_state["ano"]

df_diario, df_asesores, df_horas, df_cvx, df_team = datos_a_dataframes(d)

st.title(f"📊 Reporte Upselling — {mes_actual} {ano_actual}")

# ── KPIs ─────────────────────────────────────────────────────────────────────
st.markdown("### Resumen del mes")
c1, c2, c3, c4 = st.columns(4)
c1.metric("📥 Datos totales",     f"{d['datos_total']:,}",  f"Prom: {d['datos_prom']}/día")
c2.metric("💰 Ventas brutas",     f"{d['ventas_total']:,}", f"Prom: {d['ventas_prom']}/día")
c3.metric("✅ Finalizadas",       f"{d['final_total']:,}",  f"Prom: {d['final_prom']}/día")
c4.metric("📈 Efectividad",       f"{d['efect_total']:.2f}%")

st.markdown("---")

# ── SERIES DIARIAS ────────────────────────────────────────────────────────────
st.markdown("### Series diarias")
tab1, tab2 = st.tabs(["Volúmenes", "Efectividad"])

with tab1:
    if not df_diario.empty:
        fig = go.Figure()
        fig.add_trace(go.Bar(
            x=df_diario["Fecha"], y=df_diario["Datos"],
            name="Datos", marker_color=COLORES["light"]
        ))
        fig.add_trace(go.Bar(
            x=df_diario["Fecha"], y=df_diario["Ventas"],
            name="Ventas brutas", marker_color=COLORES["med"]
        ))
        fig.add_trace(go.Bar(
            x=df_diario["Fecha"], y=df_diario["Finalizadas"],
            name="Finalizadas", marker_color=COLORES["teal"]
        ))
        fig.update_layout(
            barmode="group", plot_bgcolor="white",
            xaxis_tickangle=-45, height=380,
            legend=dict(orientation="h", y=1.1),
            margin=dict(l=0, r=0, t=30, b=0),
        )
        st.plotly_chart(fig, use_container_width=True)

with tab2:
    if not df_diario.empty and df_diario["Efectividad%"].sum() > 0:
        fig2 = px.line(
            df_diario, x="Fecha", y="Efectividad%",
            markers=True, color_discrete_sequence=[COLORES["dark"]],
        )
        fig2.add_hline(
            y=d['efect_total'], line_dash="dash",
            line_color=COLORES["teal"],
            annotation_text=f"Prom: {d['efect_total']:.2f}%",
        )
        fig2.update_layout(
            plot_bgcolor="white", height=380,
            xaxis_tickangle=-45,
            margin=dict(l=0, r=0, t=30, b=0),
        )
        st.plotly_chart(fig2, use_container_width=True)

st.markdown("---")

# ── ASESORES ─────────────────────────────────────────────────────────────────
st.markdown("### Ranking de asesores")
col_left, col_right = st.columns([1.2, 1])

with col_left:
    if not df_asesores.empty:
        fig3 = px.bar(
            df_asesores.sort_values("Ventas"),
            x="Ventas", y="Asesor", orientation="h",
            color="Turno",
            color_discrete_map={"TM": COLORES["med"], "TN": COLORES["teal"]},
            height=max(300, len(df_asesores) * 36),
        )
        fig3.update_layout(
            plot_bgcolor="white",
            margin=dict(l=0, r=0, t=30, b=0),
        )
        st.plotly_chart(fig3, use_container_width=True)

with col_right:
    if not df_asesores.empty:
        st.dataframe(
            df_asesores[["Asesor", "Turno", "Datos", "Ventas", "Conv%"]],
            hide_index=True,
            use_container_width=True,
        )

st.markdown("---")

# ── FRANJA HORARIA ────────────────────────────────────────────────────────────
st.markdown("### Distribución por franja horaria")
if not df_horas.empty:
    fig4 = px.bar(
        df_horas, x="Hora", y="Total",
        color_discrete_sequence=[COLORES["dark"]],
        height=320,
    )
    fig4.update_layout(
        plot_bgcolor="white",
        margin=dict(l=0, r=0, t=30, b=0),
    )
    st.plotly_chart(fig4, use_container_width=True)

st.markdown("---")

# ── DETALLE CONVXSKILL ────────────────────────────────────────────────────────
st.markdown("### Detalle por asesor (CONVXSKILL)")
if not df_cvx.empty:
    st.dataframe(df_cvx, hide_index=True, use_container_width=True)

st.markdown("---")

# ── ASISTENCIA ────────────────────────────────────────────────────────────────
st.markdown("### Asistencia del equipo")
if not df_team.empty:
    for grupo, label in [("TM", "Turno Mañana"), ("TN", "Turno Noche"),
                          ("SUP", "Supervisores"), ("BO", "Backoffice")]:
        subset = df_team[df_team["Grupo"] == grupo]
        if not subset.empty:
            with st.expander(f"**{label}** ({len(subset)} personas)", expanded=True):
                st.dataframe(
                    subset[["Rol", "Nombre", "Hs/Turno", "Días P"]],
                    hide_index=True, use_container_width=True,
                )

st.markdown("---")

# ── GENERAR PDF ───────────────────────────────────────────────────────────────
st.markdown("### Generar reporte PDF")
if st.button("📄 Generar PDF", type="primary"):
    with st.spinner("Generando PDF..."):
        try:
            pdf_bytes = generar_pdf(d, mes_actual, ano_actual)
            st.download_button(
                label="⬇️ Descargar PDF",
                data=pdf_bytes,
                file_name=f"Reporte_Upselling_{mes_actual}{ano_actual}.pdf",
                mime="application/pdf",
            )
            st.success("PDF listo para descargar.")
        except Exception as e:
            st.error(f"Error generando PDF: {e}")
