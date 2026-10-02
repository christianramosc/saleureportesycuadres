# -*- coding: utf-8 -*-
"""
app_cuadres.py – Cuadres de créditos MG Colima (Streamlit)

  · Cuadre vs Excel de seguimiento (pestañas COLIMA)
  · Cuadre vs Sale-U

Necesita en la misma carpeta:
  cuadre_seguimiento_cetelem.py  y  cuadre_saleu_cetelem.py
Ejecutar local:  streamlit run app_cuadres.py
"""

import os
import tempfile
import pandas as pd
import streamlit as st

import cuadre_seguimiento_cetelem as cs
import cuadre_saleu_cetelem as cu

st.set_page_config(page_title="Cuadres de créditos · MG Colima", page_icon="📋", layout="wide")

st.markdown("""
<style>
  .block-container {padding-top: 2rem; max-width: 1400px;}
  h1 {color: #1F3864; font-weight: 700;}
  div[data-testid="stMetric"] {background: #F5F7FB; border-left: 4px solid #1F3864;
                               padding: .6rem .9rem; border-radius: 6px;}
  div[data-testid="stMetric"].alerta {border-left-color: #C00000;}
</style>""", unsafe_allow_html=True)


# =============================================================================
# Utilidades
# =============================================================================
def guardar_temporal(archivos, carpeta):
    """Guarda los archivos subidos con su nombre original (el mes se lee del nombre)."""
    rutas = []
    for a in archivos:
        ruta = os.path.join(carpeta, a.name)
        with open(ruta, "wb") as f:
            f.write(a.getbuffer())
        rutas.append(ruta)
    return rutas


def leer_bytes(ruta):
    with open(ruta, "rb") as f:
        return f.read()


def colorear(valor):
    verde, amarillo, rojo = "#C6EFCE", "#FFEB9C", "#FFC7CE"
    v = str(valor)
    if v in ("OK", "Coincide por folio", "Coincide (teléfono y nombre)"):
        return f"background-color: {verde}"
    if v.startswith("NO ") or v == "No cuadra" or "pero NO" in v:
        return f"background-color: {rojo}"
    if v == "Revisar" or v.startswith("Coincide por") or v.startswith("Posible"):
        return f"background-color: {amarillo}"
    return ""


def tabla(df, columnas, columnas_color=(), alto=None):
    cols = [c for c in columnas if c in df.columns]
    if df.empty:
        st.info("Sin registros.")
        return
    vista = df[cols].copy()
    vista.columns = [c.replace("_", " ") for c in cols]
    estilo = vista.style
    color_cols = [c.replace("_", " ") for c in columnas_color if c in cols]
    if color_cols:
        estilo = estilo.map(colorear, subset=color_cols)
    extra = {"height": alto} if alto else {}
    st.dataframe(estilo, width="stretch", hide_index=True, **extra)


def filtro_vendedor(df, clave, columna="Vendedor"):
    if df.empty or columna not in df.columns:
        return df
    opciones = sorted(v for v in df[columna].dropna().unique() if str(v).strip())
    elegidos = st.multiselect("Filtrar por vendedor", opciones, key=clave)
    return df[df[columna].isin(elegidos)] if elegidos else df


def valor_resumen(resumen, texto):
    fila = resumen[resumen["Concepto"].str.strip().str.startswith(texto)]
    return fila["Valor"].iloc[0] if not fila.empty else "—"


def descargas(res, prefijo):
    c1, c2, _ = st.columns([1, 1, 3])
    c1.download_button("⬇️ Descargar Excel", res["xlsx_bytes"], file_name=res["xlsx_nombre"],
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       width="stretch", key=f"{prefijo}_x")
    c2.download_button("⬇️ Descargar PDF", res["pdf_bytes"], file_name=res["pdf_nombre"],
                       mime="application/pdf", width="stretch", key=f"{prefijo}_p")


def ejecutar(modulo, archivo_base, capturas):
    """Corre procesar() en una carpeta temporal y regresa tablas + bytes de los archivos."""
    with tempfile.TemporaryDirectory() as tmp:
        ruta_base = guardar_temporal([archivo_base], tmp)[0]
        rutas_cet = guardar_temporal(capturas, tmp)
        res = modulo.procesar(ruta_base, rutas_cet, carpeta=tmp)
        res["xlsx_bytes"], res["pdf_bytes"] = leer_bytes(res["xlsx"]), leer_bytes(res["pdf"])
        res["xlsx_nombre"], res["pdf_nombre"] = os.path.basename(res["xlsx"]), os.path.basename(res["pdf"])
    return res


# =============================================================================
# Barra lateral: ajustes
# =============================================================================
with st.sidebar:
    st.header("⚙️ Ajustes")
    st.caption("Qué tan estricta es la comparación de nombres (aplica a los dos cuadres).")
    fuerte = st.slider("Coincidencia solo por nombre", 0.6, 1.0, 0.80, 0.05,
                       help="Parecido mínimo para aceptar un cliente con teléfono distinto.")
    con_tel = st.slider("Nombre cuando el teléfono coincide", 0.3, 1.0, 0.50, 0.05,
                        help="Si el teléfono es igual, cuánto debe parecerse el nombre.")
    for m in (cs, cu):
        m.UMBRAL_NOMBRE_FUERTE, m.UMBRAL_NOMBRE_CON_TEL = fuerte, con_tel

    st.divider()
    st.subheader("Seguimiento")
    cs.FILTRO_HOJAS = st.text_input("Pestañas que contengan", cs.FILTRO_HOJAS).strip().upper() or "COLIMA"
    st.caption("Pestaña donde debe estar cada estatus de Cetelem:")
    opciones_hoja = {"Rechazos": "RECHAZO", "Créditos finalizados": "FINALIZADO", "No se juzga": None}
    inv = {v: k for k, v in opciones_hoja.items()}
    for est in ["APROBADO", "CONTRAPROPUESTA", "FINANCIADO", "RECHAZADO"]:
        actual = inv.get(cs.HOJA_ESPERADA.get(est))
        elegido = st.selectbox(est.capitalize(), list(opciones_hoja), list(opciones_hoja).index(actual),
                               key=f"hoja_{est}")
        cs.HOJA_ESPERADA[est] = opciones_hoja[elegido]
    cs.HOJA_ESPERADA["FINANCIADOS"] = cs.HOJA_ESPERADA["FINANCIADO"]

    st.divider()
    st.subheader("Sale-U")
    st.caption("Cómo debe verse en Sale-U una contrapropuesta de Cetelem:")
    contra = st.selectbox("Contrapropuesta", ["No se juzga (revisar)", "Aceptado", "Rechazado"])
    cu.ESTATUS_ESPERADO["CONTRAPROPUESTA"] = {"Aceptado": "ACEPTADO",
                                              "Rechazado": "RECHAZADO"}.get(contra)


# =============================================================================
# Página
# =============================================================================
st.title("Cuadres de créditos · MG Colima")
st.caption("Sube las capturas mensuales de Cetelem y el archivo contra el que quieres cuadrar.")

tab_seg, tab_su = st.tabs(["📒 Excel de seguimiento", "🟦 Sale-U"])

# ---------------------------------------------------------------- SEGUIMIENTO
with tab_seg:
    c1, c2 = st.columns(2)
    f_seg = c1.file_uploader("Excel de seguimiento", type=["xlsx"], key="up_seg",
                             help="Ej. Seguimiento_Rechazos_Financiera_MG.xlsx")
    f_cet = c2.file_uploader("Capturas mensuales Cetelem", type=["xlsx"], accept_multiple_files=True,
                             key="up_cet_seg", help="Ej. JULIO_2026_FINAL.xlsx, AGOSTO_2026_FINAL.xlsx…")

    if st.button("Hacer cuadre", type="primary", disabled=not (f_seg and f_cet), key="btn_seg"):
        with st.spinner("Cruzando folios, teléfonos y nombres…"):
            try:
                st.session_state["res_seg"] = ejecutar(cs, f_seg, f_cet)
            except Exception as e:
                st.session_state.pop("res_seg", None)
                st.error(f"No se pudo hacer el cuadre: {e}")

    res = st.session_state.get("res_seg")
    if res:
        r = res["resumen"]
        st.success(f"Capturas: {', '.join(res['meses'])}  ·  Pestañas: {', '.join(res['hojas'])}")
        m = st.columns(5)
        m[0].metric("Clientes Cetelem", valor_resumen(r, "Clientes únicos"))
        m[1].metric("Encontrados", valor_resumen(r, "Clientes Cetelem encontrados"))
        m[2].metric("Faltan capturar", valor_resumen(r, "Clientes Cetelem NO"))
        m[3].metric("% de cuadre", valor_resumen(r, "Porcentaje"))
        m[4].metric("Datos a corregir", valor_resumen(r, "Datos a corregir"))
        descargas(res, "seg")

        t1, t2, t3, t4, t5 = st.tabs([
            f"Faltan en seguimiento ({len(res['faltantes'])})",
            f"A corregir / revisar ({len(res['revisar'])})",
            f"En seguimiento sin captura ({len(res['solo_periodo'])})",
            "Cuadre completo", "Resumen"])
        with t1:
            df = filtro_vendedor(res["faltantes"], "fv_seg")
            if not df.empty:
                st.bar_chart(df.groupby("Estatus_final_Cetelem").size().rename("Clientes"),
                             color="#C00000", height=220)
            tabla(df, ["Cliente_Cetelem", "Tel_Cetelem", "Vendedor", "Folios", "Historial",
                       "Estatus_final_Cetelem", "Observaciones"])
        with t2:
            tabla(res["revisar"], ["Resultado", "Cuadre_pestana", "Cliente_Cetelem", "Tel_Cetelem",
                                   "Cliente_seguimiento", "Tel_seguimiento", "Folio_seguimiento",
                                   "Pestanas_encontrado", "Observaciones"],
                  ["Resultado", "Cuadre_pestana"])
        with t3:
            st.caption("Financiera Cetelem, con fecha dentro de los meses cargados, que no aparecen en las capturas.")
            tabla(res["solo_periodo"], ["Hoja", "Fila_excel", "Fecha_ingreso", "Folio_norm", "Nombre",
                                        "Telefono", "Asesor", "Observaciones"])
            with st.expander("Ver también fuera del periodo, sin fecha y otras financieras"):
                tabla(res["solo"], ["Clasificacion", "Hoja", "Fila_excel", "Fecha_ingreso", "Folio_norm",
                                    "Nombre", "Telefono", "Financiera", "Asesor", "Observaciones"],
                      ["Clasificacion"])
        with t4:
            df = filtro_vendedor(res["cuadre"], "fv_seg_todo")
            tabla(df, cs.COLS_CUADRE, ["Resultado", "Cuadre_pestana"], alto=550)
        with t5:
            st.dataframe(r.astype(str), hide_index=True, width="stretch")

# ---------------------------------------------------------------- SALE-U
with tab_su:
    c1, c2 = st.columns(2)
    f_su = c1.file_uploader("Reporte de Sale-U", type=["csv", "xlsx"], key="up_su",
                            help="Ej. reporteLeadsconCredito_Personalizado.csv")
    f_cet2 = c2.file_uploader("Capturas mensuales Cetelem", type=["xlsx"], accept_multiple_files=True,
                              key="up_cet_su")

    if st.button("Hacer cuadre", type="primary", disabled=not (f_su and f_cet2), key="btn_su"):
        with st.spinner("Cruzando teléfonos y nombres…"):
            try:
                st.session_state["res_su"] = ejecutar(cu, f_su, f_cet2)
            except Exception as e:
                st.session_state.pop("res_su", None)
                st.error(f"No se pudo hacer el cuadre: {e}")

    res = st.session_state.get("res_su")
    if res:
        r = res["resumen"]
        st.success(f"Capturas: {', '.join(res['meses'])}")
        m = st.columns(5)
        m[0].metric("Clientes Cetelem", valor_resumen(r, "Clientes únicos"))
        m[1].metric("En Sale-U", valor_resumen(r, "Clientes Cetelem encontrados"))
        m[2].metric("Faltan en Sale-U", valor_resumen(r, "Clientes Cetelem NO"))
        m[3].metric("% de cuadre", valor_resumen(r, "Porcentaje"))
        m[4].metric("Estatus no cuadra", valor_resumen(r, "Estatus que no cuadra"))
        descargas(res, "su")

        t1, t2, t3, t4, t5 = st.tabs([
            f"Faltan en Sale-U ({len(res['faltantes'])})",
            f"A revisar ({len(res['revisar'])})",
            f"Solo en Sale-U ({len(res['solo_cet'])})",
            "Cuadre completo", "Resumen"])
        with t1:
            df = filtro_vendedor(res["faltantes"], "fv_su")
            if not df.empty:
                st.bar_chart(df.groupby("Vendedor").size().rename("Clientes"), color="#C00000",
                             height=220, horizontal=True)
            tabla(df, ["Cliente_Cetelem", "Tel_Cetelem", "Vendedor", "Folios", "Historial",
                       "Estatus_final_Cetelem"])
        with t2:
            tabla(res["revisar"], ["Resultado", "Cuadre_estatus", "Cliente_Cetelem", "Tel_Cetelem",
                                   "Cliente_SaleU", "Tel_SaleU", "Estatus_SaleU", "Financiera_SaleU",
                                   "Observaciones"], ["Resultado", "Cuadre_estatus"])
        with t3:
            st.caption("Registros de Cetelem en Sale-U que no aparecen en las capturas cargadas "
                       "(pueden ser de meses anteriores).")
            tabla(res["solo_cet"], ["Nombre", "Telefono", "Asesor", "Estatus_SaleU", "Financiera",
                                    "Fecha_Alta_Credito", "Observaciones"])
        with t4:
            df = filtro_vendedor(res["cuadre"], "fv_su_todo")
            tabla(df, ["Resultado", "Cuadre_estatus", "Cliente_Cetelem", "Tel_Cetelem", "Vendedor",
                       "Historial", "Cliente_SaleU", "Tel_SaleU", "Asesor_SaleU", "Estatus_SaleU",
                       "Financiera_SaleU", "Observaciones"], ["Resultado", "Cuadre_estatus"], alto=550)
        with t5:
            st.dataframe(r.astype(str), hide_index=True, width="stretch")
