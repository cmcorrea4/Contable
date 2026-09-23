import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
from io import BytesIO
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter
import datetime


if not st.session_state.get("authentication_status"):
    st.warning("⚠️ Debe iniciar sesión primero.")
    st.page_link("Inicio.py", label="Ir al login", icon="🔐")
    st.stop()

# ── Configuración de página ──────────────────────────────────────────────────
st.set_page_config(
    page_title="Anexos EEFF – ERI y ESF",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
    .stApp { background-color: #F5F7FA; }
    .header-block {
        background: linear-gradient(135deg, #1E3A5F 0%, #2E6DA4 100%);
        border-radius: 12px; padding: 28px 36px; margin-bottom: 28px; color: white;
    }
    .header-block h1 { font-size: 1.9rem; margin: 0; font-weight: 700; }
    .header-block p  { margin: 6px 0 0; opacity: .85; font-size: .95rem; }
    .metric-card {
        background: white; border-radius: 10px; padding: 20px 24px;
        box-shadow: 0 2px 8px rgba(0,0,0,.08); text-align: center; height: 100%;
    }
    .metric-card .number { font-size: 2.2rem; font-weight: 700; line-height: 1.1; }
    .metric-card .label  { font-size: .82rem; color: #666; margin-top: 4px;
                           text-transform: uppercase; letter-spacing: .05em; }
    .green  { color: #1A9E5C; } .red    { color: #D63B3B; }
    .blue   { color: #2E6DA4; } .orange { color: #E07B20; }
    .purple { color: #7C3AED; }
    .upload-section { background: white; border-radius: 10px; padding: 24px;
        box-shadow: 0 2px 8px rgba(0,0,0,.08); margin-bottom: 20px; }
    .section-title { font-size: 1rem; font-weight: 600; color: #1E3A5F; margin-bottom: 12px;
        padding-bottom: 8px; border-bottom: 2px solid #E8EDF3; }
    .result-block { background: white; border-radius: 10px; padding: 24px;
        box-shadow: 0 2px 8px rgba(0,0,0,.08); margin-top: 20px; }
    #MainMenu, footer, header { visibility: hidden; }
    .block-container { padding-top: 1.5rem; }
</style>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# CONSTANTES
# ══════════════════════════════════════════════════════════════════════════════
MESES_ORDEN = ["ENERO","FEBRERO","MARZO","ABRIL","MAYO","JUNIO",
               "JULIO","AGOSTO","SEPTIEMBRE","OCTUBRE","NOVIEMBRE","DICIEMBRE"]
MESES_ABREV = {"ENERO":"Ene","FEBRERO":"Feb","MARZO":"Mar","ABRIL":"Abr",
               "MAYO":"May","JUNIO":"Jun","JULIO":"Jul","AGOSTO":"Ago",
               "SEPTIEMBRE":"Sep","OCTUBRE":"Oct","NOVIEMBRE":"Nov","DICIEMBRE":"Dic"}

GRUPOS_ERI = ["INGRESOS DE ACTIVIDADES ORDINARIAS","COSTO DE VENTAS","OTROS INGRESOS",
              "GASTOS DE ADMINISTRACION","GASTOS DE VENTA","OTROS GASTOS",
              "INGRESOS FINANCIEROS","GASTOS FINANCIEROS","DIFERENCIA EN CAMBIO NETA",
              "PROVISION DE IMPUESTOS"]
# NOTA: "COSTO DE VENTAS", "GASTOS DE VENTA" y "DIFERENCIA EN CAMBIO NETA" son conceptos
# del catálogo CONCEPTOS_EEFF.xlsx que hoy no existen en ningún archivo de prueba real.
# Si tu script que arma la columna "Grupo" en terceros_ usa un texto distinto para estos
# conceptos, solo hay que ajustar el string aquí (y en NOMBRE_ERI más abajo) para que
# coincida exactamente.

GRUPOS_ESF_ACTIVO = [
    "EFECTIVO Y EQUIVALENTE AL EFECTIVO",
    "OTROS ACTIVOS FINANCIEROS CTE",
    "INVENTARIOS",
    "ACTIVOS POR IMPUESTOS",
    "CUENTAS COMERCIALES POR COBRAR Y OTRAS CUENTAS POR COBRAR CORRIENTES",
    "CUENTAS POR COBRAR A PARTES RELACIONADAS",
    "OTROS ACTIVOS NO FINANCIEROS",
    "OTROS ACTIVOS FINANCIEROS NO CTE",
    "ACTIVOS POR IMPUESTOS DIFERIDOS",
    "ACTIVOS INTANGIBLES DISTINTOS DE LA PLUSVALIA",
    "PROPIEDADES PLANTA Y EQUIPO",
]
GRUPOS_ESF_PASIVO = [
    "OTROS PASIVOS FINANCIEROS",
    "PASIVOS POR IMPUESTOS",
    "CUENTAS COMERCIALES POR PAGAR Y OTRAS CUENTAS POR PAGAR",
    "CUENTAS POR PAGAR A PARTES RELACIONADAS",
    "BENEFICIOS A LOS EMPLEADOS",
    "OTROS PASIVOS NO FINANCIEROS ",
    "CUENTAS COMERCIALES POR PAGAR Y OTRAS CUENTAS POR PAGAR NO CORRIENTES",
    "PASIVOS POR IMPUESTOS DIFERIDOS",
]
GRUPOS_ESF_ORDEN = GRUPOS_ESF_ACTIVO + GRUPOS_ESF_PASIVO
# NOTA: "OTROS ACTIVOS FINANCIEROS CTE", "INVENTARIOS", "OTROS ACTIVOS NO FINANCIEROS",
# "ACTIVOS INTANGIBLES DISTINTOS DE LA PLUSVALIA", "PROPIEDADES PLANTA Y EQUIPO" y
# "OTROS PASIVOS FINANCIEROS" son conceptos del catálogo CONCEPTOS_EEFF.xlsx que hoy no
# existen en ningún archivo de prueba real. Si tu script de clasificación usa otro texto
# exacto para la columna "Grupo", ajusta el string aquí (y en NOMBRE_ESF más abajo).

# ── Catch-all "Sin clasificar" ────────────────────────────────────────────────
# Cualquier cuenta de nivel de detalle (código 1x, 2x, 3x, 4x o 5x) cuyo "Grupo"
# NO esté en las listas de arriba se suma aquí. Esto GARANTIZA que Total Activos
# = Total Pasivos + Patrimonio sin importar qué conceptos tenga cada empresa,
# en vez de perder silenciosamente cuentas no clasificadas (que es lo que
# rompía el cuadre al probar con otra empresa).
OTROS_ACTIVOS_SIN_CLASIFICAR    = "OTROS ACTIVOS SIN CLASIFICAR"
OTROS_PASIVOS_SIN_CLASIFICAR    = "OTROS PASIVOS SIN CLASIFICAR"
OTRAS_PATRIMONIO_SIN_CLASIFICAR = "OTRAS PARTIDAS PATRIMONIALES SIN CLASIFICAR"
OTROS_INGRESOS_SIN_CLASIFICAR   = "OTROS INGRESOS SIN CLASIFICAR"
OTROS_GASTOS_SIN_CLASIFICAR     = "OTROS GASTOS SIN CLASIFICAR"

GRUPOS_ESF_ACTIVO.append(OTROS_ACTIVOS_SIN_CLASIFICAR)
GRUPOS_ESF_PASIVO.append(OTROS_PASIVOS_SIN_CLASIFICAR)
GRUPOS_ESF_ORDEN = GRUPOS_ESF_ACTIVO + GRUPOS_ESF_PASIVO
GRUPOS_ERI.append(OTROS_INGRESOS_SIN_CLASIFICAR)
GRUPOS_ERI.append(OTROS_GASTOS_SIN_CLASIFICAR)

GRUPOS_LABEL_ESF = {
    "EFECTIVO Y EQUIVALENTE AL EFECTIVO":"🟢 Efectivo y Equiv.",
    "OTROS ACTIVOS FINANCIEROS CTE":"🟢 Otros Activos Fin. Cte",
    "INVENTARIOS":"🟢 Inventarios",
    "CUENTAS COMERCIALES POR COBRAR Y OTRAS CUENTAS POR COBRAR CORRIENTES":"🟢 CxC Corrientes",
    "CUENTAS POR COBRAR A PARTES RELACIONADAS":"🟢 CxC Partes Rel.",
    "ACTIVOS POR IMPUESTOS":"🟢 Activos x Impuestos",
    "OTROS ACTIVOS NO FINANCIEROS":"🟢 Otros Activos No Fin.",
    "ACTIVOS POR IMPUESTOS DIFERIDOS":"🟢 Impuestos Diferidos A.",
    "OTROS ACTIVOS FINANCIEROS NO CTE":"🟢 Otros Activos Fin.",
    "ACTIVOS INTANGIBLES DISTINTOS DE LA PLUSVALIA":"🟢 Intangibles",
    "PROPIEDADES PLANTA Y EQUIPO":"🟢 PP&E",
    "OTROS PASIVOS FINANCIEROS":"🔴 Otros Pasivos Fin.",
    "CUENTAS COMERCIALES POR PAGAR Y OTRAS CUENTAS POR PAGAR":"🔴 CxP Corrientes",
    "CUENTAS COMERCIALES POR PAGAR Y OTRAS CUENTAS POR PAGAR NO CORRIENTES":"🔴 CxP No Corrientes",
    "CUENTAS POR PAGAR A PARTES RELACIONADAS":"🔴 CxP Partes Rel.",
    "PASIVOS POR IMPUESTOS":"🔴 Pasivos x Impuestos",
    "PASIVOS POR IMPUESTOS DIFERIDOS":"🔴 Impuestos Diferidos P.",
    "OTROS PASIVOS NO FINANCIEROS ":"🔴 Otros Pasivos",
    "BENEFICIOS A LOS EMPLEADOS":"🔴 Beneficios Empleados",
}
COLORES_ESF = {
    "EFECTIVO Y EQUIVALENTE AL EFECTIVO":"#2ecc71",
    "OTROS ACTIVOS FINANCIEROS CTE":"#40c977",
    "INVENTARIOS":"#66bb6a",
    "CUENTAS COMERCIALES POR COBRAR Y OTRAS CUENTAS POR COBRAR CORRIENTES":"#27ae60",
    "CUENTAS POR COBRAR A PARTES RELACIONADAS":"#1a9850",
    "ACTIVOS POR IMPUESTOS":"#52b788","ACTIVOS POR IMPUESTOS DIFERIDOS":"#74c69d",
    "OTROS ACTIVOS NO FINANCIEROS":"#86d19a",
    "OTROS ACTIVOS FINANCIEROS NO CTE":"#95d5b2",
    "ACTIVOS INTANGIBLES DISTINTOS DE LA PLUSVALIA":"#b7e4c7",
    "PROPIEDADES PLANTA Y EQUIPO":"#d8f3dc",
    "OTROS PASIVOS FINANCIEROS":"#f1948a",
    "CUENTAS COMERCIALES POR PAGAR Y OTRAS CUENTAS POR PAGAR":"#e74c3c",
    "CUENTAS COMERCIALES POR PAGAR Y OTRAS CUENTAS POR PAGAR NO CORRIENTES":"#c0392b",
    "CUENTAS POR PAGAR A PARTES RELACIONADAS":"#e57373",
    "PASIVOS POR IMPUESTOS":"#ef9a9a","PASIVOS POR IMPUESTOS DIFERIDOS":"#e67e22",
    "OTROS PASIVOS NO FINANCIEROS ":"#d35400","BENEFICIOS A LOS EMPLEADOS":"#9b59b6",
}

# Nombres para hojas formateadas (alineados con el archivo de referencia)
NOMBRE_ESF = {
    "EFECTIVO Y EQUIVALENTE AL EFECTIVO":                                   "  Efectivo y equivalentes al efectivo",
    "OTROS ACTIVOS FINANCIEROS CTE":                                        "  Otros activos financieros",
    "INVENTARIOS":                                                          "  Inventarios",
    "CUENTAS COMERCIALES POR COBRAR Y OTRAS CUENTAS POR COBRAR CORRIENTES": "  Cuentas comerciales por cobrar y otras cuentas por cobrar",
    "CUENTAS POR COBRAR A PARTES RELACIONADAS":                             "  Cuentas por cobrar a partes relacionadas",
    "ACTIVOS POR IMPUESTOS":                                                "  Activos por impuestos",
    "OTROS ACTIVOS NO FINANCIEROS":                                         "  Otros activos no financieros",
    "ACTIVOS POR IMPUESTOS DIFERIDOS":                                      "  Activos por impuestos diferidos",
    "OTROS ACTIVOS FINANCIEROS NO CTE":                                     "  Otros activos financieros",
    "ACTIVOS INTANGIBLES DISTINTOS DE LA PLUSVALIA":                        "  Activos intangibles distintos de la plusvalía",
    "PROPIEDADES PLANTA Y EQUIPO":                                          "  Propiedades, planta y equipo",
    "OTROS PASIVOS FINANCIEROS":                                            "  Otros pasivos financieros",
    "CUENTAS COMERCIALES POR PAGAR Y OTRAS CUENTAS POR PAGAR":              "  Cuentas comerciales por pagar y otras cuentas por pagar",
    "CUENTAS COMERCIALES POR PAGAR Y OTRAS CUENTAS POR PAGAR NO CORRIENTES":"  Cuentas comerciales por pagar y otras cuentas por pagar (NC)",
    "CUENTAS POR PAGAR A PARTES RELACIONADAS":                              "  Cuentas por pagar a partes relacionadas",
    "PASIVOS POR IMPUESTOS":                                                "  Pasivos por impuestos",
    "PASIVOS POR IMPUESTOS DIFERIDOS":                                      "  Pasivos por impuestos diferidos",
    "OTROS PASIVOS NO FINANCIEROS ":                                        "  Otros pasivos no financieros",
    "BENEFICIOS A LOS EMPLEADOS":                                           "  Beneficios a los empleados",
    "OTROS ACTIVOS SIN CLASIFICAR":                                         "  Otros activos sin clasificar",
    "OTROS PASIVOS SIN CLASIFICAR":                                         "  Otros pasivos sin clasificar",
}
NOMBRE_ERI = {
    "INGRESOS DE ACTIVIDADES ORDINARIAS": "Ingresos de actividades ordinarias",
    "COSTO DE VENTAS":                    "Costo de ventas",
    "OTROS INGRESOS":                     "Otros ingresos",
    "GASTOS DE ADMINISTRACION":           "Gastos de administración",
    "GASTOS DE VENTA":                    "Gastos de venta",
    "OTROS GASTOS":                       "Otros gastos",
    "INGRESOS FINANCIEROS":               "Ingresos financieros",
    "GASTOS FINANCIEROS":                 "Gastos financieros",
    "DIFERENCIA EN CAMBIO NETA":          "Diferencia en cambio neta",
    "PROVISION DE IMPUESTOS":             "Ingreso (gasto) por impuesto",
    "OTROS INGRESOS SIN CLASIFICAR":      "Otros ingresos sin clasificar",
    "OTROS GASTOS SIN CLASIFICAR":        "Otros gastos sin clasificar",
}
GRUPOS_LABEL_ERI = {
    "INGRESOS DE ACTIVIDADES ORDINARIAS":"🟢 Ing. Ordinarios","COSTO DE VENTAS":"🔴 Costo Ventas",
    "OTROS INGRESOS":"🟢 Otros Ingresos","GASTOS DE ADMINISTRACION":"🔴 Gtos. Admón.",
    "GASTOS DE VENTA":"🔴 Gtos. Venta","OTROS GASTOS":"🔴 Otros Gastos",
    "INGRESOS FINANCIEROS":"🔵 Ing. Financieros","GASTOS FINANCIEROS":"🟠 Gtos. Financieros",
    "DIFERENCIA EN CAMBIO NETA":"🟣 Dif. Cambio","PROVISION DE IMPUESTOS":"🟣 Prov. Impuestos",
}
COLORES_ERI = {
    "INGRESOS DE ACTIVIDADES ORDINARIAS":"#2ecc71","COSTO DE VENTAS":"#943126",
    "OTROS INGRESOS":"#27ae60","GASTOS DE ADMINISTRACION":"#e74c3c","GASTOS DE VENTA":"#cb4335",
    "OTROS GASTOS":"#c0392b","INGRESOS FINANCIEROS":"#3498db","GASTOS FINANCIEROS":"#e67e22",
    "DIFERENCIA EN CAMBIO NETA":"#8e44ad","PROVISION DE IMPUESTOS":"#9b59b6","Total general":"#2c3e50",
}

# ══════════════════════════════════════════════════════════════════════════════
# FUNCIONES DE DATOS
# ══════════════════════════════════════════════════════════════════════════════
def fmt_cop(val):
    if pd.isna(val): return "-"
    if val < 0: return f"($ {abs(val):,.0f})"
    return f"$ {val:,.0f}"


def _codigos_hoja(codigos):
    """
    De una lista/serie de códigos contables (como strings), devuelve el set de
    los que son 'cuenta hoja' (nivel de detalle): un código NO es hoja si algún
    OTRO código de la misma lista lo tiene como prefijo (es decir, es una
    cuenta de control/rollup que ya está representada por sus hijos). Se usa
    para sumar el balance real de la cuenta de control (p.ej. "1"=Activo total)
    sin duplicar montos.
    """
    codigos_u = sorted(set(codigos), key=len)
    hoja = set(codigos_u)
    for c in codigos_u:
        for otro in codigos_u:
            if otro != c and len(otro) > len(c) and otro.startswith(c):
                hoja.discard(c)
                break
    return hoja


def _calcular_patrimonio(df_periodo, totales_eri_periodo):
    """Extrae capital, superávit, utilidad acumulada, reservas, catch-all y
    utilidad del periodo para UN periodo dado (df ya filtrado a ese Mes)."""
    df_m = df_periodo.copy()
    df_m["c_clean"] = df_m["Codigo"].astype(str).str.strip().str.replace(r"\.0$", "", regex=True)

    def _extraer(grupo_regex, codigos_exactos, prefijo):
        g = df_m[df_m["Grupo"].astype(str).str.upper().str.contains(grupo_regex, na=False)]
        if not g.empty:
            return abs(g["Saldo Mes"].sum())
        m = df_m[df_m["c_clean"].isin(codigos_exactos)]
        if not m.empty:
            return abs(m["Saldo Mes"].iloc[0])
        m_sub = df_m[df_m["c_clean"].str.startswith(prefijo) & (df_m["c_clean"].str.len() >= 6)]
        return abs(m_sub["Saldo Mes"].sum())

    cap_emitido   = _extraer("CAPITAL EMITIDO|CAPITAL SOCIAL", ["31", "3105"], "31")
    superavit_cap = _extraer("SUPERAVIT", ["32", "3205"], "32")
    util_acum     = _extraer("UTILIDAD ACUMULADA|RESULTADOS DE EJERCICIOS ANTERIORES",
                              ["37", "3705", "36", "3605"], ("36", "37"))
    reservas      = _extraer("RESERVAS", ["33", "3305"], "33")

    # Catch-all: cualquier cuenta hoja de clase "3" (patrimonio) que no haya
    # quedado capturada en capital/superávit/utilidad acumulada/reservas.
    leaf3 = _codigos_hoja(df_m[df_m["c_clean"].str.match(r"^3")]["c_clean"])
    total_patrim_control = abs(df_m[df_m["c_clean"].isin(leaf3)]["Saldo Mes"].sum())
    otras_patrimonio = max(0.0, total_patrim_control - (cap_emitido + superavit_cap + util_acum + reservas))

    ing_ord    = abs(totales_eri_periodo.get("INGRESOS DE ACTIVIDADES ORDINARIAS", 0))
    costo_vtas = abs(totales_eri_periodo.get("COSTO DE VENTAS", 0))
    otros_ing  = abs(totales_eri_periodo.get("OTROS INGRESOS", 0))
    otros_ing_sc = abs(totales_eri_periodo.get(OTROS_INGRESOS_SIN_CLASIFICAR, 0))
    gtos_adm   = abs(totales_eri_periodo.get("GASTOS DE ADMINISTRACION", 0))
    gtos_venta = abs(totales_eri_periodo.get("GASTOS DE VENTA", 0))
    otros_gto  = abs(totales_eri_periodo.get("OTROS GASTOS", 0))
    otros_gto_sc = abs(totales_eri_periodo.get(OTROS_GASTOS_SIN_CLASIFICAR, 0))
    ing_fin    = abs(totales_eri_periodo.get("INGRESOS FINANCIEROS", 0))
    gto_fin    = abs(totales_eri_periodo.get("GASTOS FINANCIEROS", 0))
    dif_cambio = -totales_eri_periodo.get("DIFERENCIA EN CAMBIO NETA", 0)
    provision  = abs(totales_eri_periodo.get("PROVISION DE IMPUESTOS", 0))
    ganancia_bruta = ing_ord - costo_vtas
    util_ai   = (ganancia_bruta + otros_ing + otros_ing_sc - gtos_adm - gtos_venta - otros_gto - otros_gto_sc
                 + ing_fin - gto_fin + dif_cambio)
    util_per  = util_ai - provision

    total_patrimonio = cap_emitido + superavit_cap + util_acum + util_per + reservas + otras_patrimonio

    return {
        "capital_emitido": cap_emitido,
        "superavit_capital": superavit_cap,
        "utilidad_acumulada": util_acum,
        "utilidad_periodo": util_per,
        "reservas": reservas,
        "otras_partidas": otras_patrimonio,
        "total_patrimonio": total_patrimonio,
    }


@st.cache_data(show_spinner="Procesando archivo…")
def procesar_archivo(file_bytes: bytes):
    xls = pd.ExcelFile(BytesIO(file_bytes))
    if "terceros_" not in xls.sheet_names:
        st.error("El archivo no contiene la hoja 'terceros_'.")
        st.stop()

    df = pd.read_excel(BytesIO(file_bytes), sheet_name="terceros_", header=0)
    df.columns = [c.strip() for c in df.columns]
    df["Saldo Mes"] = pd.to_numeric(df["Saldo Mes"], errors="coerce").fillna(0)
    df["c_clean"] = df["Codigo"].astype(str).str.strip().str.replace(r"\.0$", "", regex=True)

    # ── Reclasificación catch-all ────────────────────────────────────────────
    # Cualquier cuenta HOJA (de detalle, sin hijos) de clase 1/2/4/5 cuyo Grupo
    # no esté en las listas conocidas se reetiqueta a "...SIN CLASIFICAR", para
    # que nunca desaparezca del total y el balance siempre cuadre.
    _GA = GRUPOS_ESF_ACTIVO[:-1]   # sin el catch-all, para el isin()
    _GP = GRUPOS_ESF_PASIVO[:-1]
    _GI = ["INGRESOS DE ACTIVIDADES ORDINARIAS","COSTO DE VENTAS","OTROS INGRESOS","INGRESOS FINANCIEROS","DIFERENCIA EN CAMBIO NETA"]
    _GG = ["GASTOS DE ADMINISTRACION","GASTOS DE VENTA","OTROS GASTOS","GASTOS FINANCIEROS","PROVISION DE IMPUESTOS"]
    for clase, grupos_conocidos, etiqueta in [
        ("1", _GA, OTROS_ACTIVOS_SIN_CLASIFICAR),
        ("2", _GP, OTROS_PASIVOS_SIN_CLASIFICAR),
        ("4", _GI, OTROS_INGRESOS_SIN_CLASIFICAR),
        ("5", _GG, OTROS_GASTOS_SIN_CLASIFICAR),
    ]:
        de_esta_clase = df["c_clean"].str.match(r"^\d+$", na=False) & df["c_clean"].str.startswith(clase)
        leaf = _codigos_hoja(df.loc[de_esta_clase, "c_clean"])
        # Un código de detalle está "cubierto" si ÉL MISMO o algún ANCESTRO
        # suyo (un código más corto que sea prefijo) ya tiene un Grupo válido
        # en alguna fila. Es común que el Grupo se etiquete en un nivel
        # intermedio (p.ej. la cuenta "112005") mientras el desglose por
        # tercero más profundo ("11200505") no repite la etiqueta: ese
        # detalle NO debe ir al catch-all, porque su saldo ya está sumado
        # a través del código padre.
        tagged = sorted(set(df.loc[df["Grupo"].isin(grupos_conocidos), "c_clean"]), key=len)
        def _cubierto(codigo, _tagged=tagged):
            return any(codigo.startswith(t) for t in _tagged)
        leaf_no_cubierto = {c for c in leaf if not _cubierto(c)}
        sin_clasificar = de_esta_clase & df["c_clean"].isin(leaf_no_cubierto)
        df.loc[sin_clasificar, "Grupo"] = etiqueta

    def build_pivot(df_src, grupos):
        df_f = df_src[df_src["Grupo"].isin(grupos)].copy()
        pivot = df_f.groupby(["Grupo","Mes"])["Saldo Mes"].sum().unstack(fill_value=0)
        meses_d = [m for m in MESES_ORDEN if m in pivot.columns]
        pivot = pivot.reindex(columns=meses_d, fill_value=0)
        # Reindex con TODOS los grupos posibles (no solo los que ya tienen datos),
        # para que conceptos sin movimiento en este archivo aparezcan igual con $0.
        pivot = pivot.reindex(grupos, fill_value=0)
        pivot["Total general"] = pivot.sum(axis=1)
        total_row = pivot.sum().to_frame().T
        total_row.index = ["Total general"]
        return pd.concat([pivot, total_row]), df_f

    pivot_eri, df_eri_raw = build_pivot(df, GRUPOS_ERI)

    # ESF
    df_esf_raw = df[df["Grupo"].isin(GRUPOS_ESF_ORDEN)].copy()
    df_esf_raw["Saldo Mes"] = pd.to_numeric(df_esf_raw["Saldo Mes"], errors="coerce").fillna(0)
    pivot_base = df_esf_raw.groupby(["Grupo","Mes"])["Saldo Mes"].sum().unstack(fill_value=0)
    meses_d = [m for m in MESES_ORDEN if m in pivot_base.columns]
    pivot_base = pivot_base.reindex(columns=meses_d, fill_value=0)
    # Igual que en ERI: reindex con TODOS los grupos, no solo los presentes,
    # así los conceptos sin saldo en este archivo quedan en $0 en vez de desaparecer.
    pivot_base = pivot_base.reindex(GRUPOS_ESF_ORDEN, fill_value=0)
    pivot_base["Total general"] = pivot_base.sum(axis=1)

    act_r = [g for g in GRUPOS_ESF_ACTIVO if g in pivot_base.index]
    pas_r = [g for g in GRUPOS_ESF_PASIVO if g in pivot_base.index]
    ta = pivot_base.loc[act_r].sum().to_frame().T; ta.index = ["Total Activo"]
    tp = pivot_base.loc[pas_r].sum().to_frame().T; tp.index = ["Total Pasivo"]
    pivot_esf = pd.concat([pivot_base.loc[act_r], ta, pivot_base.loc[pas_r], tp])

    # Saldo de cierre (último mes) para hoja ESF formateada
    ultimo_mes = meses_d[-1] if meses_d else None
    saldos_esf = {}
    if ultimo_mes:
        for g in GRUPOS_ESF_ORDEN:
            saldos_esf[g] = pivot_base.loc[g, ultimo_mes] if g in pivot_base.index else 0.0

    # Totales ERI acumulados
    totales_eri = {g: pivot_eri.loc[g, "Total general"] if g in pivot_eri.index else 0.0
                   for g in GRUPOS_ERI}

    # Extracción de valores de Patrimonio (último mes)
    saldos_patrimonio = {
        "capital_emitido": 0.0, "superavit_capital": 0.0, "utilidad_acumulada": 0.0,
        "reservas": 0.0, "otras_partidas": 0.0, "utilidad_periodo": 0.0, "total_patrimonio": 0.0,
    }
    if ultimo_mes:
        df_m = df[df["Mes"] == ultimo_mes]
        saldos_patrimonio = _calcular_patrimonio(df_m, totales_eri)

    # ── Período anterior ─────────────────────────────────────────────────────
    # En la hoja terceros_, un valor de "Mes" igual a "PERIODO ANTERIOR" (no
    # sensible a mayúsculas/espacios) representa el corte anterior a comparar.
    # Se calculan sus mismos saldos/totales para diligenciar las columnas
    # "Período anterior" del ESF y del ERI.
    mes_norm = df["Mes"].astype(str).str.strip().str.upper()
    df_ant = df[mes_norm == "PERIODO ANTERIOR"]
    saldos_esf_anterior = None
    totales_eri_anterior = None
    saldos_patrimonio_anterior = None
    if not df_ant.empty:
        saldos_esf_anterior = {g: 0.0 for g in GRUPOS_ESF_ORDEN}
        totales_eri_anterior = {g: 0.0 for g in GRUPOS_ERI}
        g_esf_ant = df_ant[df_ant["Grupo"].isin(GRUPOS_ESF_ORDEN)].groupby("Grupo")["Saldo Mes"].sum()
        saldos_esf_anterior.update(g_esf_ant.to_dict())
        g_eri_ant = df_ant[df_ant["Grupo"].isin(GRUPOS_ERI)].groupby("Grupo")["Saldo Mes"].sum()
        totales_eri_anterior.update(g_eri_ant.to_dict())
        saldos_patrimonio_anterior = _calcular_patrimonio(df_ant, totales_eri_anterior)

    return (df_eri_raw, pivot_eri, df_esf_raw, pivot_esf, saldos_esf, saldos_patrimonio, totales_eri, ultimo_mes,
            saldos_esf_anterior, totales_eri_anterior, saldos_patrimonio_anterior)


# ══════════════════════════════════════════════════════════════════════════════
# GENERACIÓN EXCEL FORMATEADO
# ══════════════════════════════════════════════════════════════════════════════
FMT_COP_XL  = '_-"$" * #,##0_-;\\-"$" * #,##0_-;_-"$" * "-"_-;_-@_-'
FMT_SUB_XL  = '#,##0_);\\(#,##0\\);"-       "'

def F(bold=False, size=13, name="Calibri"):
    return Font(bold=bold, size=size, name=name)

def A(h="general", v="center"):
    return Alignment(horizontal=h, vertical=v)

def thin():   return Side(border_style="thin")
def double(): return Side(border_style="double")

def set_val(ws, cell_ref, value, bold=False, size=13, fmt=None,
            halign="general", border_bottom=None):
    c = ws[cell_ref]
    c.value = value if value != 0 else (value if value is not None else None)
    c.font  = F(bold=bold, size=size)
    if halign != "general": c.alignment = A(halign)
    if fmt: c.number_format = fmt
    if border_bottom == "thin":   c.border = Border(bottom=thin())
    if border_bottom == "double": c.border = Border(bottom=double())


def generar_hoja_esf(ws, empresa, nit, periodo, saldos, saldos_patrimonio=None,
                      saldos_anterior=None, saldos_patrimonio_anterior=None):
    if saldos_patrimonio is None:
        saldos_patrimonio = {}
    tiene_anterior = bool(saldos_anterior) or bool(saldos_patrimonio_anterior)
    if saldos_anterior is None:
        saldos_anterior = {}
    if saldos_patrimonio_anterior is None:
        saldos_patrimonio_anterior = {}

    anchos = {"A":4,"B":49.6,"C":6.1,"D":17.7,"E":2.6,"F":17.7,
              "G":2.6,"H":12.1,"I":49.6,"J":6.1,"K":17.7,"L":2.6,"M":17.7,"N":12}
    for col, w in anchos.items():
        ws.column_dimensions[col].width = w
    for r in range(1, 42):
        ws.row_dimensions[r].height = 15.0
    ws.row_dimensions[1].height = 26.1
    for r in [2,3,4,5]: ws.row_dimensions[r].height = 20.1

    # Encabezado
    for r, txt, bold in [(1,empresa,True),(2,nit,False),
                         (3,"ESTADO DE SITUACIÓN FINANCIERA",True),
                         (4,periodo,False),(5,"(En pesos colombianos - $)",False)]:
        c = ws[f"B{r}"]; c.value = txt; c.font = F(bold=bold, size=18)

    # Cabecera columnas fila 7-8
    for col, val in [("C","NOTA"),("D",periodo),("F","Período anterior"),
                     ("J","NOTA"),("K",periodo),("M","Período anterior")]:
        ws[f"{col}7"].value = val
        ws[f"{col}7"].font  = F(bold=True, size=13)
        ws[f"{col}7"].alignment = A("center")
    for col in ["D","F","K","M"]:
        ws[f"{col}8"].value = "$"; ws[f"{col}8"].font = F(bold=True,size=13)
        ws[f"{col}8"].alignment = A("center")
    if not tiene_anterior:
        # Nota visible si el archivo no trae un "Mes" = "PERIODO ANTERIOR" en terceros_
        ws["F5"].value = "(Período anterior: agrega en terceros_ un Mes = \"PERIODO ANTERIOR\" para diligenciarlo)"
        ws["F5"].font = Font(italic=True, size=9, color="9AA5B1")

    # Títulos sección
    ws["B9"].value = "ACTIVOS";              ws["B9"].font = F(bold=True, size=14)
    ws["I9"].value = "PASIVOS Y PATRIMONIO"; ws["I9"].font = F(bold=True, size=14)
    ws["B11"].value = "Activos corrientes:"; ws["B11"].font = F(size=13)
    ws["I11"].value = "Pasivos corrientes:"; ws["I11"].font = F(size=13)

    # Siempre devuelve el valor (aunque sea 0). El formato contable FMT_COP_XL
    # ya se encarga de mostrar "-" cuando el saldo es cero, para que el concepto
    # aparezca en la plantilla en vez de quedar oculto.
    def v(g): return abs(saldos.get(g, 0))
    def v_ant(g): return abs(saldos_anterior.get(g, 0))

    def _escribir_fila(row_n, col_actual, col_anterior, label_col, g, nombre=None):
        ws[f"{label_col}{row_n}"].value = nombre if nombre else NOMBRE_ESF[g]
        ws[f"{label_col}{row_n}"].font = F(size=13); ws[f"{label_col}{row_n}"].alignment = A("left")
        ws[f"{col_actual}{row_n}"].value = v(g); ws[f"{col_actual}{row_n}"].number_format = FMT_COP_XL
        ws[f"{col_actual}{row_n}"].font = F(size=13)
        ws[f"{col_anterior}{row_n}"].value = v_ant(g); ws[f"{col_anterior}{row_n}"].number_format = FMT_COP_XL
        ws[f"{col_anterior}{row_n}"].font = F(size=13)

    # ── Activos corrientes (col B/D, anterior en F) ───────────────────────────
    act_cte_map = [
        (12, "EFECTIVO Y EQUIVALENTE AL EFECTIVO"),
        (13, "OTROS ACTIVOS FINANCIEROS CTE"),
        (14, "ACTIVOS POR IMPUESTOS"),
        (15, "CUENTAS COMERCIALES POR COBRAR Y OTRAS CUENTAS POR COBRAR CORRIENTES"),
        (16, "CUENTAS POR COBRAR A PARTES RELACIONADAS"),
        (17, "INVENTARIOS"),
        (18, "OTROS ACTIVOS NO FINANCIEROS"),
    ]
    for row_n, g in act_cte_map:
        _escribir_fila(row_n, "D", "F", "B", g)

    total_act_cte = sum(abs(saldos.get(g,0)) for _,g in act_cte_map)
    total_act_cte_ant = sum(abs(saldos_anterior.get(g,0)) for _,g in act_cte_map)
    ws["B19"].value = "Total activos corrientes"; ws["B19"].font = F(size=13)
    for col, val in [("D",total_act_cte), ("F",total_act_cte_ant)]:
        ws[f"{col}19"].value = val; ws[f"{col}19"].number_format = FMT_SUB_XL
        ws[f"{col}19"].font = F(size=13); ws[f"{col}19"].border = Border(bottom=thin())

    # ── Activos no corrientes (col B/D, anterior en F) ────────────────────────
    ws["B24"].value = "Activos no corrientes:"; ws["B24"].font = F(size=13)
    act_nct_map = [
        (25, "OTROS ACTIVOS FINANCIEROS NO CTE"),
        (26, "ACTIVOS POR IMPUESTOS DIFERIDOS"),
        (27, "ACTIVOS INTANGIBLES DISTINTOS DE LA PLUSVALIA"),
        (28, "PROPIEDADES PLANTA Y EQUIPO"),
        (29, "OTROS ACTIVOS SIN CLASIFICAR"),
    ]
    for row_n, g in act_nct_map:
        _escribir_fila(row_n, "D", "F", "B", g)

    total_act_nct = sum(abs(saldos.get(g,0)) for _,g in act_nct_map)
    total_act_nct_ant = sum(abs(saldos_anterior.get(g,0)) for _,g in act_nct_map)
    ws["B30"].value = "Total activos no corrientes"; ws["B30"].font = F(size=13)
    for col, val in [("D",total_act_nct), ("F",total_act_nct_ant)]:
        ws[f"{col}30"].value = val; ws[f"{col}30"].number_format = FMT_SUB_XL
        ws[f"{col}30"].font = F(size=13); ws[f"{col}30"].border = Border(bottom=thin())

    total_act = total_act_cte + total_act_nct
    total_act_ant = total_act_cte_ant + total_act_nct_ant
    ws["B38"].value = "TOTAL ACTIVOS"; ws["B38"].font = F(bold=True, size=14)
    for col, val in [("D",total_act), ("F",total_act_ant)]:
        ws[f"{col}38"].value = val; ws[f"{col}38"].number_format = FMT_COP_XL
        ws[f"{col}38"].font = F(bold=True, size=14); ws[f"{col}38"].border = Border(bottom=double())

    # ── Pasivos corrientes (col I/K, anterior en M) ───────────────────────────
    pas_cte_map = [
        (12, "OTROS PASIVOS NO FINANCIEROS "),
        (13, "OTROS PASIVOS FINANCIEROS"),
        (14, "PASIVOS POR IMPUESTOS"),
        (15, "CUENTAS COMERCIALES POR PAGAR Y OTRAS CUENTAS POR PAGAR"),
        (16, "CUENTAS POR PAGAR A PARTES RELACIONADAS"),
        (18, "BENEFICIOS A LOS EMPLEADOS"),
        (19, "OTROS PASIVOS SIN CLASIFICAR"),
    ]
    for row_n, g in pas_cte_map:
        _escribir_fila(row_n, "K", "M", "I", g)

    total_pas_cte = sum(abs(saldos.get(g,0)) for _,g in pas_cte_map)
    total_pas_cte_ant = sum(abs(saldos_anterior.get(g,0)) for _,g in pas_cte_map)
    ws["I20"].value = "Total pasivos corrientes"; ws["I20"].font = F(size=13)
    for col, val in [("K",total_pas_cte), ("M",total_pas_cte_ant)]:
        ws[f"{col}20"].value = val; ws[f"{col}20"].number_format = FMT_SUB_XL
        ws[f"{col}20"].font = F(size=13); ws[f"{col}20"].border = Border(bottom=thin())

    # ── Pasivos no corrientes (col I/K, anterior en M) ────────────────────────
    ws["I22"].value = "Pasivos no corrientes:"; ws["I22"].font = F(size=13)
    pas_nct_map = [
        (23, "CUENTAS COMERCIALES POR PAGAR Y OTRAS CUENTAS POR PAGAR NO CORRIENTES"),
        (24, "PASIVOS POR IMPUESTOS DIFERIDOS"),
    ]
    for row_n, g in pas_nct_map:
        _escribir_fila(row_n, "K", "M", "I", g)

    total_pas_nct = sum(abs(saldos.get(g,0)) for _,g in pas_nct_map)
    total_pas_nct_ant = sum(abs(saldos_anterior.get(g,0)) for _,g in pas_nct_map)
    ws["I26"].value = "Total pasivos no corrientes"; ws["I26"].font = F(size=13)
    for col, val in [("K",total_pas_nct), ("M",total_pas_nct_ant)]:
        ws[f"{col}26"].value = val; ws[f"{col}26"].number_format = FMT_SUB_XL
        ws[f"{col}26"].font = F(size=13); ws[f"{col}26"].border = Border(bottom=thin())

    total_pas = total_pas_cte + total_pas_nct
    total_pas_ant = total_pas_cte_ant + total_pas_nct_ant
    ws["I28"].value = "TOTAL PASIVOS"; ws["I28"].font = F(bold=True, size=14)
    for col, val in [("K",total_pas), ("M",total_pas_ant)]:
        ws[f"{col}28"].value = val; ws[f"{col}28"].number_format = FMT_COP_XL
        ws[f"{col}28"].font = F(bold=True, size=14); ws[f"{col}28"].border = Border(bottom=double())

    # ── Patrimonio (col I/K, anterior en M) ───────────────────────────────────
    ws["I30"].value = "Patrimonio:"; ws["I30"].font = F(size=13)
    cap_emitido  = saldos_patrimonio.get("capital_emitido", 0.0)
    superavit    = saldos_patrimonio.get("superavit_capital", 0.0)
    util_acum    = saldos_patrimonio.get("utilidad_acumulada", 0.0)
    util_periodo = saldos_patrimonio.get("utilidad_periodo", 0.0)
    reservas     = saldos_patrimonio.get("reservas", 0.0)
    otras_patr   = saldos_patrimonio.get("otras_partidas", 0.0)
    total_patrimonio = saldos_patrimonio.get(
        "total_patrimonio", cap_emitido + superavit + util_acum + util_periodo + reservas + otras_patr)

    cap_a  = saldos_patrimonio_anterior.get("capital_emitido", 0.0)
    sup_a  = saldos_patrimonio_anterior.get("superavit_capital", 0.0)
    ua_a   = saldos_patrimonio_anterior.get("utilidad_acumulada", 0.0)
    up_a   = saldos_patrimonio_anterior.get("utilidad_periodo", 0.0)
    res_a  = saldos_patrimonio_anterior.get("reservas", 0.0)
    otr_a  = saldos_patrimonio_anterior.get("otras_partidas", 0.0)
    total_patrimonio_ant = saldos_patrimonio_anterior.get(
        "total_patrimonio", cap_a + sup_a + ua_a + up_a + res_a + otr_a)

    patrimonio_map = [
        (31, "  Capital emitido", cap_emitido, cap_a),
        (32, "  Superávit de capital", superavit, sup_a),
        (33, "  Utilidad acumulada", util_acum, ua_a),
        (34, "  Utilidad del periodo", util_periodo, up_a),
        (35, "  Reservas", reservas, res_a),
        (36, "  Otras partidas patrimoniales", otras_patr, otr_a),
    ]

    for r, lbl, val, val_a in patrimonio_map:
        ws[f"I{r}"].value = lbl; ws[f"I{r}"].font = F(size=13); ws[f"I{r}"].alignment = A("left")
        ws[f"K{r}"].value = val; ws[f"K{r}"].number_format = FMT_COP_XL; ws[f"K{r}"].font = F(size=13)
        ws[f"M{r}"].value = val_a; ws[f"M{r}"].number_format = FMT_COP_XL; ws[f"M{r}"].font = F(size=13)

    ws["I37"].value = "TOTAL PATRIMONIO"; ws["I37"].font = F(bold=True, size=14)
    for col, val in [("K",total_patrimonio), ("M",total_patrimonio_ant)]:
        ws[f"{col}37"].value = val if val != 0 else None
        ws[f"{col}37"].number_format = FMT_SUB_XL
        ws[f"{col}37"].font = F(bold=True, size=14); ws[f"{col}37"].border = Border(bottom=thin())

    total_pas_patrimonio = total_pas + total_patrimonio
    total_pas_patrimonio_ant = total_pas_ant + total_patrimonio_ant
    ws["I38"].value = "TOTAL PASIVOS Y PATRIMONIO"; ws["I38"].font = F(bold=True, size=14)
    for col, val in [("K",total_pas_patrimonio), ("M",total_pas_patrimonio_ant)]:
        ws[f"{col}38"].value = val; ws[f"{col}38"].number_format = FMT_COP_XL
        ws[f"{col}38"].font = F(bold=True, size=14); ws[f"{col}38"].border = Border(bottom=double())

    ws["B40"].value = "Las notas adjuntas forman parte integral de estos estados financieros."
    ws["B40"].font = F(size=11)


def _calc_lineas_eri(totales):
    """Calcula todas las líneas del ERI a partir del dict de totales por Grupo.
    Reutilizable para el periodo actual y para el periodo anterior."""
    ing_ord     = abs(totales.get("INGRESOS DE ACTIVIDADES ORDINARIAS", 0))
    costo_vtas  = abs(totales.get("COSTO DE VENTAS", 0))
    otros_ing   = abs(totales.get("OTROS INGRESOS", 0))
    otros_ing_sc= abs(totales.get(OTROS_INGRESOS_SIN_CLASIFICAR, 0))
    gtos_adm    = abs(totales.get("GASTOS DE ADMINISTRACION", 0))
    gtos_venta  = abs(totales.get("GASTOS DE VENTA", 0))
    otros_gto   = abs(totales.get("OTROS GASTOS", 0))
    otros_gto_sc= abs(totales.get(OTROS_GASTOS_SIN_CLASIFICAR, 0))
    ing_fin     = abs(totales.get("INGRESOS FINANCIEROS", 0))
    gto_fin     = abs(totales.get("GASTOS FINANCIEROS", 0))
    dif_cambio  = -totales.get("DIFERENCIA EN CAMBIO NETA", 0)
    provision   = abs(totales.get("PROVISION DE IMPUESTOS", 0))
    ganancia    = ing_ord - costo_vtas
    util_ai     = (ganancia + otros_ing + otros_ing_sc - gtos_adm - gtos_venta - otros_gto - otros_gto_sc
                   + ing_fin - gto_fin + dif_cambio)
    util_per    = util_ai - provision
    return {
        10: ing_ord, 11: -costo_vtas, 12: ganancia, 13: -gtos_venta, 14: otros_ing,
        15: -gtos_adm, 16: -otros_gto, 17: ing_fin, 18: -gto_fin, 19: dif_cambio,
        20: otros_ing_sc, 21: -otros_gto_sc, 22: util_ai, 24: -provision,
        26: util_per, 29: util_per,
    }


def generar_hoja_eri(ws, empresa, nit, periodo, totales, totales_anterior=None):
    tiene_anterior = bool(totales_anterior)
    if totales_anterior is None:
        totales_anterior = {}
    anchos = {"A":0.9,"B":43.9,"C":6.1,"D":16.6,"E":2.6,"F":16.6,"G":9.0}
    for col, w in anchos.items():
        ws.column_dimensions[col].width = w
    for r in range(1, 35):
        ws.row_dimensions[r].height = 15.0
    for r, h in [(1,26.1),(2,20.1),(3,20.1),(4,20.1),(5,20.1),(6,18.0)]:
        ws.row_dimensions[r].height = h

    for r, txt, bold in [(1,empresa,True),(2,nit,False),
                         (3,"ESTADO DE RESULTADOS INTEGRAL",True),
                         (4,periodo,False),(5,"(En pesos colombianos - $)",False)]:
        c = ws[f"B{r}"]; c.value = txt; c.font = F(bold=bold, size=18)
    if not tiene_anterior:
        ws["F5"].value = "(Período anterior: agrega en terceros_ un Mes = \"PERIODO ANTERIOR\" para diligenciarlo)"
        ws["F5"].font = Font(italic=True, size=9, color="9AA5B1")

    for col, val in [("D","Acumulado"),("F","Período anterior")]:
        ws[f"{col}6"].value = val; ws[f"{col}6"].font = F(bold=True, size=14)
        ws[f"{col}6"].alignment = A("center")

    ws["C7"].value = "NOTA"; ws["C7"].font = F(bold=True, size=13); ws["C7"].alignment = A("center")
    for col, val in [("D",2025),("F",2024)]:
        ws[f"{col}7"].value = val; ws[f"{col}7"].font = F(bold=True, size=13)
        ws[f"{col}7"].alignment = A("center")
    for col in ["D","F"]:
        ws[f"{col}8"].value = "$"; ws[f"{col}8"].font = F(bold=True, size=13)
        ws[f"{col}8"].alignment = A("center")

    vals = _calc_lineas_eri(totales)
    vals_ant = _calc_lineas_eri(totales_anterior)

    lineas = [
        (10, "Ingresos de actividades ordinarias", 13,   False, None),
        (11, "Costo de ventas",                    None, False, None),
        (12, "Ganancia bruta",                      None, True,  None),
        (13, "Gastos de venta",                     None, False, None),
        (14, "Otros ingresos",                      14,  False, None),
        (15, "Gastos de administración",            15, False, None),
        (16, "Otros gastos",                        16, False, None),
        (17, "Ingresos financieros",                17,  False, None),
        (18, "Gastos financieros",                  16, False, None),
        (19, "Diferencia en cambio neta",           None, False, None),
        (20, "Otros ingresos sin clasificar",       None, False, None),
        (21, "Otros gastos sin clasificar",         None, False, None),
        (22, "Utilidad antes de impuesto",          None, True,  None),
        (24, "Ingreso (gasto) por impuesto",        19, False, None),
        (26, "Utilidad (pérdida) del periodo",      None, True,  None),
        (29, "Resultado integral total",            None, True,  "double"),
    ]

    for row_n, label, nota, bold, border in lineas:
        ws[f"B{row_n}"].value = label; ws[f"B{row_n}"].font = F(bold=bold, size=13)
        if nota:
            ws[f"C{row_n}"].value = nota; ws[f"C{row_n}"].font = F(size=13)
            ws[f"C{row_n}"].alignment = A("center")
        # Se muestra siempre el valor (incluido 0) para que el concepto aparezca
        # en la plantilla; FMT_COP_XL despliega "-" cuando el saldo es cero.
        for col, valores in [("D", vals), ("F", vals_ant)]:
            c = ws[f"{col}{row_n}"]
            c.value = valores.get(row_n, 0)
            c.number_format = FMT_COP_XL; c.font = F(bold=bold, size=13)
            if border == "double": c.border = Border(bottom=double())

    ws["B31"].value = "Las notas adjuntas forman parte integral de estos estados financieros."
    ws["B31"].font = F(size=12)


def generar_hoja_anexo_desplegable(ws, title, empresa, df_raw, grupos_orden, meses_cols):
    """
    Genera un Anexo ESF/ERI DESPLEGABLE: una fila resumen (bold) por cada Grupo,
    con las cuentas de detalle agrupadas debajo usando el 'outline' nativo de
    Excel (los botones +/- para expandir/colapsar aparecen a la izquierda de
    las filas). Los grupos sin movimiento en este archivo igual aparecen, en $0.
    """
    c_header_fill  = PatternFill(start_color="1E3A5F", end_color="1E3A5F", fill_type="solid")
    c_subhead_fill = PatternFill(start_color="2E6DA4", end_color="2E6DA4", fill_type="solid")
    c_grupo_fill   = PatternFill(start_color="D9E6F2", end_color="D9E6F2", fill_type="solid")
    c_alt_fill     = PatternFill(start_color="F4F8FA", end_color="F4F8FA", fill_type="solid")
    c_total_fill   = PatternFill(start_color="1E3A5F", end_color="1E3A5F", fill_type="solid")

    font_title  = Font(name="Calibri", size=14, bold=True, color="FFFFFF")
    font_header = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    font_grupo  = Font(name="Calibri", size=11, bold=True, color="1E3A5F")
    font_data   = Font(name="Calibri", size=10.5, color="425466")
    font_total  = Font(name="Calibri", size=11, bold=True, color="FFFFFF")

    border_thin = Border(
        left=Side(style="thin", color="E2E8F0"), right=Side(style="thin", color="E2E8F0"),
        top=Side(style="thin", color="E2E8F0"), bottom=Side(style="thin", color="E2E8F0"))
    border_grupo = Border(top=Side(style="thin", color="1E3A5F"))
    border_total = Border(top=Side(style="thin", color="1E3A5F"), bottom=Side(style="double", color="FFFFFF"))
    align_left  = Alignment(horizontal="left", vertical="center")
    align_right = Alignment(horizontal="right", vertical="center")
    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)

    cols = list(meses_cols) + ["Total general"]
    num_cols = len(cols) + 3  # Grupo, Nombre cuenta, Nombre tercero + meses
    last_col_letter = get_column_letter(num_cols)
    COL_GRUPO, COL_CUENTA, COL_TERCERO = 1, 2, 3
    COL_MESES_INICIO = 4

    # Habilitar el outline con el resumen ARRIBA de los detalles (para que el
    # Grupo quede visible y las cuentas se desplieguen hacia abajo).
    ws.sheet_properties.outlinePr.summaryBelow = False
    ws.sheet_properties.outlinePr.summaryRight = False

    # Fila 1: título
    ws.row_dimensions[1].height = 28.0
    ws.cell(row=1, column=1).value = f"{empresa.upper()} - {title.upper()}"
    ws.cell(row=1, column=1).font = font_title
    ws.cell(row=1, column=1).alignment = align_left
    for c in range(1, num_cols + 1):
        ws.cell(row=1, column=c).fill = c_header_fill
    ws.row_dimensions[2].height = 10.0
    ws.cell(row=3, column=1).value = "Haz clic en los botones [+] / [-] de la izquierda para desplegar el detalle de cuentas y terceros de cada concepto."
    ws.cell(row=3, column=1).font = Font(name="Calibri", size=9.5, italic=True, color="708090")

    # Fila 4: cabecera
    ws.row_dimensions[4].height = 26.0
    headers = ["Grupo", "Nombre cuenta", "Nombre tercero"] + cols
    for col_idx, h in enumerate(headers, 1):
        c = ws.cell(row=4, column=col_idx)
        c.value = h; c.font = font_header; c.fill = c_subhead_fill; c.alignment = align_center

    # Agregación de detalle por Grupo + Cuenta + Tercero (3er nivel, como en el
    # archivo de referencia: Grupo -> Cuenta -> Tercero -> "Total <cuenta>").
    if not df_raw.empty:
        df_raw = df_raw.copy()
        # pivot_table descarta silenciosamente las filas cuyo nivel de índice
        # es NaN; como muchas cuentas no tienen desglose por tercero, hay que
        # rellenar antes de pivotear o esas cuentas desaparecen del Anexo.
        df_raw["Nombre tercero"] = df_raw["Nombre tercero"].fillna("(en blanco)").astype(str).str.strip()
        df_raw.loc[df_raw["Nombre tercero"] == "", "Nombre tercero"] = "(en blanco)"
        det = (df_raw.groupby(["Grupo", "Codigo", "Nombre cuenta", "Nombre tercero", "Mes"], dropna=False)["Saldo Mes"]
                .sum().reset_index())
        pivot_det = det.pivot_table(index=["Grupo", "Codigo", "Nombre cuenta", "Nombre tercero"],
                                     columns="Mes", values="Saldo Mes", fill_value=0)
        pivot_det = pivot_det.reindex(columns=meses_cols, fill_value=0)
        pivot_det["Total general"] = pivot_det.sum(axis=1)
    else:
        pivot_det = pd.DataFrame(columns=meses_cols + ["Total general"])

    row = 5
    total_general = pd.Series(0.0, index=cols)
    font_cuenta = Font(name="Calibri", size=10.5, bold=True, color="1E3A5F")
    font_total_cuenta = Font(name="Calibri", size=10.5, bold=True, italic=True, color="1E3A5F")

    for grupo in grupos_orden:
        try:
            sub_grupo = pivot_det.xs(grupo, level="Grupo")
        except KeyError:
            sub_grupo = pd.DataFrame(columns=cols)

        grupo_vals = sub_grupo.sum(axis=0) if not sub_grupo.empty else pd.Series(0.0, index=cols)
        for c in cols:
            if c not in grupo_vals.index:
                grupo_vals[c] = 0.0
        total_general = total_general.add(grupo_vals, fill_value=0)

        # Fila resumen del Grupo (siempre visible, nivel 0)
        cA = ws.cell(row=row, column=COL_GRUPO)
        cA.value = NOMBRE_ESF.get(grupo, NOMBRE_ERI.get(grupo, grupo)).strip()
        cA.font = font_grupo; cA.fill = c_grupo_fill; cA.border = border_grupo; cA.alignment = align_left
        for c_idx in (COL_CUENTA, COL_TERCERO):
            ws.cell(row=row, column=c_idx).fill = c_grupo_fill
            ws.cell(row=row, column=c_idx).border = border_grupo
        for i, c in enumerate(cols):
            cell = ws.cell(row=row, column=COL_MESES_INICIO + i)
            cell.value = abs(float(grupo_vals.get(c, 0)))
            cell.number_format = FMT_COP_XL; cell.font = font_grupo
            cell.fill = c_grupo_fill; cell.border = border_grupo; cell.alignment = align_right
        ws.row_dimensions[row].height = 20.0
        ws.row_dimensions[row].outlineLevel = 0
        row += 1

        if sub_grupo.empty:
            continue

        # Cuentas dentro del Grupo, ordenadas por código
        cuentas = sorted(sub_grupo.index.droplevel("Nombre tercero").unique(),
                          key=lambda x: str(x[0]))
        for codigo, nombre in cuentas:
            sub_cuenta = sub_grupo.xs((codigo, nombre), level=("Codigo", "Nombre cuenta"))
            cod_txt = str(codigo).rstrip("0").rstrip(".") if isinstance(codigo, float) else str(codigo)

            # Fila CUENTA (siempre visible, nivel 1 — el ⊟/⊞ despliega los terceros)
            ws.cell(row=row, column=COL_CUENTA).value = f"{cod_txt} · {nombre}"
            ws.cell(row=row, column=COL_CUENTA).font = font_cuenta
            ws.cell(row=row, column=COL_CUENTA).alignment = align_left
            ws.cell(row=row, column=COL_CUENTA).border = border_thin
            ws.cell(row=row, column=COL_GRUPO).border = border_thin
            ws.cell(row=row, column=COL_TERCERO).border = border_thin
            for i, c in enumerate(cols):
                cell = ws.cell(row=row, column=COL_MESES_INICIO + i)
                cell.value = abs(float(sub_cuenta[c].sum())) if c in sub_cuenta.columns else 0.0
                cell.number_format = FMT_COP_XL; cell.font = font_cuenta
                cell.alignment = align_right; cell.border = border_thin
            ws.row_dimensions[row].height = 18.0
            ws.row_dimensions[row].outlineLevel = 1
            row += 1

            # Filas TERCERO (colapsadas por defecto, nivel 2) — el nombre del
            # tercero va en su propia columna, junto a la cuenta a la que pertenece
            for tercero, vals in sub_cuenta.iterrows():
                ws.cell(row=row, column=COL_CUENTA).value = nombre
                ws.cell(row=row, column=COL_CUENTA).font = font_data
                ws.cell(row=row, column=COL_CUENTA).alignment = align_left
                ws.cell(row=row, column=COL_TERCERO).value = str(tercero)
                ws.cell(row=row, column=COL_TERCERO).font = font_data
                ws.cell(row=row, column=COL_TERCERO).alignment = align_left
                for c_idx in (COL_GRUPO, COL_CUENTA, COL_TERCERO):
                    ws.cell(row=row, column=c_idx).border = border_thin
                    if row % 2 == 0: ws.cell(row=row, column=c_idx).fill = c_alt_fill
                for i, c in enumerate(cols):
                    cell = ws.cell(row=row, column=COL_MESES_INICIO + i)
                    cell.value = abs(float(vals.get(c, 0)))
                    cell.number_format = FMT_COP_XL; cell.font = font_data
                    cell.alignment = align_right; cell.border = border_thin
                    if row % 2 == 0: cell.fill = c_alt_fill
                ws.row_dimensions[row].height = 17.0
                ws.row_dimensions[row].outlineLevel = 2
                ws.row_dimensions[row].hidden = True   # colapsado por defecto
                row += 1

            # Fila "Total <cuenta>" (siempre visible, nivel 1, subtotal de la cuenta)
            ws.cell(row=row, column=COL_CUENTA).value = f"Total {nombre}"
            ws.cell(row=row, column=COL_CUENTA).font = font_total_cuenta
            ws.cell(row=row, column=COL_CUENTA).alignment = align_left
            for c_idx in (COL_GRUPO, COL_CUENTA, COL_TERCERO):
                ws.cell(row=row, column=c_idx).border = border_thin
            for i, c in enumerate(cols):
                cell = ws.cell(row=row, column=COL_MESES_INICIO + i)
                cell.value = abs(float(sub_cuenta[c].sum())) if c in sub_cuenta.columns else 0.0
                cell.number_format = FMT_COP_XL; cell.font = font_total_cuenta
                cell.alignment = align_right; cell.border = border_thin
            ws.row_dimensions[row].height = 17.0
            ws.row_dimensions[row].outlineLevel = 1
            row += 1

    # Fila de Total general
    ws.cell(row=row, column=COL_GRUPO).value = "Total general"
    ws.cell(row=row, column=COL_GRUPO).font = font_total
    for c_idx in (COL_GRUPO, COL_CUENTA, COL_TERCERO):
        ws.cell(row=row, column=c_idx).fill = c_total_fill
        ws.cell(row=row, column=c_idx).border = border_total
    ws.cell(row=row, column=COL_GRUPO).alignment = align_left
    for i, c in enumerate(cols):
        cell = ws.cell(row=row, column=COL_MESES_INICIO + i)
        cell.value = abs(float(total_general.get(c, 0)))
        cell.number_format = FMT_COP_XL; cell.font = font_total
        cell.fill = c_total_fill; cell.border = border_total; cell.alignment = align_right
    ws.row_dimensions[row].outlineLevel = 0

    # AutoFilter, anchos
    max_row = row
    ws.auto_filter.ref = f"A4:{last_col_letter}{max_row}"
    ws.column_dimensions["A"].width = 34.0
    ws.column_dimensions["B"].width = 44.0
    ws.column_dimensions["C"].width = 34.0
    for col_idx in range(4, num_cols + 1):
        ws.column_dimensions[get_column_letter(col_idx)].width = 20.0

    # Mostrar los botones de agrupación en el borde izquierdo, colapsados
    ws.sheet_format.outlineLevelRow = 1


def generar_hoja_anexo(ws, title, empresa, df):
    """Genera hojas de Anexo ESF y ERI con diseño profesional, autofiltros y anchos amplios."""
    c_header_fill = PatternFill(start_color="1E3A5F", end_color="1E3A5F", fill_type="solid")
    c_subhead_fill = PatternFill(start_color="2E6DA4", end_color="2E6DA4", fill_type="solid")
    c_alt_fill     = PatternFill(start_color="F4F8FA", end_color="F4F8FA", fill_type="solid")
    c_total_fill   = PatternFill(start_color="1E3A5F", end_color="1E3A5F", fill_type="solid")

    font_title  = Font(name="Calibri", size=14, bold=True, color="FFFFFF")
    font_header = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    font_data   = Font(name="Calibri", size=11, color="1E3A5F")
    font_total  = Font(name="Calibri", size=11, bold=True, color="FFFFFF")

    border_thin = Border(
        left=Side(style="thin", color="E2E8F0"),
        right=Side(style="thin", color="E2E8F0"),
        top=Side(style="thin", color="E2E8F0"),
        bottom=Side(style="thin", color="E2E8F0")
    )
    border_total = Border(
        top=Side(style="thin", color="1E3A5F"),
        bottom=Side(style="double", color="FFFFFF")
    )

    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    align_left   = Alignment(horizontal="left", vertical="center")
    align_right  = Alignment(horizontal="right", vertical="center")

    num_cols = len(df.columns) + 1  # 1 para Col A (Grupo)
    last_col_letter = get_column_letter(num_cols)

    # Fila 1: Banner de Título
    ws.row_dimensions[1].height = 28.0
    ws["A1"].value = f"{empresa.upper()} - {title.upper()}"
    ws["A1"].font  = font_title
    ws["A1"].alignment = Alignment(horizontal="left", vertical="center")

    for c in range(1, num_cols + 1):
        cell = ws.cell(row=1, column=c)
        cell.fill = c_header_fill

    # Fila 2: Separador
    ws.row_dimensions[2].height = 10.0

    # Fila 4: Cabecera de tabla con Autofiltros
    ws.row_dimensions[4].height = 26.0
    headers = ["Grupo"] + list(df.columns)
    for col_idx, h in enumerate(headers, 1):
        c = ws.cell(row=4, column=col_idx)
        c.value = h
        c.font  = font_header
        c.fill  = c_subhead_fill
        c.alignment = align_center

    # Filas de datos desde fila 5
    for r_idx, (idx_val, row) in enumerate(df.iterrows(), start=5):
        ws.row_dimensions[r_idx].height = 20.0
        is_total_row = str(idx_val) in ["Total general", "Total Activo", "Total Pasivo"]

        # Col A: Grupo
        cA = ws.cell(row=r_idx, column=1)
        cA.value = str(idx_val)
        cA.alignment = align_left

        if is_total_row:
            cA.font   = font_total
            cA.fill   = c_total_fill
            cA.border = border_total
        else:
            cA.font   = font_data
            if r_idx % 2 == 0:
                cA.fill = c_alt_fill
            cA.border = border_thin

        # Columnas de valores
        for c_idx, val in enumerate(row, start=2):
            cell = ws.cell(row=r_idx, column=c_idx)
            is_num = isinstance(val, (int, float)) and not pd.isna(val)
            cell.value = float(val) if is_num else (val if not pd.isna(val) else None)
            cell.number_format = FMT_COP_XL
            cell.alignment = align_right

            if is_total_row:
                cell.font   = font_total
                cell.fill   = c_total_fill
                cell.border = border_total
            else:
                cell.font   = font_data
                if r_idx % 2 == 0:
                    cell.fill = c_alt_fill
                cell.border = border_thin

    # Habilitar Autofiltro en la fila 4
    max_row = 4 + len(df)
    ws.auto_filter.ref = f"A4:{last_col_letter}{max_row}"

    # Anchos de columna amplios para legibilidad perfecta de números sin que se corten
    ws.column_dimensions["A"].width = 68.0
    for col_idx in range(2, num_cols + 1):
        col_letter = get_column_letter(col_idx)
        ws.column_dimensions[col_letter].width = 22.0


def _escribir_df_en_hoja(ws, df, index=False):
    """Escribe un DataFrame de Detalle con formato, autofiltros y anchos ajustados."""
    c_subhead_fill = PatternFill(start_color="2E6DA4", end_color="2E6DA4", fill_type="solid")
    c_alt_fill     = PatternFill(start_color="F4F8FA", end_color="F4F8FA", fill_type="solid")

    font_header = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    font_data   = Font(name="Calibri", size=10, color="1E3A5F")

    border_thin = Border(
        left=Side(style="thin", color="E2E8F0"),
        right=Side(style="thin", color="E2E8F0"),
        top=Side(style="thin", color="E2E8F0"),
        bottom=Side(style="thin", color="E2E8F0")
    )

    align_center = Alignment(horizontal="center", vertical="center")
    align_left   = Alignment(horizontal="left", vertical="center")
    align_right  = Alignment(horizontal="right", vertical="center")

    cols = list(df.columns)
    ws.row_dimensions[1].height = 25.0

    # Header
    for c_idx, col_name in enumerate(cols, 1):
        cell = ws.cell(row=1, column=c_idx)
        cell.value = str(col_name)
        cell.font  = font_header
        cell.fill  = c_subhead_fill
        cell.alignment = align_center

    # Data
    for r_idx, (_, row) in enumerate(df.iterrows(), start=2):
        ws.row_dimensions[r_idx].height = 19.0
        for c_idx, val in enumerate(row, 1):
            cell = ws.cell(row=r_idx, column=c_idx)
            col_name = cols[c_idx - 1]
            if isinstance(val, (int, float)) and not pd.isna(val):
                cell.value = float(val)
                if "Saldo" in col_name or "Monto" in col_name or "Valor" in col_name:
                    cell.number_format = FMT_COP_XL
                    cell.alignment = align_right
                else:
                    cell.alignment = align_left
            else:
                cell.value = str(val) if not pd.isna(val) else None
                cell.alignment = align_left

            cell.font = font_data
            if r_idx % 2 == 0:
                cell.fill = c_alt_fill
            cell.border = border_thin

    # AutoFilter
    max_col_letter = get_column_letter(len(cols))
    max_row = 1 + len(df)
    ws.auto_filter.ref = f"A1:{max_col_letter}{max_row}"

    # Auto Column Widths
    for c_idx, col_name in enumerate(cols, 1):
        col_letter = get_column_letter(c_idx)
        max_len = max(len(str(col_name)), int(df[col_name].astype(str).str.len().max()) if not df.empty else 10)
        ws.column_dimensions[col_letter].width = min(max(max_len + 4, 15), 55)


def generar_hoja_consol(ws, empresa, nit, df_eri_raw, meses_d):
    """
    Hoja 'CONSOL': réplica de la plantilla de referencia — detalle de gastos
    por CUENTA (no por tercero) y por mes, en dos bloques: "GASTOS DE
    ADMINISTRACIÓN" (incluye Gastos de venta) y "GASTOS NO OPERACIONALES"
    (financieros + otros gastos + diferencia en cambio), con su subtotal cada
    uno y un total general "TOTAL COSTOS Y GASTOS DE LA CIA." al final.
    Incluye columnas ACUMULADO y PROMEDIO.
    """
    c_header_fill  = PatternFill(start_color="1E3A5F", end_color="1E3A5F", fill_type="solid")
    c_subhead_fill = PatternFill(start_color="2E6DA4", end_color="2E6DA4", fill_type="solid")
    c_seccion_fill = PatternFill(start_color="D9E6F2", end_color="D9E6F2", fill_type="solid")
    c_total_fill   = PatternFill(start_color="D9E6F2", end_color="D9E6F2", fill_type="solid")
    font_title  = Font(name="Calibri", size=14, bold=True, color="FFFFFF")
    font_header = Font(name="Calibri", size=10.5, bold=True, color="FFFFFF")
    font_data   = Font(name="Calibri", size=10, color="1E3A5F")
    font_seccion= Font(name="Calibri", size=11, bold=True, color="1E3A5F")
    font_total  = Font(name="Calibri", size=10.5, bold=True, color="1E3A5F")
    border_thin = Border(bottom=Side(style="thin", color="8AA4C0"))
    align_left  = Alignment(horizontal="left", vertical="center")
    align_right = Alignment(horizontal="right", vertical="center")
    align_center= Alignment(horizontal="center", vertical="center")

    MESES_CORTO = {"ENERO":"ENE","FEBRERO":"FEB","MARZO":"MAR","ABRIL":"ABR","MAYO":"MAY",
                   "JUNIO":"JUN","JULIO":"JUL","AGOSTO":"AGO","SEPTIEMBRE":"SEP",
                   "OCTUBRE":"OCT","NOVIEMBRE":"NOV","DICIEMBRE":"DIC"}
    cols = list(meses_d) + ["ACUMULADO", "PROMEDIO"]
    num_cols = len(cols) + 1  # + Concepto
    last_col = get_column_letter(num_cols)

    ws.row_dimensions[2].height = 22.0
    ws["B2"].value = f"ANEXOS DEL ESTADO DE RESULTADOS - CONSOLIDADO — {empresa.upper()}"
    ws["B2"].font = font_title
    for c in range(1, num_cols + 1):
        ws.cell(row=2, column=c).fill = c_header_fill

    ws.row_dimensions[4].height = 20.0
    headers = [None] + [MESES_CORTO.get(m, m) for m in meses_d] + ["ACUMULADO", "PROMEDIO"]
    for c_idx, h in enumerate(headers, 1):
        if h is None: continue
        c = ws.cell(row=4, column=c_idx)
        c.value = h; c.font = font_header; c.fill = c_subhead_fill; c.alignment = align_center

    def detalle_por_cuenta(grupos):
        sub = df_eri_raw[df_eri_raw["Grupo"].isin(grupos)]
        if sub.empty:
            return pd.DataFrame(columns=meses_d + ["ACUMULADO", "PROMEDIO"])
        det = sub.groupby(["Codigo", "Nombre cuenta", "Mes"], dropna=False)["Saldo Mes"].sum().reset_index()
        piv = det.pivot_table(index=["Codigo", "Nombre cuenta"], columns="Mes",
                               values="Saldo Mes", fill_value=0)
        piv = piv.reindex(columns=meses_d, fill_value=0)
        piv["ACUMULADO"] = piv.sum(axis=1)
        piv["PROMEDIO"] = piv["ACUMULADO"] / max(len(meses_d), 1)
        return piv

    secciones = [
        ("GASTOS DE ADMINISTRACIÓN", ["GASTOS DE ADMINISTRACION", "GASTOS DE VENTA"],
         "TOTAL GASTOS DE ADMINISTRACIÓN "),
        ("GASTOS NO OPERACIONALES",
         ["GASTOS FINANCIEROS", "OTROS GASTOS", "DIFERENCIA EN CAMBIO NETA", "OTROS GASTOS SIN CLASIFICAR"],
         "TOTAL  GTOS NO OPERACIONALES"),
    ]

    r = 5
    total_general = pd.Series(0.0, index=cols)
    for titulo_seccion, grupos, titulo_total in secciones:
        ws.cell(row=r, column=1).value = titulo_seccion
        ws.cell(row=r, column=1).font = font_seccion
        for c in range(1, num_cols + 1): ws.cell(row=r, column=c).fill = c_seccion_fill
        r += 2

        piv = detalle_por_cuenta(grupos)
        seccion_total = pd.Series(0.0, index=cols)
        for (codigo, nombre), vals in piv.iterrows():
            cod_txt = str(codigo).rstrip("0").rstrip(".") if isinstance(codigo, float) else str(codigo)
            ws.cell(row=r, column=1).value = f"{cod_txt}  {nombre}"
            ws.cell(row=r, column=1).font = font_data
            for i, c in enumerate(cols):
                v = abs(float(vals.get(c, 0)))
                ws.cell(row=r, column=2 + i).value = v
                ws.cell(row=r, column=2 + i).number_format = FMT_COP_XL
                ws.cell(row=r, column=2 + i).font = font_data
                ws.cell(row=r, column=2 + i).alignment = align_right
                seccion_total[c] += v
            r += 1
        total_general = total_general.add(seccion_total, fill_value=0)

        ws.cell(row=r, column=1).value = titulo_total
        ws.cell(row=r, column=1).font = font_total
        ws.cell(row=r, column=1).border = border_thin
        for i, c in enumerate(cols):
            cell = ws.cell(row=r, column=2 + i)
            cell.value = seccion_total[c]; cell.number_format = FMT_COP_XL
            cell.font = font_total; cell.alignment = align_right; cell.border = border_thin
            cell.fill = c_total_fill
        ws.cell(row=r, column=1).fill = c_total_fill
        r += 2

    ws.cell(row=r, column=1).value = "TOTAL COSTOS Y GASTOS DE LA CIA."
    ws.cell(row=r, column=1).font = font_seccion
    ws.cell(row=r, column=1).border = border_thin
    for i, c in enumerate(cols):
        cell = ws.cell(row=r, column=2 + i)
        cell.value = total_general[c]; cell.number_format = FMT_COP_XL
        cell.font = font_seccion; cell.alignment = align_right; cell.border = border_thin

    ws.column_dimensions["A"].width = 42.0
    for c_idx in range(2, num_cols + 1):
        ws.column_dimensions[get_column_letter(c_idx)].width = 15.0


def generar_hoja_consolidado(ws, empresa, nit, pivot_eri, totales_eri, meses_d, df_eri_raw):
    """
    Hoja 'ER MENSUALIZADO': réplica de la plantilla de referencia — el mismo
    ERI pero con una columna por mes (no acumulada) más una columna final
    "ACUMULADO" que coincide exactamente con el total que se presenta en la
    hoja ERI. Bajo "Ingresos de actividades ordinarias" se listan sus cuentas
    hoja (igual que en el archivo de referencia).
    """
    c_header_fill  = PatternFill(start_color="1E3A5F", end_color="1E3A5F", fill_type="solid")
    c_subhead_fill = PatternFill(start_color="2E6DA4", end_color="2E6DA4", fill_type="solid")
    c_alt_fill     = PatternFill(start_color="F4F8FA", end_color="F4F8FA", fill_type="solid")
    c_total_fill   = PatternFill(start_color="D9E6F2", end_color="D9E6F2", fill_type="solid")
    font_title  = Font(name="Calibri", size=14, bold=True, color="FFFFFF")
    font_header = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    font_data   = Font(name="Calibri", size=10.5, color="1E3A5F")
    font_sub    = Font(name="Calibri", size=11, bold=True, color="1E3A5F")
    border_thin = Border(bottom=Side(style="thin", color="8AA4C0"))
    align_left  = Alignment(horizontal="left", vertical="center")
    align_right = Alignment(horizontal="right", vertical="center")
    align_center= Alignment(horizontal="center", vertical="center", wrap_text=True)

    MESES_LARGO = {"ENERO":"ENERO ","FEBRERO":"FEBRERO ","MARZO":"MARZO","ABRIL":"ABRIL",
                   "MAYO":"MAYO","JUNIO":"JUNIO","JULIO":"JULIO","AGOSTO":"AGOSTO",
                   "SEPTIEMBRE":"SEPTIEMBRE","OCTUBRE":"OCTUBRE","NOVIEMBRE":"NOVIEMBRE ",
                   "DICIEMBRE":"DICIEMBRE"}
    cols = list(meses_d) + ["ACUMULADO"]
    num_cols = len(cols) + 2  # Codigo + Concepto
    last_col = get_column_letter(num_cols)

    ws.row_dimensions[1].height = 20.0
    ws["B1"].value = empresa.upper(); ws["B1"].font = Font(name="Calibri", size=12, bold=True, color="1E3A5F")
    ws.row_dimensions[2].height = 22.0
    ws["B2"].value = "ESTADO DE RESULTADOS INTEGRAL  MENSUALIZADO - CONSOLIDADO"
    ws["B2"].font = font_title
    for c in range(1, num_cols + 1):
        ws.cell(row=2, column=c).fill = c_header_fill

    ws.row_dimensions[5].height = 24.0
    headers = [None, None] + [MESES_LARGO.get(m, m) for m in meses_d] + ["ACUMULADO"]
    for c_idx, h in enumerate(headers, 1):
        if h is None: continue
        c = ws.cell(row=5, column=c_idx)
        c.value = h; c.font = font_header; c.fill = c_subhead_fill; c.alignment = align_center

    # Totales por Grupo y por mes (columna) + acumulado (usa totales_eri, igual que la hoja ERI)
    totales_por_col = {}
    for mes in meses_d:
        totales_por_col[mes] = {g: (pivot_eri.loc[g, mes] if g in pivot_eri.index else 0.0) for g in GRUPOS_ERI}
    totales_por_col["ACUMULADO"] = totales_eri
    vals = {col: _calc_lineas_eri(t) for col, t in totales_por_col.items()}

    # Cuentas hoja de "Ingresos de actividades ordinarias", desglosadas por mes
    ing_leaf = pd.DataFrame()
    if not df_eri_raw.empty:
        sub = df_eri_raw[df_eri_raw["Grupo"] == "INGRESOS DE ACTIVIDADES ORDINARIAS"]
        if not sub.empty:
            det = sub.groupby(["Codigo", "Nombre cuenta", "Mes"], dropna=False)["Saldo Mes"].sum().reset_index()
            ing_leaf = det.pivot_table(index=["Codigo", "Nombre cuenta"], columns="Mes",
                                        values="Saldo Mes", fill_value=0)
            ing_leaf = ing_leaf.reindex(columns=meses_d, fill_value=0)
            ing_leaf["ACUMULADO"] = ing_leaf.sum(axis=1)

    # Las cuentas de devolución (débito, reducen el ingreso) se separan del
    # ingreso bruto, igual que en el archivo de referencia: "Ingresos de
    # actividades ordinarias" = bruto, "VENTAS NETAS" = bruto - devoluciones.
    es_devolucion = ing_leaf.index.get_level_values("Nombre cuenta").str.upper().str.contains("DEVOLU", na=False) \
        if not ing_leaf.empty else pd.Series(dtype=bool)
    ing_bruto = {c: (-ing_leaf.loc[~es_devolucion, c].sum() if not ing_leaf.empty else vals[c][10]) for c in cols}
    devoluciones = {c: (ing_leaf.loc[es_devolucion, c].sum() if not ing_leaf.empty else 0.0) for c in cols}

    def escribir(row_n, codigo, label, valores_por_col, bold=False, total=False):
        # Los valores ya vienen con el signo final a mostrar (positivo para
        # gastos, igual que en el archivo de referencia: la resta se hace en
        # las fórmulas de los subtotales, no en el signo de cada línea).
        ws.row_dimensions[row_n].height = 18.0
        if codigo is not None:
            ws.cell(row=row_n, column=1).value = codigo
            ws.cell(row=row_n, column=1).font = font_data
        cB = ws.cell(row=row_n, column=2)
        cB.value = label; cB.font = font_sub if (bold or total) else font_data; cB.alignment = align_left
        if total: cB.fill = c_total_fill
        if bold: cB.border = border_thin
        for i, col in enumerate(cols):
            v = valores_por_col.get(col, 0)
            cell = ws.cell(row=row_n, column=3 + i)
            cell.value = v; cell.number_format = FMT_COP_XL
            cell.font = font_sub if (bold or total) else font_data
            cell.alignment = align_right
            if total: cell.fill = c_total_fill
            if bold: cell.border = border_thin

    r = 6
    escribir(r, None, "Ingresos de actividades ordinarias", ing_bruto); r += 1
    for (codigo, nombre) in ing_leaf.index if not ing_leaf.empty else []:
        cod_txt = str(codigo).rstrip("0").rstrip(".") if isinstance(codigo, float) else str(codigo)
        # Se invierte el signo (no abs()) para conservar la distinción entre
        # ingreso (crédito, saldo negativo -> se ve positivo) y una cuenta
        # contraria como "Devolución en ventas" (débito, saldo positivo ->
        # se ve negativo), igual que en el archivo de referencia.
        escribir(r, cod_txt, f"          {nombre}",
                 {c: -float(ing_leaf.loc[(codigo, nombre), c]) for c in cols}); r += 1
    r += 1
    escribir(r, None, "VENTAS NETAS", {c: vals[c][10] for c in cols}, total=True); r += 2
    escribir(r, None, "Costo del servicio", {c: abs(vals[c][11]) for c in cols}); r += 2
    escribir(r, None, "UTILIDAD BRUTA", {c: vals[c][12] for c in cols}, total=True); r += 2
    escribir(r, None, "Gastos Administración", {c: abs(vals[c][13] + vals[c][15]) for c in cols}); r += 2
    escribir(r, None, "GASTOS DE OPERACION", {c: abs(vals[c][13] + vals[c][15]) for c in cols}, total=True); r += 2
    escribir(r, None, "UTILIDAD OPERATIVA",
             {c: vals[c][12] + vals[c][13] + vals[c][15] for c in cols}, total=True); r += 2
    escribir(r, None, "Otros gastos", {c: abs(vals[c][16] + vals[c][21]) for c in cols}); r += 1
    escribir(r, None, "Gastos financieros", {c: abs(vals[c][18]) for c in cols}); r += 2
    # "Otros ingresos" incluye los ingresos financieros, igual que en el
    # archivo de referencia (que no trae una línea separada para financieros).
    escribir(r, None, "Otros ingresos", {c: vals[c][14] + vals[c][17] + vals[c][20] for c in cols}); r += 1
    escribir(r, None, "Diferencia en cambio neta", {c: vals[c][19] for c in cols}); r += 2
    escribir(r, None, "UTILIDAD ANTES DE IMPTOS", {c: vals[c][22] for c in cols}, total=True); r += 2
    escribir(r, None, "Provisión impuesto de renta", {c: abs(vals[c][24]) for c in cols}); r += 2
    escribir(r, None, "UTILIDAD (PERDIDA) NETA", {c: vals[c][26] for c in cols}, total=True); r += 1

    max_row = r
    ws.auto_filter.ref = f"A5:{last_col}{max_row}"
    ws.column_dimensions["A"].width = 10.0
    ws.column_dimensions["B"].width = 38.0
    for c_idx in range(3, num_cols + 1):
        ws.column_dimensions[get_column_letter(c_idx)].width = 16.0


def generar_excel_eeff(empresa, nit, periodo, saldos_esf, saldos_patrimonio, totales_eri,
                       pivot_eri, df_eri_raw, pivot_esf, df_esf_raw,
                       saldos_esf_anterior=None, saldos_patrimonio_anterior=None, totales_eri_anterior=None):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    # Hojas formateadas principales (con período anterior si viene diligenciado)
    generar_hoja_esf(wb.create_sheet("ESF"), empresa, nit, periodo, saldos_esf, saldos_patrimonio,
                      saldos_esf_anterior, saldos_patrimonio_anterior)
    generar_hoja_eri(wb.create_sheet("ERI"), empresa, nit, periodo, totales_eri, totales_eri_anterior)

    # Hojas de Anexo — tablas DESPLEGABLES: fila resumen por Grupo con las
    # cuentas de detalle agrupadas debajo (botones +/- de Excel).
    ws_aesf = wb.create_sheet("Anexo ESF")
    generar_hoja_anexo_desplegable(ws_aesf, "ANEXO AL ESTADO DE LA SITUACIÓN FINANCIERA",
                                    empresa, df_esf_raw, GRUPOS_ESF_ORDEN, list(pivot_esf.columns[:-1]))

    ws_aeri = wb.create_sheet("Anexo ERI")
    generar_hoja_anexo_desplegable(ws_aeri, "ANEXO AL ESTADO DE RESULTADOS INTEGRAL",
                                    empresa, df_eri_raw, GRUPOS_ERI, list(pivot_eri.columns[:-1]))

    # Hojas de detalle con formato y autofiltros
    ws_desf = wb.create_sheet("Detalle ESF")
    _escribir_df_en_hoja(ws_desf, df_esf_raw, index=False)

    ws_deri = wb.create_sheet("Detalle ERI")
    _escribir_df_en_hoja(ws_deri, df_eri_raw, index=False)

    # CONSOL: detalle de gastos por cuenta y mes (réplica del archivo de referencia)
    ws_consol = wb.create_sheet("CONSOL")
    generar_hoja_consol(ws_consol, empresa, nit, df_eri_raw, list(pivot_eri.columns[:-1]))

    # ER MENSUALIZADO: misma info del ERI, una columna por mes (no acumulada)
    ws_cons = wb.create_sheet("ER MENSUALIZADO")
    generar_hoja_consolidado(ws_cons, empresa, nit, pivot_eri, totales_eri, list(pivot_eri.columns[:-1]), df_eri_raw)

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


# ══════════════════════════════════════════════════════════════════════════════
# UI
# ══════════════════════════════════════════════════════════════════════════════
st.markdown("""
<div class="header-block">
    <h1>📊 Anexos EEFF – ERI y ESF</h1>
    <p>Genera los estados financieros formateados (ESF y ERI) desde la hoja <strong>terceros_</strong></p>
</div>""", unsafe_allow_html=True)

st.markdown('<div class="upload-section">', unsafe_allow_html=True)
st.markdown('<p class="section-title">📁 Archivo Excel de Anexos EEFF</p>', unsafe_allow_html=True)
st.caption("Debe contener la hoja **terceros_** con columnas: Codigo, Nombre cuenta, Nit, Nombre tercero, Mes, Saldo Mes, Cuenta, Grupo")
uploaded = st.file_uploader("Sube el archivo Excel", type=["xlsx"], label_visibility="collapsed")
st.markdown('</div>', unsafe_allow_html=True)

if not uploaded:
    st.markdown("""<div style="text-align:center;padding:60px 20px;color:#999;">
        <div style="font-size:3rem;margin-bottom:16px;">📂</div>
        <p style="font-size:1.1rem;font-weight:600;">Sube el archivo Excel para generar los estados financieros</p>
    </div>""", unsafe_allow_html=True)
    st.stop()

file_bytes = uploaded.read()
(df_eri_raw, pivot_eri, df_esf_raw, pivot_esf, saldos_esf, saldos_patrimonio, totales_eri, ultimo_mes,
 saldos_esf_anterior, totales_eri_anterior, saldos_patrimonio_anterior) = procesar_archivo(file_bytes)

meses_disp     = [m for m in MESES_ORDEN if m in pivot_eri.columns]
meses_disp_esf = [m for m in MESES_ORDEN if m in pivot_esf.columns]

# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("""<div style="background:linear-gradient(135deg,#1E3A5F,#2E6DA4);padding:14px 18px;
                border-radius:10px;margin-bottom:16px;color:white;text-align:center;">
        <div style="font-size:1.1rem;font-weight:700;">🏢 Datos de la empresa</div></div>""",
        unsafe_allow_html=True)
    empresa = st.text_input("Nombre empresa", value="MI EMPRESA S.A.S")
    nit     = st.text_input("NIT", value="NIT 000.000.000-0")
    periodo = st.text_input("Período", value="AL 31 DE DICIEMBRE DE 2025 Y 31 DE DICIEMBRE DE 2024")

    st.divider()
    st.markdown("**⚙️ Filtros**")
    todos_meses = st.toggle("Todos los meses", value=True, key="tog_meses")
    if todos_meses:
        meses_sel = meses_disp
    else:
        meses_sel = []
        filas = [meses_disp[i:i+4] for i in range(0, len(meses_disp), 4)]
        for fila in filas:
            cols_sb = st.columns(len(fila))
            for col_sb, mes in zip(cols_sb, fila):
                if col_sb.checkbox(MESES_ABREV[mes], value=True, key=f"m_{mes}"):
                    meses_sel.append(mes)

if not meses_sel:
    st.warning("Selecciona al menos un mes."); st.stop()

# ── Tabs principales ──────────────────────────────────────────────────────────
tab_eri, tab_esf, tab_exportar = st.tabs([
    "📈 Estado de Resultados (ERI)",
    "🏦 Estado de Situación Financiera (ESF)",
    "📄 Exportar EEFF Formateado",
])

# ══════════════════════════════════════════════════════════════════════════════
# TAB ERI
# ══════════════════════════════════════════════════════════════════════════════
with tab_eri:
    cols_eri  = meses_sel + ["Total general"]
    pivot_f   = pivot_eri.loc[
        [g for g in GRUPOS_ERI if g in pivot_eri.index] + ["Total general"], cols_eri]

    ingresos = pivot_f.loc[
        [g for g in ["INGRESOS DE ACTIVIDADES ORDINARIAS","OTROS INGRESOS","INGRESOS FINANCIEROS"]
         if g in pivot_f.index], "Total general"].sum()
    gastos   = pivot_f.loc[
        [g for g in ["COSTO DE VENTAS","GASTOS DE ADMINISTRACION","GASTOS DE VENTA","OTROS GASTOS",
                      "GASTOS FINANCIEROS","PROVISION DE IMPUESTOS"]
         if g in pivot_f.index], "Total general"].sum()
    # "Diferencia en cambio neta" no se suma aquí porque puede ser ganancia o
    # pérdida (signo variable); se ve correctamente en la tabla y en la hoja ERI.
    resultado = ingresos + gastos

    st.markdown("---")
    m1, m2, m3 = st.columns(3)
    with m1:
        st.markdown(f"""<div class="metric-card"><div class="number green">{fmt_cop(abs(ingresos))}</div>
            <div class="label">Total Ingresos</div></div>""", unsafe_allow_html=True)
    with m2:
        st.markdown(f"""<div class="metric-card"><div class="number red">{fmt_cop(gastos)}</div>
            <div class="label">Total Gastos</div></div>""", unsafe_allow_html=True)
    with m3:
        rc = "green" if resultado < 0 else "red"
        rl = "Utilidad" if resultado < 0 else "Pérdida"
        st.markdown(f"""<div class="metric-card"><div class="number {rc}">{fmt_cop(abs(resultado))}</div>
            <div class="label">Resultado · {rl}</div></div>""", unsafe_allow_html=True)

    st.markdown("---")
    st.markdown('<div class="result-block">', unsafe_allow_html=True)
    st.markdown('<p class="section-title">📋 Resumen ERI por Grupo y Mes</p>', unsafe_allow_html=True)
    st.dataframe(pivot_f.style.format(fmt_cop), use_container_width=True,
                 height=min(60+40*len(pivot_f), 500))
    st.markdown('</div>', unsafe_allow_html=True)
    st.markdown("---")

    gt1, gt2, gt3 = st.tabs(["📈 Evolución mensual","🥧 Composición","🔍 Detalle por tercero"])
    with gt1:
        fig = go.Figure()
        for grupo in [g for g in GRUPOS_ERI if g in pivot_f.index]:
            vals = [pivot_f.loc[grupo, m] if m in pivot_f.columns else 0 for m in meses_sel]
            fig.add_trace(go.Bar(name=grupo, x=meses_sel, y=[abs(v) for v in vals],
                marker_color=COLORES_ERI.get(grupo,"#95a5a6"),
                customdata=[fmt_cop(v) for v in vals],
                hovertemplate=f"<b>{grupo}</b><br>%{{x}}: %{{customdata}}<extra></extra>"))
        fig.update_layout(barmode="group", title="Evolución mensual por grupo",
            xaxis_title="Mes", yaxis_title="COP", height=420,
            legend=dict(orientation="h",yanchor="bottom",y=1.02,xanchor="right",x=1),
            plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)")
        st.plotly_chart(fig, use_container_width=True)
    with gt2:
        tot_pie = pivot_f.loc[[g for g in GRUPOS_ERI if g in pivot_f.index],"Total general"]
        fig2 = px.pie(values=tot_pie.abs().values, names=tot_pie.index,
            title="Composición por grupo", color=tot_pie.index,
            color_discrete_map=COLORES_ERI, hole=0.4)
        fig2.update_traces(textposition="outside", textinfo="percent+label")
        fig2.update_layout(height=450, showlegend=False)
        st.plotly_chart(fig2, use_container_width=True)
    with gt3:
        st.markdown('<div class="result-block">', unsafe_allow_html=True)
        gd = st.selectbox("Grupo", GRUPOS_ERI, key="eri_det_g")
        md = st.selectbox("Mes",   ["Todos"]+meses_sel, key="eri_det_m")
        mask = df_eri_raw["Grupo"] == gd
        if md != "Todos": mask &= df_eri_raw["Mes"] == md
        df_d = (df_eri_raw[mask & df_eri_raw["Nombre tercero"].notna()]
            [["Mes","Codigo","Nombre cuenta","Nit","Nombre tercero","Saldo Mes"]]
            .sort_values("Saldo Mes", key=abs, ascending=False).reset_index(drop=True))
        df_d["Saldo Mes"] = df_d["Saldo Mes"].apply(fmt_cop)
        st.dataframe(df_d, use_container_width=True, height=400)
        st.markdown('</div>', unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# TAB ESF
# ══════════════════════════════════════════════════════════════════════════════
with tab_esf:
    act_r = [g for g in GRUPOS_ESF_ACTIVO if g in pivot_esf.index]
    pas_r = [g for g in GRUPOS_ESF_PASIVO if g in pivot_esf.index]
    meses_esf_f = [m for m in meses_disp_esf if m in meses_sel] or meses_disp_esf
    cols_esf_f  = meses_esf_f + ["Total general"]

    total_act_v = pivot_esf.loc["Total Activo","Total general"] if "Total Activo" in pivot_esf.index else 0
    total_pas_v = pivot_esf.loc["Total Pasivo","Total general"] if "Total Pasivo" in pivot_esf.index else 0
    saldo_act   = sum(abs(saldos_esf.get(g,0)) for g in GRUPOS_ESF_ACTIVO)
    saldo_pas   = sum(abs(saldos_esf.get(g,0)) for g in GRUPOS_ESF_PASIVO)

    st.markdown("---")
    mc1, mc2, mc3 = st.columns(3)
    with mc1:
        st.markdown(f"""<div class="metric-card"><div class="number green">{fmt_cop(abs(total_act_v))}</div>
            <div class="label">Total Activo (acumulado)</div></div>""", unsafe_allow_html=True)
    with mc2:
        st.markdown(f"""<div class="metric-card"><div class="number red">{fmt_cop(abs(total_pas_v))}</div>
            <div class="label">Total Pasivo (acumulado)</div></div>""", unsafe_allow_html=True)
    with mc3:
        saldo_pat = saldos_patrimonio.get("total_patrimonio", 0.0)
        pc = "purple" if saldo_pat >= 0 else "red"
        st.markdown(f"""<div class="metric-card"><div class="number {pc}">{fmt_cop(saldo_pat)}</div>
            <div class="label">Total Patrimonio ({ultimo_mes})</div></div>""", unsafe_allow_html=True)

    st.markdown("---")
    st.markdown('<div class="result-block">', unsafe_allow_html=True)
    st.markdown('<p class="section-title">📋 ESF por Grupo y Mes</p>', unsafe_allow_html=True)
    frames = []
    if act_r: frames.append(pivot_esf.loc[act_r, cols_esf_f])
    if "Total Activo" in pivot_esf.index: frames.append(pivot_esf.loc[["Total Activo"], cols_esf_f])
    if pas_r: frames.append(pivot_esf.loc[pas_r, cols_esf_f])
    if "Total Pasivo" in pivot_esf.index: frames.append(pivot_esf.loc[["Total Pasivo"], cols_esf_f])
    pivot_esf_f = pd.concat(frames) if frames else pd.DataFrame()

    def style_esf(row):
        if row.name in ("Total Activo","Total Pasivo"):
            return ["font-weight:bold;background-color:#EBF5FB;color:#1E3A5F"]*len(row)
        if row.name in GRUPOS_ESF_ACTIVO: return ["color:#1A9E5C"]*len(row)
        return ["color:#D63B3B"]*len(row)

    st.dataframe(pivot_esf_f.style.format(fmt_cop).apply(style_esf, axis=1),
                 use_container_width=True, height=min(60+40*len(pivot_esf_f), 600))
    st.markdown('</div>', unsafe_allow_html=True)
    st.markdown("---")

    et1, et2, et3 = st.tabs(["📊 Activo vs Pasivo mensual","🥧 Composición ESF","🔍 Detalle por tercero"])
    with et1:
        a_v = [pivot_esf.loc[act_r, m].sum() if act_r and m in pivot_esf.columns else 0 for m in meses_disp_esf]
        p_v = [pivot_esf.loc[pas_r, m].sum() if pas_r and m in pivot_esf.columns else 0 for m in meses_disp_esf]
        fig_e1 = go.Figure()
        fig_e1.add_trace(go.Scatter(x=meses_disp_esf, y=a_v, mode="lines+markers", name="Total Activo",
            line=dict(color="#1A9E5C",width=3), customdata=[fmt_cop(v) for v in a_v],
            hovertemplate="<b>Activo</b><br>%{x}: %{customdata}<extra></extra>"))
        fig_e1.add_trace(go.Scatter(x=meses_disp_esf, y=p_v, mode="lines+markers", name="Total Pasivo",
            line=dict(color="#D63B3B",width=3), customdata=[fmt_cop(v) for v in p_v],
            hovertemplate="<b>Pasivo</b><br>%{x}: %{customdata}<extra></extra>"))
        fig_e1.update_layout(title="Evolución mensual Activo vs Pasivo",
            xaxis_title="Mes", yaxis_title="COP", height=420,
            legend=dict(orientation="h",yanchor="bottom",y=1.02,xanchor="right",x=1),
            plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)")
        st.plotly_chart(fig_e1, use_container_width=True)
    with et2:
        grupos_pie = [g for g in GRUPOS_ESF_ORDEN if g in pivot_esf_f.index]
        tot_pie_e = pivot_esf_f.loc[grupos_pie,"Total general"].abs()
        fig_e2 = go.Figure(go.Pie(
            labels=[GRUPOS_LABEL_ESF.get(g,g) for g in tot_pie_e.index],
            values=tot_pie_e.values,
            marker_colors=[COLORES_ESF.get(g,"#aaa") for g in tot_pie_e.index],
            hole=0.4, textposition="outside", textinfo="percent+label"))
        fig_e2.update_layout(title="Composición ESF", height=500, showlegend=False)
        st.plotly_chart(fig_e2, use_container_width=True)
    with et3:
        st.markdown('<div class="result-block">', unsafe_allow_html=True)
        opc = [g for g in GRUPOS_ESF_ORDEN if g in df_esf_raw["Grupo"].unique()]
        gde = st.selectbox("Grupo ESF", opc, key="esf_det_g")
        mde = st.selectbox("Mes",   ["Todos"]+meses_disp_esf, key="esf_det_m")
        mask_e = df_esf_raw["Grupo"] == gde
        if mde != "Todos": mask_e &= df_esf_raw["Mes"] == mde
        df_de = (df_esf_raw[mask_e & df_esf_raw["Nombre tercero"].notna()]
            [["Mes","Codigo","Nombre cuenta","Nit","Nombre tercero","Saldo Mes"]]
            .sort_values("Saldo Mes", key=abs, ascending=False).reset_index(drop=True))
        df_de["Saldo Mes"] = df_de["Saldo Mes"].apply(fmt_cop)
        st.dataframe(df_de, use_container_width=True, height=400)
        st.markdown('</div>', unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# TAB EXPORTAR EEFF FORMATEADO
# ══════════════════════════════════════════════════════════════════════════════
with tab_exportar:
    st.markdown("---")
    st.markdown('<div class="result-block">', unsafe_allow_html=True)
    st.markdown('<p class="section-title">📄 Generar EEFF con formato profesional</p>', unsafe_allow_html=True)
    st.info(
        "Genera un archivo Excel con las hojas **ESF** y **ERI** formateadas según el "
        "estándar contable colombiano, más las hojas de Anexos y Detalle. "
        "Configura los datos de la empresa en el panel lateral antes de exportar."
    )

    c1, c2 = st.columns(2)
    with c1:
        st.markdown(f"**Vista previa ESF — saldo de cierre ({ultimo_mes})**")
        prev_esf_list = [
            {"Cuenta": NOMBRE_ESF.get(g,g).strip(), "Saldo": fmt_cop(abs(saldos_esf.get(g,0)))}
            for g in GRUPOS_ESF_ORDEN if g in saldos_esf
        ]
        prev_esf_list.append({"Cuenta": "─── PATRIMONIO ───", "Saldo": ""})
        prev_esf_list.append({"Cuenta": "Capital emitido", "Saldo": fmt_cop(saldos_patrimonio.get("capital_emitido",0))})
        prev_esf_list.append({"Cuenta": "Superávit de capital", "Saldo": fmt_cop(saldos_patrimonio.get("superavit_capital",0))})
        prev_esf_list.append({"Cuenta": "Utilidad acumulada", "Saldo": fmt_cop(saldos_patrimonio.get("utilidad_acumulada",0))})
        prev_esf_list.append({"Cuenta": "Utilidad del periodo", "Saldo": fmt_cop(saldos_patrimonio.get("utilidad_periodo",0))})
        prev_esf_list.append({"Cuenta": "TOTAL PATRIMONIO", "Saldo": fmt_cop(saldos_patrimonio.get("total_patrimonio",0))})

        prev_esf = pd.DataFrame(prev_esf_list)
        st.dataframe(prev_esf, use_container_width=True, hide_index=True, height=370)
        tot_pas_pat = saldo_pas + saldos_patrimonio.get("total_patrimonio", 0)
        st.markdown(
            f"<p><b>Total Activo:</b> {fmt_cop(saldo_act)} &nbsp;|&nbsp; <b>Total Pasivo:</b> {fmt_cop(saldo_pas)} &nbsp;|&nbsp; <b>Pasivo + Patrimonio:</b> {fmt_cop(tot_pas_pat)}</p>",
            unsafe_allow_html=True,
        )
    with c2:
        st.markdown("**Vista previa ERI — acumulado**")
        ing_ord    = abs(totales_eri.get("INGRESOS DE ACTIVIDADES ORDINARIAS",0))
        costo_vtas = abs(totales_eri.get("COSTO DE VENTAS",0))
        otros_ing  = abs(totales_eri.get("OTROS INGRESOS",0))
        gtos_adm   = abs(totales_eri.get("GASTOS DE ADMINISTRACION",0))
        gtos_venta = abs(totales_eri.get("GASTOS DE VENTA",0))
        otros_gto  = abs(totales_eri.get("OTROS GASTOS",0))
        ing_fin    = abs(totales_eri.get("INGRESOS FINANCIEROS",0))
        gto_fin    = abs(totales_eri.get("GASTOS FINANCIEROS",0))
        dif_cambio = -totales_eri.get("DIFERENCIA EN CAMBIO NETA",0)
        provision  = abs(totales_eri.get("PROVISION DE IMPUESTOS",0))
        ganancia_bruta = ing_ord - costo_vtas
        util_ai  = (ganancia_bruta + otros_ing - gtos_adm - gtos_venta - otros_gto
                    + ing_fin - gto_fin + dif_cambio)
        util_per = util_ai - provision
        prev_eri = pd.DataFrame({
            "Línea": [NOMBRE_ERI.get(g,g) for g in GRUPOS_ERI] +
                     ["─────────────────","Utilidad antes de impuesto","─────────────────","Utilidad del periodo"],
            "Valor": [fmt_cop(-totales_eri.get(g,0)) if g=="DIFERENCIA EN CAMBIO NETA"
                      else fmt_cop(abs(totales_eri.get(g,0))) for g in GRUPOS_ERI] +
                     ["",fmt_cop(util_ai),"",fmt_cop(util_per)],
        })
        st.dataframe(prev_eri, use_container_width=True, hide_index=True, height=370)

    st.markdown("---")

    # El resultado se guarda en session_state para que NO desaparezca cuando
    # Streamlit vuelve a ejecutar el script tras presionar "Descargar"
    # (download_button dispara un rerun completo; si el archivo generado
    # solo vive dentro del `if st.button("Generar")`, en ese rerun el botón
    # vuelve a ser False y todo el bloque desaparece — eso es lo que se
    # sentía como "se sale y envía a otro lugar").
    if st.button("⚙️ Generar archivo EEFF formateado", type="primary", use_container_width=True):
        with st.spinner("Generando archivo Excel…"):
            buf_eeff = generar_excel_eeff(
                empresa, nit, periodo,
                saldos_esf, saldos_patrimonio, totales_eri,
                pivot_eri, df_eri_raw, pivot_esf, df_esf_raw,
                saldos_esf_anterior, saldos_patrimonio_anterior, totales_eri_anterior,
            )
        st.session_state["eeff_buffer"] = buf_eeff.getvalue()
        st.session_state["eeff_filename"] = f"EEFF_Formateado_{empresa.strip().replace(' ','_')}.xlsx"

    if st.session_state.get("eeff_buffer"):
        st.success("✅ Archivo generado — 8 hojas: ESF, ERI, Anexo ESF, Anexo ERI, Detalle ESF, Detalle ERI, CONSOL, ER MENSUALIZADO")
        st.download_button(
            label="📥 Descargar EEFF_Formateado.xlsx",
            data=st.session_state["eeff_buffer"],
            file_name=st.session_state.get("eeff_filename", "EEFF_Formateado.xlsx"),
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            key="dl_eeff_btn",
        )
    st.markdown('</div>', unsafe_allow_html=True)


