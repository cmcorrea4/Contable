import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
from io import BytesIO
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side
import datetime

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

GRUPOS_ERI = ["INGRESOS DE ACTIVIDADES ORDINARIAS","OTROS INGRESOS",
              "GASTOS DE ADMINISTRACION","OTROS GASTOS",
              "INGRESOS FINANCIEROS","GASTOS FINANCIEROS","PROVISION DE IMPUESTOS"]

GRUPOS_LABEL_ERI = {
    "INGRESOS DE ACTIVIDADES ORDINARIAS":"🟢 Ing. Ordinarios",
    "OTROS INGRESOS":"🟢 Otros Ingresos","GASTOS DE ADMINISTRACION":"🔴 Gtos. Admón.",
    "OTROS GASTOS":"🔴 Otros Gastos","INGRESOS FINANCIEROS":"🔵 Ing. Financieros",
    "GASTOS FINANCIEROS":"🟠 Gtos. Financieros","PROVISION DE IMPUESTOS":"🟣 Prov. Impuestos",
}
COLORES_ERI = {
    "INGRESOS DE ACTIVIDADES ORDINARIAS":"#2ecc71","OTROS INGRESOS":"#27ae60",
    "GASTOS DE ADMINISTRACION":"#e74c3c","OTROS GASTOS":"#c0392b",
    "INGRESOS FINANCIEROS":"#3498db","GASTOS FINANCIEROS":"#e67e22",
    "PROVISION DE IMPUESTOS":"#9b59b6","Total general":"#2c3e50",
}

GRUPOS_ESF_ACTIVO = [
    "EFECTIVO Y EQUIVALENTE AL EFECTIVO",
    "CUENTAS COMERCIALES POR COBRAR Y OTRAS CUENTAS POR COBRAR CORRIENTES",
    "CUENTAS POR COBRAR A PARTES RELACIONADAS",
    "ACTIVOS POR IMPUESTOS",
    "ACTIVOS POR IMPUESTOS DIFERIDOS",
    "OTROS ACTIVOS FINANCIEROS NO CTE",
]
GRUPOS_ESF_PASIVO = [
    "CUENTAS COMERCIALES POR PAGAR Y OTRAS CUENTAS POR PAGAR",
    "CUENTAS COMERCIALES POR PAGAR Y OTRAS CUENTAS POR PAGAR NO CORRIENTES",
    "CUENTAS POR PAGAR A PARTES RELACIONADAS",
    "PASIVOS POR IMPUESTOS",
    "PASIVOS POR IMPUESTOS DIFERIDOS",
    "OTROS PASIVOS NO FINANCIEROS ",
    "BENEFICIOS A LOS EMPLEADOS",
]
GRUPOS_ESF_ORDEN = GRUPOS_ESF_ACTIVO + GRUPOS_ESF_PASIVO

GRUPOS_LABEL_ESF = {
    "EFECTIVO Y EQUIVALENTE AL EFECTIVO":"🟢 Efectivo y Equiv.",
    "CUENTAS COMERCIALES POR COBRAR Y OTRAS CUENTAS POR COBRAR CORRIENTES":"🟢 CxC Corrientes",
    "CUENTAS POR COBRAR A PARTES RELACIONADAS":"🟢 CxC Partes Rel.",
    "ACTIVOS POR IMPUESTOS":"🟢 Activos x Impuestos",
    "ACTIVOS POR IMPUESTOS DIFERIDOS":"🟢 Impuestos Diferidos A.",
    "OTROS ACTIVOS FINANCIEROS NO CTE":"🟢 Otros Activos Fin.",
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
    "CUENTAS COMERCIALES POR COBRAR Y OTRAS CUENTAS POR COBRAR CORRIENTES":"#27ae60",
    "CUENTAS POR COBRAR A PARTES RELACIONADAS":"#1a9850",
    "ACTIVOS POR IMPUESTOS":"#52b788","ACTIVOS POR IMPUESTOS DIFERIDOS":"#74c69d",
    "OTROS ACTIVOS FINANCIEROS NO CTE":"#95d5b2",
    "CUENTAS COMERCIALES POR PAGAR Y OTRAS CUENTAS POR PAGAR":"#e74c3c",
    "CUENTAS COMERCIALES POR PAGAR Y OTRAS CUENTAS POR PAGAR NO CORRIENTES":"#c0392b",
    "CUENTAS POR PAGAR A PARTES RELACIONADAS":"#e57373",
    "PASIVOS POR IMPUESTOS":"#ef9a9a","PASIVOS POR IMPUESTOS DIFERIDOS":"#e67e22",
    "OTROS PASIVOS NO FINANCIEROS ":"#d35400","BENEFICIOS A LOS EMPLEADOS":"#9b59b6",
}

# Nombres para hojas formateadas (alineados con el archivo de referencia)
NOMBRE_ESF = {
    "EFECTIVO Y EQUIVALENTE AL EFECTIVO":                                   "  Efectivo y equivalentes al efectivo",
    "CUENTAS COMERCIALES POR COBRAR Y OTRAS CUENTAS POR COBRAR CORRIENTES": "  Cuentas comerciales por cobrar y otras cuentas por cobrar",
    "CUENTAS POR COBRAR A PARTES RELACIONADAS":                             "  Cuentas por cobrar a partes relacionadas",
    "ACTIVOS POR IMPUESTOS":                                                "  Activos por impuestos",
    "ACTIVOS POR IMPUESTOS DIFERIDOS":                                      "  Activos por impuestos diferidos",
    "OTROS ACTIVOS FINANCIEROS NO CTE":                                     "  Otros activos financieros",
    "CUENTAS COMERCIALES POR PAGAR Y OTRAS CUENTAS POR PAGAR":              "  Cuentas comerciales por pagar y otras cuentas por pagar",
    "CUENTAS COMERCIALES POR PAGAR Y OTRAS CUENTAS POR PAGAR NO CORRIENTES":"  Cuentas comerciales por pagar y otras cuentas por pagar (NC)",
    "CUENTAS POR PAGAR A PARTES RELACIONADAS":                              "  Cuentas por pagar a partes relacionadas",
    "PASIVOS POR IMPUESTOS":                                                "  Pasivos por impuestos",
    "PASIVOS POR IMPUESTOS DIFERIDOS":                                      "  Pasivos por impuestos diferidos",
    "OTROS PASIVOS NO FINANCIEROS ":                                        "  Otros pasivos no financieros",
    "BENEFICIOS A LOS EMPLEADOS":                                           "  Beneficios a los empleados",
}
NOMBRE_ERI = {
    "INGRESOS DE ACTIVIDADES ORDINARIAS": "Ingresos de actividades ordinarias",
    "OTROS INGRESOS":                     "Otros ingresos",
    "GASTOS DE ADMINISTRACION":           "Gastos de administración",
    "OTROS GASTOS":                       "Otros gastos",
    "INGRESOS FINANCIEROS":               "Ingresos financieros",
    "GASTOS FINANCIEROS":                 "Gastos financieros",
    "PROVISION DE IMPUESTOS":             "Ingreso (gasto) por impuesto",
}

# ══════════════════════════════════════════════════════════════════════════════
# FUNCIONES DE DATOS
# ══════════════════════════════════════════════════════════════════════════════
def fmt_cop(val):
    if pd.isna(val): return "-"
    if val < 0: return f"($ {abs(val):,.0f})"
    return f"$ {val:,.0f}"


# Nombres que indican "período anterior" en la columna Mes del archivo fuente
_MES_ANT_TOKENS = {"PERIODO ANTERIOR", "PERÍODO ANTERIOR", "PERIOD ANTERIOR",
                   "ANO ANTERIOR", "AÑO ANTERIOR", "ANTERIOR"}

def _es_mes_anterior(nombre_mes: str) -> bool:
    return str(nombre_mes).strip().upper() in _MES_ANT_TOKENS


@st.cache_data(show_spinner="Procesando archivo…")
def procesar_archivo(file_bytes: bytes):
    xls = pd.ExcelFile(BytesIO(file_bytes))
    if "terceros_" not in xls.sheet_names:
        st.error("El archivo no contiene la hoja 'terceros_'.")
        st.stop()

    df = pd.read_excel(BytesIO(file_bytes), sheet_name="terceros_", header=0)
    df.columns = [c.strip() for c in df.columns]
    df["Saldo Mes"] = pd.to_numeric(df["Saldo Mes"], errors="coerce").fillna(0)

    # ── Separar filas de "período anterior" si existen en la misma hoja ─────
    mask_ant = df["Mes"].apply(_es_mes_anterior)
    df_ant   = df[mask_ant].copy()
    df       = df[~mask_ant].copy()

    # Calcular saldos_esf_ant y totales_eri_ant a partir de df_ant
    saldos_esf_ant_auto  = None
    totales_eri_ant_auto = None

    if not df_ant.empty:
        # ESF anterior — saldo por grupo (igual que saldos_esf del período actual)
        esf_ant_grp = df_ant[df_ant["Grupo"].isin(GRUPOS_ESF_ORDEN)]
        saldos_esf_ant_auto = {
            g: esf_ant_grp[esf_ant_grp["Grupo"] == g]["Saldo Mes"].sum()
            for g in GRUPOS_ESF_ORDEN
        }
        # ERI anterior — acumulado por grupo
        eri_ant_grp = df_ant[df_ant["Grupo"].isin(GRUPOS_ERI)]
        totales_eri_ant_auto = {
            g: eri_ant_grp[eri_ant_grp["Grupo"] == g]["Saldo Mes"].sum()
            for g in GRUPOS_ERI
        }

    def build_pivot(df_src, grupos):
        df_f = df_src[df_src["Grupo"].isin(grupos)].copy()
        pivot = df_f.groupby(["Grupo","Mes"])["Saldo Mes"].sum().unstack(fill_value=0)
        meses_d = [m for m in MESES_ORDEN if m in pivot.columns]
        pivot = pivot.reindex(columns=meses_d, fill_value=0)
        pivot = pivot.reindex([g for g in grupos if g in pivot.index], fill_value=0)
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
    pivot_base = pivot_base.reindex([g for g in GRUPOS_ESF_ORDEN if g in pivot_base.index], fill_value=0)
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

    return (df_eri_raw, pivot_eri, df_esf_raw, pivot_esf, saldos_esf, totales_eri,
            ultimo_mes, saldos_esf_ant_auto, totales_eri_ant_auto)


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


def generar_hoja_esf(ws, empresa, nit, periodo, saldos, saldos_ant=None):
    """saldos_ant: dict grupo→valor para la columna 'Período anterior' (F/M).
       Si es None o vacío, las columnas F/M quedan en blanco."""
    if saldos_ant is None:
        saldos_ant = {}

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

    # Títulos sección
    ws["B9"].value = "ACTIVOS";              ws["B9"].font = F(bold=True, size=14)
    ws["I9"].value = "PASIVOS Y PATRIMONIO"; ws["I9"].font = F(bold=True, size=14)
    ws["B11"].value = "Activos corrientes:"; ws["B11"].font = F(size=13)
    ws["I11"].value = "Pasivos corrientes:"; ws["I11"].font = F(size=13)

    def v(g, src):    return abs(src.get(g, 0)) or None
    def va(g):        return abs(saldos_ant.get(g, 0)) or None   # período anterior

    # ── Activos corrientes (col B/D/F) ────────────────────────────────────────
    act_cte_map = [
        (12, "EFECTIVO Y EQUIVALENTE AL EFECTIVO"),
        (14, "ACTIVOS POR IMPUESTOS"),
        (15, "CUENTAS COMERCIALES POR COBRAR Y OTRAS CUENTAS POR COBRAR CORRIENTES"),
        (16, "CUENTAS POR COBRAR A PARTES RELACIONADAS"),
    ]
    for row_n, g in act_cte_map:
        ws[f"B{row_n}"].value = NOMBRE_ESF[g]; ws[f"B{row_n}"].font = F(size=13)
        ws[f"B{row_n}"].alignment = A("left")
        ws[f"D{row_n}"].value = v(g, saldos);  ws[f"D{row_n}"].number_format = FMT_COP_XL
        ws[f"D{row_n}"].font = F(size=13)
        ws[f"F{row_n}"].value = va(g);          ws[f"F{row_n}"].number_format = FMT_COP_XL
        ws[f"F{row_n}"].font = F(size=13)

    total_act_cte     = sum(abs(saldos.get(g,0))     for _,g in act_cte_map)
    total_act_cte_ant = sum(abs(saldos_ant.get(g,0)) for _,g in act_cte_map)
    ws["B19"].value = "Total activos corrientes"; ws["B19"].font = F(size=13)
    ws["D19"].value = total_act_cte;     ws["D19"].number_format = FMT_SUB_XL
    ws["D19"].font = F(size=13); ws["D19"].border = Border(bottom=thin())
    if total_act_cte_ant:
        ws["F19"].value = total_act_cte_ant; ws["F19"].number_format = FMT_SUB_XL
        ws["F19"].font = F(size=13); ws["F19"].border = Border(bottom=thin())

    # ── Activos no corrientes (col B/D/F) ────────────────────────────────────
    ws["B24"].value = "Activos no corrientes:"; ws["B24"].font = F(size=13)
    act_nct_map = [
        (25, "OTROS ACTIVOS FINANCIEROS NO CTE"),
        (26, "ACTIVOS POR IMPUESTOS DIFERIDOS"),
    ]
    for row_n, g in act_nct_map:
        ws[f"B{row_n}"].value = NOMBRE_ESF[g]; ws[f"B{row_n}"].font = F(size=13)
        ws[f"B{row_n}"].alignment = A("left")
        ws[f"D{row_n}"].value = v(g, saldos);  ws[f"D{row_n}"].number_format = FMT_COP_XL
        ws[f"D{row_n}"].font = F(size=13)
        ws[f"F{row_n}"].value = va(g);          ws[f"F{row_n}"].number_format = FMT_COP_XL
        ws[f"F{row_n}"].font = F(size=13)

    total_act_nct     = sum(abs(saldos.get(g,0))     for _,g in act_nct_map)
    total_act_nct_ant = sum(abs(saldos_ant.get(g,0)) for _,g in act_nct_map)
    ws["B28"].value = "Total activos no corrientes"; ws["B28"].font = F(size=13)
    ws["D28"].value = total_act_nct;     ws["D28"].number_format = FMT_SUB_XL
    ws["D28"].font = F(size=13); ws["D28"].border = Border(bottom=thin())
    if total_act_nct_ant:
        ws["F28"].value = total_act_nct_ant; ws["F28"].number_format = FMT_SUB_XL
        ws["F28"].font = F(size=13); ws["F28"].border = Border(bottom=thin())

    total_act     = total_act_cte     + total_act_nct
    total_act_ant = total_act_cte_ant + total_act_nct_ant
    ws["B38"].value = "TOTAL ACTIVOS"; ws["B38"].font = F(bold=True, size=14)
    ws["D38"].value = total_act;     ws["D38"].number_format = FMT_COP_XL
    ws["D38"].font = F(bold=True, size=14); ws["D38"].border = Border(bottom=double())
    if total_act_ant:
        ws["F38"].value = total_act_ant; ws["F38"].number_format = FMT_COP_XL
        ws["F38"].font = F(bold=True, size=14); ws["F38"].border = Border(bottom=double())

    # ── Pasivos corrientes (col I/K/M) ────────────────────────────────────────
    pas_cte_map = [
        (12, "OTROS PASIVOS NO FINANCIEROS "),
        (14, "PASIVOS POR IMPUESTOS"),
        (15, "CUENTAS COMERCIALES POR PAGAR Y OTRAS CUENTAS POR PAGAR"),
        (16, "CUENTAS POR PAGAR A PARTES RELACIONADAS"),
        (18, "BENEFICIOS A LOS EMPLEADOS"),
    ]
    for row_n, g in pas_cte_map:
        ws[f"I{row_n}"].value = NOMBRE_ESF[g]; ws[f"I{row_n}"].font = F(size=13)
        ws[f"I{row_n}"].alignment = A("left")
        ws[f"K{row_n}"].value = v(g, saldos);  ws[f"K{row_n}"].number_format = FMT_COP_XL
        ws[f"K{row_n}"].font = F(size=13)
        ws[f"M{row_n}"].value = va(g);          ws[f"M{row_n}"].number_format = FMT_COP_XL
        ws[f"M{row_n}"].font = F(size=13)

    total_pas_cte     = sum(abs(saldos.get(g,0))     for _,g in pas_cte_map)
    total_pas_cte_ant = sum(abs(saldos_ant.get(g,0)) for _,g in pas_cte_map)
    ws["I20"].value = "Total pasivos corrientes"; ws["I20"].font = F(size=13)
    ws["K20"].value = total_pas_cte;     ws["K20"].number_format = FMT_SUB_XL
    ws["K20"].font = F(size=13); ws["K20"].border = Border(bottom=thin())
    if total_pas_cte_ant:
        ws["M20"].value = total_pas_cte_ant; ws["M20"].number_format = FMT_SUB_XL
        ws["M20"].font = F(size=13); ws["M20"].border = Border(bottom=thin())

    # ── Pasivos no corrientes (col I/K/M) ────────────────────────────────────
    ws["I22"].value = "Pasivos no corrientes:"; ws["I22"].font = F(size=13)
    pas_nct_map = [
        (23, "CUENTAS COMERCIALES POR PAGAR Y OTRAS CUENTAS POR PAGAR NO CORRIENTES"),
        (24, "PASIVOS POR IMPUESTOS DIFERIDOS"),
    ]
    for row_n, g in pas_nct_map:
        ws[f"I{row_n}"].value = NOMBRE_ESF[g]; ws[f"I{row_n}"].font = F(size=13)
        ws[f"I{row_n}"].alignment = A("left")
        ws[f"K{row_n}"].value = v(g, saldos);  ws[f"K{row_n}"].number_format = FMT_COP_XL
        ws[f"K{row_n}"].font = F(size=13)
        ws[f"M{row_n}"].value = va(g);          ws[f"M{row_n}"].number_format = FMT_COP_XL
        ws[f"M{row_n}"].font = F(size=13)

    total_pas_nct     = sum(abs(saldos.get(g,0))     for _,g in pas_nct_map)
    total_pas_nct_ant = sum(abs(saldos_ant.get(g,0)) for _,g in pas_nct_map)
    ws["I26"].value = "Total pasivos no corrientes"; ws["I26"].font = F(size=13)
    ws["K26"].value = total_pas_nct;     ws["K26"].number_format = FMT_SUB_XL
    ws["K26"].font = F(size=13); ws["K26"].border = Border(bottom=thin())
    if total_pas_nct_ant:
        ws["M26"].value = total_pas_nct_ant; ws["M26"].number_format = FMT_SUB_XL
        ws["M26"].font = F(size=13); ws["M26"].border = Border(bottom=thin())

    total_pas     = total_pas_cte     + total_pas_nct
    total_pas_ant = total_pas_cte_ant + total_pas_nct_ant
    ws["I28"].value = "TOTAL PASIVOS"; ws["I28"].font = F(bold=True, size=14)
    ws["K28"].value = total_pas;     ws["K28"].number_format = FMT_COP_XL
    ws["K28"].font = F(bold=True, size=14); ws["K28"].border = Border(bottom=double())
    if total_pas_ant:
        ws["M28"].value = total_pas_ant; ws["M28"].number_format = FMT_COP_XL
        ws["M28"].font = F(bold=True, size=14); ws["M28"].border = Border(bottom=double())

    # ── Patrimonio (col I/K) ──────────────────────────────────────────────────
    ws["I30"].value = "Patrimonio:"; ws["I30"].font = F(size=13)
    for r, lbl in [(31,"  Capital emitido"),(32,"  Superávit de capital"),
                   (33,"  Utilidad acumulada"),(34,"  Utilidad del periodo")]:
        ws[f"I{r}"].value = lbl; ws[f"I{r}"].font = F(size=13); ws[f"I{r}"].alignment = A("left")

    patrimonio     = total_act - total_pas
    patrimonio_ant = total_act_ant - total_pas_ant
    ws["K36"].value = patrimonio;     ws["K36"].number_format = FMT_COP_XL
    ws["K36"].font = F(bold=True, size=14); ws["K36"].border = Border(bottom=double())
    if total_act_ant:
        ws["M36"].value = patrimonio_ant; ws["M36"].number_format = FMT_COP_XL
        ws["M36"].font = F(bold=True, size=14); ws["M36"].border = Border(bottom=double())

    ws["I36"].value = "TOTAL PATRIMONIO"; ws["I36"].font = F(bold=True, size=14)
    ws["I38"].value = "TOTAL PASIVOS Y PATRIMONIO"; ws["I38"].font = F(bold=True, size=14)
    # TOTAL PASIVOS Y PATRIMONIO = TOTAL ACTIVOS (ecuación fundamental)
    ws["K38"].value = total_act; ws["K38"].number_format = FMT_COP_XL
    ws["K38"].font = F(bold=True, size=14); ws["K38"].border = Border(bottom=double())
    if total_act_ant:
        ws["M38"].value = total_act_ant; ws["M38"].number_format = FMT_COP_XL
        ws["M38"].font = F(bold=True, size=14); ws["M38"].border = Border(bottom=double())

    ws["B40"].value = "Las notas adjuntas forman parte integral de estos estados financieros."
    ws["B40"].font = F(size=11)


def generar_hoja_eri(ws, empresa, nit, periodo, totales, totales_ant=None):
    """totales_ant: dict grupo→valor para la columna F 'Período anterior'.
       Si es None o vacío, la columna F queda en blanco."""
    if totales_ant is None:
        totales_ant = {}

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

    for col, val in [("D","Acumulado"),("F","Acumulado")]:
        ws[f"{col}6"].value = val; ws[f"{col}6"].font = F(bold=True, size=14)
        ws[f"{col}6"].alignment = A("center")

    ws["C7"].value = "NOTA"; ws["C7"].font = F(bold=True, size=13); ws["C7"].alignment = A("center")
    for col, val in [("D",2025),("F",2024)]:
        ws[f"{col}7"].value = val; ws[f"{col}7"].font = F(bold=True, size=13)
        ws[f"{col}7"].alignment = A("center")
    for col in ["D","F"]:
        ws[f"{col}8"].value = "$"; ws[f"{col}8"].font = F(bold=True, size=13)
        ws[f"{col}8"].alignment = A("center")

    # Valores con signo contable — período actual
    ing_ord  = abs(totales.get("INGRESOS DE ACTIVIDADES ORDINARIAS", 0))
    otros_ing= abs(totales.get("OTROS INGRESOS", 0))
    gtos_adm = abs(totales.get("GASTOS DE ADMINISTRACION", 0))
    otros_gto= abs(totales.get("OTROS GASTOS", 0))
    ing_fin  = abs(totales.get("INGRESOS FINANCIEROS", 0))
    gto_fin  = abs(totales.get("GASTOS FINANCIEROS", 0))
    provision= abs(totales.get("PROVISION DE IMPUESTOS", 0))
    ganancia  = ing_ord
    util_ai   = ing_ord + otros_ing - gtos_adm - otros_gto + ing_fin - gto_fin
    util_per  = util_ai - provision

    # Valores período anterior
    ing_ord_a  = abs(totales_ant.get("INGRESOS DE ACTIVIDADES ORDINARIAS", 0))
    otros_ing_a= abs(totales_ant.get("OTROS INGRESOS", 0))
    gtos_adm_a = abs(totales_ant.get("GASTOS DE ADMINISTRACION", 0))
    otros_gto_a= abs(totales_ant.get("OTROS GASTOS", 0))
    ing_fin_a  = abs(totales_ant.get("INGRESOS FINANCIEROS", 0))
    gto_fin_a  = abs(totales_ant.get("GASTOS FINANCIEROS", 0))
    provision_a= abs(totales_ant.get("PROVISION DE IMPUESTOS", 0))
    ganancia_a  = ing_ord_a
    util_ai_a   = ing_ord_a + otros_ing_a - gtos_adm_a - otros_gto_a + ing_fin_a - gto_fin_a
    util_per_a  = util_ai_a - provision_a

    hay_ant = bool(totales_ant)  # sólo escribir columna F si se cargó archivo anterior

    # (row, label, nota, val_actual, val_anterior, bold, border)
    lineas = [
        (10, "Ingresos de actividades ordinarias", 13,  ing_ord,    ing_ord_a,   False, None),
        (12, "Ganancia bruta",                     None,ganancia,   ganancia_a,  True,  None),
        (14, "Otros ingresos",                     14,  otros_ing,  otros_ing_a, False, None),
        (15, "Gastos de administración",           15, -gtos_adm,  -gtos_adm_a,  False, None),
        (16, "Otros gastos",                       16, -otros_gto, -otros_gto_a, False, None),
        (17, "Ingresos financieros",               17,  ing_fin,    ing_fin_a,   False, None),
        (18, "Gastos financieros",                 16, -gto_fin,   -gto_fin_a,   False, None),
        (21, "Utilidad antes de impuesto",         None,util_ai,    util_ai_a,   True,  None),
        (23, "Ingreso (gasto) por impuesto",       19, -provision, -provision_a, False, None),
        (25, "Utilidad (pérdida) del periodo",     None,util_per,   util_per_a,  True,  None),
        (29, "Resultado integral total",           None,util_per,   util_per_a,  True,  "double"),
    ]

    for row_n, label, nota, val, val_a, bold, border in lineas:
        ws[f"B{row_n}"].value = label; ws[f"B{row_n}"].font = F(bold=bold, size=13)
        if nota:
            ws[f"C{row_n}"].value = nota; ws[f"C{row_n}"].font = F(size=13)
            ws[f"C{row_n}"].alignment = A("center")
        # Columna D — período actual
        c = ws[f"D{row_n}"]
        c.value = val if val != 0 else None
        c.number_format = FMT_COP_XL; c.font = F(bold=bold, size=13)
        if border == "double": c.border = Border(bottom=double())
        # Columna F — período anterior
        if hay_ant:
            cf = ws[f"F{row_n}"]
            cf.value = val_a if val_a != 0 else None
            cf.number_format = FMT_COP_XL; cf.font = F(bold=bold, size=13)
            if border == "double": cf.border = Border(bottom=double())

    ws["B31"].value = "Las notas adjuntas forman parte integral de estos estados financieros."
    ws["B31"].font = F(size=12)


def generar_hoja_er_mensualizado(ws, empresa, df_eri_raw):
    """Hoja ER MENSUALIZADO: P&L por mes con detalle de cuentas."""
    from openpyxl.styles import PatternFill, colors
    from openpyxl.utils import get_column_letter

    MESES = [m for m in MESES_ORDEN if m in df_eri_raw["Mes"].unique()]
    MESES_ABR = [MESES_ABREV.get(m, m[:3]) for m in MESES]

    # Pivots por grupo y por cuenta
    df_raw = df_eri_raw.copy()
    df_raw["Cuenta_llave"] = df_raw["Codigo"].astype(str) + "  " + df_raw["Nombre cuenta"].str.strip()

    pivot_grupo = (df_raw.groupby(["Grupo","Mes"])["Saldo Mes"].sum()
                   .unstack(fill_value=0).reindex(columns=MESES, fill_value=0))
    pivot_cuenta = (df_raw.groupby(["Cuenta_llave","Grupo","Mes"])["Saldo Mes"].sum()
                    .unstack(fill_value=0).reindex(columns=MESES, fill_value=0))
    pivot_cuenta = pivot_cuenta.reset_index()

    # Estilos
    azul_osc  = PatternFill("solid", fgColor="1E3A5F")
    azul_med  = PatternFill("solid", fgColor="2E6DA4")
    azul_cla  = PatternFill("solid", fgColor="D6E4F0")
    blanco    = Font(color="FFFFFF", bold=True, size=11, name="Calibri")
    negro_b   = Font(bold=True, size=10, name="Calibri")
    negro_n   = Font(bold=False, size=10, name="Calibri")
    centro    = Alignment(horizontal="center", vertical="center")
    izq       = Alignment(horizontal="left",   vertical="center")

    def fmt_num(ws, cell):
        cell.number_format = '#,##0_);(#,##0);"-"'
        cell.alignment = Alignment(horizontal="right", vertical="center")

    # Anchos
    ws.column_dimensions["A"].width = 12
    ws.column_dimensions["B"].width = 42
    for i in range(len(MESES)):
        ws.column_dimensions[get_column_letter(i+3)].width = 16
    ws.column_dimensions[get_column_letter(len(MESES)+3)].width = 18  # ACUMULADO

    # Fila 1 — empresa
    ws.row_dimensions[1].height = 20
    c = ws.cell(1, 2, empresa); c.font = Font(bold=True, size=13, name="Calibri")

    # Fila 2 — título con fondo azul oscuro
    ws.row_dimensions[2].height = 22
    ws.merge_cells(start_row=2, start_column=2, end_row=2,
                   end_column=len(MESES)+3)
    c = ws.cell(2, 2, "ESTADO DE RESULTADOS INTEGRAL  MENSUALIZADO - CONSOLIDADO")
    c.font = blanco; c.fill = azul_osc; c.alignment = centro

    # Fila 4 — cabecera meses
    ws.row_dimensions[4].height = 18
    ws.cell(4, 2, "").font = negro_b
    for i, abr in enumerate(MESES_ABR):
        c = ws.cell(4, i+3, abr.upper() + " ")
        c.font = blanco; c.fill = azul_med; c.alignment = centro
    c = ws.cell(4, len(MESES)+3, "ACUMULADO")
    c.font = blanco; c.fill = azul_med; c.alignment = centro

    # Mapeo grupos → filas del P&L
    # (etiqueta_grupo_en_data, label_display, es_ingreso, es_subtotal_label)
    ESTRUCTURA_PL = [
        # Ingresos
        ("INGRESOS DE ACTIVIDADES ORDINARIAS", "Ingresos de actividades ordinarias", True,  False),
        (None, "VENTAS NETAS",                                                        True,  True),
        (None, "Costo del servicio",                                                  False, False),
        (None, "UTILIDAD BRUTA",                                                      True,  True),
        # Gastos operación
        ("GASTOS DE ADMINISTRACION",          "Gastos Administración",               False, False),
        (None, "GASTOS DE OPERACION",                                                 False, True),
        (None, "UTILIDAD OPERATIVA",                                                  True,  True),
        # Otros
        ("OTROS GASTOS",                      "Otros gastos",                        False, False),
        ("GASTOS FINANCIEROS",                "Gastos financieros",                  False, False),
        ("OTROS INGRESOS",                    "Otros ingresos",                      True,  False),
        (None, "Diferencia en cambio neta",                                           True,  False),
        (None, "UTILIDAD ANTES DE IMPTOS",                                            True,  True),
        ("PROVISION DE IMPUESTOS",            "Provisión impuesto de renta",         False, False),
        (None, "UTILIDAD (PERDIDA) NETA",                                             True,  True),
    ]

    # Precalcular acumulados por grupo
    def gacc(g):
        if g in pivot_grupo.index:
            return {m: pivot_grupo.loc[g, m] for m in MESES}
        return {m: 0 for m in MESES}

    ing_m   = gacc("INGRESOS DE ACTIVIDADES ORDINARIAS")
    gto_m   = gacc("GASTOS DE ADMINISTRACION")
    otrg_m  = gacc("OTROS GASTOS")
    gfin_m  = gacc("GASTOS FINANCIEROS")
    oing_m  = gacc("OTROS INGRESOS")
    prov_m  = gacc("PROVISION DE IMPUESTOS")

    def ub(m):   return ing_m[m]
    def uop(m):  return ing_m[m] - abs(gto_m[m])
    def uai(m):  return uop(m) - abs(otrg_m[m]) - abs(gfin_m[m]) + abs(oing_m[m])
    def unet(m): return uai(m) - abs(prov_m[m])

    subtotales = {
        "VENTAS NETAS":           {m: ing_m[m]  for m in MESES},
        "Costo del servicio":     {m: 0          for m in MESES},
        "UTILIDAD BRUTA":         {m: ub(m)      for m in MESES},
        "GASTOS DE OPERACION":    {m: abs(gto_m[m]) for m in MESES},
        "UTILIDAD OPERATIVA":     {m: uop(m)     for m in MESES},
        "Diferencia en cambio neta": {m: 0        for m in MESES},
        "UTILIDAD ANTES DE IMPTOS":  {m: uai(m)  for m in MESES},
        "UTILIDAD (PERDIDA) NETA":   {m: unet(m) for m in MESES},
    }

    fila = 5
    for grupo_key, label, es_ing, es_sub in ESTRUCTURA_PL:
        ws.row_dimensions[fila].height = 16

        if es_sub:
            # Fila subtotal — fondo azul claro + negrita
            c = ws.cell(fila, 2, label); c.font = negro_b; c.fill = azul_cla
            vals = subtotales.get(label, {m: 0 for m in MESES})
            for i, m in enumerate(MESES):
                cell = ws.cell(fila, i+3, vals[m] if vals[m] != 0 else None)
                cell.fill = azul_cla; fmt_num(ws, cell); cell.font = negro_b
            acum = sum(vals.values())
            cell = ws.cell(fila, len(MESES)+3, acum if acum != 0 else None)
            cell.fill = azul_cla; fmt_num(ws, cell); cell.font = negro_b
            fila += 1

            # Filas de detalle de cuentas del grupo anterior
            if grupo_key is None and es_sub:
                pass  # subtotales sin grupo no tienen detalle
        else:
            if grupo_key and grupo_key in pivot_grupo.index:
                # Fila de grupo (encabezado)
                c = ws.cell(fila, 2, label); c.font = negro_n
                vals = {m: pivot_grupo.loc[grupo_key, m] for m in MESES}
                for i, m in enumerate(MESES):
                    cell = ws.cell(fila, i+3, abs(vals[m]) if vals[m] != 0 else None)
                    fmt_num(ws, cell); cell.font = negro_n
                acum = sum(abs(v) for v in vals.values())
                cell = ws.cell(fila, len(MESES)+3, acum if acum != 0 else None)
                fmt_num(ws, cell); cell.font = negro_n
                fila += 1

                # Cuentas detalle
                cuentas_grp = pivot_cuenta[pivot_cuenta["Grupo"] == grupo_key]
                for _, cr in cuentas_grp.iterrows():
                    ws.row_dimensions[fila].height = 15
                    lbl = "          " + str(cr["Cuenta_llave"])
                    c = ws.cell(fila, 2, lbl); c.font = negro_n; c.alignment = izq
                    for i, m in enumerate(MESES):
                        v = cr[m] if m in cr.index else 0
                        cell = ws.cell(fila, i+3, abs(v) if v != 0 else None)
                        fmt_num(ws, cell); cell.font = negro_n
                    acum_c = sum(abs(cr[m]) for m in MESES if m in cr.index)
                    cell = ws.cell(fila, len(MESES)+3, acum_c if acum_c != 0 else None)
                    fmt_num(ws, cell); cell.font = negro_n
                    fila += 1
            else:
                # Fila sin datos (ej. Costo del servicio, Diferencia en cambio)
                c = ws.cell(fila, 2, label); c.font = negro_n
                for i in range(len(MESES)+1):
                    ws.cell(fila, i+3, None)
                fila += 1

    # Autofilter en fila 4
    ws.auto_filter.ref = f"B4:{get_column_letter(len(MESES)+3)}4"


def generar_hoja_consol(ws, empresa, df_eri_raw):
    """Hoja CONSOL: todos los grupos ERI, detalle de cuentas por mes + ACUMULADO.

    Cubre TODOS los GRUPOS_ERI (no sólo Gastos).
    Columnas: Nombre cuenta | Ene | Feb | ... | ACUMULADO
    ACUMULADO de cada grupo = suma YTD de todas sus cuentas (coincide con ERI).
    """
    from openpyxl.styles import PatternFill
    from openpyxl.utils import get_column_letter

    MESES = [m for m in MESES_ORDEN if m in df_eri_raw["Mes"].unique()]
    MESES_ABR = [MESES_ABREV.get(m, m[:3]).upper() for m in MESES]
    n_mes = len(MESES)
    # total columnas: col A (descripción) + meses + ACUMULADO
    N_COLS = 1 + n_mes + 1

    azul_osc  = PatternFill("solid", fgColor="1E3A5F")
    azul_med  = PatternFill("solid", fgColor="2E6DA4")
    azul_cla  = PatternFill("solid", fgColor="D6E4F0")
    gris_alt  = PatternFill("solid", fgColor="F2F4F7")
    blanco    = Font(color="FFFFFF", bold=True, size=11, name="Calibri")
    negro_b   = Font(bold=True,  size=10, name="Calibri")
    negro_n   = Font(bold=False, size=10, name="Calibri")
    izq       = Alignment(horizontal="left",  vertical="center")
    der       = Alignment(horizontal="right", vertical="center")

    thin_s  = Side(border_style="thin",   color="CCCCCC")
    thick_s = Side(border_style="medium", color="2E6DA4")
    borde_n = Border(left=thin_s,  right=thin_s,  top=thin_s,  bottom=thin_s)
    borde_g = Border(left=thick_s, right=thick_s, top=thick_s, bottom=thick_s)

    def fmt_num(cell, bold=False):
        cell.number_format = FMT_COP_XL
        cell.alignment = der
        cell.font = negro_b if bold else negro_n

    # Anchos de columna
    ws.column_dimensions["A"].width = 50
    for j in range(2, N_COLS):
        ws.column_dimensions[get_column_letter(j)].width = 15
    ws.column_dimensions[get_column_letter(N_COLS)].width = 18  # ACUMULADO

    # Fila 1: título fusionado
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=N_COLS)
    c = ws.cell(1, 1, f"CONSOLIDADO ESTADO DE RESULTADOS — {empresa}")
    c.font = blanco; c.fill = azul_osc
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 22

    # Fila 2: encabezados
    ws.row_dimensions[2].height = 18
    c = ws.cell(2, 1, "Grupo / Cuenta"); c.font = blanco; c.fill = azul_med
    c.alignment = Alignment(horizontal="center", vertical="center")
    for j, abr in enumerate(MESES_ABR, 2):
        c = ws.cell(2, j, abr); c.font = blanco; c.fill = azul_med
        c.alignment = Alignment(horizontal="center", vertical="center")
    c = ws.cell(2, N_COLS, "ACUMULADO"); c.font = blanco; c.fill = azul_med
    c.alignment = Alignment(horizontal="center", vertical="center")

    # Preparar pivot por cuenta × mes
    df_raw = df_eri_raw.copy()
    df_raw["_cuenta"] = df_raw["Codigo"].astype(str).str.strip() + "  " + df_raw["Nombre cuenta"].str.strip()
    pivot_cta = (df_raw.groupby(["_cuenta", "Grupo", "Mes"])["Saldo Mes"].sum()
                 .unstack(fill_value=0).reindex(columns=MESES, fill_value=0)
                 .reset_index())

    # Mapeo grupo → (etiqueta sección, etiqueta total)
    LABEL_GRUPO = {
        "INGRESOS DE ACTIVIDADES ORDINARIAS": ("INGRESOS DE ACTIVIDADES ORDINARIAS",  "TOTAL INGRESOS ORDINARIOS"),
        "OTROS INGRESOS":                     ("OTROS INGRESOS",                      "TOTAL OTROS INGRESOS"),
        "GASTOS DE ADMINISTRACION":           ("GASTOS DE ADMINISTRACIÓN",            "TOTAL GASTOS DE ADMINISTRACIÓN"),
        "OTROS GASTOS":                       ("OTROS GASTOS",                        "TOTAL OTROS GASTOS"),
        "INGRESOS FINANCIEROS":               ("INGRESOS FINANCIEROS",                "TOTAL INGRESOS FINANCIEROS"),
        "GASTOS FINANCIEROS":                 ("GASTOS NO OPERACIONALES",             "TOTAL GASTOS NO OPERACIONALES"),
        "PROVISION DE IMPUESTOS":             ("PROVISIÓN DE IMPUESTOS",              "TOTAL PROVISIÓN DE IMPUESTOS"),
    }

    fila = 3
    totales_gran_mes = {m: 0.0 for m in MESES}

    for grupo_key in GRUPOS_ERI:
        cuentas_grp = pivot_cta[pivot_cta["Grupo"] == grupo_key]
        if cuentas_grp.empty:
            continue

        titulo_secc, titulo_total = LABEL_GRUPO.get(grupo_key, (grupo_key, f"TOTAL {grupo_key}"))

        # Fila encabezado de grupo
        ws.row_dimensions[fila].height = 16
        c = ws.cell(fila, 1, titulo_secc)
        c.font = negro_b; c.fill = azul_cla; c.alignment = izq; c.border = borde_g
        for j in range(2, N_COLS + 1):
            cell = ws.cell(fila, j); cell.fill = azul_cla; cell.border = borde_n
        fila += 1

        totales_sec_mes = {m: 0.0 for m in MESES}

        for ri, (_, cr) in enumerate(cuentas_grp.iterrows()):
            ws.row_dimensions[fila].height = 15
            lbl = str(cr["_cuenta"])
            fill_row = gris_alt if ri % 2 == 0 else None

            c = ws.cell(fila, 1, "  " + lbl)
            c.font = negro_n; c.alignment = izq; c.border = borde_n
            if fill_row: c.fill = fill_row

            row_acum = 0.0
            for j, m in enumerate(MESES, 2):
                v = abs(float(cr[m])) if m in cr.index else 0.0
                totales_sec_mes[m] += v
                cell = ws.cell(fila, j, v if v != 0 else None)
                fmt_num(cell); cell.border = borde_n
                if fill_row: cell.fill = fill_row
                row_acum += v

            cell = ws.cell(fila, N_COLS, row_acum if row_acum != 0 else None)
            fmt_num(cell); cell.border = borde_n
            if fill_row: cell.fill = fill_row
            fila += 1

        # Fila total de sección
        ws.row_dimensions[fila].height = 16
        c = ws.cell(fila, 1, titulo_total)
        c.font = negro_b; c.fill = azul_cla; c.alignment = izq; c.border = borde_g
        sec_acum = 0.0
        for j, m in enumerate(MESES, 2):
            v = totales_sec_mes[m]
            totales_gran_mes[m] += v
            cell = ws.cell(fila, j, v if v != 0 else None)
            fmt_num(cell, bold=True); cell.fill = azul_cla; cell.border = borde_n
            sec_acum += v
        cell = ws.cell(fila, N_COLS, sec_acum if sec_acum != 0 else None)
        fmt_num(cell, bold=True); cell.fill = azul_cla; cell.border = borde_n
        fila += 1

        # Espacio entre secciones
        ws.row_dimensions[fila].height = 6
        fila += 1

    # Fila gran total
    ws.row_dimensions[fila].height = 18
    c = ws.cell(fila, 1, "TOTAL GENERAL")
    c.font = Font(bold=True, size=11, color="FFFFFF", name="Calibri"); c.fill = azul_osc
    c.alignment = izq
    gt_acum = 0.0
    for j, m in enumerate(MESES, 2):
        v = totales_gran_mes[m]
        cell = ws.cell(fila, j, v if v != 0 else None)
        cell.number_format = FMT_COP_XL; cell.alignment = der
        cell.font = Font(bold=True, size=11, color="FFFFFF", name="Calibri")
        cell.fill = azul_osc
        gt_acum += v
    cell = ws.cell(fila, N_COLS, gt_acum if gt_acum != 0 else None)
    cell.number_format = FMT_COP_XL; cell.alignment = der
    cell.font = Font(bold=True, size=11, color="FFFFFF", name="Calibri")
    cell.fill = azul_osc

    ws.freeze_panes = "B3"


def _escribir_df_en_hoja(ws, df, index=True):
    """Escribe un DataFrame en un Worksheet de openpyxl sin pandas ExcelWriter."""
    from openpyxl.utils import get_column_letter
    if index:
        ws.append([""] + list(df.columns))
        for row_idx, (idx_val, row) in enumerate(df.iterrows()):
            ws.append([str(idx_val)] + [
                float(v) if isinstance(v, (int, float)) and not pd.isna(v) else str(v)
                for v in row
            ])
    else:
        ws.append(list(df.columns))
        for _, row in df.iterrows():
            ws.append([
                float(v) if isinstance(v, (int, float)) and not pd.isna(v) else str(v)
                for v in row
            ])


def generar_hoja_anexo(ws, df_raw_src, titulo, es_pivot=True, grupos_orden=None):
    """Formatea una hoja de Anexo con jerarquía Grupo → Cuenta → Tercero por mes.

    es_pivot=True  → expansión jerárquica: Grupo (subtotal) > Cuenta > Tercero
    es_pivot=False → tabla de detalle plano (Codigo, Nombre cuenta, Nit, Nombre tercero, Mes, Saldo Mes…)
    """
    from openpyxl.styles import PatternFill
    from openpyxl.utils import get_column_letter

    azul_osc  = PatternFill("solid", fgColor="1E3A5F")
    azul_med  = PatternFill("solid", fgColor="2E6DA4")
    azul_cla  = PatternFill("solid", fgColor="D6E4F0")
    gris_alt  = PatternFill("solid", fgColor="F2F4F7")

    blanco    = Font(color="FFFFFF", bold=True, size=11, name="Calibri")
    negro_b   = Font(bold=True,  size=10, name="Calibri")
    negro_n   = Font(bold=False, size=10, name="Calibri")
    negro_it  = Font(bold=False, size=9,  name="Calibri", italic=True, color="444444")
    centro    = Alignment(horizontal="center", vertical="center", wrap_text=True)
    izq       = Alignment(horizontal="left",   vertical="center", wrap_text=False)
    der       = Alignment(horizontal="right",  vertical="center")

    thin_s  = Side(border_style="thin",   color="CCCCCC")
    thick_s = Side(border_style="medium", color="2E6DA4")
    borde_n = Border(left=thin_s, right=thin_s, top=thin_s, bottom=thin_s)
    borde_g = Border(left=thick_s, right=thin_s, top=thick_s, bottom=thick_s)  # grupos

    def fmt_num(cell, bold=False):
        cell.number_format = FMT_COP_XL
        cell.alignment = der
        cell.font = negro_b if bold else negro_n

    # ────────────────────────────────────────────────────────────────────────────
    if es_pivot:
        # Pivot jerárquico: Grupo → Cuenta → Tercero × Meses
        df = df_raw_src.copy()
        MESES = [m for m in MESES_ORDEN if m in df["Mes"].unique()]
        n_mes = len(MESES)
        n_cols = 1 + n_mes + 1   # col Descripción + meses + Total general

        # Anchos
        ws.column_dimensions["A"].width = 52
        for j in range(2, n_cols + 1):
            ws.column_dimensions[get_column_letter(j)].width = 20 if j == n_cols else 15

        # Fila 1: título
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=n_cols)
        c = ws.cell(1, 1, titulo); c.font = blanco; c.fill = azul_osc
        c.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 22

        # Fila 2: encabezados
        ws.row_dimensions[2].height = 18
        c = ws.cell(2, 1, "Grupo / Cuenta / Tercero"); c.font = blanco; c.fill = azul_med; c.alignment = centro
        for j, m in enumerate(MESES, 2):
            c = ws.cell(2, j, MESES_ABREV.get(m, m[:3]).upper())
            c.font = blanco; c.fill = azul_med; c.alignment = centro
        c = ws.cell(2, n_cols, "TOTAL"); c.font = blanco; c.fill = azul_med; c.alignment = centro

        # Pivot base: Grupo × Mes
        piv_grp = (df.groupby(["Grupo", "Mes"])["Saldo Mes"].sum()
                   .unstack(fill_value=0).reindex(columns=MESES, fill_value=0))
        # Pivot cuenta: Cuenta × Mes (clave = Codigo + Nombre cuenta)
        df["_cuenta"] = df["Codigo"].astype(str).str.strip() + "  " + df["Nombre cuenta"].str.strip()
        piv_cta = (df.groupby(["_cuenta", "Grupo", "Mes"])["Saldo Mes"].sum()
                   .unstack(fill_value=0).reindex(columns=MESES, fill_value=0).reset_index())
        # Pivot tercero: Tercero × Mes (clave = Nit + Nombre tercero)
        df["_tercero"] = df["Nit"].astype(str).str.strip() + "  " + df["Nombre tercero"].fillna("").str.strip()
        piv_ter = (df.groupby(["_tercero", "_cuenta", "Grupo", "Mes"])["Saldo Mes"].sum()
                   .unstack(fill_value=0).reindex(columns=MESES, fill_value=0).reset_index())

        orden_grp = grupos_orden if grupos_orden else list(piv_grp.index)

        fila = 3
        for grp in orden_grp:
            if grp not in piv_grp.index:
                continue
            # ── Fila grupo (subtotal) ──
            ws.row_dimensions[fila].height = 16
            c = ws.cell(fila, 1, grp); c.font = negro_b; c.fill = azul_cla; c.alignment = izq; c.border = borde_g
            vals_g = [piv_grp.loc[grp, m] for m in MESES]
            tot_g = sum(abs(v) for v in vals_g)
            for j, v in enumerate(vals_g, 2):
                cell = ws.cell(fila, j, abs(v) if v != 0 else None)
                fmt_num(cell, bold=True); cell.fill = azul_cla; cell.border = borde_n
            cell = ws.cell(fila, n_cols, tot_g if tot_g != 0 else None)
            fmt_num(cell, bold=True); cell.fill = azul_cla; cell.border = borde_n
            fila += 1

            # ── Filas de cuentas dentro del grupo ──
            ctas = piv_cta[piv_cta["Grupo"] == grp]
            for _, cr in ctas.iterrows():
                cta_lbl = str(cr["_cuenta"])
                vals_c = [cr[m] if m in cr.index else 0 for m in MESES]
                tot_c  = sum(abs(v) for v in vals_c)

                ws.row_dimensions[fila].height = 15
                c = ws.cell(fila, 1, "  " + cta_lbl); c.font = negro_n; c.alignment = izq; c.border = borde_n
                for j, v in enumerate(vals_c, 2):
                    cell = ws.cell(fila, j, abs(v) if v != 0 else None)
                    fmt_num(cell); cell.border = borde_n
                cell = ws.cell(fila, n_cols, tot_c if tot_c != 0 else None)
                fmt_num(cell); cell.border = borde_n
                fila += 1

                # ── Filas de terceros dentro de la cuenta ──
                terceros = piv_ter[(piv_ter["Grupo"] == grp) & (piv_ter["_cuenta"] == cta_lbl)]
                for _, tr in terceros.iterrows():
                    ter_lbl = str(tr["_tercero"])
                    if ter_lbl.strip() in ("", "nan  ", "  "):
                        continue
                    vals_t = [tr[m] if m in tr.index else 0 for m in MESES]
                    tot_t  = sum(abs(v) for v in vals_t)
                    ws.row_dimensions[fila].height = 14
                    fill_t = gris_alt if fila % 2 == 0 else None
                    c = ws.cell(fila, 1, "      " + ter_lbl); c.font = negro_it; c.alignment = izq; c.border = borde_n
                    if fill_t: c.fill = fill_t
                    for j, v in enumerate(vals_t, 2):
                        cell = ws.cell(fila, j, abs(v) if v != 0 else None)
                        fmt_num(cell); cell.border = borde_n
                        if fill_t: cell.fill = fill_t
                    cell = ws.cell(fila, n_cols, tot_t if tot_t != 0 else None)
                    fmt_num(cell); cell.border = borde_n
                    if fill_t: cell.fill = fill_t
                    fila += 1

        ws.freeze_panes = "B3"

    # ────────────────────────────────────────────────────────────────────────────
    else:
        # Tabla de detalle plano
        df = df_raw_src.copy()
        columnas = list(df.columns)
        n_cols   = len(columnas)

        # Fila 1: título fusionado
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=n_cols)
        c = ws.cell(1, 1, titulo); c.font = blanco; c.fill = azul_osc
        c.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 22

        # Fila 2: encabezados
        ws.row_dimensions[2].height = 18
        for j, col in enumerate(columnas, 1):
            c = ws.cell(2, j, col); c.font = blanco; c.fill = azul_med; c.alignment = centro

        # Anchos por columna
        anchos_det = {
            "Codigo": 12, "Nombre cuenta": 36, "Nit": 15,
            "Nombre tercero": 36, "Mes": 14, "Saldo Mes": 18,
            "Cuenta": 12, "Grupo": 46,
        }
        for j, col in enumerate(columnas, 1):
            ws.column_dimensions[get_column_letter(j)].width = anchos_det.get(col, 16)

        # Filas de datos
        for i, (_, row) in enumerate(df.iterrows(), 3):
            ws.row_dimensions[i].height = 14
            fill = gris_alt if i % 2 == 0 else None

            for j, col in enumerate(columnas, 1):
                val = row[col]
                c = ws.cell(i, j)
                if col == "Saldo Mes":
                    c.value = float(val) if not pd.isna(val) else None
                    c.number_format = FMT_COP_XL; c.alignment = der
                elif col == "Mes":
                    c.value = str(val) if val is not None else ""; c.alignment = centro
                elif col in ("Codigo", "Nit", "Cuenta"):
                    c.value = str(val) if val is not None else ""; c.alignment = centro
                else:
                    c.value = str(val) if val is not None else ""; c.alignment = izq
                c.font = negro_n
                if fill: c.fill = fill
                c.border = borde_n

        ws.freeze_panes = "A3"
        ws.auto_filter.ref = f"A2:{get_column_letter(n_cols)}2"


def generar_excel_eeff(empresa, nit, periodo, saldos_esf, totales_eri,
                       pivot_eri, df_eri_raw, pivot_esf, df_esf_raw,
                       saldos_esf_ant=None, totales_eri_ant=None):
    """saldos_esf_ant / totales_eri_ant: dicts con datos del período anterior
       para poblar las columnas 'Período anterior' en ESF y ERI."""
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    # Hojas formateadas
    generar_hoja_esf(wb.create_sheet("ESF"), empresa, nit, periodo,
                     saldos_esf, saldos_ant=saldos_esf_ant)
    generar_hoja_eri(wb.create_sheet("ERI"), empresa, nit, periodo,
                     totales_eri, totales_ant=totales_eri_ant)

    # Hojas de Anexo (jerarquía Grupo → Cuenta → Tercero) con formato corporativo
    generar_hoja_anexo(wb.create_sheet("Anexo ESF"), df_esf_raw,
                       titulo="ANEXO ESF — SALDOS MENSUALES POR GRUPO", es_pivot=True,
                       grupos_orden=GRUPOS_ESF_ORDEN)
    generar_hoja_anexo(wb.create_sheet("Anexo ERI"), df_eri_raw,
                       titulo="ANEXO ERI — MOVIMIENTOS MENSUALES POR GRUPO", es_pivot=True,
                       grupos_orden=GRUPOS_ERI)

    # Hojas de detalle con formato corporativo
    generar_hoja_anexo(wb.create_sheet("Detalle ESF"), df_esf_raw,
                       titulo="DETALLE ESF — MOVIMIENTOS POR TERCERO", es_pivot=False)
    generar_hoja_anexo(wb.create_sheet("Detalle ERI"), df_eri_raw,
                       titulo="DETALLE ERI — MOVIMIENTOS POR TERCERO", es_pivot=False)

    # Hojas adicionales: CONSOL y ER MENSUALIZADO
    generar_hoja_consol(wb.create_sheet("CONSOL"), empresa, df_eri_raw)
    generar_hoja_er_mensualizado(wb.create_sheet("ER MENSUALIZADO"), empresa, df_eri_raw)

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
(df_eri_raw, pivot_eri, df_esf_raw, pivot_esf, saldos_esf, totales_eri,
 ultimo_mes, _saldos_esf_ant_auto, _totales_eri_ant_auto) = procesar_archivo(file_bytes)

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
        [g for g in ["GASTOS DE ADMINISTRACION","OTROS GASTOS","GASTOS FINANCIEROS","PROVISION DE IMPUESTOS"]
         if g in pivot_f.index], "Total general"].sum()
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
        pc = "blue" if saldo_act >= saldo_pas else "red"
        st.markdown(f"""<div class="metric-card"><div class="number {pc}">{fmt_cop(saldo_act-saldo_pas)}</div>
            <div class="label">Activo − Pasivo ({ultimo_mes})</div></div>""", unsafe_allow_html=True)

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

    # ── Período anterior ─────────────────────────────────────────────────────
    # Opción 1: detectado automáticamente en el mismo archivo (mes "PERIODO ANTERIOR")
    # Opción 2: archivo separado del año anterior
    saldos_esf_ant  = _saldos_esf_ant_auto
    totales_eri_ant = _totales_eri_ant_auto

    if _saldos_esf_ant_auto:
        st.success("✅ Período anterior detectado en el mismo archivo (fila con Mes = 'PERIODO ANTERIOR')")
    else:
        st.markdown("**📂 Archivo del período anterior** *(opcional — o agrega filas con Mes='PERIODO ANTERIOR' en el mismo archivo)*")
        uploaded_ant = st.file_uploader(
            "Sube el archivo Excel del año anterior",
            type=["xlsx"],
            key="uploader_anterior",
            label_visibility="collapsed",
            help="Debe tener la misma estructura (hoja 'terceros_'). Alternativa: agrega filas con Mes='PERIODO ANTERIOR' en el mismo archivo.",
        )
        if uploaded_ant:
            try:
                ant_bytes = uploaded_ant.read()
                res_ant = procesar_archivo(ant_bytes)
                saldos_esf_ant  = res_ant[4]   # saldos_esf del archivo anterior
                totales_eri_ant = res_ant[5]   # totales_eri del archivo anterior
                st.success(f"✅ Archivo año anterior cargado: **{uploaded_ant.name}**")
            except Exception as ex:
                st.warning(f"⚠️ No se pudo procesar el archivo anterior: {ex}")
                saldos_esf_ant  = None
                totales_eri_ant = None

    c1, c2 = st.columns(2)
    with c1:
        st.markdown(f"**Vista previa ESF — saldo de cierre ({ultimo_mes})**")
        prev_esf = pd.DataFrame({
            "Cuenta": [NOMBRE_ESF.get(g,g).strip() for g in GRUPOS_ESF_ORDEN if g in saldos_esf],
            "Saldo":  [fmt_cop(abs(saldos_esf.get(g,0))) for g in GRUPOS_ESF_ORDEN if g in saldos_esf],
        })
        st.dataframe(prev_esf, use_container_width=True, hide_index=True, height=370)
        st.markdown(
            f"<p><b>Total Activo:</b> {fmt_cop(saldo_act)} &nbsp;|&nbsp; <b>Total Pasivo:</b> {fmt_cop(saldo_pas)}</p>",
            unsafe_allow_html=True,
        )
    with c2:
        st.markdown("**Vista previa ERI — acumulado**")
        ing_ord  = abs(totales_eri.get("INGRESOS DE ACTIVIDADES ORDINARIAS",0))
        otros_ing= abs(totales_eri.get("OTROS INGRESOS",0))
        gtos_adm = abs(totales_eri.get("GASTOS DE ADMINISTRACION",0))
        otros_gto= abs(totales_eri.get("OTROS GASTOS",0))
        ing_fin  = abs(totales_eri.get("INGRESOS FINANCIEROS",0))
        gto_fin  = abs(totales_eri.get("GASTOS FINANCIEROS",0))
        provision= abs(totales_eri.get("PROVISION DE IMPUESTOS",0))
        util_ai  = ing_ord + otros_ing - gtos_adm - otros_gto + ing_fin - gto_fin
        util_per = util_ai - provision
        prev_eri = pd.DataFrame({
            "Línea": [NOMBRE_ERI.get(g,g) for g in GRUPOS_ERI] +
                     ["─────────────────","Utilidad antes de impuesto","─────────────────","Utilidad del periodo"],
            "Valor": [fmt_cop(abs(totales_eri.get(g,0))) for g in GRUPOS_ERI] +
                     ["",fmt_cop(util_ai),"",fmt_cop(util_per)],
        })
        st.dataframe(prev_eri, use_container_width=True, hide_index=True, height=370)

    st.markdown("---")
    if st.button("⚙️ Generar archivo EEFF formateado", type="primary", use_container_width=True):
        with st.spinner("Generando archivo Excel…"):
            buf_eeff = generar_excel_eeff(
                empresa, nit, periodo,
                saldos_esf, totales_eri,
                pivot_eri, df_eri_raw, pivot_esf, df_esf_raw,
                saldos_esf_ant=saldos_esf_ant,
                totales_eri_ant=totales_eri_ant,
            )
        periodo_ant_msg = " (con período anterior)" if saldos_esf_ant else ""
        st.success(f"✅ Archivo generado{periodo_ant_msg} — 8 hojas: ESF, ERI, Anexo ESF, Anexo ERI, Detalle ESF, Detalle ERI, CONSOL, ER MENSUALIZADO")
        st.download_button(
            label="📥 Descargar EEFF_Formateado.xlsx",
            data=buf_eeff,
            file_name="EEFF_Formateado.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    st.markdown('</div>', unsafe_allow_html=True)
