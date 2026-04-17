import streamlit as st
import pandas as pd
import io
from datetime import datetime

if not st.session_state.get("authentication_status"):
    st.warning("⚠️ Debe iniciar sesión primero.")
    st.page_link("Inicio.py", label="Ir al login", icon="🔐")
    st.stop()


# ── Configuración de página ──────────────────────────────────────────────────
st.set_page_config(
    page_title="Comparativa Contable",
    page_icon="🔍",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Estilos ──────────────────────────────────────────────────────────────────
st.markdown("""
<style>
    /* Fondo general */
    .stApp { background-color: #F5F7FA; }

    /* Encabezado */
    .header-block {
        background: linear-gradient(135deg, #1E3A5F 0%, #2E6DA4 100%);
        border-radius: 12px;
        padding: 28px 36px;
        margin-bottom: 28px;
        color: white;
    }
    .header-block h1 { font-size: 1.9rem; margin: 0; font-weight: 700; }
    .header-block p  { margin: 6px 0 0; opacity: .85; font-size: .95rem; }

    /* Tarjetas de métricas */
    .metric-card {
        background: white;
        border-radius: 10px;
        padding: 20px 24px;
        box-shadow: 0 2px 8px rgba(0,0,0,.08);
        text-align: center;
        height: 100%;
    }
    .metric-card .number {
        font-size: 2.2rem;
        font-weight: 700;
        line-height: 1.1;
    }
    .metric-card .label {
        font-size: .82rem;
        color: #666;
        margin-top: 4px;
        text-transform: uppercase;
        letter-spacing: .05em;
    }
    .green  { color: #1A9E5C; }
    .red    { color: #D63B3B; }
    .blue   { color: #2E6DA4; }
    .orange { color: #E07B20; }

    /* Sección de upload */
    .upload-section {
        background: white;
        border-radius: 10px;
        padding: 24px;
        box-shadow: 0 2px 8px rgba(0,0,0,.08);
        margin-bottom: 20px;
    }
    .section-title {
        font-size: 1rem;
        font-weight: 600;
        color: #1E3A5F;
        margin-bottom: 12px;
        padding-bottom: 8px;
        border-bottom: 2px solid #E8EDF3;
    }

    /* Badges en tabla */
    .badge-ok   { background:#D4EDDA; color:#155724; padding:3px 10px; border-radius:12px; font-size:.8rem; font-weight:600; }
    .badge-fail { background:#F8D7DA; color:#721C24; padding:3px 10px; border-radius:12px; font-size:.8rem; font-weight:600; }

    /* Ocultar elementos Streamlit genéricos */
    #MainMenu, footer, header { visibility: hidden; }
    .block-container { padding-top: 1.5rem; }

    /* Tabla resultado */
    .result-block {
        background: white;
        border-radius: 10px;
        padding: 24px;
        box-shadow: 0 2px 8px rgba(0,0,0,.08);
        margin-top: 20px;
    }
</style>
""", unsafe_allow_html=True)


# ── Funciones ────────────────────────────────────────────────────────────────

def redondear(valor, decimales=2):
    """Redondea un valor numérico."""
    try:
        return round(float(valor), decimales)
    except Exception:
        return None


def cargar_extracto(archivo) -> tuple:
    """
    Lee el extracto bancario (CSV, TXT o Excel) y devuelve:
      - serie_col_i: columna I (índice 8) con valores numéricos
      - desc_col_g:  columna G (índice 6) con descripciones
    """
    nombre = archivo.name.lower()
    es_excel = nombre.endswith(".xlsx") or nombre.endswith(".xls")

    if es_excel:
        contenido = archivo.read()
        df = pd.read_excel(io.BytesIO(contenido), sheet_name=0, header=None)
    else:
        contenido = archivo.read()
        df = None
        for enc in ["latin1", "utf-8", "cp1252"]:
            try:
                df = pd.read_csv(
                    io.BytesIO(contenido),
                    encoding=enc,
                    sep=None,
                    engine="python",
                    header=None,
                )
                break
            except Exception:
                continue

    if df is None or df.shape[1] < 9:
        st.error(f"El archivo no tiene suficientes columnas (se necesitan al menos 9). Columnas detectadas: {df.shape[1] if df is not None else 0}")
        return None, None

    col_i = pd.to_numeric(df[8], errors="coerce").dropna()
    desc_col_g = df[6].fillna("").astype(str).str.strip()

    return col_i.reset_index(drop=True), desc_col_g


def cargar_excel(archivo) -> pd.DataFrame:
    """Lee el Excel, detecta la fila de encabezados y devuelve Débito y Crédito."""
    df_raw = pd.read_excel(archivo, sheet_name=0, header=None)

    # Detectar fila de encabezado buscando 'Crédito'
    header_row = None
    for i, row in df_raw.iterrows():
        vals = [str(v).strip().lower() for v in row.values]
        if any("cr" in v and "dito" in v for v in vals):
            header_row = i
            break

    if header_row is None:
        st.error("No se encontró la fila de encabezados en el Excel.")
        return None

    df = df_raw.iloc[header_row + 1:].copy()
    df.columns = [str(c).strip() for c in df_raw.iloc[header_row].values]

    # Buscar columnas flexiblemente
    col_debito = next(
        (c for c in df.columns if "d" in c.lower() and "bito" in c.lower()), None
    )
    col_credito = next(
        (c for c in df.columns if "cr" in c.lower() and "dito" in c.lower()), None
    )

    if not col_debito or not col_credito:
        st.error(f"No se encontraron columnas Débito/Crédito. Columnas disponibles: {list(df.columns)}")
        return None

    col_tercero = next(
        (c for c in df.columns if "tercero" in c.lower()), None
    )

    cols = [col_debito, col_credito]
    if col_tercero:
        cols.append(col_tercero)

    resultado = df[cols].copy()
    resultado[col_debito]  = pd.to_numeric(resultado[col_debito],  errors="coerce")
    resultado[col_credito] = pd.to_numeric(resultado[col_credito], errors="coerce")
    resultado = resultado.dropna(subset=[col_debito, col_credito], how="all")
    resultado.reset_index(drop=True, inplace=True)
    return resultado, col_debito, col_credito, col_tercero


def comparar(serie_csv: pd.Series, desc_csv: pd.Series,
             df_excel: pd.DataFrame, col_debito: str, col_credito: str,
             col_tercero: str, tolerancia: float):
    """
    Regla de negocio:
      - Valores NEGATIVOS del CSV  → buscar (valor absoluto) en Crédito del Excel
      - Valores POSITIVOS del CSV  → buscar en Débito del Excel
    Enriquecimiento:
      - Encontrados    → muestra Nombre del tercero del Excel
      - No encontrados → muestra Descripción (col G) del CSV
    """
    # Construir mapas valor → nombre tercero
    map_debito  = {}
    map_credito = {}
    if col_tercero:
        for _, row in df_excel.iterrows():
            nombre = str(row[col_tercero]).strip() if pd.notna(row[col_tercero]) else ""
            if pd.notna(row[col_debito]):
                v = round(float(row[col_debito]), 2)
                if v not in map_debito:
                    map_debito[v] = nombre
            if pd.notna(row[col_credito]):
                v = round(float(row[col_credito]), 2)
                if v not in map_credito:
                    map_credito[v] = nombre

    set_debito  = set(map_debito.keys())
    set_credito = set(map_credito.keys())

    filas = []
    for idx, val in serie_csv.items():
        val_r       = round(float(val), 2)
        es_negativo = val_r < 0
        val_abs     = round(abs(val_r), 2)
        desc        = str(desc_csv.get(idx, "")).strip()

        if es_negativo:
            match_v = next((v for v in set_credito if abs(val_abs - v) <= tolerancia), None)
            encontrado_col  = match_v is not None
            columna_destino = col_credito
            nombre_tercero  = map_credito.get(match_v, "") if match_v else ""
        else:
            match_v = next((v for v in set_debito if abs(val_r - v) <= tolerancia), None)
            encontrado_col  = match_v is not None
            columna_destino = col_debito
            nombre_tercero  = map_debito.get(match_v, "") if match_v else ""

        if encontrado_col:
            estado = "✅ CORRECTO"
        else:
            estado = "❌ NO ENCONTRADO"

        fila = {
            "Valor CSV (Col I)":  val_r,
            "Descripción CSV":    desc,
            "Tipo":               "Negativo → Crédito" if es_negativo else "Positivo → Débito",
            f"En {columna_destino}": "✅" if encontrado_col else "❌",

            "Nombre del tercero": nombre_tercero if encontrado_col else "",
            "Estado":             estado,
        }
        filas.append(fila)

    return pd.DataFrame(filas)


def exportar_excel(df_resultado: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df_resultado.to_excel(writer, index=False, sheet_name="Comparativa")
    return buf.getvalue()


# ── UI ───────────────────────────────────────────────────────────────────────

st.markdown("""
<div class="header-block">
    <h1>🔍 Comparativa Contable</h1>
    <p>Verifica si los valores de la columna I del extracto bancario (CSV) están registrados en el libro contable (Excel)</p>
</div>
""", unsafe_allow_html=True)

# ── Carga de archivos ────────────────────────────────────────────────────────
col1, col2 = st.columns(2)

with col1:
    st.markdown('<div class="upload-section">', unsafe_allow_html=True)
    st.markdown('<p class="section-title">📄 Extracto Bancario (CSV o Excel)</p>', unsafe_allow_html=True)
    st.caption("Columna I · valores numéricos del extracto · acepta CSV, TXT o Excel")
    archivo_csv = st.file_uploader("Subir extracto", type=["csv", "txt", "xlsx", "xls"], key="csv", label_visibility="collapsed")
    st.markdown('</div>', unsafe_allow_html=True)

with col2:
    st.markdown('<div class="upload-section">', unsafe_allow_html=True)
    st.markdown('<p class="section-title">📊 Libro Contable (Excel)</p>', unsafe_allow_html=True)
    st.caption("Columnas: Crédito · Saldo Movimiento")
    archivo_xlsx = st.file_uploader("Subir Excel", type=["xlsx", "xls"], key="xlsx", label_visibility="collapsed")
    st.markdown('</div>', unsafe_allow_html=True)

# ── Opciones avanzadas ───────────────────────────────────────────────────────
with st.expander("⚙️ Opciones avanzadas"):
    tolerancia = st.number_input(
        "Tolerancia de comparación (diferencia máxima permitida)",
        min_value=0.0,
        max_value=100.0,
        value=0.01,
        step=0.01,
        format="%.2f",
        help="Útil para diferencias por redondeo. Por defecto 0.01"
    )

# ── Procesamiento ────────────────────────────────────────────────────────────
if archivo_csv and archivo_xlsx:
    with st.spinner("Procesando archivos..."):
        serie_csv, desc_csv = cargar_extracto(archivo_csv)
        resultado_excel = cargar_excel(archivo_xlsx)

    if serie_csv is None or resultado_excel is None:
        st.stop()

    df_excel, col_debito, col_credito, col_tercero = resultado_excel

    df_resultado = comparar(serie_csv, desc_csv, df_excel, col_debito, col_credito, col_tercero, tolerancia)

    # ── Métricas ─────────────────────────────────────────────────────────────
    total        = len(df_resultado)
    encontrados  = df_resultado["Estado"].str.contains("CORRECTO").sum()
    no_encontrados = total - encontrados
    pct          = round(encontrados / total * 100, 1) if total > 0 else 0

    st.markdown("---")
    m1, m2, m3, m4 = st.columns(4)

    with m1:
        st.markdown(f"""
        <div class="metric-card">
            <div class="number blue">{total}</div>
            <div class="label">Valores CSV analizados</div>
        </div>""", unsafe_allow_html=True)
    with m2:
        st.markdown(f"""
        <div class="metric-card">
            <div class="number green">{encontrados}</div>
            <div class="label">Encontrados en Excel</div>
        </div>""", unsafe_allow_html=True)
    with m3:
        st.markdown(f"""
        <div class="metric-card">
            <div class="number red">{no_encontrados}</div>
            <div class="label">No encontrados</div>
        </div>""", unsafe_allow_html=True)
    with m4:
        color = "green" if pct == 100 else ("orange" if pct >= 80 else "red")
        st.markdown(f"""
        <div class="metric-card">
            <div class="number {color}">{pct}%</div>
            <div class="label">Coincidencia</div>
        </div>""", unsafe_allow_html=True)

    # ── Tabla resultado ───────────────────────────────────────────────────────
    st.markdown('<div class="result-block">', unsafe_allow_html=True)
    st.markdown('<p class="section-title">📋 Resultado detallado</p>', unsafe_allow_html=True)

    # Filtro
    filtro = st.radio(
        "Mostrar:",
        ["Todos", "Solo NO ENCONTRADOS", "Solo ✅ CORRECTOS"],
        horizontal=True,
    )

    df_mostrar = df_resultado.copy()
    if filtro == "Solo NO ENCONTRADOS":
        df_mostrar = df_mostrar[df_mostrar["Estado"] == "❌ NO ENCONTRADO"]
    elif filtro == "Solo ✅ CORRECTOS":
        df_mostrar = df_mostrar[df_mostrar["Estado"] == "✅ CORRECTO"]

    # Colorear por estado
    def colorear_fila(row):
        estado = str(row["Estado"])
        if "CORRECTO" in estado:
            return ["background-color: #F0FFF4"] * len(row)
        else:
            return ["background-color: #FFF5F5"] * len(row)

    st.dataframe(
        df_mostrar.style.apply(colorear_fila, axis=1),
        use_container_width=True,
        height=420,
    )
    st.markdown('</div>', unsafe_allow_html=True)

    # ── Exportar ──────────────────────────────────────────────────────────────
    st.markdown("---")
    col_exp1, col_exp2, _ = st.columns([1, 1, 3])

    with col_exp1:
        xlsx_bytes = exportar_excel(df_resultado)
        st.download_button(
            label="⬇️ Descargar Excel",
            data=xlsx_bytes,
            file_name=f"comparativa_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    with col_exp2:
        csv_bytes = df_resultado.to_csv(index=False).encode("utf-8")
        st.download_button(
            label="⬇️ Descargar CSV",
            data=csv_bytes,
            file_name=f"comparativa_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
            mime="text/csv",
            use_container_width=True,
        )

    # ── Vista previa de archivos fuente ──────────────────────────────────────
    with st.expander("🔎 Vista previa de archivos cargados"):
        t1, t2 = st.tabs(["CSV — Columna I", "Excel — Débito & Crédito"])
        with t1:
            st.dataframe(serie_csv.rename("Valor (Col I)"), use_container_width=True, height=300)
        with t2:
            cols_preview = [c for c in [col_tercero, col_debito, col_credito] if c]
            st.dataframe(df_excel[cols_preview], use_container_width=True, height=300)

else:
    # Estado vacío
    st.markdown("""
    <div style="text-align:center; padding: 60px 20px; color: #999;">
        <div style="font-size: 3rem; margin-bottom: 16px;">📂</div>
        <p style="font-size: 1.1rem; font-weight: 600;">Carga los dos archivos para iniciar la comparativa</p>
        <p style="font-size: .9rem;">CSV con la columna I · Excel con Crédito y Saldo Movimiento</p>
    </div>
    """, unsafe_allow_html=True)
