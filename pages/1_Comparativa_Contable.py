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
    .stApp { background-color: #F5F7FA; }

    .header-block {
        background: linear-gradient(135deg, #1E3A5F 0%, #2E6DA4 100%);
        border-radius: 12px;
        padding: 28px 36px;
        margin-bottom: 28px;
        color: white;
    }
    .header-block h1 { font-size: 1.9rem; margin: 0; font-weight: 700; }
    .header-block p  { margin: 6px 0 0; opacity: .85; font-size: .95rem; }

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
    .purple { color: #7B3FA0; }

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

    #MainMenu, footer, header { visibility: hidden; }
    .block-container { padding-top: 1.5rem; }

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

def cargar_extracto(archivo) -> tuple:
    """
    Lee el extracto bancario (CSV, TXT o Excel).
    Devuelve:
      - serie_col_i : columna I (índice 8) con los montos
      - desc_col_g  : columna G (índice 6) con descripciones
    Excluye filas de saldo (SALDO INICIAL, SALDO DIA, SALDO FINAL).
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
        st.error(
            f"El archivo no tiene suficientes columnas (se necesitan al menos 9). "
            f"Columnas detectadas: {df.shape[1] if df is not None else 0}"
        )
        return None, None

    # Excluir filas de saldo que no son transacciones reales
    desc_raw = df[6].fillna("").astype(str).str.strip().str.upper()
    mask_saldo = desc_raw.str.startswith("SALDO")
    df = df[~mask_saldo].copy()

    col_i    = pd.to_numeric(df[8], errors="coerce").dropna()
    desc_col = df[6].fillna("").astype(str).str.strip()

    return col_i.reset_index(drop=True), desc_col.reset_index(drop=True)


def cargar_excel(archivo) -> tuple:
    """
    Lee el libro contable Excel.
    Detecta la fila de encabezados buscando 'Crédito' o 'Credito'.
    Devuelve (df, col_debito, col_credito, col_tercero).
    """
    df_raw = pd.read_excel(archivo, sheet_name=0, header=None)

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

    col_debito = next(
        (c for c in df.columns if "d" in c.lower() and "bito" in c.lower()), None
    )
    col_credito = next(
        (c for c in df.columns if "cr" in c.lower() and "dito" in c.lower()), None
    )

    if not col_debito or not col_credito:
        st.error(
            f"No se encontraron columnas Débito/Crédito. "
            f"Columnas disponibles: {list(df.columns)}"
        )
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


def comparar(
    serie_csv: pd.Series,
    desc_csv: pd.Series,
    df_excel: pd.DataFrame,
    col_debito: str,
    col_credito: str,
    col_tercero: str,
    tolerancia: float,
) -> pd.DataFrame:
    """
    Para CADA valor del CSV (en valor absoluto) busca coincidencia
    en la columna Débito Y en la columna Crédito del Excel.

    Casos de estado:
      ✅ En Débito        → encontrado solo en Débito
      ✅ En Crédito       → encontrado solo en Crédito
      ✅ En Débito y Crédito → encontrado en ambas columnas
      ❌ No encontrado    → no está en ninguna

    También reporta el Nombre del Tercero de cada coincidencia.
    """
    # Construir mapas  valor_redondeado → nombre_tercero   para cada columna
    def build_map(col_name):
        m = {}
        for _, row in df_excel.iterrows():
            v = row[col_name]
            if pd.isna(v) or v == 0:
                continue
            vr = round(float(v), 2)
            if vr not in m:
                nombre = (
                    str(row[col_tercero]).strip()
                    if col_tercero and pd.notna(row[col_tercero])
                    else ""
                )
                m[vr] = nombre
        return m

    map_debito  = build_map(col_debito)
    map_credito = build_map(col_credito)

    set_debito  = set(map_debito.keys())
    set_credito = set(map_credito.keys())

    filas = []
    for idx in range(len(serie_csv)):
        val   = serie_csv.iloc[idx]
        desc  = str(desc_csv.iloc[idx]).strip() if idx < len(desc_csv) else ""

        val_r   = round(float(val), 2)
        val_abs = round(abs(val_r), 2)

        # Buscar en Débito
        match_d = next(
            (v for v in set_debito if abs(val_abs - v) <= tolerancia), None
        )
        # Buscar en Crédito
        match_c = next(
            (v for v in set_credito if abs(val_abs - v) <= tolerancia), None
        )

        en_debito  = match_d is not None
        en_credito = match_c is not None

        nombre_d = map_debito.get(match_d, "")  if en_debito  else ""
        nombre_c = map_credito.get(match_c, "") if en_credito else ""

        # Estado consolidado
        if en_debito and en_credito:
            estado = "✅ En Débito y Crédito"
        elif en_debito:
            estado = "✅ En Débito"
        elif en_credito:
            estado = "✅ En Crédito"
        else:
            estado = "❌ No encontrado"

        # Nombre único (si coincide en ambas, mostrar el más informativo)
        nombre_tercero = nombre_d or nombre_c

        filas.append(
            {
                "Valor CSV (Col I)":   val_r,
                "Descripción CSV":     desc,
                f"En {col_debito}":    "✅" if en_debito  else "❌",
                f"En {col_credito}":   "✅" if en_credito else "❌",
                "Nombre del tercero":  nombre_tercero,
                "Estado":              estado,
            }
        )

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
    <p>Verifica si los valores del extracto bancario (CSV) están registrados en <strong>Débito</strong>
    o <strong>Crédito</strong> del libro contable (Excel)</p>
</div>
""", unsafe_allow_html=True)

# ── Carga de archivos ────────────────────────────────────────────────────────
col1, col2 = st.columns(2)

with col1:
    st.markdown('<div class="upload-section">', unsafe_allow_html=True)
    st.markdown('<p class="section-title">📄 Extracto Bancario (CSV o Excel)</p>', unsafe_allow_html=True)
    st.caption("Columna I · valores del extracto · acepta CSV, TXT o Excel")
    archivo_csv = st.file_uploader(
        "Subir extracto", type=["csv", "txt", "xlsx", "xls"],
        key="csv", label_visibility="collapsed"
    )
    st.markdown('</div>', unsafe_allow_html=True)

with col2:
    st.markdown('<div class="upload-section">', unsafe_allow_html=True)
    st.markdown('<p class="section-title">📊 Libro Contable (Excel)</p>', unsafe_allow_html=True)
    st.caption("Columnas: Débito · Crédito · Nombre del tercero")
    archivo_xlsx = st.file_uploader(
        "Subir Excel", type=["xlsx", "xls"],
        key="xlsx", label_visibility="collapsed"
    )
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
        help="Útil para diferencias por redondeo. Por defecto 0.01",
    )

# ── Procesamiento ────────────────────────────────────────────────────────────
if archivo_csv and archivo_xlsx:
    with st.spinner("Procesando archivos..."):
        serie_csv, desc_csv = cargar_extracto(archivo_csv)
        resultado_excel = cargar_excel(archivo_xlsx)

    if serie_csv is None or resultado_excel is None:
        st.stop()

    df_excel, col_debito, col_credito, col_tercero = resultado_excel

    df_resultado = comparar(
        serie_csv, desc_csv,
        df_excel, col_debito, col_credito, col_tercero,
        tolerancia,
    )

    # ── Métricas ─────────────────────────────────────────────────────────────
    total            = len(df_resultado)
    en_debito        = df_resultado["Estado"].str.contains("Débito").sum()
    en_credito       = df_resultado["Estado"].str.contains("Crédito").sum()
    en_ambos         = df_resultado["Estado"].str.contains("Débito y Crédito").sum()
    no_encontrados   = (df_resultado["Estado"] == "❌ No encontrado").sum()
    encontrados_tot  = total - no_encontrados
    pct              = round(encontrados_tot / total * 100, 1) if total > 0 else 0

    st.markdown("---")
    m1, m2, m3, m4, m5 = st.columns(5)

    with m1:
        st.markdown(f"""
        <div class="metric-card">
            <div class="number blue">{total}</div>
            <div class="label">Valores analizados</div>
        </div>""", unsafe_allow_html=True)
    with m2:
        st.markdown(f"""
        <div class="metric-card">
            <div class="number green">{en_debito}</div>
            <div class="label">Encontrados en Débito</div>
        </div>""", unsafe_allow_html=True)
    with m3:
        st.markdown(f"""
        <div class="metric-card">
            <div class="number orange">{en_credito}</div>
            <div class="label">Encontrados en Crédito</div>
        </div>""", unsafe_allow_html=True)
    with m4:
        st.markdown(f"""
        <div class="metric-card">
            <div class="number red">{no_encontrados}</div>
            <div class="label">No encontrados</div>
        </div>""", unsafe_allow_html=True)
    with m5:
        color = "green" if pct == 100 else ("orange" if pct >= 80 else "red")
        st.markdown(f"""
        <div class="metric-card">
            <div class="number {color}">{pct}%</div>
            <div class="label">Coincidencia total</div>
        </div>""", unsafe_allow_html=True)

    # ── Tabla resultado ───────────────────────────────────────────────────────
    st.markdown('<div class="result-block">', unsafe_allow_html=True)
    st.markdown('<p class="section-title">📋 Resultado detallado</p>', unsafe_allow_html=True)

    filtro = st.radio(
        "Mostrar:",
        [
            "Todos",
            "❌ No encontrados",
            "✅ En Débito",
            "✅ En Crédito",
            "✅ En Débito y Crédito",
        ],
        horizontal=True,
    )

    df_mostrar = df_resultado.copy()
    if filtro == "❌ No encontrados":
        df_mostrar = df_mostrar[df_mostrar["Estado"] == "❌ No encontrado"]
    elif filtro == "✅ En Débito":
        df_mostrar = df_mostrar[df_mostrar["Estado"] == "✅ En Débito"]
    elif filtro == "✅ En Crédito":
        df_mostrar = df_mostrar[df_mostrar["Estado"] == "✅ En Crédito"]
    elif filtro == "✅ En Débito y Crédito":
        df_mostrar = df_mostrar[df_mostrar["Estado"] == "✅ En Débito y Crédito"]

    def colorear_fila(row):
        estado = str(row["Estado"])
        if "No encontrado" in estado:
            return ["background-color: #FFF5F5"] * len(row)
        elif "Débito y Crédito" in estado:
            return ["background-color: #EEF4FF"] * len(row)
        elif "Débito" in estado:
            return ["background-color: #F0FFF4"] * len(row)
        elif "Crédito" in estado:
            return ["background-color: #FFFBF0"] * len(row)
        return [""] * len(row)

    st.dataframe(
        df_mostrar.style.apply(colorear_fila, axis=1),
        use_container_width=True,
        height=440,
    )
    st.markdown('</div>', unsafe_allow_html=True)

    # ── Leyenda de colores ────────────────────────────────────────────────────
    st.caption(
        "🟢 Verde = encontrado en Débito · 🟡 Amarillo = encontrado en Crédito · "
        "🔵 Azul = encontrado en ambas · 🔴 Rojo = no encontrado"
    )

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
        t1, t2 = st.tabs(["CSV — Valores analizados", "Excel — Débito & Crédito"])
        with t1:
            prev_csv = pd.DataFrame(
                {"Descripción (Col G)": desc_csv, "Valor (Col I)": serie_csv}
            )
            st.dataframe(prev_csv, use_container_width=True, height=300)
        with t2:
            cols_preview = [c for c in [col_tercero, col_debito, col_credito] if c]
            st.dataframe(df_excel[cols_preview], use_container_width=True, height=300)

else:
    st.markdown("""
    <div style="text-align:center; padding: 60px 20px; color: #999;">
        <div style="font-size: 3rem; margin-bottom: 16px;">📂</div>
        <p style="font-size: 1.1rem; font-weight: 600;">Carga los dos archivos para iniciar la comparativa</p>
        <p style="font-size: .9rem;">CSV con la columna I · Excel con Débito, Crédito y Nombre del tercero</p>
    </div>
    """, unsafe_allow_html=True)
