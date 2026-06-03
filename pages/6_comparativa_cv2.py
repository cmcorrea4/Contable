import streamlit as st
import pandas as pd
import io
import base64
import json
from datetime import datetime
from openai import OpenAI
import fitz  # PyMuPDF

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

    .pdf-badge {
        display: inline-block;
        background: #EEF2FF;
        color: #3730A3;
        border: 1px solid #C7D2FE;
        border-radius: 6px;
        padding: 3px 10px;
        font-size: .78rem;
        font-weight: 600;
        margin-bottom: 10px;
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


# ── Configuración del modelo ─────────────────────────────────────────────────
openai_api_key = st.secrets["settings"]["key"]

OPENAI_MODEL = "gpt-4o"

PROMPT_OCR_PDF = """Eres un extractor contable experto. Se te entrega una o varias imágenes de páginas de un extracto bancario de Bancolombia.

Tu tarea es extraer TODAS las transacciones reales (movimientos de dinero) y devolverlas como JSON.

REGLAS ESTRICTAS:
1. Excluye filas de saldo: SALDO INICIAL, SALDO DIA, SALDO FINAL, SALDO ANTERIOR, SALDO ACTUAL.
2. Incluye TODOS los débitos (valores negativos) y créditos (valores positivos).
3. El campo "valor" debe ser numérico con signo (negativo si es cargo, positivo si es abono).
4. El campo "descripcion" es el texto de la columna DESCRIPCIÓN del extracto.
5. No incluyas separadores de miles en los números, solo el valor numérico puro.

Devuelve ÚNICAMENTE un JSON válido con este formato, sin texto adicional, sin bloques markdown:
[
  {"descripcion": "PAGO A PROV wework colombia sa", "valor": -24257870.00},
  {"descripcion": "ABONO INTERESES AHORROS", "valor": 8161.21}
]"""


# ── Funciones PDF OCR con OpenAI ─────────────────────────────────────────────

def pdf_a_imagenes_b64(archivo) -> list[str]:
    """
    Convierte cada página del PDF a PNG en base64 usando PyMuPDF.
    Devuelve lista de strings base64, una por página.
    """
    contenido = archivo.read()
    archivo.seek(0)
    doc = fitz.open(stream=contenido, filetype="pdf")
    imagenes = []
    for pagina in doc:
        # 2x de resolución para que el texto sea legible por GPT-4o
        mat = fitz.Matrix(2.0, 2.0)
        pix = pagina.get_pixmap(matrix=mat)
        png_bytes = pix.tobytes("png")
        imagenes.append(base64.b64encode(png_bytes).decode("utf-8"))
    doc.close()
    return imagenes


def extraer_transacciones_pdf(archivo) -> tuple:
    """
    Convierte el PDF a imágenes y las envía a GPT-4o para OCR y extracción.
    Devuelve (serie_valores, serie_descripciones) igual que cargar_extracto().
    """
    client = OpenAI(api_key=openai_api_key)

    imagenes_b64 = pdf_a_imagenes_b64(archivo)
    n_paginas = len(imagenes_b64)

    # Construir content: primero el prompt, luego una imagen por página
    content = [{"type": "text", "text": PROMPT_OCR_PDF}]
    for i, img_b64 in enumerate(imagenes_b64):
        content.append({
            "type": "image_url",
            "image_url": {
                "url": f"data:image/png;base64,{img_b64}",
                "detail": "high",
            },
        })

    try:
        respuesta = client.chat.completions.create(
            model=OPENAI_MODEL,
            messages=[{"role": "user", "content": content}],
            max_tokens=8192,
            temperature=0,
        )
    except Exception as e:
        st.error(f"❌ Error al llamar la API de OpenAI: {e}")
        return None, None

    texto = respuesta.choices[0].message.content or ""

    # Limpiar posibles bloques markdown
    texto_limpio = texto.strip()
    if texto_limpio.startswith("```"):
        lineas = texto_limpio.split("\n")
        texto_limpio = "\n".join(
            l for l in lineas if not l.startswith("```")
        ).strip()

    try:
        transacciones = json.loads(texto_limpio)
    except json.JSONDecodeError as e:
        st.error(f"❌ No se pudo parsear el JSON devuelto por GPT-4o: {e}")
        with st.expander("Ver respuesta raw"):
            st.code(texto[:3000])
        return None, None

    if not transacciones:
        st.warning("⚠️ GPT-4o no encontró transacciones en el PDF.")
        return None, None

    valores       = pd.Series([float(t["valor"])      for t in transacciones])
    descripciones = pd.Series([str(t["descripcion"])  for t in transacciones])

    return valores, descripciones


# ── Funciones existentes (sin cambios) ──────────────────────────────────────

def cargar_extracto(archivo) -> tuple:
    """
    Lee el extracto bancario (CSV, TXT o Excel).
    Columna I (índice 8) = montos · Columna G (índice 6) = descripción.
    Excluye filas de saldo.
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

    desc_raw   = df[6].fillna("").astype(str).str.strip().str.upper()
    mask_saldo = desc_raw.str.startswith("SALDO")
    df         = df[~mask_saldo].copy()

    col_i    = pd.to_numeric(df[8], errors="coerce").dropna()
    desc_col = df[6].fillna("").astype(str).str.strip()

    return col_i.reset_index(drop=True), desc_col.reset_index(drop=True)


def cargar_excel(archivo) -> tuple:
    """
    Lee el libro contable Excel.
    Detecta la fila de encabezados buscando 'Crédito' o 'Credito'.
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

    resultado          = df[cols].copy()
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
    Para CADA valor del extracto (en valor absoluto) busca coincidencia
    en Débito Y Crédito del libro contable.
    """
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
        val  = serie_csv.iloc[idx]
        desc = str(desc_csv.iloc[idx]).strip() if idx < len(desc_csv) else ""

        val_r   = round(float(val), 2)
        val_abs = round(abs(val_r), 2)

        match_d = next((v for v in set_debito  if abs(val_abs - v) <= tolerancia), None)
        match_c = next((v for v in set_credito if abs(val_abs - v) <= tolerancia), None)

        en_debito  = match_d is not None
        en_credito = match_c is not None

        nombre_d = map_debito.get(match_d,  "") if en_debito  else ""
        nombre_c = map_credito.get(match_c, "") if en_credito else ""

        if en_debito and en_credito:
            estado = "✅ En Débito y Crédito"
        elif en_debito:
            estado = "✅ En Débito"
        elif en_credito:
            estado = "✅ En Crédito"
        else:
            estado = "❌ No encontrado"

        filas.append({
            "Valor Extracto":       val_r,
            "Descripción Extracto": desc,
            f"En {col_debito}":     "✅" if en_debito  else "❌",
            f"En {col_credito}":    "✅" if en_credito else "❌",
            "Nombre del tercero":   nombre_d or nombre_c,
            "Estado":               estado,
        })

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
    <p>Verifica si los valores del extracto bancario están registrados en
    <strong>Débito</strong> o <strong>Crédito</strong> del libro contable.
    Acepta CSV, Excel o <strong>PDF de Bancolombia</strong> (OCR con GPT-4o).</p>
</div>
""", unsafe_allow_html=True)

# ── Selector de tipo de extracto ─────────────────────────────────────────────
tipo_extracto = st.radio(
    "Tipo de archivo del extracto bancario:",
    ["📄 CSV / Excel  (columnas fijas)", "🤖 PDF Bancolombia (OCR con GPT-4o)"],
    horizontal=True,
    key="tipo_extracto",
)
es_pdf = tipo_extracto.startswith("🤖")

st.markdown("---")

# ── Carga de archivos ────────────────────────────────────────────────────────
col1, col2 = st.columns(2)

with col1:
    st.markdown('<div class="upload-section">', unsafe_allow_html=True)
    if es_pdf:
        st.markdown('<p class="section-title">📑 Extracto Bancario (PDF)</p>', unsafe_allow_html=True)
        st.markdown('<span class="pdf-badge">🤖 OCR con GPT-4o</span>', unsafe_allow_html=True)
        st.caption("Estado de cuenta o Informe consolidado Bancolombia en PDF")
        archivo_extracto = st.file_uploader(
            "Subir PDF", type=["pdf"],
            key="pdf_extracto", label_visibility="collapsed"
        )
    else:
        st.markdown('<p class="section-title">📄 Extracto Bancario (CSV o Excel)</p>', unsafe_allow_html=True)
        st.caption("Columna I (índice 8) = valores · Columna G (índice 6) = descripción")
        archivo_extracto = st.file_uploader(
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
        min_value=0.0, max_value=100.0, value=0.01, step=0.01, format="%.2f",
        help="Útil para diferencias por redondeo. Por defecto 0.01",
    )

# ── Procesamiento ────────────────────────────────────────────────────────────
if archivo_extracto and archivo_xlsx:
    with st.spinner("Procesando archivos..."):

        if es_pdf:
            with st.status("🤖 Extrayendo transacciones del PDF con GPT-4o...", expanded=True) as status:
                st.write("Convirtiendo páginas del PDF a imágenes…")
                serie_csv, desc_csv = extraer_transacciones_pdf(archivo_extracto)
                if serie_csv is not None:
                    st.write(f"✅ {len(serie_csv)} transacciones extraídas correctamente.")
                    status.update(label="✅ PDF procesado", state="complete")
                else:
                    status.update(label="❌ Error procesando PDF", state="error")
        else:
            serie_csv, desc_csv = cargar_extracto(archivo_extracto)

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
    total           = len(df_resultado)
    en_debito       = df_resultado["Estado"].str.contains("Débito").sum()
    en_credito      = df_resultado["Estado"].str.contains("Crédito").sum()
    no_encontrados  = (df_resultado["Estado"] == "❌ No encontrado").sum()
    encontrados_tot = total - no_encontrados
    pct             = round(encontrados_tot / total * 100, 1) if total > 0 else 0

    if es_pdf:
        st.info(
            f"📑 PDF procesado: **{len(serie_csv)} transacciones** extraídas con GPT-4o",
            icon="🤖",
        )

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
        ["Todos", "❌ No encontrados", "✅ En Débito", "✅ En Crédito", "✅ En Débito y Crédito"],
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

    # ── Vista previa ──────────────────────────────────────────────────────────
    with st.expander("🔎 Vista previa de archivos cargados"):
        t1, t2 = st.tabs(["Extracto — Valores analizados", "Excel — Débito & Crédito"])
        with t1:
            label1 = "Descripción (GPT-4o OCR)" if es_pdf else "Descripción (Col G)"
            label2 = "Valor (GPT-4o OCR)"        if es_pdf else "Valor (Col I)"
            st.dataframe(
                pd.DataFrame({label1: desc_csv, label2: serie_csv}),
                use_container_width=True, height=300,
            )
        with t2:
            cols_preview = [c for c in [col_tercero, col_debito, col_credito] if c]
            st.dataframe(df_excel[cols_preview], use_container_width=True, height=300)

else:
    st.markdown("""
    <div style="text-align:center; padding: 60px 20px; color: #999;">
        <div style="font-size: 3rem; margin-bottom: 16px;">📂</div>
        <p style="font-size: 1.1rem; font-weight: 600;">Carga los dos archivos para iniciar la comparativa</p>
        <p style="font-size: .9rem;">CSV / Excel o PDF del extracto bancario · Excel con Débito, Crédito y Nombre del tercero</p>
    </div>
    """, unsafe_allow_html=True)
