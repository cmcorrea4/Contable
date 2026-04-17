import streamlit as st
import pandas as pd
import io
from datetime import datetime

if not st.session_state.get("authentication_status"):
    st.warning("⚠️ Debe iniciar sesión primero.")
    st.page_link("Inicio.py", label="Ir al login", icon="🔐")
    st.stop()

st.set_page_config(page_title="Comparativa IVA Facturas", page_icon="🧾", layout="wide")

st.markdown("""
<style>
    .stApp { background-color: #F5F7FA; }
    .header-block { background: linear-gradient(135deg, #1E3A5F 0%, #2E6DA4 100%); border-radius: 12px; padding: 28px 36px; margin-bottom: 28px; color: white; }
    .header-block h1 { font-size: 1.9rem; margin: 0; font-weight: 700; }
    .metric-card { background: white; border-radius: 10px; padding: 20px 24px; box-shadow: 0 2px 8px rgba(0,0,0,.08); text-align: center; }
    .metric-card .number { font-size: 2.2rem; font-weight: 700; }
    .metric-card .label { font-size: .82rem; color: #666; text-transform: uppercase; }
    .green { color: #1A9E5C; } .red { color: #D63B3B; } .blue { color: #2E6DA4; } .orange { color: #E07B20; } .purple { color: #7B2EA4; }
    .upload-section { background: white; border-radius: 10px; padding: 24px; box-shadow: 0 2px 8px rgba(0,0,0,.08); margin-bottom: 20px; }
    .section-title { font-size: 1rem; font-weight: 600; color: #1E3A5F; margin-bottom: 12px; padding-bottom: 8px; border-bottom: 2px solid #E8EDF3; }
    .result-block { background: white; border-radius: 10px; padding: 24px; box-shadow: 0 2px 8px rgba(0,0,0,.08); margin-top: 20px; }
    .tab-header { font-size: 1.05rem; font-weight: 700; color: #1E3A5F; margin-bottom: 4px; }
    #MainMenu, footer, header { visibility: hidden; }
</style>
""", unsafe_allow_html=True)


# ── Loaders ─────────────────────────────────────────────────────────────────

def cargar_facturas(archivo):
    """Carga facturas electrónicas. Devuelve dos DataFrames: facturas normales y notas de crédito."""
    df = pd.read_excel(archivo, header=0, engine="openpyxl")
    df.columns = [str(c).strip() for c in df.columns]

    col_tipo   = next((c for c in df.columns if "tipo" in c.lower() and "documento" in c.lower()), df.columns[0])
    col_nit    = next((c for c in df.columns if "nit" in c.lower() and "emisor" in c.lower()), None)
    col_nombre = next((c for c in df.columns if "nombre" in c.lower() and "emisor" in c.lower()), None)
    col_iva    = next((c for c in df.columns if c.strip().upper() == "IVA"), None)

    if not all([col_nit, col_nombre, col_iva]):
        st.error(f"Columnas no encontradas en Facturas. Disponibles: {list(df.columns)}")
        return None, None, None, None, None, None

    df[col_nit]    = df[col_nit].astype(str).str.strip().str.replace(r"\.0$", "", regex=True)
    df[col_nombre] = df[col_nombre].astype(str).str.strip()
    df[col_iva]    = pd.to_numeric(df[col_iva], errors="coerce")

    mask_nc = df[col_tipo].astype(str).str.strip().str.lower().str.contains("nota de cr")
    df_fact = df[~mask_nc][[col_nit, col_nombre, col_iva]].copy()
    df_nc   = df[mask_nc][[col_nit, col_nombre, col_iva]].copy()

    return df_fact, df_nc, col_nit, col_nombre, col_iva


def cargar_libro_iva(archivo):
    """Carga el libro auxiliar IVA. Devuelve dos DataFrames: valores positivos y negativos."""
    df_raw = pd.read_excel(archivo, header=None, engine="openpyxl")
    header_row = next((i for i, row in df_raw.iterrows()
                       if any("identificaci" in str(v).lower() for v in row.values)), None)
    if header_row is None:
        st.error("No se encontró fila de encabezados en el Libro IVA.")
        return None, None, None, None

    df = df_raw.iloc[header_row + 1:].copy()
    df.columns = [str(c).strip() for c in df_raw.iloc[header_row].values]

    col_id     = next((c for c in df.columns if "identificaci" in c.lower()), None)
    col_nombre = next((c for c in df.columns if "nombre" in c.lower() and "tercero" in c.lower()), None)
    col_vic    = next((c for c in df.columns if "valor" in c.lower() and "impuesto" in c.lower() and "compra" in c.lower()), None)

    if not all([col_id, col_nombre, col_vic]):
        st.error(f"Columnas no encontradas en Libro IVA. Disponibles: {list(df.columns)}")
        return None, None, None, None

    df[col_id]     = pd.to_numeric(df[col_id], errors="coerce")
    df             = df[df[col_id].notna()].copy()
    df[col_id]     = df[col_id].astype(int).astype(str)
    df[col_nombre] = df[col_nombre].astype(str).str.strip()
    df[col_vic]    = pd.to_numeric(df[col_vic], errors="coerce")
    df             = df.dropna(subset=[col_vic]).reset_index(drop=True)

    df_pos = df[df[col_vic] > 0][[col_id, col_nombre, col_vic]].copy()
    df_neg = df[df[col_vic] < 0][[col_id, col_nombre, col_vic]].copy()
    # Para comparar, trabajamos con el valor absoluto de los negativos
    df_neg = df_neg.copy()
    df_neg[col_vic] = df_neg[col_vic].abs()

    return df_pos, df_neg, col_id, col_nombre, col_vic


# ── Comparador genérico ──────────────────────────────────────────────────────

def comparar(df_fact, col_nit, col_nombre_f, col_iva,
             df_libro, col_id, col_nombre_l, col_vic, tolerancia):
    idx_nit_iva    = {}
    idx_nit_nombre = {}
    for _, row in df_libro.iterrows():
        nit = str(row[col_id]).strip()
        vic = round(float(row[col_vic]), 2) if pd.notna(row[col_vic]) else None
        if vic is not None:
            idx_nit_iva.setdefault(nit, []).append(vic)
        if nit not in idx_nit_nombre:
            idx_nit_nombre[nit] = str(row[col_nombre_l]).strip()

    filas = []
    for _, row in df_fact.iterrows():
        nit      = str(row[col_nit]).strip()
        nombre_f = str(row[col_nombre_f]).strip()
        iva      = round(float(row[col_iva]), 2) if pd.notna(row[col_iva]) else 0.0
        nit_ok   = nit in idx_nit_iva
        iva_ok   = nit_ok and any(abs(iva - v) <= tolerancia for v in idx_nit_iva[nit])
        nombre_l = idx_nit_nombre.get(nit, "") if nit_ok else ""

        if nit_ok and iva_ok:
            estado = "✅ CORRECTO"
        elif nit_ok:
            estado = "⚠️ NIT OK / IVA NO ENCONTRADO"
        else:
            estado = "❌ NIT NO ENCONTRADO"

        filas.append({
            "NIT Emisor":    nit,
            "Nombre Emisor": nombre_f,
            "IVA Documento": iva,
            "Nombre Tercero (Libro)": nombre_l,
            "IVA en Libro":  "✅" if iva_ok else "❌",
            "Estado":        estado,
        })
    return pd.DataFrame(filas)


def exportar_excel(df_fact_res, df_nc_res):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        df_fact_res.to_excel(w, index=False, sheet_name="Facturas")
        df_nc_res.to_excel(w, index=False, sheet_name="Notas de Crédito")
    return buf.getvalue()


def mostrar_resultado(df_res, key_prefix):
    total     = len(df_res)
    correctos = (df_res["Estado"] == "✅ CORRECTO").sum()
    iva_no    = (df_res["Estado"] == "⚠️ NIT OK / IVA NO ENCONTRADO").sum()
    nit_no    = (df_res["Estado"] == "❌ NIT NO ENCONTRADO").sum()
    pct       = round(correctos / total * 100, 1) if total > 0 else 0

    m1, m2, m3, m4, m5 = st.columns(5)
    with m1: st.markdown(f'<div class="metric-card"><div class="number blue">{total}</div><div class="label">Documentos analizados</div></div>', unsafe_allow_html=True)
    with m2: st.markdown(f'<div class="metric-card"><div class="number green">{correctos}</div><div class="label">✅ Correctos</div></div>', unsafe_allow_html=True)
    with m3: st.markdown(f'<div class="metric-card"><div class="number orange">{iva_no}</div><div class="label">⚠️ NIT OK / IVA no coincide</div></div>', unsafe_allow_html=True)
    with m4: st.markdown(f'<div class="metric-card"><div class="number red">{nit_no}</div><div class="label">❌ NIT no encontrado</div></div>', unsafe_allow_html=True)
    with m5:
        color = "green" if pct == 100 else ("orange" if pct >= 80 else "red")
        st.markdown(f'<div class="metric-card"><div class="number {color}">{pct}%</div><div class="label">Coincidencia</div></div>', unsafe_allow_html=True)

    st.markdown('<div class="result-block">', unsafe_allow_html=True)
    st.markdown('<p class="section-title">📋 Resultado detallado</p>', unsafe_allow_html=True)

    filtro = st.radio(
        "Mostrar:",
        ["Todos", "Solo ✅ CORRECTOS", "Solo ⚠️ NIT OK / IVA NO", "Solo ❌ NIT NO ENCONTRADO"],
        horizontal=True,
        key=f"filtro_{key_prefix}"
    )
    df_m = df_res.copy()
    if filtro == "Solo ✅ CORRECTOS":           df_m = df_m[df_m["Estado"] == "✅ CORRECTO"]
    elif filtro == "Solo ⚠️ NIT OK / IVA NO":  df_m = df_m[df_m["Estado"] == "⚠️ NIT OK / IVA NO ENCONTRADO"]
    elif filtro == "Solo ❌ NIT NO ENCONTRADO": df_m = df_m[df_m["Estado"] == "❌ NIT NO ENCONTRADO"]

    def color_fila(row):
        e = str(row["Estado"])
        if "CORRECTO" in e: return ["background-color:#F0FFF4"] * len(row)
        elif "NIT OK" in e: return ["background-color:#FFF8E1"] * len(row)
        else:               return ["background-color:#FFF5F5"] * len(row)

    st.dataframe(df_m.style.apply(color_fila, axis=1), use_container_width=True, height=440)
    st.markdown('</div>', unsafe_allow_html=True)


# ── UI ───────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="header-block">
    <h1>🧾 Comparativa IVA — Facturas & Notas de Crédito vs Libro Contable</h1>
    <p>Verifica NIT Emisor, Nombre e IVA de facturas y notas de crédito electrónicas contra el libro auxiliar de IVA</p>
</div>""", unsafe_allow_html=True)

col1, col2 = st.columns(2)
with col1:
    st.markdown('<div class="upload-section"><p class="section-title">📄 Facturas Electrónicas (Excel)</p>', unsafe_allow_html=True)
    st.caption("Tipo de documento · NIT Emisor · Nombre Emisor · IVA")
    archivo_fact = st.file_uploader("Facturas", type=["xlsx", "xls"], key="fact", label_visibility="collapsed")
    st.markdown('</div>', unsafe_allow_html=True)
with col2:
    st.markdown('<div class="upload-section"><p class="section-title">📊 Libro Auxiliar IVA (Excel)</p>', unsafe_allow_html=True)
    st.caption("Identificación · Nombre tercero · Valor impuesto compras (positivos y negativos)")
    archivo_libro = st.file_uploader("Libro IVA", type=["xlsx", "xls"], key="libro", label_visibility="collapsed")
    st.markdown('</div>', unsafe_allow_html=True)

with st.expander("⚙️ Opciones avanzadas"):
    tolerancia = st.number_input(
        "Tolerancia IVA", min_value=0.0, max_value=1000.0, value=1.0, step=0.5, format="%.2f",
        help="Diferencia máxima aceptable entre el IVA del documento y el valor en el libro."
    )

if archivo_fact and archivo_libro:
    with st.spinner("Procesando..."):
        df_fact, df_nc, col_nit, col_nombre_f, col_iva = cargar_facturas(archivo_fact)
        df_libro_pos, df_libro_neg, col_id, col_nombre_l, col_vic = cargar_libro_iva(archivo_libro)

    if df_fact is None or df_libro_pos is None:
        st.stop()

    df_fact_res = comparar(df_fact, col_nit, col_nombre_f, col_iva,
                           df_libro_pos, col_id, col_nombre_l, col_vic, tolerancia)
    df_nc_res   = comparar(df_nc, col_nit, col_nombre_f, col_iva,
                           df_libro_neg, col_id, col_nombre_l, col_vic, tolerancia)

    st.markdown("---")
    tab1, tab2 = st.tabs(["🧾 Facturas Electrónicas", "🔄 Notas de Crédito Electrónicas"])

    with tab1:
        st.markdown('<p class="tab-header">Facturas vs. Valores positivos en Libro IVA</p>', unsafe_allow_html=True)
        if len(df_fact_res) == 0:
            st.info("No se encontraron facturas electrónicas en el archivo.")
        else:
            mostrar_resultado(df_fact_res, "fact")

    with tab2:
        st.markdown('<p class="tab-header">Notas de Crédito vs. Valores negativos en Libro IVA (comparados en valor absoluto)</p>', unsafe_allow_html=True)
        if len(df_nc_res) == 0:
            st.info("No se encontraron notas de crédito electrónicas en el archivo.")
        else:
            mostrar_resultado(df_nc_res, "nc")

    st.markdown("---")
    c1, c2, _ = st.columns([1, 1, 3])
    with c1:
        st.download_button(
            "⬇️ Descargar Excel (ambas hojas)",
            exportar_excel(df_fact_res, df_nc_res),
            f"comparativa_iva_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    with c2:
        csv_combined = pd.concat(
            [df_fact_res.assign(Tipo="Factura"), df_nc_res.assign(Tipo="Nota de Crédito")],
            ignore_index=True
        )
        st.download_button(
            "⬇️ Descargar CSV",
            csv_combined.to_csv(index=False).encode("utf-8"),
            f"comparativa_iva_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
            mime="text/csv",
            use_container_width=True
        )

    with st.expander("🔎 Vista previa archivos cargados"):
        t1, t2, t3, t4 = st.tabs([
            "Facturas (archivo)", "Notas de Crédito (archivo)",
            "Libro IVA — Positivos", "Libro IVA — Negativos (abs)"
        ])
        with t1: st.dataframe(df_fact, use_container_width=True, height=300)
        with t2: st.dataframe(df_nc, use_container_width=True, height=300)
        with t3: st.dataframe(df_libro_pos, use_container_width=True, height=300)
        with t4: st.dataframe(df_libro_neg, use_container_width=True, height=300)

else:
    st.markdown("""
    <div style="text-align:center;padding:60px 20px;color:#999;">
        <div style="font-size:3rem;margin-bottom:16px;">📂</div>
        <p style="font-size:1.1rem;font-weight:600;">Carga los dos archivos para iniciar</p>
        <p style="font-size:.9rem;">
            Facturas (Tipo · NIT Emisor · Nombre Emisor · IVA) &nbsp;·&nbsp;
            Libro IVA (Identificación · Nombre tercero · Valor impuesto compras)
        </p>
        <p style="font-size:.85rem;color:#bbb;margin-top:8px;">
            Las <strong>Notas de Crédito</strong> se detectan automáticamente en la columna "Tipo de documento"
            y se comparan contra los valores <em>negativos</em> del Libro IVA.
        </p>
    </div>""", unsafe_allow_html=True)
