import re
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


# ── Utilidades NIT ───────────────────────────────────────────────────────────

def normalizar_nit(nit):
    """Deja solo dígitos (quita puntos, guiones, .0 de floats, etc.)."""
    return re.sub(r"[^0-9]", "", str(nit or ""))


def claves_nit(nit):
    """Variantes del NIT (con y sin posible dígito de verificación) para
    poder cruzar NIT del documento (sin DV) vs Identificación del libro
    (a veces con el DV pegado al final)."""
    n = normalizar_nit(nit)
    claves = {n} if n else set()
    if len(n) > 1:
        claves.add(n[:-1])
    return claves


# ── Loaders ─────────────────────────────────────────────────────────────────

def cargar_facturas(archivo):
    """Carga facturas electrónicas. Devuelve dos DataFrames (facturas y notas
    de crédito) con columnas NIT / Nombre / IVA / Direccion, donde Direccion
    indica si el documento fue 'venta' (emitido por la empresa) o 'compra'
    (recibido), determinado automáticamente comparando NIT Emisor/Receptor
    contra el NIT propio detectado."""
    df = pd.read_excel(archivo, header=0, engine="openpyxl")
    df.columns = [str(c).strip() for c in df.columns]

    col_tipo   = next((c for c in df.columns if "tipo" in c.lower() and "documento" in c.lower()), df.columns[0])
    col_nit    = next((c for c in df.columns if "nit" in c.lower() and "emisor" in c.lower()), None)
    col_nombre = next((c for c in df.columns if "nombre" in c.lower() and "emisor" in c.lower()), None)
    col_nit_r    = next((c for c in df.columns if "nit" in c.lower() and "receptor" in c.lower()), None)
    col_nombre_r = next((c for c in df.columns if "nombre" in c.lower() and "receptor" in c.lower()), None)
    col_iva    = next((c for c in df.columns if c.strip().upper() == "IVA"), None)
    col_grupo  = next((c for c in df.columns if "grupo" in c.lower()), None)

    if not all([col_nit, col_nombre, col_iva]):
        st.error(f"Columnas no encontradas en Facturas. Disponibles: {list(df.columns)}")
        return None, None, None, None, None, None

    df[col_nit]    = df[col_nit].apply(normalizar_nit)
    df[col_nombre] = df[col_nombre].astype(str).str.strip()
    df[col_iva]    = pd.to_numeric(df[col_iva], errors="coerce")

    tiene_receptor = bool(col_nit_r and col_nombre_r)
    nit_propio = None

    if tiene_receptor:
        df[col_nit_r]    = df[col_nit_r].apply(normalizar_nit)
        df[col_nombre_r] = df[col_nombre_r].astype(str).str.strip()

        # NIT propio = el que más se repite entre Emisor + Receptor
        conteo = pd.concat([df[col_nit], df[col_nit_r]]).value_counts()
        if len(conteo) > 0:
            nit_propio = conteo.index[0]

        if col_grupo:
            es_venta = df[col_grupo].astype(str).str.strip().str.lower().eq("emitido")
        else:
            es_venta = df[col_nit] == nit_propio

        df["_NIT"]       = df[col_nit_r].where(es_venta, df[col_nit])
        df["_Nombre"]    = df[col_nombre_r].where(es_venta, df[col_nombre])
        df["_Direccion"] = es_venta.map({True: "venta", False: "compra"})
    else:
        # Sin columna de receptor: se mantiene el comportamiento anterior (solo compras)
        df["_NIT"]       = df[col_nit]
        df["_Nombre"]    = df[col_nombre]
        df["_Direccion"] = "compra"

    df["_IVA"] = df[col_iva]

    mask_nc = df[col_tipo].astype(str).str.strip().str.lower().str.contains("nota de cr")
    cols = ["_NIT", "_Nombre", "_IVA", "_Direccion"]
    ren  = {"_NIT": "NIT", "_Nombre": "Nombre", "_IVA": "IVA", "_Direccion": "Direccion"}
    df_fact = df[~mask_nc][cols].rename(columns=ren).copy()
    df_nc   = df[mask_nc][cols].rename(columns=ren).copy()

    return df_fact, df_nc, nit_propio, tiene_receptor, None, None


def cargar_libro_iva(archivo):
    """Carga el libro auxiliar de IVA y construye índices NIT -> [valores]
    tanto para 'Valor impuesto ventas' como 'Valor impuesto compras'
    (y sus columnas de devolución, si existen), con el NIT indexado en sus
    dos variantes (con/sin dígito de verificación)."""
    df_raw = pd.read_excel(archivo, header=None, engine="openpyxl")
    header_row = next((i for i, row in df_raw.iterrows()
                       if any("identificaci" in str(v).lower() for v in row.values)), None)
    if header_row is None:
        st.error("No se encontró fila de encabezados en el Libro IVA.")
        return None

    df = df_raw.iloc[header_row + 1:].copy()
    df.columns = [str(c).strip() for c in df_raw.iloc[header_row].values]

    col_id     = next((c for c in df.columns if "identificaci" in c.lower()), None)
    col_nombre = next((c for c in df.columns if "nombre" in c.lower() and "tercero" in c.lower()), None)
    col_vic    = next((c for c in df.columns if "valor" in c.lower() and "impuesto" in c.lower() and "compra" in c.lower() and "devoluci" not in c.lower()), None)
    col_viv    = next((c for c in df.columns if "valor" in c.lower() and "impuesto" in c.lower() and "venta" in c.lower() and "devoluci" not in c.lower()), None)
    col_vidc   = next((c for c in df.columns if "valor" in c.lower() and "devoluci" in c.lower() and "compra" in c.lower()), None)
    col_vidv   = next((c for c in df.columns if "valor" in c.lower() and "devoluci" in c.lower() and "venta" in c.lower()), None)

    if not all([col_id, col_nombre, col_vic]):
        st.error(f"Columnas no encontradas en Libro IVA. Disponibles: {list(df.columns)}")
        return None

    if col_viv is None:
        st.warning("⚠️ No se encontró la columna 'Valor impuesto ventas' en el Libro IVA. "
                   "Los documentos emitidos (ventas) no podrán conciliarse.")

    df[col_id]     = df[col_id].apply(normalizar_nit)
    df             = df[df[col_id] != ""].copy()
    df[col_nombre] = df[col_nombre].astype(str).str.strip()
    for c in [col_vic, col_viv, col_vidc, col_vidv]:
        if c is not None:
            df[c] = pd.to_numeric(df[c], errors="coerce")

    def construir_indice(col_valor):
        idx_valores, idx_nombres = {}, {}
        if col_valor is None:
            return idx_valores, idx_nombres
        for _, row in df.iterrows():
            v = row[col_valor]
            if pd.isna(v) or v == 0:
                continue
            v = round(abs(float(v)), 2)
            for k in claves_nit(row[col_id]):
                idx_valores.setdefault(k, []).append(v)
                idx_nombres.setdefault(k, row[col_nombre])
        return idx_valores, idx_nombres

    idx_compras, nom_compras   = construir_indice(col_vic)
    idx_ventas,  nom_ventas    = construir_indice(col_viv)
    idx_dev_compras, nom_dc    = construir_indice(col_vidc)
    idx_dev_ventas,  nom_dv    = construir_indice(col_vidv)

    return {
        "compras": (idx_compras, nom_compras),
        "ventas": (idx_ventas, nom_ventas),
        "dev_compras": (idx_dev_compras, nom_dc),
        "dev_ventas": (idx_dev_ventas, nom_dv),
        "tiene_ventas": col_viv is not None,
        "tiene_devoluciones": (col_vidc is not None) or (col_vidv is not None),
    }


# ── Comparador genérico ──────────────────────────────────────────────────────

def comparar(df_fact, libro, tolerancia, es_nota_credito=False):
    filas = []
    for _, row in df_fact.iterrows():
        nit       = str(row["NIT"]).strip()
        nombre_f  = str(row["Nombre"]).strip()
        iva       = round(float(row["IVA"]), 2) if pd.notna(row["IVA"]) else 0.0
        direccion = row["Direccion"]

        if es_nota_credito:
            clave_idx = "dev_ventas" if direccion == "venta" else "dev_compras"
        else:
            clave_idx = "ventas" if direccion == "venta" else "compras"
        idx_nit_iva, idx_nit_nombre = libro[clave_idx]

        nit_ok = False
        valores = None
        for k in claves_nit(nit):
            if k in idx_nit_iva:
                nit_ok = True
                valores = idx_nit_iva[k]
                break
        nombre_l = ""
        for k in claves_nit(nit):
            if k in idx_nit_nombre:
                nombre_l = idx_nit_nombre[k]
                break

        # ── Lógica IVA = 0: no hay nada que cruzar contra el libro ──
        if iva == 0.0:
            estado = "✅ CORRECTO (IVA $0)"
            iva_ok = True
        elif nit_ok and any(abs(iva - v) <= tolerancia for v in valores):
            estado = "✅ CORRECTO"
            iva_ok = True
        elif nit_ok:
            estado = "⚠️ NIT OK / IVA NO ENCONTRADO"
            iva_ok = False
        else:
            estado = "❌ NIT NO ENCONTRADO"
            iva_ok = False

        filas.append({
            "NIT":                    nit,
            "Nombre":                 nombre_f,
            "Dirección":              "Venta (emitido)" if direccion == "venta" else "Compra (recibido)",
            "IVA Documento":          iva,
            "Nombre Tercero (Libro)": nombre_l,
            "IVA en Libro":           "✅" if iva_ok else "❌",
            "Estado":                 estado,
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
    correctos = df_res["Estado"].str.startswith("✅").sum()
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
    if filtro == "Solo ✅ CORRECTOS":
        df_m = df_m[df_m["Estado"].str.startswith("✅")]
    elif filtro == "Solo ⚠️ NIT OK / IVA NO":
        df_m = df_m[df_m["Estado"] == "⚠️ NIT OK / IVA NO ENCONTRADO"]
    elif filtro == "Solo ❌ NIT NO ENCONTRADO":
        df_m = df_m[df_m["Estado"] == "❌ NIT NO ENCONTRADO"]

    def color_fila(row):
        e = str(row["Estado"])
        if e.startswith("✅"):  return ["background-color:#F0FFF4"] * len(row)
        elif "NIT OK" in e:     return ["background-color:#FFF8E1"] * len(row)
        else:                   return ["background-color:#FFF5F5"] * len(row)

    st.dataframe(df_m.style.apply(color_fila, axis=1), use_container_width=True, height=440)
    st.markdown('</div>', unsafe_allow_html=True)


# ── UI ───────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="header-block">
    <h1>🧾 Comparativa IVA — Facturas & Notas de Crédito vs Libro Contable</h1>
    <p>Verifica NIT, Nombre e IVA de facturas y notas de crédito electrónicas (emitidas y recibidas) contra el libro auxiliar de IVA</p>
</div>""", unsafe_allow_html=True)

col1, col2 = st.columns(2)
with col1:
    st.markdown('<div class="upload-section"><p class="section-title">📄 Facturas Electrónicas (Excel)</p>', unsafe_allow_html=True)
    st.caption("Tipo de documento · NIT/Nombre Emisor · NIT/Nombre Receptor (opcional) · IVA")
    archivo_fact = st.file_uploader("Facturas", type=["xlsx", "xls"], key="fact", label_visibility="collapsed")
    st.markdown('</div>', unsafe_allow_html=True)
with col2:
    st.markdown('<div class="upload-section"><p class="section-title">📊 Libro Auxiliar IVA (Excel)</p>', unsafe_allow_html=True)
    st.caption("Identificación · Nombre tercero · Valor impuesto ventas / compras")
    archivo_libro = st.file_uploader("Libro IVA", type=["xlsx", "xls"], key="libro", label_visibility="collapsed")
    st.markdown('</div>', unsafe_allow_html=True)

with st.expander("⚙️ Opciones avanzadas"):
    tolerancia = st.number_input(
        "Tolerancia IVA", min_value=0.0, max_value=1000.0, value=1.0, step=0.5, format="%.2f",
        help="Diferencia máxima aceptable entre el IVA del documento y el valor en el libro."
    )

if archivo_fact and archivo_libro:
    with st.spinner("Procesando..."):
        df_fact, df_nc, nit_propio, tiene_receptor, _, _ = cargar_facturas(archivo_fact)
        libro = cargar_libro_iva(archivo_libro)

    if df_fact is None or libro is None:
        st.stop()

    if tiene_receptor and nit_propio:
        st.caption(f"🏢 NIT propio detectado automáticamente: **{nit_propio}** "
                   f"(usado para saber si cada documento es venta o compra)")
    elif not tiene_receptor:
        st.info("El archivo de facturas no tiene columna de NIT Receptor: todo se tratará como compra "
                "(comportamiento anterior).")
    if not libro["tiene_ventas"]:
        st.info("El libro no tiene columna de ventas: solo se conciliarán documentos de compra.")

    df_fact_res = comparar(df_fact, libro, tolerancia, es_nota_credito=False)
    df_nc_res   = comparar(df_nc, libro, tolerancia, es_nota_credito=True) if libro["tiene_devoluciones"] else pd.DataFrame(columns=df_fact_res.columns)

    st.markdown("---")
    tab1, tab2 = st.tabs(["🧾 Facturas Electrónicas", "🔄 Notas de Crédito Electrónicas"])

    with tab1:
        st.markdown('<p class="tab-header">Facturas (ventas → Valor impuesto ventas · compras → Valor impuesto compras)</p>', unsafe_allow_html=True)
        if len(df_fact_res) == 0:
            st.info("No se encontraron facturas electrónicas en el archivo.")
        else:
            mostrar_resultado(df_fact_res, "fact")

    with tab2:
        st.markdown('<p class="tab-header">Notas de Crédito (contra columnas de devolución en ventas/compras)</p>', unsafe_allow_html=True)
        if not libro["tiene_devoluciones"]:
            st.warning("El libro cargado no tiene columnas de devolución (ventas/compras); no se pueden conciliar notas de crédito.")
        elif len(df_nc_res) == 0:
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
        t1, t2 = st.tabs(["Facturas (normalizado)", "Notas de Crédito (normalizado)"])
        with t1: st.dataframe(df_fact, use_container_width=True, height=300)
        with t2: st.dataframe(df_nc, use_container_width=True, height=300)

else:
    st.markdown("""
    <div style="text-align:center;padding:60px 20px;color:#999;">
        <div style="font-size:3rem;margin-bottom:16px;">📂</div>
        <p style="font-size:1.1rem;font-weight:600;">Carga los dos archivos para iniciar</p>
        <p style="font-size:.9rem;">
            Facturas (Tipo · NIT/Nombre Emisor · NIT/Nombre Receptor · IVA) &nbsp;·&nbsp;
            Libro IVA (Identificación · Nombre tercero · Valor impuesto ventas/compras)
        </p>
        <p style="font-size:.85rem;color:#bbb;margin-top:8px;">
            La app detecta automáticamente el <strong>NIT propio</strong> de la empresa para saber
            si cada documento fue <em>emitido</em> (venta, se compara contra "Valor impuesto ventas")
            o <em>recibido</em> (compra, contra "Valor impuesto compras"). Las
            <strong>Notas de Crédito</strong> se detectan en "Tipo de documento" y se comparan
            contra las columnas de devolución correspondientes.
        </p>
    </div>""", unsafe_allow_html=True)
