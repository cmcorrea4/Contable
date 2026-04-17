import streamlit as st
import streamlit_authenticator as stauth
import yaml
from yaml.loader import SafeLoader

# ── Configuración de página ──────────────────────────────────────────────────
st.set_page_config(
    page_title="Portal Contable",
    page_icon="🏢",
    layout="wide",
    initial_sidebar_ebar_state="expanded"
)

# ── Estilos CSS ──────────────────────────────────────────────────────────────
st.markdown("""
<style>
    .stApp { background-color: #F5F7FA; }
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    .block-container { padding-top: 2rem; padding-bottom: 2rem; }

    .hero-container {
        background: linear-gradient(135deg, #0A2342 0%, #175676 100%);
        border-radius: 16px;
        padding: 40px 50px;
        margin-bottom: 40px;
        color: white;
        text-align: center;
        box-shadow: 0 10px 25px rgba(0,0,0,0.1);
    }
    .hero-title {
        font-size: 2.8rem;
        font-weight: 800;
        margin-bottom: 15px;
        letter-spacing: -0.02em;
    }
    .hero-subtitle {
        font-size: 1.2rem;
        font-weight: 400;
        opacity: 0.9;
        max-width: 800px;
        margin: 0 auto;
        line-height: 1.6;
    }
    .section-header {
        text-align: center;
        font-size: 1.8rem;
        color: #0A2342;
        font-weight: 700;
        margin-bottom: 30px;
        position: relative;
    }
    .section-header::after {
        content: '';
        display: block;
        width: 60px;
        height: 4px;
        background: #4BA3E3;
        margin: 10px auto 0;
        border-radius: 2px;
    }
    .feature-card {
        background: white;
        border-radius: 12px;
        padding: 30px;
        height: 100%;
        box-shadow: 0 4px 15px rgba(0,0,0,0.05);
        border-top: 4px solid #4BA3E3;
        transition: transform 0.3s ease, box-shadow 0.3s ease;
        text-align: center;
    }
    .feature-card:hover {
        transform: translateY(-5px);
        box-shadow: 0 12px 25px rgba(0,0,0,0.1);
    }
    .feature-icon-container {
        width: 60px;
        height: 60px;
        background: #E8F4FC;
        border-radius: 50%;
        display: flex;
        align-items: center;
        justify-content: center;
        margin: 0 auto 20px;
        font-size: 1.8rem;
    }
    .feature-title {
        font-size: 1.3rem;
        font-weight: 700;
        color: #1E293B;
        margin-bottom: 15px;
    }
    .feature-desc {
        font-size: 0.95rem;
        color: #64748B;
        line-height: 1.5;
        margin-bottom: 20px;
    }
    .feature-chip {
        display: inline-block;
        background: #F1F5F9;
        color: #475569;
        font-size: 0.75rem;
        font-weight: 600;
        padding: 4px 12px;
        border-radius: 20px;
        border: 1px solid #E2E8F0;
        margin: 2px;
    }
    .info-footer {
        margin-top: 60px;
        text-align: center;
        color: #94A3B8;
        font-size: 0.85rem;
        border-top: 1px solid #E2E8F0;
        padding-top: 20px;
    }

    /* Login form styling */
    .login-wrapper {
        max-width: 420px;
        margin: 60px auto;
        background: white;
        border-radius: 16px;
        padding: 40px;
        box-shadow: 0 10px 30px rgba(0,0,0,0.1);
        border-top: 5px solid #4BA3E3;
    }
    .login-title {
        text-align: center;
        color: #0A2342;
        font-size: 1.6rem;
        font-weight: 700;
        margin-bottom: 8px;
    }
    .login-subtitle {
        text-align: center;
        color: #64748B;
        font-size: 0.9rem;
        margin-bottom: 30px;
    }
</style>
""", unsafe_allow_html=True)


# ── Carga de credenciales desde st.secrets ───────────────────────────────────
# En Streamlit Cloud: App Settings > Secrets > pega el contenido de secrets.toml
credentials = {
    "usernames": {}
}

for username, data in st.secrets.get("credentials", {}).items():
    credentials["usernames"][username] = {
        "name": data["name"],
        "password": data["password"],  # debe ser hash bcrypt
        "email": data.get("email", ""),
    }

cookie_config = st.secrets.get("cookie", {
    "name": "portal_contable_auth",
    "key": "clave_secreta_cambiar",
    "expiry_days": 1
})

# ── Inicializar autenticador ──────────────────────────────────────────────────
authenticator = stauth.Authenticate(
    credentials,
    cookie_config["name"],
    cookie_config["key"],
    cookie_config["expiry_days"],
)

# ── Pantalla de Login ─────────────────────────────────────────────────────────
if not st.session_state.get("authentication_status"):
    st.markdown("""
    <div style="text-align:center; padding: 30px 0 10px;">
        <div style="font-size:3rem;">🏢</div>
        <div style="font-size:1.8rem; font-weight:800; color:#0A2342;">Portal Contable</div>
        <div style="color:#64748B; margin-top:6px;">Ingrese sus credenciales para continuar</div>
    </div>
    """, unsafe_allow_html=True)

authenticator.login(location="main", key="login_portal")

auth_status = st.session_state.get("authentication_status")

if auth_status is False:
    st.error("❌ Usuario o contraseña incorrectos.")
    st.stop()

elif auth_status is None:
    st.stop()

# ── Usuario autenticado: mostrar app ─────────────────────────────────────────
# Botón de logout en sidebar
with st.sidebar:
    st.markdown(f"👤 **{st.session_state.get('name', '')}**")
    authenticator.logout("Cerrar sesión", location="sidebar", key="logout_sidebar")
    st.divider()

# ── Contenido Principal ───────────────────────────────────────────────────────

st.markdown("""
<div class="hero-container">
    <div class="hero-title">Automatización y Conciliación Contable</div>
    <div class="hero-subtitle">
        Bienvenido al portal centralizado para la validación de extractos bancarios, 
        conciliación de IVA en facturación electrónica y preparación de traslados de nómina.
        Seleccione una herramienta en el menú lateral para comenzar.
    </div>
</div>
""", unsafe_allow_html=True)

st.markdown('<div class="section-header">Nuestras Herramientas</div>', unsafe_allow_html=True)

col1, col2 = st.columns(2)

with col1:
    st.markdown("""
    <div class="feature-card">
        <div class="feature-icon-container">🔍</div>
        <div class="feature-title">Comparativa Contable</div>
        <div class="feature-desc">
            Audita las transacciones bancarias cruzando automáticamente los valores del extracto (CSV) contra los registros de débitos y créditos en el libro contable (Excel).
        </div>
        <div>
            <span class="feature-chip">CSV Bancario</span>
            <span class="feature-chip">Excel Contable</span>
            <span class="feature-chip">Tolerancias</span>
        </div>
    </div>
    """, unsafe_allow_html=True)

with col2:
    st.markdown("""
    <div class="feature-card">
        <div class="feature-icon-container" style="background:#EBF9F1; color:#1A9E5C;">🧾</div>
        <div class="feature-title">Comparativa de IVA</div>
        <div class="feature-desc">
            Verifique rápida y confiablemente que el NIT Emisor, el Nombre y el IVA reportado en el reporte de Facturación Electrónica coincidan exactamente con el Libro Auxiliar.
        </div>
        <div>
            <span class="feature-chip">NIT Validación</span>
            <span class="feature-chip">Reportes DIAN</span>
        </div>
    </div>
    """, unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)
col3, col4 = st.columns(2)

with col3:
    st.markdown("""
    <div class="feature-card">
        <div class="feature-icon-container" style="background:#FFF3E0; color:#E65100;">📋</div>
        <div class="feature-title">Traslado de Nómina</div>
        <div class="feature-desc">
            Herramienta para agilizar la preparación de la interfaz de nómina. Traspasa conceptos y valores desde el reporte crudo (3.3) hacia la plantilla formateada (3.2).
        </div>
        <div>
            <span class="feature-chip">Mapeo Dinámico</span>
            <span class="feature-chip">Deduplicación</span>
        </div>
    </div>
    """, unsafe_allow_html=True)

with col4:
    st.markdown("""
    <div class="feature-card">
        <div class="feature-icon-container" style="background:#F3E8FF; color:#7C3AED;">📊</div>
        <div class="feature-title">Anexo ERI</div>
        <div class="feature-desc">
            Genera el anexo al Estado de Resultados Integral desde la hoja terceros_ del archivo EEFF, con desglose por grupo, métricas de ingresos y gastos y gráficos interactivos.
        </div>
        <div>
            <span class="feature-chip">Ingresos / Gastos</span>
            <span class="feature-chip">Pivote Mensual</span>
            <span class="feature-chip">Plotly</span>
        </div>
    </div>
    """, unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)
col5, col6 = st.columns(2)

with col5:
    st.markdown("""
    <div class="feature-card">
        <div class="feature-icon-container" style="background:#E1F0FF; color:#0A2342;">🤖</div>
        <div class="feature-title">Agente de Datos</div>
        <div class="feature-desc">
            Analiza archivos Excel y CSV interactuando mediante lenguaje natural y un asistente IA avanzado para explorar datos, generar estadísticas y obtener respuestas sobre la información.
        </div>
        <div>
            <span class="feature-chip">OpenAI GPT</span>
            <span class="feature-chip">CSV / Excel</span>
            <span class="feature-chip">Análisis IA</span>
        </div>
    </div>
    """, unsafe_allow_html=True)

st.markdown("""
<div class="info-footer">
    Gestión Contable Inteligente v2.1 • Diseñado para precisión y ahorro de tiempo operativo.
</div>
""", unsafe_allow_html=True)
