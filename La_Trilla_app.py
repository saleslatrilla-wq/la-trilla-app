import streamlit as st
from pathlib import Path

st.set_page_config(
    page_title="La Trilla · Sistema Integral",
    layout="wide",
    page_icon="☕"
)

# ===================== ESTILOS GLOBALES =====================
st.markdown("""
<link href="https://fonts.googleapis.com/css2?family=Playfair+Display:wght@400;600;700&family=DM+Sans:wght@300;400;500;600&display=swap" rel="stylesheet">
<style>
/* ── RESET & BASE ── */
*, *::before, *::after { box-sizing: border-box; }

html, body, [data-testid="stAppViewContainer"] {
    background-color: #0f0e0c !important;
    color: #e8e0d0 !important;
    font-family: 'DM Sans', sans-serif !important;
}

/* ── FONDO PRINCIPAL ── */
[data-testid="stAppViewContainer"] > .main {
    background: linear-gradient(160deg, #0f0e0c 0%, #1a1714 60%, #0f0e0c 100%) !important;
}
[data-testid="stMain"] {
    background: transparent !important;
}

/* ── SIDEBAR ── */
[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #141210 0%, #1c1916 100%) !important;
    border-right: 1px solid #2e2820 !important;
}
[data-testid="stSidebar"] * {
    color: #d4c9b5 !important;
    font-family: 'DM Sans', sans-serif !important;
}
[data-testid="stSidebar"] .stSelectbox label,
[data-testid="stSidebar"] .stMarkdown p {
    color: #a89880 !important;
    font-size: 0.78rem !important;
    letter-spacing: 0.08em !important;
    text-transform: uppercase !important;
}
[data-testid="stSidebar"] hr {
    border-color: #2e2820 !important;
    margin: 0.8rem 0 !important;
}
[data-testid="stSidebar"] .stCaption {
    color: #5a5040 !important;
    font-size: 0.72rem !important;
}

/* ── SELECTBOX SIDEBAR ── */
[data-testid="stSidebar"] [data-baseweb="select"] > div {
    background: #1f1c18 !important;
    border: 1px solid #3a3028 !important;
    border-radius: 8px !important;
    color: #e8e0d0 !important;
}
[data-testid="stSidebar"] [data-baseweb="select"] > div:hover {
    border-color: #c8a84b !important;
}

/* ── TÍTULOS ── */
h1, h2, h3 {
    font-family: 'Playfair Display', serif !important;
    color: #f0e6d0 !important;
    letter-spacing: -0.01em !important;
}
h1 { 
    font-size: 2.2rem !important;
    font-weight: 700 !important;
    border-bottom: 1px solid #2e2820;
    padding-bottom: 0.5rem;
    margin-bottom: 1.2rem !important;
}
h2 { font-size: 1.5rem !important; font-weight: 600 !important; }
h3 { font-size: 1.15rem !important; }

/* ── MÉTRICAS ── */
[data-testid="metric-container"] {
    background: #1a1714 !important;
    border: 1px solid #2e2820 !important;
    border-radius: 12px !important;
    padding: 1rem !important;
}
[data-testid="stMetricValue"] {
    color: #c8a84b !important;
    font-family: 'Playfair Display', serif !important;
    font-size: 1.8rem !important;
}
[data-testid="stMetricLabel"] {
    color: #a89880 !important;
    font-size: 0.78rem !important;
    letter-spacing: 0.06em !important;
    text-transform: uppercase !important;
}

/* ── BOTONES PRIMARIOS ── */
.stButton > button[kind="primary"] {
    background: linear-gradient(135deg, #c8a84b 0%, #a8882b 100%) !important;
    color: #0f0e0c !important;
    border: none !important;
    border-radius: 8px !important;
    font-family: 'DM Sans', sans-serif !important;
    font-weight: 600 !important;
    font-size: 0.9rem !important;
    letter-spacing: 0.04em !important;
    padding: 0.55rem 1.4rem !important;
    transition: all 0.2s ease !important;
    box-shadow: 0 4px 14px rgba(200, 168, 75, 0.25) !important;
}
.stButton > button[kind="primary"]:hover {
    background: linear-gradient(135deg, #d4b85c 0%, #b8982b 100%) !important;
    box-shadow: 0 6px 20px rgba(200, 168, 75, 0.4) !important;
    transform: translateY(-1px) !important;
}

/* ── BOTONES SECUNDARIOS ── */
.stButton > button:not([kind="primary"]) {
    background: transparent !important;
    border: 1px solid #3a3028 !important;
    color: #d4c9b5 !important;
    border-radius: 8px !important;
    font-family: 'DM Sans', sans-serif !important;
    transition: all 0.2s ease !important;
}
.stButton > button:not([kind="primary"]):hover {
    border-color: #c8a84b !important;
    color: #c8a84b !important;
}

/* ── INPUTS / TEXT ── */
.stTextInput > div > div > input,
.stTextInput > div > div > input[type="password"] {
    background: #1a1714 !important;
    border: 1px solid #3a3028 !important;
    border-radius: 8px !important;
    color: #e8e0d0 !important;
    font-family: 'DM Sans', sans-serif !important;
    padding: 0.65rem 1rem !important;
    transition: border-color 0.2s !important;
}
.stTextInput > div > div > input:focus {
    border-color: #c8a84b !important;
    box-shadow: 0 0 0 2px rgba(200, 168, 75, 0.15) !important;
}
.stTextInput label {
    color: #a89880 !important;
    font-size: 0.82rem !important;
    letter-spacing: 0.05em !important;
    text-transform: uppercase !important;
}

/* ── NUMBER INPUT ── */
.stNumberInput > div > div > input {
    background: #1a1714 !important;
    border: 1px solid #3a3028 !important;
    border-radius: 8px !important;
    color: #e8e0d0 !important;
}

/* ── DATE INPUT ── */
.stDateInput > div > div > input {
    background: #1a1714 !important;
    border: 1px solid #3a3028 !important;
    border-radius: 8px !important;
    color: #e8e0d0 !important;
}

/* ── SELECTBOX GENERAL ── */
[data-baseweb="select"] > div {
    background: #1a1714 !important;
    border: 1px solid #3a3028 !important;
    border-radius: 8px !important;
    color: #e8e0d0 !important;
}
[data-baseweb="select"] > div:hover {
    border-color: #c8a84b !important;
}
[data-baseweb="popover"] {
    background: #1f1c18 !important;
    border: 1px solid #3a3028 !important;
}
[data-baseweb="menu"] {
    background: #1f1c18 !important;
}
[data-baseweb="menu"] li:hover {
    background: #2a2520 !important;
}

/* ── TABS ── */
[data-baseweb="tab-list"] {
    background: #141210 !important;
    border-radius: 10px !important;
    padding: 4px !important;
    gap: 2px !important;
    border: 1px solid #2e2820 !important;
}
[data-baseweb="tab"] {
    background: transparent !important;
    color: #a89880 !important;
    border-radius: 7px !important;
    font-family: 'DM Sans', sans-serif !important;
    font-size: 0.85rem !important;
    font-weight: 500 !important;
    border: none !important;
    padding: 0.5rem 1.2rem !important;
    transition: all 0.2s !important;
}
[data-baseweb="tab"][aria-selected="true"] {
    background: #c8a84b !important;
    color: #0f0e0c !important;
    font-weight: 600 !important;
}
[data-baseweb="tab-highlight"] { display: none !important; }

/* ── DATAFRAME ── */
[data-testid="stDataFrame"] {
    border-radius: 12px !important;
    overflow: hidden !important;
    border: 1px solid #2e2820 !important;
}
[data-testid="stDataFrameResizable"] {
    background: #141210 !important;
}

/* ── ALERTAS ── */
[data-testid="stAlert"] {
    border-radius: 10px !important;
    border-left-width: 4px !important;
    font-family: 'DM Sans', sans-serif !important;
}
div[data-baseweb="notification"][kind="positive"],
.stSuccess {
    background: rgba(34, 85, 34, 0.25) !important;
    border-left-color: #4caf50 !important;
    color: #90cc90 !important;
}
div[data-baseweb="notification"][kind="negative"],
.stError {
    background: rgba(100, 30, 30, 0.3) !important;
    border-left-color: #e05050 !important;
    color: #f09090 !important;
}
div[data-baseweb="notification"][kind="info"],
.stInfo {
    background: rgba(30, 60, 100, 0.25) !important;
    border-left-color: #c8a84b !important;
    color: #c8a84b !important;
}
div[data-baseweb="notification"][kind="warning"],
.stWarning {
    background: rgba(100, 70, 20, 0.3) !important;
    border-left-color: #f0a030 !important;
    color: #f0c060 !important;
}

/* ── FILE UPLOADER ── */
[data-testid="stFileUploader"] {
    background: #1a1714 !important;
    border: 2px dashed #3a3028 !important;
    border-radius: 12px !important;
    transition: border-color 0.2s !important;
}
[data-testid="stFileUploader"]:hover {
    border-color: #c8a84b !important;
}
[data-testid="stFileUploader"] label {
    color: #a89880 !important;
}

/* ── EXPANDER ── */
[data-testid="stExpander"] {
    background: #1a1714 !important;
    border: 1px solid #2e2820 !important;
    border-radius: 10px !important;
}
[data-testid="stExpander"] summary {
    color: #d4c9b5 !important;
    font-family: 'DM Sans', sans-serif !important;
}

/* ── CONTAINER CON BORDE ── */
[data-testid="stVerticalBlockBorderWrapper"] {
    background: #1a1714 !important;
    border: 1px solid #2e2820 !important;
    border-radius: 12px !important;
}

/* ── PROGRESS BAR ── */
[data-testid="stProgressBar"] > div {
    background: #2e2820 !important;
    border-radius: 99px !important;
}
[data-testid="stProgressBar"] > div > div {
    background: linear-gradient(90deg, #c8a84b, #a8882b) !important;
    border-radius: 99px !important;
}

/* ── CHECKBOX ── */
[data-baseweb="checkbox"] span {
    border-color: #3a3028 !important;
    background: #1a1714 !important;
}
[data-baseweb="checkbox"][aria-checked="true"] span {
    background: #c8a84b !important;
    border-color: #c8a84b !important;
}

/* ── SCROLLBAR ── */
::-webkit-scrollbar { width: 6px; height: 6px; }
::-webkit-scrollbar-track { background: #141210; }
::-webkit-scrollbar-thumb { background: #3a3028; border-radius: 3px; }
::-webkit-scrollbar-thumb:hover { background: #c8a84b; }

/* ── CAPTION ── */
.stCaption {
    color: #5a5040 !important;
    font-size: 0.75rem !important;
}

/* ── DIVIDER ── */
hr {
    border-color: #2e2820 !important;
}

/* ── DOWNLOAD BUTTON ── */
[data-testid="stDownloadButton"] > button {
    background: #1a1714 !important;
    border: 1px solid #c8a84b !important;
    color: #c8a84b !important;
    border-radius: 8px !important;
    font-family: 'DM Sans', sans-serif !important;
    font-weight: 500 !important;
    transition: all 0.2s !important;
}
[data-testid="stDownloadButton"] > button:hover {
    background: #c8a84b !important;
    color: #0f0e0c !important;
}

/* ── SPINNER ── */
[data-testid="stSpinner"] {
    color: #c8a84b !important;
}
</style>
""", unsafe_allow_html=True)

# ===================== CARPETAS =====================
Path("historial_precios").mkdir(parents=True, exist_ok=True)
Path("Etiquetas_base").mkdir(parents=True, exist_ok=True)

# ===================== GOOGLE SHEETS =====================
@st.cache_resource
def get_google_client():
    import gspread
    from google.oauth2.service_account import Credentials
    creds = Credentials.from_service_account_info(
        st.secrets["google_service_account"],
        scopes=[
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive"
        ]
    )
    return gspread.authorize(creds)

@st.cache_resource
def get_spreadsheet():
    client = get_google_client()
    return client.open_by_key("1S3fpi4UYQoNe0kJzic4fhipZqI6BzsmhY9wPRYFd5FM")

st.session_state.google_spreadsheet = get_spreadsheet()

# ===================== CONTRASEÑAS =====================
GLOBAL_PASSWORD = st.secrets["global"]["password"]
CALC_ADMIN_PASSWORD = st.secrets["calculadora"]["password"]

# ===================== LOGIN GLOBAL =====================
if "global_logged_in" not in st.session_state:
    st.session_state.global_logged_in = False

if not st.session_state.global_logged_in:
    st.markdown("""
    <style>
    [data-testid="stAppViewContainer"] > .main {
        background: radial-gradient(ellipse at 30% 20%, #1f1a10 0%, #0f0e0c 60%) !important;
    }
    </style>
    """, unsafe_allow_html=True)

    col1, col2, col3 = st.columns([1, 1.4, 1])
    with col2:
        st.markdown("<div style='height:3rem'></div>", unsafe_allow_html=True)
        try:
            st.image("LOGO Blanco sin Fondo.png", use_container_width=True)
        except:
            st.markdown("<h1 style='text-align:center;font-family:Playfair Display,serif;color:#f0e6d0;'>La Trilla</h1>", unsafe_allow_html=True)

        st.markdown("""
        <div style='text-align:center;color:#5a5040;font-size:0.78rem;letter-spacing:0.25em;
        text-transform:uppercase;margin-bottom:2.5rem;font-family:DM Sans,sans-serif;'>
        La Tostaduria &middot; Sistema Integral</div>
        <div style='background:#141210;border:1px solid #2e2820;border-radius:16px;
        padding:2rem;box-shadow:0 20px 60px rgba(0,0,0,0.6);'>
        """, unsafe_allow_html=True)

        password = st.text_input("Contraseña", type="password", placeholder="Ingresa la contraseña del sistema")

        if st.button("Entrar al Sistema", type="primary", use_container_width=True):
            if password == GLOBAL_PASSWORD:
                st.session_state.global_logged_in = True
                st.success("✅ Acceso concedido")
                st.rerun()
            else:
                st.error("❌ Contraseña incorrecta")

        st.markdown("</div>", unsafe_allow_html=True)
        st.markdown("""
        <div style='text-align:center;color:#3a3028;font-size:0.7rem;letter-spacing:0.1em;
        font-family:DM Sans,sans-serif;margin-top:1.5rem;'>
        &copy; La Trilla &middot; Acceso Restringido</div>
        """, unsafe_allow_html=True)
    st.stop()

# ===================== ADMIN CALCULADORA =====================
if "calc_admin_mode" not in st.session_state:
    st.session_state.calc_admin_mode = False

# ===================== BARRA LATERAL =====================
try:
    st.sidebar.image("LOGO Blanco sin Fondo.png", use_container_width=True)
except:
    st.sidebar.markdown("<h2 style='font-family:Playfair Display,serif;color:#f0e6d0;text-align:center;'>La Trilla</h2>", unsafe_allow_html=True)

st.sidebar.markdown("""
<div style='text-align:center;color:#5a5040;font-size:0.7rem;letter-spacing:0.2em;
text-transform:uppercase;padding-bottom:0.8rem;border-bottom:1px solid #2e2820;
font-family:DM Sans,sans-serif;'>Sistema Integral</div>
""", unsafe_allow_html=True)

modulo = st.sidebar.selectbox(
    "Módulo",
    [
        "📋 1. Lista de Precios",
        "🏭 2. Producción / Envasado",
        "📊 3. Calculadora de Precios"
    ]
)

st.sidebar.markdown("---")
st.sidebar.caption("v1.0 · Datos en Google Sheets")

# ===================== EJECUCIÓN DE MÓDULOS =====================
if modulo == "📋 1. Lista de Precios":
    with open("lista_precios_app.py", "r", encoding="utf-8") as f:
        exec(f.read())

elif modulo == "🏭 2. Producción / Envasado":
    with open("Produccion_app.py", "r", encoding="utf-8") as f:
        exec(f.read())

elif modulo == "📊 3. Calculadora de Precios":
    with st.sidebar.expander("🔑 Modo Administrador - Calculadora"):
        if not st.session_state.calc_admin_mode:
            pwd = st.text_input("Contraseña Admin Calculadora", type="password")
            if st.button("Activar Calculadora", use_container_width=True):
                if pwd == CALC_ADMIN_PASSWORD:
                    st.session_state.calc_admin_mode = True
                    st.success("✅ Calculadora desbloqueada")
                    st.rerun()
                else:
                    st.error("❌ Contraseña incorrecta")
        else:
            st.success("🔓 Modo Admin Activo")
            if st.button("Bloquear Calculadora", use_container_width=True):
                st.session_state.calc_admin_mode = False
                st.rerun()

    if not st.session_state.calc_admin_mode:
        st.error("🔒 **Calculadora bloqueada**")
        st.info("Usa el 'Modo Administrador' en la barra lateral para desbloquearla.")
        st.stop()

    with open("Calculadora_Precios_app.py", "r", encoding="utf-8") as f:
        exec(f.read())