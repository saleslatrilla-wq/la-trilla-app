import streamlit as st
from pathlib import Path

st.set_page_config(
    page_title="🧩 Sistema Integral",
    layout="wide",
    page_icon="📊"
)

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
        scopes=["https://www.googleapis.com/auth/spreadsheets"]
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
    st.title("🔐 La Trilla - Sistema Integral")
    st.subheader("Iniciar Sesión")
    
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        # Imagen con manejo de error
        try:
            st.image("LOGO Blanco sin Fondo.png", use_container_width=True)
        except:
            st.markdown("### 🌾 La Trilla")
        
        password = st.text_input("Contraseña del sistema", type="password", placeholder="Ingresa la contraseña")
        
        if st.button("Entrar al Sistema", type="primary", use_container_width=True):
            if password == GLOBAL_PASSWORD:
                st.session_state.global_logged_in = True
                st.success("✅ Acceso concedido")
                st.rerun()
            else:
                st.error("❌ Contraseña incorrecta")
    st.stop()

# ===================== ADMIN CALCULADORA =====================
if "calc_admin_mode" not in st.session_state:
    st.session_state.calc_admin_mode = False

# ===================== BARRA LATERAL =====================
try:
    st.sidebar.image("LOGO Blanco sin Fondo.png", use_container_width=True)
except:
    st.sidebar.markdown("### 🌾 La Trilla")

st.sidebar.markdown("#### Sistema Integral")
st.sidebar.markdown("---")

modulo = st.sidebar.selectbox(
    "Selecciona el módulo:",
    [
        "📋 1. Lista de Precios",
        "🏭 2. Producción / Envasado",
        "📊 3. Calculadora de Precios"
    ]
)

st.sidebar.markdown("---")
st.sidebar.caption("Versión Integral - Datos en Google Sheets")

# ===================== EJECUCIÓN DE MÓDULOS =====================
if modulo == "📋 1. Lista de Precios":
    with open("lista_precios_app.py", "r", encoding="utf-8") as f:
        exec(f.read())

elif modulo == "🏭 2. Producción / Envasado":
    with open("Producción_app.py", "r", encoding="utf-8") as f:
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
