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
# Clave única de administrador: en Secrets como
#   [modo_admin]
#   password = "..."
# (también acepta ["modo admin"] con comillas)
def _admin_secret():
    for k in ("modo_admin", "modo admin"):
        try:
            return str(st.secrets[k]["password"])
        except Exception:
            pass
    return None

ADMIN_PASSWORD = _admin_secret()

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

# ===================== MODO ADMINISTRADOR (único) =====================
if "admin_mode" not in st.session_state:
    st.session_state.admin_mode = False

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
        "📊 3. Calculadora de Precios",
        "📦 4. Pedidos",
        "🛒 5. Compras"
    ]
)

st.sidebar.markdown("---")
st.sidebar.caption("Versión Integral - Datos en Google Sheets")

# ----- Activar Modo Administrador (al final de la barra lateral) -----
import hmac as _hmac
with st.sidebar.expander("🔓 Modo Administrador activo" if st.session_state.admin_mode else "🔑 Activar Modo Administrador"):
    if not st.session_state.admin_mode:
        with st.form("form_modo_admin", clear_on_submit=True):
            _pwd = st.text_input("Contraseña de administrador", type="password")
            _ok = st.form_submit_button("Activar", use_container_width=True)
        if _ok:
            if ADMIN_PASSWORD is None:
                st.error("Falta [modo_admin] password en los Secrets.")
            elif _hmac.compare_digest(_pwd.encode(), ADMIN_PASSWORD.encode()):
                st.session_state.admin_mode = True
                st.rerun()
            else:
                st.error("❌ Contraseña incorrecta")
    else:
        st.success("Acceso a: Administración (Lista de Precios), Calculadora y Compras")
        if st.button("🔒 Desactivar Modo Administrador", use_container_width=True):
            st.session_state.admin_mode = False
            st.rerun()

def _requiere_admin(nombre):
    if not st.session_state.admin_mode:
        st.error(f"🔒 **{nombre}** requiere Modo Administrador")
        st.info("Actívalo al final de la barra lateral: 🔑 Activar Modo Administrador.")
        st.stop()

# ===================== EJECUCIÓN DE MÓDULOS =====================
if modulo == "📋 1. Lista de Precios":
    with open("lista_precios_app.py", "r", encoding="utf-8") as f:
        exec(f.read())

elif modulo == "🏭 2. Producción / Envasado":
    with open("Produccion_app.py", "r", encoding="utf-8") as f:
        exec(f.read())

elif modulo == "📊 3. Calculadora de Precios":
    _requiere_admin("Calculadora de Precios")

    with open("Calculadora_Precios_app.py", "r", encoding="utf-8") as f:
        exec(f.read())

elif modulo == "📦 4. Pedidos":
    with open("Pedidos_app.py", "r", encoding="utf-8") as f:
        exec(f.read())

elif modulo == "🛒 5. Compras":
    _requiere_admin("Compras")
    with open("Compras_app.py", "r", encoding="utf-8") as f:
        exec(f.read())
