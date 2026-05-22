import streamlit as st
import pandas as pd
import json
from datetime import datetime
import xml.etree.ElementTree as ET
from io import BytesIO

st.set_page_config(page_title="La Trilla - Lista de Precios", layout="wide", page_icon="📋")

# ===================== CONEXIÓN A GOOGLE SHEETS =====================
spreadsheet = st.session_state.google_spreadsheet

def load_sheet(sheet_name):
    try:
        worksheet = spreadsheet.worksheet(sheet_name)
        data = worksheet.get_all_records()
        return pd.DataFrame(data)
    except:
        try:
            worksheet = spreadsheet.add_worksheet(title=sheet_name, rows="1000", cols="20")
        except:
            pass
        return pd.DataFrame()

def save_sheet(sheet_name, df):
    try:
        worksheet = spreadsheet.worksheet(sheet_name)
    except:
        worksheet = spreadsheet.add_worksheet(title=sheet_name, rows="1000", cols="20")
    worksheet.clear()
    df = df.replace([float('nan'), float('inf'), -float('inf')], [None, None, None])
    if not df.empty:
        worksheet.update([df.columns.values.tolist()] + df.values.tolist())
    else:
        worksheet.update([[]])

# ==================== ETIQUETAS DESDE GOOGLE DRIVE ====================
ETIQUETAS_FOLDER_ID = "1T8T3Lt3VyckBwPWkZISpjR_KgYo1aJR7"

def _get_drive_service():
    from googleapiclient.discovery import build
    gc = st.session_state.google_spreadsheet.client
    credentials = gc.auth
    return build('drive', 'v3', credentials=credentials)

def buscar_etiqueta_base(nombre_producto):
    try:
        service = _get_drive_service()
        query = f"'{ETIQUETAS_FOLDER_ID}' in parents and name contains '{nombre_producto}' and trashed=false"
        resultados = service.files().list(q=query, fields="files(id, name)").execute()
        archivos = resultados.get("files", [])
        return archivos[0] if archivos else None
    except Exception as e:
        st.error(f"Error buscando etiqueta en Drive: {e}")
        return None

def descargar_etiqueta_drive(file_id):
    try:
        from googleapiclient.http import MediaIoBaseDownload
        service = _get_drive_service()
        request = service.files().get_media(fileId=file_id)
        buf = BytesIO()
        downloader = MediaIoBaseDownload(buf, request)
        done = False
        while not done:
            _, done = downloader.next_chunk()
        buf.seek(0)
        return buf
    except Exception as e:
        st.error(f"Error descargando etiqueta: {e}")
        return None

# ==================== FORMATO FECHA CHILENA ====================
def format_fecha_chilena(fecha):
    if isinstance(fecha, str):
        try:
            fecha = datetime.strptime(fecha, "%Y-%m-%d").date()
        except:
            return fecha
    return fecha.strftime("%d/%m/%Y")

# ==================== FORMATO CLP ====================
def format_clp(valor):
    if isinstance(valor, (int, float)):
        return f"${int(valor):,}".replace(",", ".")
    return valor

# ==================== ESTILO TABLA ====================
def style_by_numero(df):
    df = df.copy()
    if "N°" not in df.columns:
        df["N°"] = range(1, len(df) + 1)
    unique_n = sorted(df["N°"].unique())
    colors = ["#f8f9fa", "#e6f3ff"]
    color_map = {n: colors[i % len(colors)] for i, n in enumerate(unique_n)}
    def apply_color(row):
        color = color_map.get(row["N°"], "#ffffff")
        return [f'background-color: {color}' for _ in row]
    return df.style.apply(apply_color, axis=1).set_properties(**{'text-align': 'left', 'font-size': '14px'})

# ==================== CARGAR PRECIOS ====================
def cargar_precios(uploaded_file):
    if uploaded_file is None:
        return None
    df = pd.read_excel(uploaded_file)
    df = df.rename(columns={"Subproducto": "Productos"})
    columnas = ["N°", "Lote", "Productos", "Precio Venta Bruto", "Precio KG"]
    for col in columnas:
        if col not in df.columns:
            st.error(f"Falta columna: {col}")
            return None
    df["Precio Venta Bruto"] = df["Precio Venta Bruto"].apply(format_clp)
    df["Precio KG"] = df["Precio KG"].apply(format_clp)
    return df

# ==================== HISTORIAL EN GOOGLE SHEETS ====================
HISTORIAL_SHEET = "historial_precios"

def guardar_lote_historial(df, fecha_llegada):
    fecha = fecha_llegada.strftime("%Y-%m-%d")
    fecha_recepcion = datetime.now().strftime("%Y-%m-%d")
    productos_json = json.dumps(df[["N°", "Productos", "Precio Venta Bruto", "Precio KG"]].to_dict(orient="records"), ensure_ascii=False)
    nuevo_registro = pd.DataFrame([{
        "lote": df["Lote"].iloc[0],
        "fecha_llegada": fecha,
        "fecha_recepcion": fecha_recepcion,
        "productos_json": productos_json
    }])
    df_actual = load_sheet(HISTORIAL_SHEET)
    df_actual = pd.concat([df_actual, nuevo_registro], ignore_index=True)
    save_sheet(HISTORIAL_SHEET, df_actual)
    st.success(f"✅ Lote {df['Lote'].iloc[0]} guardado en Google Sheets")

def listar_lotes_recibidos():
    df = load_sheet(HISTORIAL_SHEET)
    if df.empty:
        return []
    return sorted(df["lote"].unique().tolist(), reverse=True)

def cargar_historial_lote(lote):
    df = load_sheet(HISTORIAL_SHEET)
    if df.empty:
        return None
    fila = df[df["lote"] == lote]
    if fila.empty:
        return None
    data = fila.iloc[0].to_dict()
    data["productos"] = json.loads(data["productos_json"])
    return data

# ==================== MODIFICAR ETIQUETA ====================
def modificar_etiqueta_ezpx(archivo_buf, lote_nuevo, precio_bruto, precio_kg, fecha_prod, fecha_venc, origen):
    try:
        tree = ET.parse(archivo_buf)
        root = tree.getroot()
        for elem in root.iter("GraphicShape"):
            data = elem.find("Data")
            if data is None or data.text is None:
                continue
            txt = (data.text or "").strip()
            if txt == "09ABL604":
                data.text = lote_nuevo
            elif txt == "$1.400":
                data.text = precio_bruto
            elif txt == "$22.800":
                data.text = precio_kg
            elif txt == "CANADÁ":
                data.text = origen.upper()
            elif txt == "02/2026":
                data.text = fecha_prod
            elif txt == "02/2028":
                data.text = fecha_venc
        buf = BytesIO()
        tree.write(buf, encoding="utf-8", xml_declaration=True)
        buf.seek(0)
        return buf.getvalue()
    except Exception as e:
        st.error(f"Error al modificar etiqueta: {e}")
        return None

# ==================== APP ====================
st.title("📋 Lista de Precios")

tab2, tab1 = st.tabs(["🔎 Consultar Lotes", "🆕 Nueva Recepción"])

# ====================== TAB 1 - NUEVA RECEPCIÓN ======================
with tab1:
    st.subheader("📤 Nueva Recepción de Precios")

    if "uploader_counter" not in st.session_state:
        st.session_state.uploader_counter = 0

    if st.session_state.get("recepcion_exitosa", False):
        st.success("✅ Recepción exitosa")
        st.session_state.recepcion_exitosa = False

    uploaded_file = st.file_uploader(
        "Subir archivo Excel de precios",
        type=["xlsx"],
        key=f"uploader_key_{st.session_state.uploader_counter}"
    )

    if uploaded_file:
        df = cargar_precios(uploaded_file)
        if df is not None:
            fecha_llegada = st.date_input("Fecha de llegada a Bodega", value=None)
            st.subheader("Vista previa del lote")
            columnas_mostrar = ["Lote", "Productos", "Precio Venta Bruto", "Precio KG"]
            df_preview = df[columnas_mostrar + ["N°"]].copy()
            styled_df = style_by_numero(df_preview)
            st.dataframe(styled_df, use_container_width=True, hide_index=True, column_order=columnas_mostrar)
            if st.button("📦 Recepcionar Lotes", type="primary", disabled=fecha_llegada is None):
                guardar_lote_historial(df, fecha_llegada)
                st.session_state.recepcion_exitosa = True
                st.session_state.uploader_counter += 1
                st.rerun()

# ====================== TAB 2 - CONSULTAR LOTES ======================
with tab2:
    st.subheader("🔎 Consultar Lotes Recibidos")
    lotes_disponibles = listar_lotes_recibidos()
    if lotes_disponibles:
        lote_seleccionado = st.selectbox("Seleccionar Lote", lotes_disponibles)
        if st.button("Cargar Información del Lote", type="primary"):
            data = cargar_historial_lote(lote_seleccionado)
            if data:
                st.session_state.data_lote = data
                st.success(f"✅ Lote cargado: **{data['lote']}**")

        if "data_lote" in st.session_state and st.session_state.data_lote:
            data = st.session_state.data_lote
            st.success(f"**Lote:** {data['lote']} | **Fecha llegada:** {format_fecha_chilena(data['fecha_llegada'])}")

            df_lote = pd.DataFrame(data["productos"])
            busqueda_precio = st.text_input("🔎 Buscar producto", placeholder="Ej: Lenteja, Poroto...", key="busqueda_precio")
            if busqueda_precio:
                df_lote = df_lote[df_lote["Productos"].str.contains(busqueda_precio, case=False, na=False)]
            styled_lote = style_by_numero(df_lote)

            selection = st.dataframe(
                styled_lote,
                use_container_width=True,
                hide_index=True,
                column_order=["Productos", "Precio Venta Bruto", "Precio KG"],
                on_select="rerun",
                selection_mode="single-row",
                key=f"tabla_lote_{lote_seleccionado}"
            )

            if len(selection["selection"]["rows"]) > 0:
                fila_idx = selection["selection"]["rows"][0]
                fila = df_lote.iloc[fila_idx]
                producto = fila["Productos"]
                precio_bruto = fila["Precio Venta Bruto"]
                precio_kg = fila["Precio KG"]

                if "current_producto" not in st.session_state or st.session_state.current_producto != producto:
                    st.session_state.current_producto = producto
                    if "ruta_base" in st.session_state:
                        del st.session_state.ruta_base
                    if "etiqueta_bytes" in st.session_state:
                        del st.session_state.etiqueta_bytes

                st.info(f"**Producto seleccionado:** {producto}")

                if st.button("🔍 Cargar Etiqueta Base", type="primary"):
                    archivo = buscar_etiqueta_base(producto)
                    if archivo:
                        st.session_state.ruta_base = archivo
                        st.success(f"✅ Etiqueta base encontrada: {archivo['name']}")
                    else:
                        st.error(f"❌ No se encontró etiqueta para '{producto}' en Google Drive")

                if "ruta_base" in st.session_state:
                    st.subheader("📋 Datos para la etiqueta")

                    paises = ["Seleccione un país..."] + ["Afganistán", "Albania", "Alemania", "Andorra", "Angola", "Antigua y Barbuda", "Arabia Saudita", "Argelia", "Argentina", "Armenia", "Australia", "Austria", "Azerbaiyán", "Bahamas", "Bangladés", "Barbados", "Baréin", "Bélgica", "Belice", "Benín", "Bielorrusia", "Birmania", "Bolivia", "Bosnia y Herzegovina", "Botsuana", "Brasil", "Brunéi", "Bulgaria", "Burkina Faso", "Burundi", "Bután", "Cabo Verde", "Camboya", "Camerún", "Canadá", "Catar", "Chad", "Chile", "China", "Chipre", "Ciudad del Vaticano", "Colombia", "Comoras", "Corea del Norte", "Corea del Sur", "Costa de Marfil", "Costa Rica", "Croacia", "Cuba", "Dinamarca", "Dominica", "Ecuador", "Egipto", "El Salvador", "Emiratos Árabes Unidos", "Eritrea", "Eslovaquia", "Eslovenia", "España", "Estados Unidos", "Estonia", "Etiopía", "Filipinas", "Finlandia", "Fiyi", "Francia", "Gabón", "Gambia", "Georgia", "Ghana", "Granada", "Grecia", "Guatemala", "Guinea", "Guinea-Bisáu", "Guinea Ecuatorial", "Guyana", "Haití", "Honduras", "Hungría", "India", "Indonesia", "Irak", "Irán", "Irlanda", "Islandia", "Islas Marshall", "Islas Salomón", "Israel", "Italia", "Jamaica", "Japón", "Jordania", "Kazajistán", "Kenia", "Kirguistán", "Kiribati", "Kuwait", "Laos", "Lesoto", "Letonia", "Líbano", "Liberia", "Libia", "Liechtenstein", "Lituania", "Luxemburgo", "Macedonia del Norte", "Madagascar", "Malasia", "Malaui", "Maldivas", "Malí", "Malta", "Marruecos", "Mauricio", "Mauritania", "México", "Micronesia", "Moldavia", "Mónaco", "Mongolia", "Montenegro", "Mozambique", "Namibia", "Nauru", "Nepal", "Nicaragua", "Níger", "Nigeria", "Noruega", "Nueva Zelanda", "Omán", "Países Bajos", "Pakistán", "Palaos", "Panamá", "Papúa Nueva Guinea", "Paraguay", "Perú", "Polonia", "Portugal", "Reino Unido", "República Centroafricana", "República Checa", "República Democrática del Congo", "República Dominicana", "República del Congo", "Ruanda", "Rumania", "Rusia", "Samoa", "San Cristóbal y Nieves", "San Marino", "San Vicente y las Granadinas", "Santa Lucía", "Santo Tomé y Príncipe", "Senegal", "Serbia", "Seychelles", "Sierra Leona", "Singapur", "Siria", "Somalia", "Sri Lanka", "Suazilandia", "Sudáfrica", "Sudán", "Sudán del Sur", "Suecia", "Suiza", "Surinam", "Tailandia", "Tanzania", "Tayikistán", "Timor Oriental", "Togo", "Tonga", "Trinidad y Tobago", "Túnez", "Turkmenistán", "Turquía", "Tuvalu", "Ucrania", "Uganda", "Uruguay", "Uzbekistán", "Vanuatu", "Venezuela", "Vietnam", "Yemen", "Yibuti", "Zambia", "Zimbabue"]
                    origen = st.selectbox("País de Origen", options=paises, index=0, key=f"origen_{producto}")
                    fecha_prod = st.date_input("Fecha de Producción", value=None, key=f"prod_{producto}", format="DD/MM/YYYY")
                    fecha_venc = st.date_input("Fecha de Vencimiento", value=None, key=f"venc_{producto}", format="DD/MM/YYYY")

                    disabled = (origen == "Seleccione un país..." or fecha_prod is None or fecha_venc is None)

                    if st.button("🚀 Confirmar y Descargar", type="primary", disabled=disabled):
                        archivo_buf = descargar_etiqueta_drive(st.session_state.ruta_base['id'])
                        if archivo_buf:
                            bytes_etiqueta = modificar_etiqueta_ezpx(
                                archivo_buf,
                                data['lote'],
                                precio_bruto,
                                precio_kg,
                                fecha_prod.strftime("%m/%Y"),
                                fecha_venc.strftime("%m/%Y"),
                                origen
                            )
                            if bytes_etiqueta:
                                st.session_state.etiqueta_bytes = bytes_etiqueta
                                st.success("✅ Etiqueta generada correctamente")

                    if "etiqueta_bytes" in st.session_state and st.session_state.etiqueta_bytes:
                        st.download_button(
                            label="⬇️ Descargar etiqueta modificada",
                            data=st.session_state.etiqueta_bytes,
                            file_name=f"{producto} - {data['lote']}.ezpx",
                            mime="application/xml"
                        )
    else:
        st.info("Aún no hay lotes recepcionados")

st.caption("Desarrollado para La Trilla con ❤️ • Datos en Google Sheets v1.1")
