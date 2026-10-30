import streamlit as st
import pandas as pd
import json
from datetime import datetime
import xml.etree.ElementTree as ET
from io import BytesIO
import re
import zipfile
import hmac

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
    if pd.isna(valor) or valor is None:
        return ""
    if isinstance(valor, (int, float)):
        try:
            return f"${int(valor):,}".replace(",", ".")
        except (ValueError, TypeError):
            return ""
    return valor

# ==================== ESTILO TABLA ====================
def style_by_numero(df):
    df = df.copy()
    if "N°" not in df.columns:
        df["N°"] = range(1, len(df) + 1)
    df["N°"] = pd.to_numeric(df["N°"], errors="coerce").fillna(0).astype(int)
    unique_n = sorted(df["N°"].unique())
    colors = ["#f8f9fa", "#e6f3ff"]
    color_map = {n: colors[i % len(colors)] for i, n in enumerate(unique_n)}
    color_series = df["N°"].map(color_map).reset_index(drop=True)
    df = df.drop(columns=["N°"]).reset_index(drop=True)
    def apply_color(row):
        color = color_series.iloc[row.name]
        return [f'background-color: {color}' for _ in row]
    return df.style.apply(apply_color, axis=1).set_properties(**{'text-align': 'left', 'font-size': '14px'})

# ==================== CARGAR PRECIOS ====================
def cargar_precios(uploaded_file):
    if uploaded_file is None:
        return None
    df = pd.read_excel(uploaded_file, engine="xlrd" if uploaded_file.name.endswith(".xls") else "openpyxl")
    if "Precio Venta Bruto" in df.columns and "Redondeo" in df.columns:
        df = df.drop(columns=["Precio Venta Bruto"])
    df = df.rename(columns={"Subproducto": "Productos", "Redondeo": "Precio Venta Bruto"})
    columnas = ["N°", "Lote", "Productos", "Precio Venta Bruto", "Precio KG"]
    for col in columnas:
        if col not in df.columns:
            st.error(f"Falta columna: {col}")
            return None

    # Filtrar filas sin precios válidos (ej: productos a granel/Saco sin margen configurado -> Redondeo/Precio KG vacíos)
    df = df.dropna(subset=["Precio Venta Bruto", "Precio KG"])
    if df.empty:
        st.error("❌ No hay productos con precios de venta calculados en este archivo.\n"
                 "Las filas de productos a granel (Saco/Bolsa) sin margen configurado se omiten automáticamente.\n"
                 "Asegúrate de configurar márgenes > 0% en la Calculadora de Precios para los productos que quieres recepcionar.")
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
    # Ordenar por fecha_recepcion descendente (más reciente primero)
    df["fecha_recepcion_dt"] = pd.to_datetime(df["fecha_recepcion"], errors="coerce")
    df_sorted = df.sort_values("fecha_recepcion_dt", ascending=False)
    # Obtener lotes únicos manteniendo el orden (el más reciente primero)
    lotes_unicos = df_sorted["lote"].drop_duplicates().tolist()
    return lotes_unicos

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
    import re
    try:
        tree = ET.parse(archivo_buf)
        root = tree.getroot()

        # Detectar patrones dinámicamente en la etiqueta base
        patron_precio = re.compile(r"^\$[\d\.]+$")       # $1.200 / $37.500
        patron_fecha  = re.compile(r"^\d{2}/\d{4}$")     # 05/2025 / 06/2026
        patron_lote   = re.compile(r"^\d{2}[A-Z]{3}\d{3}$")  # 22CPL606

        precios_encontrados = []  # [(elem_data, valor)]
        fechas_encontradas  = []
        
        for elem in root.iter("GraphicShape"):
            data = elem.find("Data")
            if data is None or data.text is None:
                continue
            txt = data.text.strip()
            if patron_precio.match(txt):
                precios_encontrados.append((data, txt))
            elif patron_fecha.match(txt):
                fechas_encontradas.append((data, txt))
            elif patron_lote.match(txt):
                data.text = lote_nuevo
            elif txt.upper() in [p.upper() for p in ["CANADÁ","CANADA","ANDORRA","ARGENTINA","BRASIL","CHILE","CHINA","COLOMBIA","ESPAÑA","ESTADOS UNIDOS","FRANCIA","ITALIA","MÉXICO","PERU","PORTUGAL","URUGUAY"] + [origen.upper()]]:
                data.text = origen.upper()

        # El precio menor = precio bruto unitario; el mayor = precio por KG
        if len(precios_encontrados) >= 2:
            precios_encontrados.sort(key=lambda x: float(x[1].replace("$","").replace(".","").replace(",",".")))
            precios_encontrados[0][0].text = precio_bruto
            precios_encontrados[1][0].text = precio_kg
        elif len(precios_encontrados) == 1:
            precios_encontrados[0][0].text = precio_bruto

        # La fecha menor = producción; la mayor = vencimiento
        if len(fechas_encontradas) >= 2:
            fechas_encontradas.sort(key=lambda x: x[1])
            fechas_encontradas[0][0].text = fecha_prod
            fechas_encontradas[1][0].text = fecha_venc
        elif len(fechas_encontradas) == 1:
            fechas_encontradas[0][0].text = fecha_prod

        buf = BytesIO()
        tree.write(buf, encoding="utf-8", xml_declaration=True)
        buf.seek(0)
        return buf.getvalue()
    except Exception as e:
        st.error(f"Error al modificar etiqueta: {e}")
        return None

# ==================== CONFIGURACIÓN ADMIN ====================
# Define la clave en .streamlit/secrets.toml  ->  admin_password = "tu_clave"
# Si no existe en secrets, se usa esta (CÁMBIALA).
ADMIN_PASSWORD_DEFAULT = "trilla"

def _admin_password():
    try:
        return str(st.secrets["admin_password"])
    except Exception:
        return ADMIN_PASSWORD_DEFAULT

# ==================== DATOS DE PRODUCTOS (ORIGEN Y FECHAS) ====================
DATOS_SHEET = "datos_productos"
COLS_DATOS = ["producto", "origen", "fecha_produccion", "fecha_vencimiento", "ultimo_uso", "ultimo_lote"]
REGEX_FECHA = r"^(0[1-9]|1[0-2])/\d{4}$"   # MM/AAAA
PAISES = ["Afganistán", "Albania", "Alemania", "Andorra", "Angola", "Antigua y Barbuda", "Arabia Saudita", "Argelia", "Argentina", "Armenia", "Australia", "Austria", "Azerbaiyán", "Bahamas", "Bangladés", "Barbados", "Baréin", "Bélgica", "Belice", "Benín", "Bielorrusia", "Birmania", "Bolivia", "Bosnia y Herzegovina", "Botsuana", "Brasil", "Brunéi", "Bulgaria", "Burkina Faso", "Burundi", "Bután", "Cabo Verde", "Camboya", "Camerún", "Canadá", "Catar", "Chad", "Chile", "China", "Chipre", "Ciudad del Vaticano", "Colombia", "Comoras", "Corea del Norte", "Corea del Sur", "Costa de Marfil", "Costa Rica", "Croacia", "Cuba", "Dinamarca", "Dominica", "Ecuador", "Egipto", "El Salvador", "Emiratos Árabes Unidos", "Eritrea", "Eslovaquia", "Eslovenia", "España", "Estados Unidos", "Estonia", "Etiopía", "Filipinas", "Finlandia", "Fiyi", "Francia", "Gabón", "Gambia", "Georgia", "Ghana", "Granada", "Grecia", "Guatemala", "Guinea", "Guinea-Bisáu", "Guinea Ecuatorial", "Guyana", "Haití", "Honduras", "Hungría", "India", "Indonesia", "Irak", "Irán", "Irlanda", "Islandia", "Islas Marshall", "Islas Salomón", "Israel", "Italia", "Jamaica", "Japón", "Jordania", "Kazajistán", "Kenia", "Kirguistán", "Kiribati", "Kuwait", "Laos", "Lesoto", "Letonia", "Líbano", "Liberia", "Libia", "Liechtenstein", "Lituania", "Luxemburgo", "Macedonia del Norte", "Madagascar", "Malasia", "Malaui", "Maldivas", "Malí", "Malta", "Marruecos", "Mauricio", "Mauritania", "México", "Micronesia", "Moldavia", "Mónaco", "Mongolia", "Montenegro", "Mozambique", "Namibia", "Nauru", "Nepal", "Nicaragua", "Níger", "Nigeria", "Noruega", "Nueva Zelanda", "Omán", "Países Bajos", "Pakistán", "Palaos", "Panamá", "Papúa Nueva Guinea", "Paraguay", "Perú", "Polonia", "Portugal", "Reino Unido", "República Centroafricana", "República Checa", "República Democrática del Congo", "República Dominicana", "República del Congo", "Ruanda", "Rumania", "Rusia", "Samoa", "San Cristóbal y Nieves", "San Marino", "San Vicente y las Granadinas", "Santa Lucía", "Santo Tomé y Príncipe", "Senegal", "Serbia", "Seychelles", "Sierra Leona", "Singapur", "Siria", "Somalia", "Sri Lanka", "Suazilandia", "Sudáfrica", "Sudán", "Sudán del Sur", "Suecia", "Suiza", "Surinam", "Tailandia", "Tanzania", "Tayikistán", "Timor Oriental", "Togo", "Tonga", "Trinidad y Tobago", "Túnez", "Turkmenistán", "Turquía", "Tuvalu", "Ucrania", "Uganda", "Uruguay", "Uzbekistán", "Vanuatu", "Venezuela", "Vietnam", "Yemen", "Yibuti", "Zambia", "Zimbabue"]

def cargar_datos_productos():
    df = load_sheet(DATOS_SHEET)
    for c in COLS_DATOS:
        if c not in df.columns:
            df[c] = ""
    return df[COLS_DATOS].fillna("").astype(str).reset_index(drop=True)

@st.cache_data(ttl=120, show_spinner=False)
def _datos_productos_cache():
    return cargar_datos_productos()

def obtener_datos_producto(producto, df=None):
    if df is None:
        df = _datos_productos_cache()
    mask = df["producto"].str.strip().str.lower() == str(producto).strip().lower()
    return df[mask].iloc[0].to_dict() if mask.any() else None

def datos_completos(d):
    """Devuelve d solo si tiene origen y ambas fechas válidas (MM/AAAA)."""
    if not d:
        return None
    if (d.get("origen", "").strip()
            and re.match(REGEX_FECHA, d.get("fecha_produccion", "").strip())
            and re.match(REGEX_FECHA, d.get("fecha_vencimiento", "").strip())):
        return d
    return None

def guardar_datos_productos(items, lote):
    """items: [{'producto','origen','fecha_prod','fecha_venc'}]. Actualiza datos y 'último uso'."""
    df = cargar_datos_productos()
    hoy = datetime.now().strftime("%d/%m/%Y")
    for it in items:
        reg = {"producto": it["producto"], "origen": it["origen"],
               "fecha_produccion": it["fecha_prod"], "fecha_vencimiento": it["fecha_venc"],
               "ultimo_uso": hoy, "ultimo_lote": lote}
        mask = df["producto"].str.strip().str.lower() == it["producto"].strip().lower()
        if mask.any():
            for k, v in reg.items():
                df.loc[mask, k] = v
        else:
            df = pd.concat([df, pd.DataFrame([reg])], ignore_index=True)
    save_sheet(DATOS_SHEET, df)
    _datos_productos_cache.clear()

def registrar_uso(items, lote):
    """Callback de los botones de descarga: guarda los datos usados como 'último uso'."""
    if items:
        guardar_datos_productos(items, lote)

def generar_etiqueta_producto(producto, lote, precio_bruto, precio_kg, fecha_prod, fecha_venc, origen):
    """Busca la etiqueta base en Drive, la modifica y devuelve (bytes, error)."""
    archivo = buscar_etiqueta_base(producto)
    if not archivo:
        return None, "no se encontró etiqueta base en Google Drive"
    buf = descargar_etiqueta_drive(archivo["id"])
    if not buf:
        return None, "no se pudo descargar la etiqueta base"
    data = modificar_etiqueta_ezpx(buf, lote, precio_bruto, precio_kg, fecha_prod, fecha_venc, origen)
    if not data:
        return None, "no se pudo modificar la etiqueta"
    return data, None

def productos_en_historial():
    df = load_sheet(HISTORIAL_SHEET)
    nombres = {}
    if df.empty or "productos_json" not in df.columns:
        return []
    for pj in df["productos_json"]:
        try:
            for p in json.loads(pj):
                n = str(p.get("Productos", "")).strip()
                if n:
                    nombres.setdefault(n.lower(), n)
        except Exception:
            continue
    return sorted(nombres.values(), key=str.lower)

def normalizar_precio(v):
    """Acepta '1500', '$1.500', 1500 -> '$1.500'. Devuelve None si no es válido."""
    if v is None or (not isinstance(v, str) and pd.isna(v)):
        return None
    if isinstance(v, (int, float)):
        return format_clp(int(v))
    digitos = re.sub(r"\D", "", str(v))
    return format_clp(int(digitos)) if digitos else None

# ==================== APP ====================
st.title("📋 Lista de Precios")

tab2, tab1, tab3 = st.tabs(["🔎 Consultar Lotes", "🆕 Nueva Recepción", "🔐 Administración"])

# ====================== TAB 1 - NUEVA RECEPCIÓN ======================
with tab1:
    st.subheader("📤 Nueva Recepción de Precios")

    if "uploader_counter" not in st.session_state:
        st.session_state.uploader_counter = 0

    if st.session_state.get("recepcion_exitosa", False):
        st.success("✅ Recepción exitosa")
        st.session_state.recepcion_exitosa = False

    uploaded_file = st.file_uploader(
        "Subir archivo Excel de precios (Modulo/ 3. 📊Calculadora de Precios -> 💰 Precios de Venta -> 📤Descargar Excel de Precios de Venta)",
        type=["xlsx", "xls"],
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
            st.dataframe(styled_df, use_container_width=True, hide_index=True)
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
        opciones = [""] + lotes_disponibles
        lote_seleccionado = st.selectbox(
            "Seleccionar Lote",
            opciones,
            index=0,
            placeholder="Escribe para buscar o selecciona un lote..."
        )
        if st.button("Cargar Información del Lote", type="primary", disabled=not lote_seleccionado):
            data = cargar_historial_lote(lote_seleccionado)
            if data:
                st.session_state.data_lote = data
                st.success(f"✅ Lote cargado: **{data['lote']}**")
            else:
                st.error(f"No se encontró información para el lote '{lote_seleccionado}'")

        if st.session_state.get("data_lote"):
            data = st.session_state.data_lote
            st.success(f"**Lote:** {data['lote']} | **Fecha llegada:** {format_fecha_chilena(data['fecha_llegada'])}")
            df_lote_completo = pd.DataFrame(data["productos"])

            # ---------- ETIQUETAS DE TODO EL LOTE (ZIP) ----------
            with st.expander("📦 Descargar todas las etiquetas del lote (productos con datos guardados)"):
                dp = _datos_productos_cache()
                listos, sin_datos = [], []
                for _, r in df_lote_completo.iterrows():
                    d = datos_completos(obtener_datos_producto(r["Productos"], dp))
                    if d:
                        listos.append((r, d))
                    else:
                        sin_datos.append(str(r["Productos"]))
                st.write(f"✅ **{len(listos)}** con datos listos · ⚠️ **{len(sin_datos)}** sin datos completos")
                if listos:
                    st.dataframe(pd.DataFrame([{
                        "Producto": r["Productos"], "Origen": d["origen"],
                        "Producción": d["fecha_produccion"], "Vencimiento": d["fecha_vencimiento"],
                        "Último uso": d["ultimo_uso"] or "—", "Último lote": d["ultimo_lote"] or "—",
                    } for r, d in listos]), hide_index=True, use_container_width=True)
                if sin_datos:
                    st.caption("Sin datos (se guardan solos al generar su etiqueta una vez, o en 🔐 Administración): "
                               + ", ".join(sin_datos))

                if st.button("⚡ Generar ZIP de etiquetas", disabled=not listos, key="btn_zip"):
                    barra = st.progress(0.0, text="Generando etiquetas...")
                    zbuf, usados, fallidos, nombres = BytesIO(), [], [], set()
                    with zipfile.ZipFile(zbuf, "w", zipfile.ZIP_DEFLATED) as z:
                        for i, (r, d) in enumerate(listos):
                            prod = str(r["Productos"])
                            b, err = generar_etiqueta_producto(
                                prod, data["lote"], r["Precio Venta Bruto"], r["Precio KG"],
                                d["fecha_produccion"], d["fecha_vencimiento"], d["origen"])
                            if b:
                                nombre = f"{prod} - {data['lote']}.ezpx".replace("/", "-")
                                if nombre in nombres:
                                    nombre = f"{prod} ({i + 1}) - {data['lote']}.ezpx".replace("/", "-")
                                nombres.add(nombre)
                                z.writestr(nombre, b)
                                usados.append({"producto": prod, "origen": d["origen"],
                                               "fecha_prod": d["fecha_produccion"],
                                               "fecha_venc": d["fecha_vencimiento"]})
                            else:
                                fallidos.append(f"{prod}: {err}")
                            barra.progress((i + 1) / len(listos), text=f"Generando etiquetas... {i + 1}/{len(listos)}")
                    barra.empty()
                    st.session_state.zip_lote = {"lote": data["lote"], "bytes": zbuf.getvalue(),
                                                 "usados": usados, "fallidos": fallidos}

                zip_info = st.session_state.get("zip_lote")
                if zip_info and zip_info["lote"] == data["lote"]:
                    for f in zip_info["fallidos"]:
                        st.warning(f"⚠️ {f}")
                    if zip_info["usados"]:
                        st.download_button(
                            f"⬇️ Descargar ZIP ({len(zip_info['usados'])} etiquetas)",
                            data=zip_info["bytes"],
                            file_name=f"Etiquetas - {data['lote']}.zip",
                            mime="application/zip",
                            on_click=registrar_uso,
                            args=(zip_info["usados"], data["lote"]),
                            type="primary",
                            key="dl_zip",
                        )

            # ---------- TABLA Y ETIQUETA INDIVIDUAL ----------
            df_lote = df_lote_completo.copy()
            opciones_f = ["Todos"] + sorted([str(p) for p in df_lote["Productos"].tolist() if p is not None and str(p).strip() != ""])
            busqueda_precio = st.selectbox("🔎 Filtrar producto", options=opciones_f, key="busqueda_precio")
            if busqueda_precio != "Todos":
                df_lote = df_lote[df_lote["Productos"] == busqueda_precio].reset_index(drop=True)

            styled_lote = style_by_numero(df_lote)
            selection = st.dataframe(
                styled_lote,
                use_container_width=True,
                hide_index=True,
                on_select="rerun",
                selection_mode="single-row",
                key=f"tabla_lote_{lote_seleccionado}"
            )

            if len(selection["selection"]["rows"]) > 0:
                fila = df_lote.iloc[selection["selection"]["rows"][0]]
                producto = str(fila["Productos"])
                precio_bruto = fila["Precio Venta Bruto"]
                precio_kg = fila["Precio KG"]

                if st.session_state.get("current_producto") != producto:
                    st.session_state.current_producto = producto
                    for k in ("etiqueta_bytes", "etiqueta_info", "etiqueta_firma", "etiqueta_error"):
                        st.session_state.pop(k, None)

                st.info(f"**Producto seleccionado:** {producto}")

                guardado = datos_completos(obtener_datos_producto(producto))
                usar_guardados = False
                if guardado:
                    st.success(
                        f"📌 **Datos guardados** · Origen: **{guardado['origen']}** · "
                        f"Producción: **{guardado['fecha_produccion']}** · Vencimiento: **{guardado['fecha_vencimiento']}**\n\n"
                        f"🕒 Último uso: {guardado['ultimo_uso'] or 'sin registro'}"
                        + (f" (lote {guardado['ultimo_lote']})" if guardado['ultimo_lote'] else "")
                    )
                    modo = st.radio(
                        "¿Qué datos usar para la etiqueta?",
                        ["♻️ Reutilizar datos guardados", "✏️ Ingresar datos nuevos"],
                        horizontal=True, key=f"modo_{producto}"
                    )
                    usar_guardados = modo.startswith("♻️")
                else:
                    st.caption("Este producto aún no tiene datos guardados: ingrésalos una vez y quedarán para los próximos lotes.")
                modo_actual = "guardados" if usar_guardados else "nuevos"

                if usar_guardados:
                    # Generación automática: sin pasos intermedios
                    firma = (producto, data["lote"], precio_bruto, precio_kg,
                             guardado["origen"], guardado["fecha_produccion"], guardado["fecha_vencimiento"])
                    if st.session_state.get("etiqueta_firma") != firma:
                        with st.spinner("Generando etiqueta con los datos guardados..."):
                            bytes_et, err = generar_etiqueta_producto(
                                producto, data["lote"], precio_bruto, precio_kg,
                                guardado["fecha_produccion"], guardado["fecha_vencimiento"], guardado["origen"])
                        st.session_state.etiqueta_firma = firma
                        st.session_state.etiqueta_bytes = bytes_et
                        st.session_state.etiqueta_error = err
                        st.session_state.etiqueta_info = {
                            "producto": producto, "modo": "guardados",
                            "item": {"producto": producto, "origen": guardado["origen"],
                                     "fecha_prod": guardado["fecha_produccion"],
                                     "fecha_venc": guardado["fecha_vencimiento"]},
                        } if bytes_et else None
                    if st.session_state.get("etiqueta_error"):
                        st.error(f"❌ {st.session_state.etiqueta_error} ('{producto}')")
                else:
                    st.subheader("📋 Datos para la etiqueta")
                    origen = st.selectbox("País de Origen", options=["Seleccione un país..."] + PAISES, index=0, key=f"origen_{producto}")
                    fecha_prod = st.date_input("Fecha de Producción", value=None, key=f"prod_{producto}", format="DD/MM/YYYY")
                    fecha_venc = st.date_input("Fecha de Vencimiento", value=None, key=f"venc_{producto}", format="DD/MM/YYYY")
                    disabled = (origen == "Seleccione un país..." or fecha_prod is None or fecha_venc is None)

                    if st.button("🚀 Confirmar y generar etiqueta", type="primary", disabled=disabled):
                        fp, fv = fecha_prod.strftime("%m/%Y"), fecha_venc.strftime("%m/%Y")
                        with st.spinner("Generando etiqueta..."):
                            bytes_et, err = generar_etiqueta_producto(
                                producto, data["lote"], precio_bruto, precio_kg, fp, fv, origen)
                        st.session_state.etiqueta_firma = None  # invalida la generación automática
                        st.session_state.etiqueta_bytes = bytes_et
                        st.session_state.etiqueta_error = None
                        if bytes_et:
                            st.session_state.etiqueta_info = {
                                "producto": producto, "modo": "nuevos",
                                "item": {"producto": producto, "origen": origen, "fecha_prod": fp, "fecha_venc": fv},
                            }
                        else:
                            st.session_state.etiqueta_info = None
                            st.error(f"❌ {err} ('{producto}')")

                info = st.session_state.get("etiqueta_info")
                if (info and info["producto"] == producto and info["modo"] == modo_actual
                        and st.session_state.get("etiqueta_bytes")):
                    st.success("✅ Etiqueta lista")
                    st.download_button(
                        label="⬇️ Descargar etiqueta",
                        data=st.session_state.etiqueta_bytes,
                        file_name=f"{producto} - {data['lote']}.ezpx".replace("/", "-"),
                        mime="application/xml",
                        on_click=registrar_uso,
                        args=([info["item"]], data["lote"]),
                        type="primary",
                        key=f"dl_{producto}",
                    )
    else:
        st.info("Aún no hay lotes recepcionados")

# ====================== TAB 3 - ADMINISTRACIÓN ======================
with tab3:
    st.subheader("🔐 Administración")

    if not st.session_state.get("admin_ok", False):
        with st.form("login_admin"):
            pwd = st.text_input("Contraseña", type="password")
            entrar = st.form_submit_button("Ingresar", type="primary")
        if entrar:
            if hmac.compare_digest(pwd.encode(), _admin_password().encode()):
                st.session_state.admin_ok = True
                st.rerun()
            else:
                st.error("Contraseña incorrecta")
    else:
        if st.button("🔒 Cerrar sesión", key="adm_logout"):
            st.session_state.admin_ok = False
            st.rerun()

        if st.session_state.get("adm_msg"):
            st.success(st.session_state.pop("adm_msg"))

        sub_a, sub_b = st.tabs(["✏️ Editar registros de lotes", "🏷️ Datos de productos (origen y fechas)"])

        # ---------- EDITAR LOTES ----------
        with sub_a:
            df_hist = load_sheet(HISTORIAL_SHEET)
            if df_hist.empty:
                st.info("No hay registros para editar")
            else:
                lote_edit = st.selectbox("Lote a editar", [""] + listar_lotes_recibidos(), key="adm_lote",
                                         placeholder="Selecciona un lote...")
                if lote_edit:
                    idx = df_hist.index[df_hist["lote"].astype(str) == str(lote_edit)][0]
                    fila_h = df_hist.loc[idx]
                    try:
                        prods = json.loads(fila_h["productos_json"])
                    except Exception:
                        prods = []
                    fecha_dt = pd.to_datetime(fila_h["fecha_llegada"], errors="coerce")
                    fecha_actual = fecha_dt.date() if pd.notna(fecha_dt) else datetime.now().date()

                    c1, c2 = st.columns(2)
                    nuevo_nombre = c1.text_input("Lote", value=str(fila_h["lote"]), key=f"adm_nom_{lote_edit}")
                    fecha_ll = c2.date_input("Fecha de llegada a Bodega", value=fecha_actual,
                                             format="DD/MM/YYYY", key=f"adm_fecha_{lote_edit}")

                    st.caption("Edita, agrega o elimina filas. Los precios pueden escribirse como 1500 o $1.500. "
                               "Las filas sin producto se ignoran.")
                    df_prod = pd.DataFrame(prods, columns=["N°", "Productos", "Precio Venta Bruto", "Precio KG"])
                    editado = st.data_editor(
                        df_prod, num_rows="dynamic", hide_index=True, use_container_width=True,
                        key=f"adm_ed_{lote_edit}",
                        column_config={
                            "N°": st.column_config.NumberColumn("N°", step=1),
                            "Productos": st.column_config.TextColumn("Productos"),
                            "Precio Venta Bruto": st.column_config.TextColumn("Precio Venta Bruto"),
                            "Precio KG": st.column_config.TextColumn("Precio KG"),
                        },
                    )

                    if st.button("💾 Guardar cambios", type="primary", key="adm_guardar"):
                        nombre = nuevo_nombre.strip()
                        errores, registros = [], []
                        if not nombre:
                            errores.append("El nombre del lote no puede estar vacío.")
                        elif nombre != str(lote_edit) and nombre in df_hist["lote"].astype(str).tolist():
                            errores.append(f"Ya existe un lote llamado '{nombre}'.")
                        for _, row in editado.iterrows():
                            prod = "" if pd.isna(row["Productos"]) else str(row["Productos"]).strip()
                            if not prod:
                                continue
                            pb, pk = normalizar_precio(row["Precio Venta Bruto"]), normalizar_precio(row["Precio KG"])
                            if pb is None or pk is None:
                                errores.append(f"Precio inválido o vacío en '{prod}'.")
                                continue
                            n = pd.to_numeric(row["N°"], errors="coerce")
                            registros.append({"N°": 0 if pd.isna(n) else int(n), "Productos": prod,
                                              "Precio Venta Bruto": pb, "Precio KG": pk})
                        if not registros and not errores:
                            errores.append("El lote debe tener al menos un producto.")
                        if errores:
                            for e in errores:
                                st.error(e)
                        else:
                            df_hist.loc[idx, "lote"] = nombre
                            df_hist.loc[idx, "fecha_llegada"] = fecha_ll.strftime("%Y-%m-%d")
                            df_hist.loc[idx, "productos_json"] = json.dumps(registros, ensure_ascii=False)
                            save_sheet(HISTORIAL_SHEET, df_hist)
                            st.session_state.pop("data_lote", None)
                            st.session_state.adm_msg = f"✅ Lote '{nombre}' actualizado"
                            st.rerun()

                    with st.expander("🗑️ Eliminar este lote"):
                        st.warning("Se elimina el registro completo del historial. Esta acción no se puede deshacer.")
                        confirmar = st.checkbox("Confirmo que quiero eliminar este lote", key=f"adm_conf_{lote_edit}")
                        if st.button("Eliminar lote", disabled=not confirmar, key="adm_del"):
                            save_sheet(HISTORIAL_SHEET, df_hist.drop(index=idx).reset_index(drop=True))
                            st.session_state.pop("data_lote", None)
                            st.session_state.adm_msg = f"🗑️ Lote '{lote_edit}' eliminado"
                            st.rerun()

        # ---------- DATOS DE PRODUCTOS ----------
        with sub_b:
            st.caption("Origen y fechas que se usan automáticamente al generar etiquetas. Fechas en formato MM/AAAA. "
                       "'Último uso' y 'Último lote' se actualizan solos al descargar una etiqueta.")
            if st.button("➕ Importar productos de los lotes recibidos", key="adm_import"):
                df_dp = cargar_datos_productos()
                existentes = set(df_dp["producto"].str.strip().str.lower())
                nuevos = [n for n in productos_en_historial() if n.lower() not in existentes]
                if nuevos:
                    df_dp = pd.concat([df_dp, pd.DataFrame({"producto": nuevos})], ignore_index=True).fillna("")
                    save_sheet(DATOS_SHEET, df_dp)
                    _datos_productos_cache.clear()
                    st.session_state.dp_ver = st.session_state.get("dp_ver", 0) + 1
                    st.session_state.adm_msg = f"✅ Se agregaron {len(nuevos)} productos nuevos"
                else:
                    st.session_state.adm_msg = "No hay productos nuevos para importar"
                st.rerun()

            df_dp = cargar_datos_productos()
            ed = st.data_editor(
                df_dp, num_rows="dynamic", hide_index=True, use_container_width=True,
                key=f"adm_datos_prod_{st.session_state.get('dp_ver', 0)}",
                column_config={
                    "producto": st.column_config.TextColumn("Producto", required=True),
                    "origen": st.column_config.SelectboxColumn("País de origen", options=PAISES),
                    "fecha_produccion": st.column_config.TextColumn("Producción (MM/AAAA)", validate=REGEX_FECHA),
                    "fecha_vencimiento": st.column_config.TextColumn("Vencimiento (MM/AAAA)", validate=REGEX_FECHA),
                    "ultimo_uso": st.column_config.TextColumn("Último uso", disabled=True),
                    "ultimo_lote": st.column_config.TextColumn("Último lote", disabled=True),
                },
            )

            if st.button("💾 Guardar datos de productos", type="primary", key="adm_guardar_dp"):
                ed2 = ed.fillna("").astype(str)
                ed2["producto"] = ed2["producto"].str.strip()
                ed2 = ed2[ed2["producto"] != ""]

                def _fecha_ok(x):
                    x = x.strip()
                    return x == "" or re.match(REGEX_FECHA, x) is not None

                malas = ed2[~(ed2["fecha_produccion"].map(_fecha_ok) & ed2["fecha_vencimiento"].map(_fecha_ok))]["producto"].tolist()
                if malas:
                    st.error("Fechas inválidas (usa MM/AAAA) en: " + ", ".join(malas))
                else:
                    ed2 = ed2[~ed2["producto"].str.lower().duplicated(keep="last")]
                    save_sheet(DATOS_SHEET, ed2[COLS_DATOS].reset_index(drop=True))
                    _datos_productos_cache.clear()
                    st.session_state.dp_ver = st.session_state.get("dp_ver", 0) + 1
                    st.session_state.adm_msg = "✅ Datos de productos guardados"
                    st.rerun()

st.caption("Desarrollado para La Trilla con ❤️ • Datos en Google Sheets v1.3 (administración, datos de productos y etiquetas en lote)")
