import streamlit as st
import pandas as pd
import json
import io
from datetime import datetime, date
import copy
import numpy as np

st.set_page_config(page_title="La Trilla - Envasado", layout="wide", page_icon="🥜")

# ===================== CONEXIÓN A GOOGLE SHEETS =====================
spreadsheet = st.session_state.google_spreadsheet

def load_sheet(sheet_name):
    try:
        worksheet = spreadsheet.worksheet(sheet_name)
        data = worksheet.get_all_records()
        df = pd.DataFrame(data)
        numeric_cols = ['Stock MP (kg)', 'Stock Actual', 'Unidades a Envasar', 'Kg Usados', 
                        'Stock Final', 'Cobertura (semanas)', 'Formato en Kg', 'demanda_semanal']
        for col in numeric_cols:
            if col in df.columns:
                df[col] = df[col].apply(lambda x: str(x).replace(',', '.') if isinstance(x, str) else x)
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
        return df
    except:
        return pd.DataFrame()

def save_sheet(sheet_name, df):
    try:
        worksheet = spreadsheet.worksheet(sheet_name)
    except:
        worksheet = spreadsheet.add_worksheet(title=sheet_name, rows="2000", cols="50")
    worksheet.clear()
    if not df.empty:
        worksheet.update([df.columns.values.tolist()] + df.values.tolist())
    else:
        worksheet.update([[]])

# ==================== SISTEMA DE LOGIN ====================
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
    st.session_state.username = ""

USUARIOS = {
    "Joannahan": "q1w2e3r4Trilla.",
    "Daniela": "q1w2e3r4Trilla.",
    "Gonzalo": "q1w2e3r4Trilla.",
    "envasados": "envasados253",
    "vendedor": "vendedor253",
}

is_admin = st.session_state.get("username") in ["Joannahan", "Daniela", "Gonzalo"]
is_vendedor = st.session_state.get("username") == "vendedor"

def login():
    if not st.session_state.logged_in:
        st.title("🔐 La Trilla - Envasado")
        st.subheader("Iniciar Sesión")
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            username = st.text_input("Usuario", placeholder="Ingresa tu usuario")
            password = st.text_input("Contraseña", type="password", placeholder="Ingresa tu contraseña")
            if st.button("Ingresar", type="primary", use_container_width=True):
                if username in USUARIOS and USUARIOS[username] == password:
                    st.session_state.logged_in = True
                    st.session_state.username = username
                    st.success(f"✅ Bienvenido, {username}!")
                    st.rerun()
                else:
                    st.error("❌ Usuario o contraseña incorrecta")
        return False
    return True

if not login():
    st.stop()

# ==================== LOG DE MOVIMIENTOS (Google Sheets) ====================
def cargar_movimientos():
    df = load_sheet("movimientos_log")
    return df.to_dict('records') if not df.empty else []

def guardar_movimiento(usuario, producto, movimientos, detalle_final):
    movimientos_log = cargar_movimientos()
    ahora = datetime.now()
    registro = {
        "fecha": ahora.strftime("%Y-%m-%d"),
        "hora": ahora.strftime("%H:%M:%S"),
        "timestamp": ahora.strftime("%Y-%m-%d %H:%M:%S"),
        "usuario": usuario,
        "producto": producto,
        "movimientos": movimientos,
        "detalle_final": detalle_final
    }
    movimientos_log.append(registro)
    df_log = pd.DataFrame(movimientos_log)
    save_sheet("movimientos_log", df_log)

# ==================== BOX WAREHOUSE (Google Sheets) ====================
def cargar_boxes():
    df = load_sheet("box_warehouse")
    boxes = {}
    # Inicializar todas las cajas vacías por defecto
    for letra in "ABCDEFGHIJKL":
        for num in range(1, 5):
            boxes[f"{letra}{num}"] = {}
    if df.empty or "caja" not in df.columns:
        return boxes
    for _, row in df.iterrows():
        caja = str(row.get("caja", "")).strip()
        producto = str(row.get("producto", "")).strip()
        unidades = row.get("unidades", 0)
        if caja and producto:
            if caja not in boxes:
                boxes[caja] = {}
            try:
                boxes[caja][producto] = int(unidades)
            except (ValueError, TypeError):
                boxes[caja][producto] = 0
    return boxes

def guardar_boxes(boxes):
    rows = []
    for caja, contenido in boxes.items():
        if contenido:
            for producto, unidades in contenido.items():
                rows.append({"caja": caja, "producto": producto, "unidades": unidades})
        else:
            # Guardar la caja vacía igual para que no desaparezca
            rows.append({"caja": caja, "producto": "", "unidades": 0})
    df = pd.DataFrame(rows)
    save_sheet("box_warehouse", df)
boxes_almacen = cargar_boxes()
# ==================== PROGRESO Y HISTORIAL (Google Sheets) ====================
PROGRESO_SHEET = "progreso_envasado"
HISTORIAL_SHEET = "historial_produccion"

def cargar_progreso():
    df = load_sheet(PROGRESO_SHEET)
    if df.empty:
        return {}
    progreso = {}
    for _, row in df.iterrows():
        key = row["key"]
        progreso[key] = {
            "unidades_real": int(row.get("unidades_real", 0)),
            "completado": bool(row.get("completado", False))
        }
    return progreso

def guardar_progreso(progreso):
    rows = []
    for key, data in progreso.items():
        rows.append({
            "key": key,
            "unidades_real": data["unidades_real"],
            "completado": data["completado"]
        })
    df = pd.DataFrame(rows)
    save_sheet(PROGRESO_SHEET, df)

def cargar_historial():
    df = load_sheet(HISTORIAL_SHEET)
    if df.empty:
        return {}
    return dict(zip(df["fecha"], df["unidades"]))

def guardar_historial(historial):
    df = pd.DataFrame(list(historial.items()), columns=["fecha", "unidades"])
    save_sheet(HISTORIAL_SHEET, df)

progreso_actual = cargar_progreso()
if "progreso_anterior" not in st.session_state:
    st.session_state.progreso_anterior = copy.deepcopy(progreso_actual)
if "cajas_asignadas" not in st.session_state:
    st.session_state.cajas_asignadas = {}

historial = cargar_historial()

# ==================== CARGA DE PLAN (solo si no es vendedor) ====================
if not is_vendedor:
    df_plan = load_sheet("Plan_Envasado_Actual")
else:
    df_plan = pd.DataFrame()

# ==================== ESTILOS Y SESIÓN ====================
st.markdown("""
<style>
    .stNumberInput input { inputmode: decimal !important; -webkit-appearance: none; -moz-appearance: textfield; }
</style>
""", unsafe_allow_html=True)

if "vista_seleccionada" not in st.session_state:
    st.session_state.vista_seleccionada = "Lista de Prioridad"
if "busqueda_forzada" not in st.session_state:
    st.session_state.busqueda_forzada = ""

# ==================== BARRA LATERAL ====================
st.sidebar.header("Filtros y Navegación")
opciones_vista = ["Dashboard", "Lista de Prioridad", "Gráfico Diario", "BOX warehouse"]
if is_vendedor:
    opciones_vista = ["BOX warehouse"]  # Solo BOX para vendedor
vista = st.sidebar.selectbox("Vista", opciones_vista, index=0)
busqueda = st.sidebar.text_input("Buscar producto", value=st.session_state.busqueda_forzada)
if busqueda != st.session_state.busqueda_forzada:
    st.session_state.busqueda_forzada = ""
st.sidebar.divider()
if st.sidebar.button("🚪 Cerrar Sesión"):
    st.session_state.logged_in = False
    st.session_state.username = ""
    st.rerun()

# ==================== GENERADOR DE PLAN (SOLO ADMIN) ====================
if is_admin:
    with st.sidebar.expander("📤 Generar Plan de Envasado (Admin)"):
        st.caption("Sube ProductInputData.xlsx → genera Plan_Envasado_Actual")
        uploaded_file = st.file_uploader("📁 Selecciona ProductInputData.xlsx", type=["xlsx"], key="input_uploader")
        if uploaded_file is not None:
            if st.button("🚀 Procesar y Generar Plan de Envasado", type="primary", use_container_width=True):
                with st.spinner("Analizando datos y generando el plan..."):
                    try:
                        df = pd.read_excel(uploaded_file, sheet_name="Datos_Limpios")
                        df.columns = [str(col).strip() for col in df.columns]
                        numeric_cols = ['Stock Actual', 'Formato en Kg'] + [col for col in df.columns if "Semana" in str(col)]
                        for col in numeric_cols:
                            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
                        df['demanda_semanal'] = df[[col for col in df.columns if "Semana" in str(col)]].sum(axis=1) / 4.0
                        
                        # === LÓGICA EXACTA DEL CÓDIGO ORIGINAL ===
                        df_mp = df[(df['Grupo_ID'].notna()) & (df['Sub_ID'].astype(str).str.contains(r'\.0$|^\d+$', regex=True))].copy()
                        df_env = df[(df['Grupo_ID'].notna()) & (~df['Sub_ID'].astype(str).str.contains(r'\.0$|^\d+$', regex=True))].copy()
                        df_mp['Stock_kg'] = df_mp['Stock Actual'] * df_mp['Formato en Kg']
                        
                        resultados = []
                        for grupo in sorted(df_mp['Grupo_ID'].unique()):
                            mp_rows = df_mp[df_mp['Grupo_ID'] == grupo]
                            stock_total_mp = float(mp_rows['Stock_kg'].sum())
                            if not mp_rows.empty:
                                mp_principal = mp_rows.loc[mp_rows['Stock_kg'].idxmax()]
                                nombre_mp = str(mp_principal['Producto']).strip()
                            else:
                                continue
                            if stock_total_mp <= 0.5: 
                                continue
                            envs = df_env[(df_env['Grupo_ID'] == grupo) & (df_env['demanda_semanal'] > 0.1)].copy()
                            if envs.empty: 
                                continue
                            formatos = []
                            total_demanda_kg = 0.0
                            for _, row in envs.iterrows():
                                peso = float(row['Formato en Kg'])
                                if peso <= 0: continue
                                demanda = float(row['demanda_semanal'])
                                stock_actual = float(row['Stock Actual'])
                                nombre_fmt = str(row['Formato']).strip()
                                formatos.append({'formato': nombre_fmt, 'peso': peso, 'demanda': demanda, 'stock_actual': stock_actual})
                                total_demanda_kg += demanda * peso
                            if not formatos: continue
                            cobertura_objetivo = stock_total_mp / total_demanda_kg
                            distribucion = {}
                            kg_usados = 0.0
                            for fmt in formatos:
                                stock_final_target = fmt['demanda'] * cobertura_objetivo
                                unidades = max(0, int(np.ceil(stock_final_target - fmt['stock_actual'])))
                                kg_necesario = unidades * fmt['peso']
                                if kg_usados + kg_necesario > stock_total_mp:
                                    unidades = int((stock_total_mp - kg_usados) / fmt['peso'])
                                    kg_necesario = unidades * fmt['peso']
                                stock_final = fmt['stock_actual'] + unidades
                                cobertura = round(stock_final / fmt['demanda'], 2)
                                distribucion[fmt['formato']] = {
                                    'stock_actual': int(fmt['stock_actual']),
                                    'unidades_envasar': unidades,
                                    'kg_usados': round(kg_necesario, 2),
                                    'stock_final': int(stock_final),
                                    'cobertura_semanas': cobertura,
                                    'peso': fmt['peso']
                                }
                                kg_usados += kg_necesario
                            kg_sobrante = stock_total_mp - kg_usados
                            if kg_sobrante > 0.5:
                                small_first = sorted(formatos, key=lambda x: x['peso'])
                                for fmt in small_first:
                                    extra = int(kg_sobrante / fmt['peso'])
                                    if extra > 0:
                                        key = fmt['formato']
                                        distribucion[key]['unidades_envasar'] += extra
                                        distribucion[key]['kg_usados'] += round(extra * fmt['peso'], 2)
                                        distribucion[key]['stock_final'] += extra
                                        distribucion[key]['cobertura_semanas'] = round(distribucion[key]['stock_final'] / fmt['demanda'], 2)
                                        kg_sobrante -= extra * fmt['peso']
                                    if kg_sobrante < fmt['peso']: break
                            utilizacion = round((stock_total_mp - kg_sobrante) / stock_total_mp * 100, 1)
                            resultados.append({
                                'producto': nombre_mp,
                                'stock_mp_kg': round(stock_total_mp, 2),
                                'distribucion': distribucion,
                                'kg_sobrante': round(kg_sobrante, 2),
                                'utilizacion_%': utilizacion,
                                'cobertura_objetivo': round(cobertura_objetivo, 2)
                            })
                        
                        # Construir DataFrame final (exacto al original)
                        filas = []
                        for r in resultados:
                            distrib_items = sorted(r['distribucion'].items(), key=lambda x: x[1]['peso'])
                            for fmt_name, d in distrib_items:
                                porcentaje = round((d['kg_usados'] / r['stock_mp_kg'] * 100), 1) if r['stock_mp_kg'] > 0 else 0
                                filas.append({
                                    'Producto MP': r['producto'],
                                    'Stock MP (kg)': r['stock_mp_kg'],
                                    'Formato': fmt_name,
                                    'Stock Actual': d['stock_actual'],
                                    'Unidades a Envasar': d['unidades_envasar'],
                                    'Kg Usados': d['kg_usados'],
                                    '% Stock Asignado': f"{porcentaje}%",
                                    'Stock Final': d['stock_final'],
                                    'Cobertura (semanas)': d['cobertura_semanas']
                                })
                        df_final = pd.DataFrame(filas)
                        
                        # Guardar en Google Sheets (reemplaza el Excel)
                        save_sheet("Plan_Envasado_Actual", df_final)
                        
                        st.success("✅ ¡Plan generado correctamente! Guardado en Google Sheets como 'Plan_Envasado_Actual'")
                        st.dataframe(df_final, use_container_width=True)
                        st.cache_data.clear()
                        st.session_state.vista_seleccionada = "Lista de Prioridad"
                        st.rerun()
                    except Exception as e:
                        st.error(f"❌ Error al procesar: {str(e)}")

st.sidebar.caption("Versión 1.6 - Google Sheets • Lógica 100% original")

# ==================== PROGRESO GENERAL DEL PLAN ====================
if not df_plan.empty:
    total_kg_plan = df_plan["Kg Usados"].sum()
    total_kg_real = 0
    productos_completados = 0
    for producto, group in df_plan.groupby("Producto MP"):
        key_producto = producto.replace(" ", "_").replace(".", "")
        kg_real_prod = 0
        completado = True
        for _, row in group.iterrows():
            key_formato = f"{key_producto}_{row['Formato']}"
            unidades_real = progreso_actual.get(key_formato, {}).get("unidades_real", 0)
            if row["Unidades a Envasar"] > 0:
                kg_real_prod += unidades_real * row["Kg Usados"] / row["Unidades a Envasar"]
            if not progreso_actual.get(key_formato, {}).get("completado", False):
                completado = False
        total_kg_real += kg_real_prod
        if completado:
            productos_completados += 1
    porcentaje_global = (total_kg_real / total_kg_plan * 100) if total_kg_plan > 0 else 0
else:
    total_kg_plan = 0
    total_kg_real = 0
    porcentaje_global = 0
    productos_completados = 0

st.subheader("📊 Progreso General del Plan")
col1, col2, col3 = st.columns([1, 4, 1])
with col1:
    st.metric("Completado", f"{porcentaje_global:.1f}%")
with col2:
    st.progress(min(porcentaje_global / 100, 1.0))
with col3:
    st.metric("Productos", f"{productos_completados}/{len(df_plan['Producto MP'].unique()) if not df_plan.empty else 0}")
st.caption(f"**{total_kg_real:.1f} kg** envasados de **{total_kg_plan:.1f} kg** planificados")
st.divider()

# ==================== VISTAS ====================
if not df_plan.empty:
    df_grouped = df_plan.groupby("Producto MP")
else:
    df_grouped = pd.DataFrame().groupby("Producto MP")  # vacío seguro

if vista == "Dashboard":
    for producto, group in df_grouped:
        if st.session_state.busqueda_forzada and producto != st.session_state.busqueda_forzada:
            continue
        elif busqueda and busqueda.lower() not in producto.lower():
            continue

        stock_mp = group["Stock MP (kg)"].iloc[0] if not group.empty else 0
        key_producto = producto.replace(" ", "_").replace(".", "")

        with st.container(border=True):
            col_title, col_status = st.columns([4, 1])
            with col_title:
                st.subheader(producto)
                st.markdown(f"**Stock MP: {stock_mp} kg**", unsafe_allow_html=True)
            with col_status:
                producto_completado = st.checkbox("✅ Marcar producto como COMPLETO",
                    value=all(progreso_actual.get(f"{key_producto}_{row['Formato']}", {}).get("completado", False) for _, row in group.iterrows()),
                    key=f"chk_prod_{key_producto}")

            for _, row in group.sort_values(by="Formato").iterrows():
                formato = row["Formato"]
                unidades_plan = int(row["Unidades a Envasar"])
                key_formato = f"{key_producto}_{formato}"

                if key_formato not in progreso_actual:
                    progreso_actual[key_formato] = {"unidades_real": 0, "completado": False}

                if key_formato not in st.session_state.cajas_asignadas:
                    st.session_state.cajas_asignadas[key_formato] = "A1"

                with st.container(border=True):
                    st.markdown(f"**{formato}** → Plan: **{unidades_plan}** unidades")
                    col_caja, col_num, col_status = st.columns([2, 3, 1])
                    with col_caja:
                        caja_seleccionada = st.selectbox(
                            "📦 Caja destino",
                            options=list(boxes_almacen.keys()),
                            index=list(boxes_almacen.keys()).index(st.session_state.cajas_asignadas[key_formato]),
                            key=f"caja_{key_formato}",
                            label_visibility="collapsed"
                        )
                        st.session_state.cajas_asignadas[key_formato] = caja_seleccionada
                    with col_num:
                        unidades_real = st.number_input(
                            "Unidades ya envasadas",
                            min_value=0,
                            value=progreso_actual[key_formato]["unidades_real"],
                            step=1,
                            key=f"num_{key_formato}",
                            label_visibility="collapsed"
                        )
                        progreso_actual[key_formato]["unidades_real"] = unidades_real
                    with col_status:
                        auto_completado = unidades_real >= unidades_plan
                        progreso_actual[key_formato]["completado"] = auto_completado
                        if auto_completado:
                            st.success("✅ Completado")
                        else:
                            st.write("")

            total_plan = sum(int(row["Unidades a Envasar"]) for _, row in group.iterrows() if row["Unidades a Envasar"] > 0)
            total_real = sum(progreso_actual.get(f"{key_producto}_{row['Formato']}", {}).get("unidades_real", 0) for _, row in group.iterrows())
            progreso_general = min(total_real / total_plan, 1.0) if total_plan > 0 else 0
            st.progress(progreso_general)
            st.caption(f"{progreso_general*100:.0f}% completado del producto")

            if st.button("Guardar registro", type="primary", use_container_width=True, key=f"btn_{key_producto}"):
                movimientos = []
                detalle_final = {}
                hay_movimiento = False
                for _, row in group.iterrows():
                    key_formato = f"{key_producto}_{row['Formato']}"
                    formato = row['Formato']
                    nueva_cantidad = progreso_actual.get(key_formato, {}).get("unidades_real", 0)
                    anterior_cantidad = st.session_state.progreso_anterior.get(key_formato, {}).get("unidades_real", 0)
                    detalle_final[formato] = nueva_cantidad
                    delta = nueva_cantidad - anterior_cantidad
                    if delta != 0:
                        hay_movimiento = True
                        tipo = "más" if delta > 0 else "menos"
                        cantidad = abs(delta)
                        movimientos.append({"formato": formato, "tipo": tipo, "cantidad": int(cantidad)})
                        if delta > 0:
                            caja_destino = st.session_state.cajas_asignadas[key_formato]
                            if producto not in boxes_almacen[caja_destino]:
                                boxes_almacen[caja_destino][producto] = 0
                            boxes_almacen[caja_destino][producto] += delta
                            guardar_boxes(boxes_almacen)
                if hay_movimiento:
                    guardar_movimiento(st.session_state.username, producto, movimientos, detalle_final)
                    st.success(f"✅ ¡Guardado! Se registraron {len(movimientos)} movimiento(s)")
                else:
                    st.warning("⚠️ No hubo cambios.")
                st.session_state.progreso_anterior = copy.deepcopy(progreso_actual)
                guardar_progreso(progreso_actual)
                st.session_state.vista_seleccionada = "Lista de Prioridad"
                st.session_state.busqueda_forzada = ""
                st.rerun()

    guardar_progreso(progreso_actual)

elif vista == "Lista de Prioridad":
    st.subheader("📋 Lista de Prioridad (qué envasar primero)")
    if df_plan.empty:
        st.info("No hay plan cargado. Usa el generador en la barra lateral.")
    else:
        prioridad_data = []
        for producto, group in df_grouped:
            if busqueda and busqueda.lower() not in producto.lower():
                continue
            stock_mp = group["Stock MP (kg)"].iloc[0]
            kg_real_total = 0
            coberturas = []
            for _, row in group.iterrows():
                key_formato = f"{producto.replace(' ', '_').replace('.', '')}_{row['Formato']}"
                unidades_real = progreso_actual.get(key_formato, {}).get("unidades_real", 0)
                if row["Unidades a Envasar"] > 0:
                    kg_real_total += unidades_real * row["Kg Usados"] / row["Unidades a Envasar"]
                coberturas.append(row["Cobertura (semanas)"])
            avance = kg_real_total / stock_mp * 100 if stock_mp > 0 else 0
            cobertura_promedio = sum(coberturas) / len(coberturas) if coberturas else 0
            if cobertura_promedio < 2:
                prioridad_emoji = "🔥🔥🔥🔥🔥"
            elif cobertura_promedio < 4:
                prioridad_emoji = "🔥🔥🔥🔥"
            elif cobertura_promedio < 6:
                prioridad_emoji = "🔥🔥🔥"
            elif cobertura_promedio < 9:
                prioridad_emoji = "🔥🔥"
            else:
                prioridad_emoji = "🔥"
            prioridad_data.append({
                "Prioridad": prioridad_emoji,
                "Producto": producto,
                "Stock MP (kg)": round(stock_mp, 1),
                "Avance (%)": round(avance, 1),
                "Cobertura Promedio": round(cobertura_promedio, 2),
                "Kg Pendientes": round(stock_mp - kg_real_total, 1)
            })
        df_prioridad = pd.DataFrame(prioridad_data)
        if not df_prioridad.empty:
            df_prioridad = df_prioridad.sort_values(by=["Cobertura Promedio", "Kg Pendientes"], ascending=[True, False])
            selected = st.dataframe(df_prioridad, use_container_width=True, hide_index=True, on_select="rerun", selection_mode="single-row")
            if len(selected["selection"]["rows"]) > 0:
                fila = selected["selection"]["rows"][0]
                producto_seleccionado = df_prioridad.iloc[fila]["Producto"]
                st.success(f"✅ Seleccionaste: **{producto_seleccionado}**")
                if st.button("→ Ir al Dashboard y filtrar este producto", type="primary"):
                    st.session_state.vista_seleccionada = "Dashboard"
                    st.session_state.busqueda_forzada = producto_seleccionado
                    st.rerun()

elif vista == "BOX warehouse":
    st.subheader("📦 BOX Warehouse")
    tab_ver, tab_modificar = st.tabs(["📋 Ver Contenido", "✏️ Modificar Cajas"])
    with tab_ver:
        busqueda_box = st.text_input("🔎 Buscar producto", placeholder="Nombre del producto...", key="busqueda_box")
        if busqueda_box:
            resultados = []
            for caja, contenido in boxes_almacen.items():
                for producto, unidades in contenido.items():
                    if busqueda_box.lower() in producto.lower():
                        resultados.append({"Caja": caja, "Producto": producto, "Unidades": unidades})
            if resultados:
                st.dataframe(pd.DataFrame(resultados), use_container_width=True, hide_index=True)
            else:
                st.warning("No se encontró el producto en ninguna caja.")
        else:
            st.write("**Contenido de las cajas:**")
            cols = st.columns(4)
            for i, (caja, contenido) in enumerate(boxes_almacen.items()):
                with cols[i % 4]:
                    with st.container(border=True):
                        st.markdown(f"**{caja}**")
                        if contenido:
                            for producto, unidades in contenido.items():
                                nombre_limpio = producto.split(" Saco ")[0].strip()
                                st.markdown(f"**{nombre_limpio}**<br><small>{unidades} uds</small>", unsafe_allow_html=True)
                        else:
                            st.caption("Vacía")
    with tab_modificar:
        st.write("**Modificar contenido de cajas**")
        caja_seleccionada = st.selectbox("Seleccionar caja", options=list(boxes_almacen.keys()), key="mod_caja")
        if caja_seleccionada in boxes_almacen:
            contenido = boxes_almacen[caja_seleccionada]
            nuevo_contenido = dict(contenido)
            if contenido:
                for producto, unidades in list(contenido.items()):
                    nueva_unidad = st.number_input(f"{producto}", value=unidades, step=1, min_value=0, key=f"mod_unid_{caja_seleccionada}_{producto}")
                    nuevo_contenido[producto] = nueva_unidad
            else:
                st.caption("Caja vacía")
            if st.button("Guardar cambios", type="primary", use_container_width=True):
                boxes_almacen[caja_seleccionada] = {p: u for p, u in nuevo_contenido.items() if u > 0}
                guardar_boxes(boxes_almacen)
                st.success("✅ Cambios guardados")
                st.rerun()

else:  # Gráfico Diario
    st.subheader("📈 Progreso Diario de Producción")
    if historial:
        df_hist = pd.DataFrame(list(historial.items()), columns=["Fecha", "Unidades Envasadas"])
        df_hist["Fecha"] = pd.to_datetime(df_hist["Fecha"])
        df_hist = df_hist.sort_values("Fecha")
        st.line_chart(df_hist.set_index("Fecha")["Unidades Envasadas"], use_container_width=True)
    else:
        st.info("Aún no hay datos.")

st.caption("Desarrollado con ❤️ para La Trilla • Versión Google Sheets 1.6 - Lógica 100% original del código de 602 líneas")
