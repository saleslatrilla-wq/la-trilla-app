import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
import copy

st.set_page_config(page_title="La Trilla - Envasado", layout="wide", page_icon="🥜")

# ===================== CONEXIÓN A GOOGLE SHEETS =====================
spreadsheet = st.session_state.google_spreadsheet

def load_sheet(sheet_name):
    try:
        worksheet = spreadsheet.worksheet(sheet_name)
        data = worksheet.get_all_records()
        return pd.DataFrame(data)
    except:
        return pd.DataFrame()

def save_sheet(sheet_name, df):
    try:
        worksheet = spreadsheet.worksheet(sheet_name)
    except:
        worksheet = spreadsheet.add_worksheet(title=sheet_name, rows="2000", cols="50")
    worksheet.clear()
    df = df.replace([np.nan, np.inf, -np.inf], [None, None, None])
    if not df.empty:
        worksheet.update([df.columns.values.tolist()] + df.values.tolist())
    else:
        worksheet.update([[]])

# ===================== LOGIN ====================
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

def login():
    if not st.session_state.logged_in:
        st.title("🔐 Envasado")
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

# ===================== PERMISOS =====================
username = st.session_state.username
is_admin = username in ["Joannahan", "Daniela", "Gonzalo"]
is_vendedor = username == "vendedor"

# ===================== INICIALIZACIÓN SESSION STATE =====================
if "cajas_asignadas" not in st.session_state:
    st.session_state.cajas_asignadas = {}
if "progreso_anterior" not in st.session_state:
    st.session_state.progreso_anterior = {}
if "vista_seleccionada" not in st.session_state:
    st.session_state.vista_seleccionada = "BOX warehouse" if is_vendedor else "Lista de Prioridad"
if "busqueda_forzada" not in st.session_state:
    st.session_state.busqueda_forzada = ""

# ===================== BOX WAREHOUSE =====================
df_boxes = load_sheet("box_warehouse")
boxes_almacen = {}

if df_boxes.empty or len(df_boxes) < 48:
    st.info("Inicializando las 48 cajas en Google Sheets...")
    data = []
    for letra in "ABCDEFGHIJKL":
        for num in range(1, 5):
            caja = f"{letra}{num}"
            boxes_almacen[caja] = {}
            data.append({"caja": caja, "producto": "", "unidades": 0})
    save_sheet("box_warehouse", pd.DataFrame(data))
    st.success("✅ Cajas A1-L4 creadas")
    st.rerun()
else:
    for _, row in df_boxes.iterrows():
        caja = str(row["caja"])
        producto = str(row.get("producto", ""))
        unidades = int(row.get("unidades", 0))
        if caja not in boxes_almacen:
            boxes_almacen[caja] = {}
        if producto and producto.strip():
            boxes_almacen[caja][producto] = unidades

# ===================== PROGRESO =====================
progreso_actual = {}
df_progreso = load_sheet("progreso_envasado")
if not df_progreso.empty:
    for _, row in df_progreso.iterrows():
        key = str(row["key"])
        progreso_actual[key] = {
            "unidades_real": int(row.get("unidades_real", 0)),
            "completado": bool(row.get("completado", False))
        }

movimientos_log = load_sheet("movimientos_log").to_dict('records') if not load_sheet("movimientos_log").empty else []

def guardar_boxes():
    data = []
    for caja, contenido in boxes_almacen.items():
        for producto, unidades in contenido.items():
            data.append({"caja": caja, "producto": producto, "unidades": unidades})
    save_sheet("box_warehouse", pd.DataFrame(data))

def guardar_progreso():
    data = []
    for key, info in progreso_actual.items():
        data.append({"key": key, "unidades_real": info["unidades_real"], "completado": info["completado"]})
    save_sheet("progreso_envasado", pd.DataFrame(data))

def guardar_movimiento(usuario, producto, movimientos, detalle_final):
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
    save_sheet("movimientos_log", pd.DataFrame(movimientos_log))

# ===================== CARGAR PLAN =====================
df_plan = load_sheet("Plan_Envasado_Actual") if not is_vendedor else pd.DataFrame()
if not df_plan.empty:
    numeric_cols = ['Stock MP (kg)', 'Kg Usados', 'Unidades a Envasar', 'Cobertura (semanas)', 'Stock Actual', 'Formato en Kg', '% Stock Asignado']
    for col in numeric_cols:
        if col in df_plan.columns:
            df_plan[col] = pd.to_numeric(df_plan[col], errors='coerce').fillna(0)

# ===================== BARRA LATERAL =====================
st.sidebar.header("Filtros y Navegación")

if is_vendedor:
    vista = "BOX warehouse"
    st.sidebar.info("👤 Modo Vendedor\nSolo tienes acceso a **BOX Warehouse**")
else:
    opciones = ["Dashboard", "Lista de Prioridad", "BOX warehouse"]
    if is_admin:
        opciones.append("Gráfico Diario")
    vista = st.sidebar.selectbox("Vista", opciones, index=1 if st.session_state.vista_seleccionada == "Lista de Prioridad" else 0)

busqueda = st.sidebar.text_input("Buscar producto", value=st.session_state.busqueda_forzada)
if busqueda != st.session_state.busqueda_forzada:
    st.session_state.busqueda_forzada = ""

st.sidebar.divider()
if st.sidebar.button("🚪 Cerrar Sesión"):
    st.session_state.logged_in = False
    st.session_state.username = ""
    st.rerun()

# ===================== GENERAR PLAN LÓGICO (CORREGIDO) =====================
if is_admin:
    with st.sidebar.expander("📤 Generar Plan de Envasado (Admin)"):
        st.caption("Sube ProductInputData.xlsx → Plan optimizado")
        uploaded_file = st.file_uploader("📁 Selecciona ProductInputData.xlsx", type=["xlsx"], key="input_uploader")
        if uploaded_file and st.button("🚀 Procesar y Generar Plan LÓGICO", type="primary", use_container_width=True):
            with st.spinner("Calculando plan correcto..."):
                try:
                    df = pd.read_excel(uploaded_file, sheet_name="Datos_Limpios")
                    df.columns = [str(col).strip() for col in df.columns]

                    df['demanda_semanal'] = df[['Unidades Vendida Semana 1',
                                               'Unidades Vendida Semana 2',
                                               'Unidades Vendida Semana 3',
                                               'Unidades Vendida Semana 4']].sum(axis=1) / 4.0

                    df['Producto MP'] = df['Producto']
                    # CORRECCIÓN: Stock MP solo se calcula correctamente
                    df['Stock MP (kg)'] = df['Stock Actual'] * df['Formato en Kg'].where(df['Formato en Kg'] > 1, 1)

                    plan_rows = []
                    for prod, group in df.groupby("Producto MP"):
                        stock_total_kg = group["Stock MP (kg)"].sum()
                        total_demand_kg = (group['demanda_semanal'] * group['Formato en Kg']).sum()
                        scale_factor = stock_total_kg / total_demand_kg if total_demand_kg > 0 else 1.0

                        for _, row in group.iterrows():
                            unidades_plan = round(row['demanda_semanal'] * scale_factor)
                            kg_usados = round(unidades_plan * row['Formato en Kg'], 2)
                            stock_final = round(stock_total_kg - kg_usados, 2)
                            pct_asignado = round((kg_usados / stock_total_kg * 100), 1) if stock_total_kg > 0 else 0
                            cobertura = round(stock_total_kg / (kg_usados / 4), 2) if kg_usados > 0 else 0

                            plan_rows.append({
                                "Producto MP": prod,
                                "Stock MP (kg)": round(stock_total_kg, 2),
                                "Formato": row["Formato"],
                                "Stock Actual": row["Stock Actual"],
                                "Unidades a Envasar": int(unidades_plan),
                                "Kg Usados": kg_usados,
                                "% Stock Asignado": pct_asignado,
                                "Stock Final": stock_final,
                                "Cobertura (semanas)": cobertura
                            })

                    df_plan_final = pd.DataFrame(plan_rows)
                    save_sheet("Plan_Envasado_Actual", df_plan_final)
                    st.success("✅ ¡Plan LÓGICO corregido y guardado!")
                    st.rerun()
                except Exception as e:
                    st.error(f"❌ Error: {str(e)}")

# ===================== PROGRESO GENERAL =====================
if not is_vendedor and not df_plan.empty:
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
    st.subheader("📊 Progreso General del Plan")
    col1, col2, col3 = st.columns([1, 4, 1])
    with col1:
        st.metric("Completado", f"{porcentaje_global:.1f}%")
    with col2:
        st.progress(min(porcentaje_global / 100, 1.0))
    with col3:
        st.metric("Productos", f"{productos_completados}/{df_plan['Producto MP'].nunique()}")
    st.caption(f"**{total_kg_real:.1f} kg** envasados de **{total_kg_plan:.1f} kg** planificados")
    st.divider()

# ===================== VISTAS =====================
df_grouped = df_plan.groupby("Producto MP") if not df_plan.empty else None

if is_vendedor or vista == "BOX warehouse":
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
                guardar_boxes()
                st.success("✅ Cambios guardados")
                st.rerun()

else:
    if vista == "Dashboard":
        for producto, group in df_grouped:
            if st.session_state.busqueda_forzada and producto != st.session_state.busqueda_forzada:
                continue
            elif busqueda and busqueda.lower() not in producto.lower():
                continue

            stock_mp = group["Stock MP (kg)"].iloc[0]
            key_producto = producto.replace(" ", "_").replace(".", "")

            with st.container(border=True):
                col_title, col_status = st.columns([4, 1])
                with col_title:
                    st.subheader(producto)
                    st.markdown(f"**Stock MP: {stock_mp} kg**", unsafe_allow_html=True)
                with col_status:
                    st.checkbox("✅ Marcar producto como COMPLETO",
                                value=all(progreso_actual.get(f"{key_producto}_{row['Formato']}", {}).get("completado", False) for _, row in group.iterrows()),
                                key=f"chk_prod_{key_producto}")

                for _, row in group.sort_values(by="Formato").iterrows():
                    formato = row["Formato"]
                    unidades_plan = int(row["Unidades a Envasar"])
                    key_formato = f"{key_producto}_{formato}"

                    if key_formato not in progreso_actual:
                        progreso_actual[key_formato] = {"unidades_real": 0, "completado": False}

                    default_caja = st.session_state.cajas_asignadas.get(key_formato, list(boxes_almacen.keys())[0] if boxes_almacen else "A1")

                    with st.container(border=True):
                        st.markdown(f"**{formato}** → Plan: **{unidades_plan}** unidades")
                        col_caja, col_num, col_status = st.columns([2, 3, 1])
                        with col_caja:
                            caja_seleccionada = st.selectbox(
                                "📦 Caja destino",
                                options=list(boxes_almacen.keys()),
                                index=list(boxes_almacen.keys()).index(default_caja),
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
                                if producto not in boxes_almacen.get(caja_destino, {}):
                                    boxes_almacen[caja_destino] = {}
                                if producto not in boxes_almacen[caja_destino]:
                                    boxes_almacen[caja_destino][producto] = 0
                                boxes_almacen[caja_destino][producto] += delta
                    if hay_movimiento:
                        guardar_movimiento(st.session_state.username, producto, movimientos, detalle_final)
                        guardar_boxes()
                        st.success(f"✅ ¡Guardado! Se registraron {len(movimientos)} movimiento(s)")
                    st.session_state.progreso_anterior = copy.deepcopy(progreso_actual)
                    guardar_progreso()
                    st.session_state.vista_seleccionada = "Lista de Prioridad"
                    st.session_state.busqueda_forzada = ""
                    st.rerun()

    elif vista == "Lista de Prioridad":
        st.subheader("📋 Lista de Prioridad (qué envasar primero)")
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
        else:
            st.info("No hay productos para mostrar")

    elif vista == "Gráfico Diario" and is_admin:
        st.subheader("📈 Progreso Diario de Producción")
        if movimientos_log:
            df_hist = pd.DataFrame(movimientos_log)
            if not df_hist.empty:
                df_hist["Fecha"] = pd.to_datetime(df_hist["fecha"])
                df_hist = df_hist.groupby("Fecha").size().reset_index(name="Unidades Envasadas")
                st.line_chart(df_hist.set_index("Fecha")["Unidades Envasadas"])
        else:
            st.info("Aún no hay datos.")

st.caption("Desarrollado para La Trilla con ❤️ • Plan Lógico v1.5 - Datos en Google Sheets")
