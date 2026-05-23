import streamlit as st
import pandas as pd
from io import BytesIO
import re
import math

st.set_page_config(page_title="Precios App", layout="wide", initial_sidebar_state="collapsed")
st.title("📊 Automatizador de Precios")

# ===================== CONEXIÓN A GOOGLE SHEETS =====================
spreadsheet = st.session_state.google_spreadsheet

def load_sheet(sheet_name):
    try:
        worksheet = spreadsheet.worksheet(sheet_name)
        data = worksheet.get_all_records()
        return pd.DataFrame(data)
    except:
        worksheet = spreadsheet.add_worksheet(title=sheet_name, rows="1000", cols="20")
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

def save_df(df, sheet_name):
    save_sheet(sheet_name, df)

def load_df(sheet_name, columns=None):
    df = load_sheet(sheet_name)
    if df.empty and columns is not None:
        df = pd.DataFrame(columns=columns)
    return df

def save_reception(df, metadata):
    save_sheet("last_reception_df", df)
    meta_df = pd.DataFrame([metadata])
    save_sheet("last_reception_metadata", meta_df)

def load_reception():
    df = load_sheet("last_reception_df")
    meta_df = load_sheet("last_reception_metadata")
    if not meta_df.empty:
        metadata = meta_df.iloc[0].to_dict()
        return df, metadata
    return None, None

tab1, tab2, tab3 = st.tabs(["📥 Recepción", "⚙️ Configuración", "💰 Precios de Venta"])

# ==================================== PESTAÑA 1: RECEPCIÓN ====================================
with tab1:
    st.subheader("Datos de Recepción")
    if "reception_df" not in st.session_state:
        df_saved, meta_saved = load_reception()
        if df_saved is not None:
            st.session_state.reception_df = df_saved
            st.session_state.reception_metadata = meta_saved

    uploaded_file = st.file_uploader("Subir reporte de factura (xls / html)", type=["xls", "xlsx", "html"])

    def clean_cost(x):
        if pd.isna(x):
            return 0.0
        if isinstance(x, (int, float)):
            return float(x)
        x = str(x).strip()
        # Corrección principal: maneja tanto 31.932,77 como 31932.77
        x = x.replace(".", "").replace(",", ".")
        try:
            return float(x)
        except:
            return 0.0

    @st.cache_data
    def parse_reporte(file):
        content = file.getvalue()
        filename = file.name.lower()
        full_df = None
        if filename.endswith(('.xlsx', '.xls')):
            try:
                full_df = pd.read_excel(BytesIO(content), sheet_name=0, header=6)
            except:
                tables = pd.read_html(BytesIO(content), encoding="utf-8")
                full_df = tables[0]
        else:
            tables = pd.read_html(BytesIO(content), encoding="utf-8")
            full_df = tables[0]

        full_df.columns = [str(col).strip().replace("\n", " ").replace(" ", " ") for col in full_df.columns]

        metadata = {"Tipo Documento": "No encontrado", "Fecha": "No encontrada", "Proveedor": "No encontrado", "Usuario": "No encontrado"}
        try:
            if filename.endswith('.xlsx'):
                if len(full_df) > 0:
                    first = full_df.iloc[0]
                    metadata["Fecha"] = str(first.get("Fecha", "")).strip()
                    metadata["Usuario"] = str(first.get("Usuario", "")).strip()
                    doc = str(first.get("Documento de Recepción", "")).strip()
                    metadata["Tipo Documento"] = doc if "Factura" in doc else f"Factura Nº {doc}"
                    prov = str(first.get("Nota", "")).strip()
                    metadata["Proveedor"] = prov
            else:
                text = content.decode('utf-8', errors='replace')
                doc_match = re.search(r'Factura\s*Nº?\s*(\d+)', text, re.I)
                if doc_match: metadata["Tipo Documento"] = f"Factura Nº {doc_match.group(1)}"
                fecha_match = re.search(r'(\d{1,2}/\d{1,2}/\d{4})', text)
                if fecha_match: metadata["Fecha"] = fecha_match.group(1)
                user_match = re.search(r'="?\d{1,2}/\d{1,2}/\d{4}"?[^<]*</td>\s*<td>([^<]+?)</td>', text, re.I)
                if user_match: metadata["Usuario"] = user_match.group(1).strip()
                prov_match = re.search(r'(NAMA INTERNACIONAL SA / \w+|INNOVACIONES, PRODUCTOS Y SERVICIOS SPA / \w+)', text)
                if prov_match: metadata["Proveedor"] = prov_match.group(0)
        except:
            pass

        producto_col = next((col for col in full_df.columns if "Producto" in col), None)
        serie_col = next((col for col in full_df.columns if "Serie" in col), None)
        costo_col = next((col for col in full_df.columns if "Costo Neto Unitario" in col), None)
        cantidad_col = next((col for col in full_df.columns if "Cantidad" in col), None)

        df = full_df[[producto_col, serie_col or "Serie", costo_col, cantidad_col]].copy()
        df.columns = ["Producto", "Serie/Lote", "Costo Neto Unitario", "Cantidad"]
        df = df.dropna(subset=["Producto"]).reset_index(drop=True)
        df["Costo Neto Unitario"] = df["Costo Neto Unitario"].apply(clean_cost)
        df["Cantidad"] = pd.to_numeric(df["Cantidad"], errors="coerce").fillna(0)
        df["Costo Neto Total"] = df["Costo Neto Unitario"] * df["Cantidad"]
        df.insert(0, "N°", range(1, len(df) + 1))
        return df, metadata

    if uploaded_file is not None:
        df_import, metadata = parse_reporte(uploaded_file)
        st.session_state.reception_df = df_import
        st.session_state.reception_metadata = metadata
        save_reception(df_import, metadata)
        st.success("✅ Archivo cargado y procesado correctamente")
    elif "reception_df" in st.session_state:
        df_import = st.session_state.reception_df
        metadata = st.session_state.reception_metadata
        st.info("📂 Mostrando datos de la última recepción guardada")
    else:
        st.info("Sube un archivo para comenzar")
        st.stop()

    st.markdown(f"## {metadata['Tipo Documento']}")
    col1, col2 = st.columns([1, 1])
    with col1:
        st.markdown(f"**🏬 Proveedor:** {metadata['Proveedor']}")
        total_factura = df_import["Costo Neto Total"].sum()
        def format_clp(v):
            txt = f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            return f"${txt}"
        st.markdown(f"**💰 Total Neto Factura:** {format_clp(total_factura)}")
    with col2:
        st.markdown(f"**📅 Fecha de Recepción:** {metadata['Fecha']}")
        st.markdown(f"**👤 Usuario:** {metadata['Usuario']}")

    df_show = df_import.copy()
    df_show["Costo Neto Unitario"] = df_show["Costo Neto Unitario"].apply(format_clp)
    df_show["Costo Neto Total"] = df_show["Costo Neto Total"].apply(format_clp)
    st.dataframe(df_show[["N°", "Producto", "Serie/Lote", "Cantidad", "Costo Neto Unitario", "Costo Neto Total"]],
                 use_container_width=True, hide_index=True,
                 column_config={col: st.column_config.TextColumn(width="auto") for col in df_show.columns})

# ==================================== PESTAÑA 2: CONFIGURACIÓN ====================================
with tab2:
    st.subheader("⚙️ Configuración")
    subtab1, subtab2, subtab3, subtab4 = st.tabs(["📋 Identificador", "🔧 Backend Precios", "📦 Costos", "📏 Márgenes"])

    # ===================== IDENTIFICADOR =====================
    with subtab1:
        st.subheader("📋 Identificador")
        st.caption("Base completa de productos y sus subproductos")
        if "bdb_full_df" not in st.session_state:
            st.session_state.bdb_full_df = load_df("productos")
        if st.session_state.bdb_full_df.empty:
            st.warning("⚠️ La pestaña 'productos' está vacía.")

        if not st.session_state.bdb_full_df.empty:
            search_term = st.text_input("🔎 Buscar producto", "", placeholder="Escribe nombre del producto...", key="search_identificador")
            df_full = st.session_state.bdb_full_df.copy().fillna("")
            df_to_show = df_full
            if search_term and search_term.strip():
                search = search_term.strip()
                df_to_show = df_full[df_full["PRODUCTO"].astype(str).str.contains(search, case=False, na=False)]
            original_indices_bdb = df_to_show.index.tolist()
            df_to_show = df_to_show.reset_index(drop=True)
            for col in df_to_show.columns:
                if col != "Seleccionar":
                    df_to_show[col] = df_to_show[col].astype(str)
            df_to_show["Seleccionar"] = False

            edit_bdb = st.toggle("Desbloquear edición de Identificador", value=False, key="toggle_bdb")

            if st.button("➕ Agregar nuevo producto", key="add_bdb"):
                st.session_state.show_add_bdb = True
            if st.session_state.get("show_add_bdb", False):
                with st.form("form_add_bdb"):
                    st.write("**Nuevo producto**")
                    new_data = {}
                    cols = st.session_state.bdb_full_df.columns.tolist()
                    for col in cols:
                        new_data[col] = st.text_input(col, "")
                    if st.form_submit_button("Guardar nuevo"):
                        new_row = pd.DataFrame([new_data])
                        st.session_state.bdb_full_df = pd.concat([st.session_state.bdb_full_df, new_row], ignore_index=True)
                        save_df(st.session_state.bdb_full_df, "productos")
                        st.success("Producto agregado")
                        st.session_state.show_add_bdb = False
                        st.rerun()

            edited = st.data_editor(
                df_to_show,
                use_container_width=True,
                num_rows="fixed",
                disabled=not edit_bdb,
                column_config={
                    **{col: st.column_config.TextColumn(width="auto") for col in df_to_show.columns if col != "Seleccionar"},
                    "Seleccionar": st.column_config.CheckboxColumn("Seleccionar", default=False, width="small")
                },
                hide_index=True,
                key="bdb_editor"
            )
            if edit_bdb:
                if st.button("💾 Guardar cambios Identificador", type="primary", key="save_bdb"):
                    cols_bdb = [c for c in edited.columns if c != "Seleccionar"]
                    edited_data = edited[cols_bdb].copy()
                    for i, orig_idx in enumerate(original_indices_bdb):
                        for col in cols_bdb:
                            st.session_state.bdb_full_df.at[orig_idx, col] = edited_data.at[i, col]
                    save_df(st.session_state.bdb_full_df, "productos")
                    st.success("✅ Cambios guardados correctamente en Google Sheets")
                    st.rerun()

            if st.button("🗑️ Eliminar seleccionados", type="secondary", key="del_btn_bdb"):
                selected = edited[edited["Seleccionar"] == True]
                if len(selected) > 0:
                    st.session_state["confirm_del_bdb"] = True
                    st.session_state["indices_to_del_bdb"] = df_to_show.index[edited["Seleccionar"] == True].tolist()
                else:
                    st.warning("No hay filas seleccionadas")
                if st.session_state.get("confirm_del_bdb", False):
                    st.warning(f"¿Estás seguro de eliminar permanentemente **{len(st.session_state['indices_to_del_bdb'])} fila(s)** de la base de datos?")
                    col1, col2 = st.columns(2)
                    if col1.button("Sí, eliminar", key="yes_del_bdb"):
                        st.session_state.bdb_full_df = st.session_state.bdb_full_df.drop(st.session_state["indices_to_del_bdb"]).reset_index(drop=True)
                        save_df(st.session_state.bdb_full_df, "productos")
                        del st.session_state["confirm_del_bdb"]
                        del st.session_state["indices_to_del_bdb"]
                        st.rerun()
                    if col2.button("Cancelar", key="cancel_del_bdb"):
                        del st.session_state["confirm_del_bdb"]
                        del st.session_state["indices_to_del_bdb"]
                        st.rerun()

    # ===================== BACKEND PRECIOS =====================
    with subtab2:
        st.subheader("🔧 Backend Precios")
        st.caption("Utilidad se calcula automáticamente si existe COD en Márgenes.")
        if "backend_df" not in st.session_state:
            st.session_state.backend_df = load_df("backend_precios", columns=["COD", "Producto", "Insumos", "MOD y MOI"])
        if "margenes_df" not in st.session_state:
            st.session_state.margenes_df = load_df("margenes", columns=["Código", "Margen (%)", "Nota"])

        def calcular_utilidad(cod, margenes_df):
            if pd.isna(cod) or str(cod).strip() == "":
                return 0.0
            match = margenes_df[margenes_df["Código"].astype(str).str.strip() == str(cod).strip()]
            return float(match.iloc[0]["Margen (%)"]) if len(match) > 0 else 0.0

        df_full = st.session_state.backend_df.copy()
        df_full["COD"] = df_full["COD"].astype(str)
        if "Utilidad" not in df_full.columns:
            df_full["Utilidad"] = 0.0
        df_display = df_full.copy()
        df_display["Utilidad"] = df_display["COD"].apply(lambda x: calcular_utilidad(x, st.session_state.margenes_df))
        cols_order = ["COD", "Utilidad", "Producto", "Insumos", "MOD y MOI"]
        df_display = df_display[cols_order]

        search_term = st.text_input("🔎 Buscar producto", "", placeholder="Escribe nombre del producto...", key="search_backend")
        df_to_show = df_display
        if search_term and search_term.strip():
            search = search_term.strip()
            df_to_show = df_display[df_display["Producto"].astype(str).str.contains(search, case=False, na=False)]
        # Guardar índices originales ANTES del reset para poder mapear de vuelta correctamente
        original_indices_backend = df_to_show.index.tolist()
        df_to_show = df_to_show.reset_index(drop=True)
        for col in df_to_show.columns:
            if col != "Seleccionar" and col != "Utilidad":
                df_to_show[col] = df_to_show[col].astype(str)
        df_to_show["Seleccionar"] = False

        edit_backend = st.toggle("Desbloquear edición de Backend Precios", value=False, key="toggle_backend")

        if st.button("➕ Agregar nuevo backend", key="add_backend"):
            st.session_state.show_add_backend = True
        if st.session_state.get("show_add_backend", False):
            with st.form("form_add_backend"):
                st.write("**Nuevo registro Backend**")
                new_data = {}
                for col in cols_order:
                    new_data[col] = st.text_input(col, "")
                if st.form_submit_button("Guardar nuevo"):
                    new_row = pd.DataFrame([new_data])
                    st.session_state.backend_df = pd.concat([st.session_state.backend_df, new_row], ignore_index=True)
                    save_df(st.session_state.backend_df, "backend_precios")
                    st.success("Registro agregado")
                    st.session_state.show_add_backend = False
                    st.rerun()

        edited = st.data_editor(
            df_to_show,
            use_container_width=True,
            num_rows="fixed",
            disabled=not edit_backend,
            column_config={
                "COD": st.column_config.TextColumn(width="auto"),
                "Utilidad": st.column_config.NumberColumn(width="auto", format="%.1f%%"),
                "Producto": st.column_config.TextColumn(width="auto"),
                "Insumos": st.column_config.TextColumn(width="auto"),
                "MOD y MOI": st.column_config.TextColumn(width="auto"),
                "Seleccionar": st.column_config.CheckboxColumn("Seleccionar", default=False, width="small")
            },
            hide_index=True,
            key="backend_editor"
        )
        # BUG FIX: Guardar solo con botón explícito (no en cada render).
        # Usar columnas reales del backend (sin "Utilidad" que es calculada).
        cols_to_save_backend = ["COD", "Producto", "Insumos", "MOD y MOI"]
        if edit_backend:
            if st.button("💾 Guardar cambios Backend", type="primary", key="save_backend"):
                edited_data = edited[cols_to_save_backend].copy()
                # Mapear filas editadas de vuelta a los índices originales del df completo
                for i, orig_idx in enumerate(original_indices_backend):
                    for col in cols_to_save_backend:
                        st.session_state.backend_df.at[orig_idx, col] = edited_data.at[i, col]
                save_df(st.session_state.backend_df, "backend_precios")
                st.success("✅ Cambios guardados correctamente en Google Sheets")
                st.rerun()

        if st.button("🗑️ Eliminar seleccionados", type="secondary", key="del_btn_backend"):
            selected = edited[edited["Seleccionar"] == True]
            if len(selected) > 0:
                st.session_state["confirm_del_backend"] = True
                st.session_state["indices_to_del_backend"] = df_to_show.index[edited["Seleccionar"] == True].tolist()
            else:
                st.warning("No hay filas seleccionadas")
            if st.session_state.get("confirm_del_backend", False):
                st.warning(f"¿Estás seguro de eliminar permanentemente **{len(st.session_state['indices_to_del_backend'])} fila(s)** de la base de datos?")
                col1, col2 = st.columns(2)
                if col1.button("Sí, eliminar", key="yes_del_backend"):
                    st.session_state.backend_df = st.session_state.backend_df.drop(st.session_state["indices_to_del_backend"]).reset_index(drop=True)
                    save_df(st.session_state.backend_df, "backend_precios")
                    del st.session_state["confirm_del_backend"]
                    del st.session_state["indices_to_del_backend"]
                    st.rerun()
                if col2.button("Cancelar", key="cancel_del_backend"):
                    del st.session_state["confirm_del_backend"]
                    del st.session_state["indices_to_del_backend"]
                    st.rerun()

    # ===================== COSTOS =====================
    with subtab3:
        st.subheader("📦 Costos")
        cost_sub1, cost_sub2 = st.tabs(["Insumos", "MOD y MOI"])
        with cost_sub1:
            st.subheader("Insumos")
            if "insumos_df" not in st.session_state:
                df = load_df("insumos", columns=["Código", "Insumo", "Costo Neto Unitario"])
                df["Código"] = df["Código"].astype(str)
                df["Costo Neto Unitario"] = pd.to_numeric(df["Costo Neto Unitario"], errors="coerce").fillna(0)
                st.session_state.insumos_df = df
            search_term = st.text_input("🔎 Buscar insumo", "", placeholder="Escribe nombre del insumo...", key="search_insumos")
            df_full = st.session_state.insumos_df.copy().fillna("")
            df_to_show = df_full
            if search_term and search_term.strip():
                search = search_term.strip()
                df_to_show = df_full[df_full["Insumo"].astype(str).str.contains(search, case=False, na=False)]
            original_indices_insumos = df_to_show.index.tolist()
            df_to_show = df_to_show.reset_index(drop=True)
            df_to_show["Código"] = df_to_show["Código"].astype(str)
            df_to_show["Insumo"] = df_to_show["Insumo"].astype(str)
            df_to_show["Seleccionar"] = False
            edit_insumos = st.toggle("Desbloquear edición de Insumos", value=False, key="toggle_insumos")
            if st.button("➕ Agregar nuevo insumo", key="add_insumos"):
                st.session_state.show_add_insumos = True
            if st.session_state.get("show_add_insumos", False):
                with st.form("form_add_insumos"):
                    st.write("**Nuevo insumo**")
                    new_data = {}
                    for col in ["Código", "Insumo", "Costo Neto Unitario"]:
                        new_data[col] = st.text_input(col, "")
                    if st.form_submit_button("Guardar nuevo"):
                        new_row = pd.DataFrame([new_data])
                        st.session_state.insumos_df = pd.concat([st.session_state.insumos_df, new_row], ignore_index=True)
                        save_df(st.session_state.insumos_df, "insumos")
                        st.success("Insumo agregado")
                        st.session_state.show_add_insumos = False
                        st.rerun()
            edited = st.data_editor(
                df_to_show,
                use_container_width=True,
                num_rows="fixed",
                disabled=not edit_insumos,
                column_config={
                    "Código": st.column_config.TextColumn(width="auto"),
                    "Insumo": st.column_config.TextColumn(width="auto"),
                    "Costo Neto Unitario": st.column_config.NumberColumn(width="auto", format="$%.2f"),
                    "Seleccionar": st.column_config.CheckboxColumn("Seleccionar", default=False, width="small")
                },
                hide_index=True,
                key="insumos_editor"
            )
            if edit_insumos:
                if st.button("💾 Guardar cambios Insumos", type="primary", key="save_insumos"):
                    cols_ins = [c for c in edited.columns if c != "Seleccionar"]
                    edited_data = edited[cols_ins].copy()
                    for i, orig_idx in enumerate(original_indices_insumos):
                        for col in cols_ins:
                            st.session_state.insumos_df.at[orig_idx, col] = edited_data.at[i, col]
                    save_df(st.session_state.insumos_df, "insumos")
                    st.success("✅ Cambios guardados correctamente en Google Sheets")
                    st.rerun()
            if st.button("🗑️ Eliminar seleccionados", type="secondary", key="del_btn_insumos"):
                selected = edited[edited["Seleccionar"] == True]
                if len(selected) > 0:
                    st.session_state["confirm_del_insumos"] = True
                    st.session_state["indices_to_del_insumos"] = df_to_show.index[edited["Seleccionar"] == True].tolist()
                else:
                    st.warning("No hay filas seleccionadas")
                if st.session_state.get("confirm_del_insumos", False):
                    st.warning(f"¿Estás seguro de eliminar permanentemente **{len(st.session_state['indices_to_del_insumos'])} fila(s)** de la base de datos?")
                    col1, col2 = st.columns(2)
                    if col1.button("Sí, eliminar", key="yes_del_insumos"):
                        st.session_state.insumos_df = st.session_state.insumos_df.drop(st.session_state["indices_to_del_insumos"]).reset_index(drop=True)
                        save_df(st.session_state.insumos_df, "insumos")
                        del st.session_state["confirm_del_insumos"]
                        del st.session_state["indices_to_del_insumos"]
                        st.rerun()
                    if col2.button("Cancelar", key="cancel_del_insumos"):
                        del st.session_state["confirm_del_insumos"]
                        del st.session_state["indices_to_del_insumos"]
                        st.rerun()
        with cost_sub2:
            st.subheader("MOD y MOI")
            if "modmoi_df" not in st.session_state:
                df = load_df("modmoi", columns=["Código", "Tipo", "Costo Neto", "Descripción"])
                df["Código"] = df["Código"].astype(str)
                st.session_state.modmoi_df = df
            search_term = st.text_input("🔎 Buscar MOD/MOI", "", placeholder="Escribe descripción...", key="search_modmoi")
            df_full = st.session_state.modmoi_df.copy().fillna("")
            df_to_show = df_full
            if search_term and search_term.strip():
                search = search_term.strip()
                df_to_show = df_full[df_full["Descripción"].astype(str).str.contains(search, case=False, na=False)]
            original_indices_modmoi = df_to_show.index.tolist()
            df_to_show = df_to_show.reset_index(drop=True)
            df_to_show["Código"] = df_to_show["Código"].astype(str)
            df_to_show["Tipo"] = df_to_show["Tipo"].astype(str)
            df_to_show["Descripción"] = df_to_show["Descripción"].astype(str)
            df_to_show["Seleccionar"] = False
            edit_modmoi = st.toggle("Desbloquear edición de MOD y MOI", value=False, key="toggle_modmoi")
            if st.button("➕ Agregar nuevo MOD/MOI", key="add_modmoi"):
                st.session_state.show_add_modmoi = True
            if st.session_state.get("show_add_modmoi", False):
                with st.form("form_add_modmoi"):
                    st.write("**Nuevo MOD/MOI**")
                    new_data = {}
                    for col in ["Código", "Tipo", "Costo Neto", "Descripción"]:
                        new_data[col] = st.text_input(col, "")
                    if st.form_submit_button("Guardar nuevo"):
                        new_row = pd.DataFrame([new_data])
                        st.session_state.modmoi_df = pd.concat([st.session_state.modmoi_df, new_row], ignore_index=True)
                        save_df(st.session_state.modmoi_df, "modmoi")
                        st.success("Registro agregado")
                        st.session_state.show_add_modmoi = False
                        st.rerun()
            edited = st.data_editor(
                df_to_show,
                use_container_width=True,
                num_rows="fixed",
                disabled=not edit_modmoi,
                column_config={
                    "Código": st.column_config.TextColumn(width="auto"),
                    "Tipo": st.column_config.TextColumn(width="auto"),
                    "Costo Neto": st.column_config.NumberColumn(width="auto", format="$%.2f"),
                    "Descripción": st.column_config.TextColumn(width="auto"),
                    "Seleccionar": st.column_config.CheckboxColumn("Seleccionar", default=False, width="small")
                },
                hide_index=True,
                key="modmoi_editor"
            )
            if edit_modmoi:
                if st.button("💾 Guardar cambios MOD y MOI", type="primary", key="save_modmoi"):
                    cols_mm = [c for c in edited.columns if c != "Seleccionar"]
                    edited_data = edited[cols_mm].copy()
                    for i, orig_idx in enumerate(original_indices_modmoi):
                        for col in cols_mm:
                            st.session_state.modmoi_df.at[orig_idx, col] = edited_data.at[i, col]
                    save_df(st.session_state.modmoi_df, "modmoi")
                    st.success("✅ Cambios guardados correctamente en Google Sheets")
                    st.rerun()
            if st.button("🗑️ Eliminar seleccionados", type="secondary", key="del_btn_modmoi"):
                selected = edited[edited["Seleccionar"] == True]
                if len(selected) > 0:
                    st.session_state["confirm_del_modmoi"] = True
                    st.session_state["indices_to_del_modmoi"] = df_to_show.index[edited["Seleccionar"] == True].tolist()
                else:
                    st.warning("No hay filas seleccionadas")
                if st.session_state.get("confirm_del_modmoi", False):
                    st.warning(f"¿Estás seguro de eliminar permanentemente **{len(st.session_state['indices_to_del_modmoi'])} fila(s)** de la base de datos?")
                    col1, col2 = st.columns(2)
                    if col1.button("Sí, eliminar", key="yes_del_modmoi"):
                        st.session_state.modmoi_df = st.session_state.modmoi_df.drop(st.session_state["indices_to_del_modmoi"]).reset_index(drop=True)
                        save_df(st.session_state.modmoi_df, "modmoi")
                        del st.session_state["confirm_del_modmoi"]
                        del st.session_state["indices_to_del_modmoi"]
                        st.rerun()
                    if col2.button("Cancelar", key="cancel_del_modmoi"):
                        del st.session_state["confirm_del_modmoi"]
                        del st.session_state["indices_to_del_modmoi"]
                        st.rerun()

    # ===================== MÁRGENES =====================
    with subtab4:
        st.subheader("📏 Márgenes")
        if "margenes_df" not in st.session_state:
            st.session_state.margenes_df = load_df("margenes", columns=["Código", "Margen (%)", "Nota"])
        st.session_state.margenes_df["Código"] = st.session_state.margenes_df["Código"].astype(str)
        search_term = st.text_input("🔎 Buscar margen", "", placeholder="Escribe nota o código...", key="search_margenes")
        df_full = st.session_state.margenes_df.copy().fillna("")
        df_to_show = df_full
        if search_term and search_term.strip():
            search = search_term.strip()
            df_to_show = df_full[df_full["Nota"].astype(str).str.contains(search, case=False, na=False) |
                                df_full["Código"].astype(str).str.contains(search, case=False, na=False)]
        original_indices_margenes = df_to_show.index.tolist()
        df_to_show = df_to_show.reset_index(drop=True)
        df_to_show["Código"] = df_to_show["Código"].astype(str)
        df_to_show["Nota"] = df_to_show["Nota"].astype(str)
        df_to_show["Seleccionar"] = False
        edit_margenes = st.toggle("Desbloquear edición de Márgenes", value=False, key="toggle_margenes")
        if st.button("➕ Agregar nuevo margen", key="add_margenes"):
            st.session_state.show_add_margenes = True
        if st.session_state.get("show_add_margenes", False):
            with st.form("form_add_margenes"):
                st.write("**Nuevo margen**")
                new_data = {}
                for col in ["Código", "Margen (%)", "Nota"]:
                    new_data[col] = st.text_input(col, "")
                if st.form_submit_button("Guardar nuevo"):
                    new_row = pd.DataFrame([new_data])
                    st.session_state.margenes_df = pd.concat([st.session_state.margenes_df, new_row], ignore_index=True)
                    save_df(st.session_state.margenes_df, "margenes")
                    st.success("Margen agregado")
                    st.session_state.show_add_margenes = False
                    st.rerun()
        edited = st.data_editor(
            df_to_show,
            use_container_width=True,
            num_rows="fixed",
            disabled=not edit_margenes,
            column_config={
                "Código": st.column_config.TextColumn(width="auto"),
                "Margen (%)": st.column_config.NumberColumn(width="auto", format="%.1f%%"),
                "Nota": st.column_config.TextColumn(width="auto"),
                "Seleccionar": st.column_config.CheckboxColumn("Seleccionar", default=False, width="small")
            },
            hide_index=True,
            key="margenes_editor"
        )
        if edit_margenes:
            if st.button("💾 Guardar cambios Márgenes", type="primary", key="save_margenes"):
                cols_mar = [c for c in edited.columns if c != "Seleccionar"]
                edited_data = edited[cols_mar].copy()
                for i, orig_idx in enumerate(original_indices_margenes):
                    for col in cols_mar:
                        st.session_state.margenes_df.at[orig_idx, col] = edited_data.at[i, col]
                save_df(st.session_state.margenes_df, "margenes")
                st.success("✅ Cambios guardados correctamente en Google Sheets")
                st.rerun()
        if st.button("🗑️ Eliminar seleccionados", type="secondary", key="del_btn_margenes"):
            selected = edited[edited["Seleccionar"] == True]
            if len(selected) > 0:
                st.session_state["confirm_del_margenes"] = True
                st.session_state["indices_to_del_margenes"] = df_to_show.index[edited["Seleccionar"] == True].tolist()
            else:
                st.warning("No hay filas seleccionadas")
            if st.session_state.get("confirm_del_margenes", False):
                st.warning(f"¿Estás seguro de eliminar permanentemente **{len(st.session_state['indices_to_del_margenes'])} fila(s)** de la base de datos?")
                col1, col2 = st.columns(2)
                if col1.button("Sí, eliminar", key="yes_del_margenes"):
                    st.session_state.margenes_df = st.session_state.margenes_df.drop(st.session_state["indices_to_del_margenes"]).reset_index(drop=True)
                    save_df(st.session_state.margenes_df, "margenes")
                    del st.session_state["confirm_del_margenes"]
                    del st.session_state["indices_to_del_margenes"]
                    st.rerun()
                if col2.button("Cancelar", key="cancel_del_margenes"):
                    del st.session_state["confirm_del_margenes"]
                    del st.session_state["indices_to_del_margenes"]
                    st.rerun()

# ==================================== PESTAÑA 3: PRECIOS DE VENTA ====================================
with tab3:
    st.subheader("💰 Precios de Venta")
    if "reception_df" not in st.session_state:
        st.warning("Primero sube un archivo en la pestaña Recepción")
        st.stop()
    reception_df = st.session_state.reception_df.copy()
    bdb_df = st.session_state.get("bdb_full_df", pd.DataFrame())
    backend_df = st.session_state.get("backend_df", pd.DataFrame())
    insumos_df = st.session_state.get("insumos_df", pd.DataFrame())
    modmoi_df = st.session_state.get("modmoi_df", pd.DataFrame())
    margenes_df = st.session_state.get("margenes_df", pd.DataFrame())

    def safe_float(val):
        if pd.isna(val) or str(val).strip() in ["", "nan", "None", "No Encontrado"]:
            return 0.0
        try:
            return float(val)
        except:
            return 0.0

    def get_additional_costs(codigos_str, df_costos, col_costo):
        if not codigos_str or pd.isna(codigos_str) or str(codigos_str).strip() == "":
            return 0.0
        codigos = [c.strip() for c in str(codigos_str).split() if c.strip()]
        total = 0.0
        for codigo in codigos:
            match = df_costos[df_costos["Código"].astype(str).str.strip() == codigo]
            if len(match) > 0:
                total += safe_float(match.iloc[0][col_costo])
        return total

    def calcular_precios_venta(reception_df, bdb_df, backend_df, insumos_df, modmoi_df, margenes_df):
        resultados = []
        for _, row in reception_df.iterrows():
            producto = str(row["Producto"]).strip()
            costo_unitario = safe_float(row["Costo Neto Unitario"])
            cantidad = safe_float(row["Cantidad"])
            lote = str(row.get("Serie/Lote", "")).strip()

            match_bdb = bdb_df[bdb_df["PRODUCTO"].astype(str).str.strip() == producto]
            if len(match_bdb) == 0:
                match_bdb = bdb_df[bdb_df["PRODUCTO"].astype(str).str.contains(producto, case=False, na=False)]
            if len(match_bdb) == 0:
                resultados.append({
                    "N°": int(row["N°"]),
                    "Lote": lote,
                    "Producto": producto,
                    "Subproducto": "No encontrado",
                    "Costo Factor": 0,
                    "Insumos + MOD y MOI": 0,
                    "Costo Neto Total": round(costo_unitario * cantidad, 2),
                    "Utilidad": "SIN UTILIDAD",
                    "Utilidad Neta": "",
                    "% Real": "",
                    "Precio Venta Neto": 0,
                    "Precio Venta Bruto": "",
                    "Precio KG": ""
                })
                continue

            main_factor = safe_float(match_bdb.iloc[0].get("SUB1_Factor", 1.0))
            if main_factor <= 0:
                main_factor = 1.0

            sub_cols = [col for col in match_bdb.columns if col.startswith("SUB") and not col.endswith("_Factor")]

            for sub_col in sub_cols:
                sub_name = str(match_bdb.iloc[0][sub_col]).strip()
                if not sub_name or sub_name == "":
                    continue
                factor_col = sub_col + "_Factor"
                sub_factor = safe_float(match_bdb.iloc[0].get(factor_col, 1.0))
                if sub_factor <= 0:
                    sub_factor = 1.0

                costo_por_kg = costo_unitario / main_factor
                costo_sub_neto = costo_por_kg * sub_factor

                backend_match = backend_df[backend_df["Producto"].astype(str).str.strip() == sub_name]
                if len(backend_match) == 0:
                    backend_match = backend_df[backend_df["Producto"].astype(str).str.contains(sub_name, case=False, na=False)]

                if len(backend_match) > 0:
                    cod = backend_match.iloc[0]["COD"]
                    utilidad = safe_float(calcular_utilidad(cod, margenes_df))
                    insumos_cod = backend_match.iloc[0].get("Insumos", "")
                    modmoi_cod = backend_match.iloc[0].get("MOD y MOI", "")
                else:
                    utilidad = 0.0
                    insumos_cod = ""
                    modmoi_cod = ""

                costo_insumos = get_additional_costs(insumos_cod, insumos_df, "Costo Neto Unitario")
                costo_modmoi = get_additional_costs(modmoi_cod, modmoi_df, "Costo Neto")

                if utilidad > 0:
                    precio_neto_factor = costo_sub_neto / (1 - utilidad / 100)
                    precio_neto_final = precio_neto_factor + costo_insumos + costo_modmoi
                    precio_bruto_temp = precio_neto_final * 1.19
                    precio_bruto = math.ceil(precio_bruto_temp / 50) * 50
                    utilidad_str = f"{utilidad:.1f}%"
                    precio_kg = round(precio_bruto / sub_factor, 2) if sub_factor > 0 else 0
                    real_pct_str = f"{round(((precio_neto_final - costo_sub_neto) / precio_neto_final * 100), 1)}%"
                else:
                    precio_neto_final = costo_sub_neto + costo_insumos + costo_modmoi
                    precio_bruto_temp = precio_neto_final * 1.19
                    precio_bruto = math.ceil(precio_bruto_temp / 50) * 50
                    utilidad_str = "SIN UTILIDAD"
                    precio_kg = ""
                    real_pct_str = ""

                costo_total_sub = round(costo_sub_neto + costo_insumos + costo_modmoi, 2)
                precio_venta_neto_calc = round(precio_bruto / 1.19, 2) if utilidad > 0 else 0
                utilidad_neta = round(precio_venta_neto_calc - costo_total_sub, 2) if utilidad > 0 else ""
                resultados.append({
                    "N°": int(row["N°"]),
                    "Lote": lote,
                    "Producto": producto,
                    "Subproducto": sub_name,
                    "Costo Factor": round(costo_sub_neto, 2),
                    "Insumos + MOD y MOI": round(costo_insumos + costo_modmoi, 2),
                    "Costo Neto Total": costo_total_sub,
                    "Utilidad": utilidad_str,
                    "Utilidad Neta": utilidad_neta,
                    "% Real": real_pct_str,
                    "Precio Venta Neto": round(precio_neto_final, 2),
                    "Precio Venta Bruto": precio_bruto if utilidad > 0 else "",
                    "Precio KG": precio_kg if utilidad > 0 else ""
                })

        df_final = pd.DataFrame(resultados)
        for col in ["Costo Factor", "Insumos + MOD y MOI", "Costo Neto Total", "Precio Venta Neto"]:
            if col in df_final.columns:
                df_final[col] = df_final[col].apply(lambda x: f"${x:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
        for col in ["Precio Venta Bruto", "Precio KG", "Utilidad Neta"]:
            if col in df_final.columns:
                df_final[col] = df_final[col].apply(lambda x: f"${int(round(x)):,.0f}".replace(",", "X").replace(".", ",").replace("X", ".") if x != "" and x != 0 else "")
        return df_final

    if st.button("🔄 Calcular Precios de Venta", type="primary"):
        with st.spinner("Aplicando lógica completa usando SUB*_Factor..."):
            df_final = calcular_precios_venta(reception_df, bdb_df, backend_df, insumos_df, modmoi_df, margenes_df)
            st.session_state.df_precios = df_final
            st.success("✅ Precios calculados correctamente según la lógica del Excel")

    if "df_precios" in st.session_state:
        df_display = st.session_state.df_precios[["N°", "Lote", "Subproducto", "Costo Factor", "Insumos + MOD y MOI", "Costo Neto Total", "Utilidad", "Utilidad Neta", "% Real", "Precio Venta Bruto", "Precio KG"]].copy()

        def style_row(row):
            n = int(row["N°"])
            bg_color = "#e5e5e5" if n % 2 == 0 else "white"
            util_str = str(row["Utilidad"]).strip()
            if util_str == "SIN UTILIDAD":
                return [f'background-color: {bg_color}'] * 11
            try:
                real_val = float(str(row["% Real"]).replace("%", ""))
                util_val = float(str(util_str).replace("%", ""))
                color_real = '#00cc00' if real_val >= util_val else '#ff4444'
                return [
                    f'background-color: {bg_color}',
                    f'background-color: {bg_color}',
                    f'background-color: {bg_color}',
                    f'background-color: {bg_color}',
                    f'background-color: {bg_color}',
                    f'background-color: {bg_color}',
                    f'background-color: {bg_color}',
                    f'color: #1a86c7; background-color: {bg_color}',
                    f'color: {color_real}; background-color: {bg_color}',
                    f'background-color: {bg_color}',
                    f'background-color: {bg_color}'
                ]
            except:
                return [f'background-color: {bg_color}'] * 11

        styled_df = df_display.style.apply(style_row, axis=1)
        st.dataframe(styled_df, use_container_width=True, hide_index=True)

        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            st.session_state.df_precios.to_excel(writer, index=False)
        st.download_button("📥 Descargar Excel de Precios de Venta", data=output.getvalue(),
                           file_name="Precios_de_Venta.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

st.caption("Desarrollado para La Trilla con ❤️ • Datos en Google Sheets v1.0")
