import math
from datetime import date, datetime, timedelta
from io import BytesIO
import re

import numpy as np
import pandas as pd
import streamlit as st

import compras_core as cc
if "Frecuencia" not in cc.COLS_PROVEEDORES or "Mín. kg compra" not in cc.COLS_PROVEEDORES or "Máx. viajes" not in cc.COLS_PROVEEDORES or "max_viajes" not in cc.distribuir_viajes.__code__.co_varnames:
    st.error("Falta actualizar el archivo **compras_core.py** en GitHub: Compras_app.py y compras_core.py siempre deben subirse juntos. "
             "Sube la última versión de compras_core.py y reinicia la app (Manage app → Reboot).")
    st.stop()

# ---------------- Contraseña del módulo ----------------
import hashlib, hmac
_PW_SHA256 = "23b1cc0f5125ec67c962e2d049f22ea0ef49133d57ca28618451df5e57b15274"      # contraseña por defecto (sólo su huella); mejor: definir COMPRAS_PASSWORD en los Secrets de Streamlit


def _pw_ok(x):
    try:
        sec = st.secrets.get("COMPRAS_PASSWORD")
    except Exception:
        sec = None
    if sec:
        return hmac.compare_digest(str(x), str(sec))
    return hmac.compare_digest(hashlib.sha256(str(x).encode()).hexdigest(), _PW_SHA256)


if not st.session_state.get("cmp_unlocked"):
    st.markdown("#### 🔒 Compras y Reposición")
    with st.form("form_pw_compras"):
        _pw = st.text_input("Contraseña del módulo", type="password")
        _go = st.form_submit_button("Desbloquear", type="primary")
    if _go:
        if _pw_ok(_pw):
            st.session_state.cmp_unlocked = True
            st.rerun()
        else:
            st.error("Contraseña incorrecta.")
    st.stop()

# Se ejecuta con exec() desde La_Trilla_app.py, igual que el resto de los módulos.
spreadsheet = st.session_state.google_spreadsheet

st.markdown("<style>.block-container{padding-top:2.4rem !important;padding-bottom:1rem !important}"
            "div[data-testid='stRadio'] label p{font-size:0.92rem}</style>", unsafe_allow_html=True)


# ===================== Google Sheets (mismo formato que Calculadora) =====================
def _cmp_load(sheet_name, columns):
    try:
        ws = spreadsheet.worksheet(sheet_name)
        vals = ws.get_all_values()
    except Exception:
        spreadsheet.add_worksheet(title=sheet_name, rows="2000", cols=str(max(10, len(columns))))
        vals = []
    if len(vals) < 2:
        return pd.DataFrame(columns=columns)
    df = pd.DataFrame(vals[1:], columns=vals[0])
    for c in columns:
        if c not in df.columns:
            df[c] = ""
    if sheet_name == cc.SH_PRECIOS:      # precios siempre como número real ("7731,08", "$3.656,30" -> 7731.08 / 3656.3)
        df["Precio Neto Kg"] = df["Precio Neto Kg"].map(lambda x: (lambda v: v if v > 0 else "")(cc.to_float(x, 0.0)))
    return df


def _cmp_save(sheet_name, df, force=False):
    try:
        ws = spreadsheet.worksheet(sheet_name)
    except Exception:
        ws = spreadsheet.add_worksheet(title=sheet_name, rows="2000", cols="12")
    if sheet_name == cc.SH_PRECIOS and not force:      # protección: no pisar la base si se perderían muchos precios
        try:
            _old = ws.get_all_values()
            _h = _old[0] if _old else []
            if "Precio Neto Kg" in _h:
                _i = _h.index("Precio Neto Kg")
                _n0 = sum(1 for r in _old[1:] if len(r) > _i and cc.to_float(r[_i], 0.0) > 0)
                _n1 = sum(1 for v in df["Precio Neto Kg"] if cc.to_float(v, 0.0) > 0)
                if _n0 - _n1 > max(5, 0.05 * _n0):
                    raise RuntimeError(f"⛔ No guardé los precios: este cambio dejaría {_n0 - _n1} precios en blanco (la base tiene {_n0} y quedarían {_n1}). "
                                       "Recarga con 🔄 y revisa.")
        except RuntimeError:
            raise
        except Exception:
            pass
    ws.clear()
    d = df.copy().astype(object).where(pd.notna(df), "")
    rows = [list(map(str, d.columns))]
    for _, r in d.iterrows():
        rows.append([v if isinstance(v, (int, float, bool)) else str(v) for v in r])
    ws.update(rows, value_input_option="RAW")


def _cmp_append(sheet_name, df, columns):
    try:
        ws = spreadsheet.worksheet(sheet_name)
        vals = ws.get_all_values()
        if not vals:
            ws.update([columns], value_input_option="RAW")
        elif len(vals[0]) < len(columns):       # hoja antigua: se agregan las columnas nuevas al encabezado
            vals[0] = list(vals[0]) + [c for c in columns if c not in vals[0]]
            ws.clear()
            ws.update(vals, value_input_option="RAW")
    except Exception:
        ws = spreadsheet.add_worksheet(title=sheet_name, rows="2000", cols="12")
        ws.update([columns], value_input_option="RAW")
    d = df[columns].astype(object).where(pd.notna(df[columns]), "")
    ws.append_rows([[v if isinstance(v, (int, float)) else str(v) for v in r] for r in d.values.tolist()],
                   value_input_option="RAW")


def _get(key, sheet, cols, force=False):
    if force or key not in st.session_state:
        st.session_state[key] = _cmp_load(sheet, cols)
    return st.session_state[key]


def clp(v):
    try:
        return "$" + f"{float(v):,.0f}".replace(",", ".")
    except Exception:
        return ""


# ===================== Estado =====================
ss = st.session_state
ss.setdefault("cmp_opciones", {})          # clave producto -> nº de opción elegida
ss.setdefault("cmp_feriados", dict(cc.FERIADOS_DEFAULT))
# estado del plan de compra: se inicializa al inicio para que los callbacks siempre lo encuentren
ss.setdefault("cmp_elegido", {})
ss.setdefault("cmp_rev", set())
ss.setdefault("cmp_excl", set())
ss.setdefault("cmp_dirty", False)
ss.setdefault("cmp_del", None)
ss.setdefault("cmp_bultos", {})

# ===================== Guardar / continuar el análisis =====================
_SNAP_KEYS = ["cmp_rep", "cmp_meta", "cmp_ventas", "cmp_cal_hist", "cmp_cal_fut", "cmp_cal_fut_sig", "cmp_feriados",
              "cmp_elegido", "cmp_rev", "cmp_excl", "cmp_bultos", "cmp_cur", "cmp_vjg_cfg", "cmp_vjg_asig",
              "_k_ini", "_k_fin", "_k_col", "_k_crit", "_k_ver", "_k_busca", "plan_tipos",
              "_k_vj_cap", "_k_vj_max", "_k_vj_buf", "_k_vj_paso", "_k_vj_llenar",
              "cmp_ventas_lin", "cmp_esp_repite", "_k_metodo", "_k_esp_mult", "_k_esp_min",
              "cmp_ly_lin", "cmp_ly_cal", "cmp_temp", "_k_temp_fin", "_k_temp_ini", "_k_temp_n",
              "cmp_siempre", "cmp_stock_manual", "cmp_stock_orig", "cmp_stock_new", "cmp_stock_base", "cmp_stock_fecha", "cmp_stock_base_fecha", "cmp_comprar_prev"]
_SNAP_MAX_SESIONES = 8


def _sesiones_listar():
    d = _cmp_load(cc.SH_SESION, cc.COLS_SESION)
    if d.empty:
        return pd.DataFrame(columns=["ID", "Fecha", "Detalle"])
    d = d.drop_duplicates("ID")[["ID", "Fecha", "Detalle"]]
    return d.sort_values("ID", ascending=False).reset_index(drop=True)


def _cb_guardar_sesion():
    if "cmp_rep" not in ss:
        ss.cmp_err_guardar = "No hay un análisis para guardar: primero carga el Reposition Stock en 📥 Datos."
        return
    try:
        est = {k: ss[k] for k in _SNAP_KEYS if k in ss}
        if ss.get("cmp_dirty"):                      # cambios de precios / quiebres aún no guardados en la base
            est["cmp_precios"], est["cmp_dirty"] = ss.cmp_precios, True
        partes = cc.partir(cc.empaquetar(est))
        ahora = datetime.now()
        sid = ahora.strftime("%Y%m%d-%H%M%S")
        f_ini, f_fin = ss.get("_k_ini"), ss.get("_k_fin")
        det = f"{len(ss.cmp_rep)} productos"
        if f_ini and f_fin:
            det += f" · cobertura {f_ini:%d/%m} al {f_fin:%d/%m}"
        det += f" · {len(ss.get('cmp_rev', ()))} revisadas"
        filas = pd.DataFrame([{"ID": sid, "Fecha": ahora.strftime("%d/%m/%Y %H:%M"), "Detalle": det,
                               "Parte": i + 1, "Partes": len(partes), "Datos": p} for i, p in enumerate(partes)])
        _cmp_save(cc.SH_SESION, filas)               # sólo se conserva el último registro (reemplaza al anterior)
        ss.pop("cmp_ses_list", None)
        ss.cmp_ultimo_guardado = ahora.strftime("%d/%m %H:%M")
        st.toast(f"✅ Análisis guardado ({ss.cmp_ultimo_guardado})", icon="✅")
        ss.pop("cmp_err_guardar", None)
    except Exception as e:
        ss.cmp_err_guardar = f"No se pudo guardar el análisis: {e}"
        st.toast("No se pudo guardar el análisis", icon="❌")


def _cb_cargar_sesion(sid):
    try:
        d = _cmp_load(cc.SH_SESION, cc.COLS_SESION)
        d = d[d["ID"].astype(str) == str(sid)].copy()
        d["Parte"] = pd.to_numeric(d["Parte"], errors="coerce")
        est = cc.desempaquetar("".join(d.sort_values("Parte")["Datos"].astype(str)))
    except Exception as e:
        ss.cmp_msg_ses = (False, f"No se pudo cargar el registro: {e}")
        return
    for k in ("cmp_plan_cache", "cmp_plan_cache_key", "cmp_pn_key", "cmp_stale", "cmp_last_rows", "cmp_nec", "cmp_base",
              "cmp_plan_sel", "cmp_del", "cmp_msg"):
        ss.pop(k, None)
    for k in [k for k in ss.keys() if isinstance(k, str) and (k.startswith("inc_") or k.startswith("vjd_") or k == "vjn")]:
        ss.pop(k, None)
    for k, v in est.items():
        ss[k] = v
    cfg = ss.get("cmp_vjg_cfg")
    if cfg:                                             # viajes definidos a mano
        ss["vjn"] = cfg["n"]
        for i, f_ in enumerate(cfg["fechas"]):
            ss[f"vjd_{i}"] = f_
    # se recalculan la demanda y la necesidad con los datos cargados
    try:
        peso_h = float(ss.cmp_cal_hist["Peso"].sum())
        ss.cmp_peso_hist = peso_h
        g, _sf = cc.demanda_por_producto(ss.cmp_rep, ss.cmp_meta, cc.build_formato_map(ss.cmp_formatos), peso_h)
        ss.cmp_base = g
        cf = ss.get("cmp_cal_fut")
        if cf is not None and len(cf):
            ss.cmp_nec = _nec_completa(g, float(cf.loc[cf["Incluir"], "Peso"].sum()), float(ss.get("_k_col", 10.0)))
    except Exception as e:
        ss.cmp_msg_ses = (False, f"Se cargó el registro pero no se pudo recalcular: {e}")
        return
    ss.cmp_msg_ses = (True, "Análisis cargado. Puedes seguir desde donde lo dejaste (revisa 🛒 Plan de compra).")


_NAVMAP = {"📥 Datos": [("Datos y ventas", 0)],
           "📊 Análisis": [("Cobertura", 1), ("Cambios de venta", "cv"), ("Productos no considerados", "nc")],
           "🛒 Plan de compra": [("Plan de compra", 2)],
           "📄 Pedidos": [("Órdenes de compra", 4), ("Logística y viajes", 5)],
           "⚙️ Configuración": [("Proveedores y precios", 3), ("Procesados", 6)]}
hc1, hc2, hc3, hc4 = st.columns([1.3, 6.4, 1.0, 0.7], vertical_alignment="center")
hc1.markdown("#### 🛒 Compras")
_g = hc2.radio("Sección", list(_NAVMAP), horizontal=True, key="cmp_grp", label_visibility="collapsed")
hc3.button("💾 Guardar", on_click=_cb_guardar_sesion, use_container_width=True,
           help="Guarda en Google Sheets (hoja compras_sesion) todo lo que llevas: datos, cobertura, proveedores elegidos, "
                "cantidades, productos revisados y viajes, para continuar otro día con «Cargar último registro» en 📥 Datos.")
if hc4.button("🔄", use_container_width=True, help="Recargar bases: vuelve a leer proveedores, precios y formatos desde Google Sheets"):
    for k in ("cmp_precios", "cmp_formatos", "cmp_provs", "cmp_pesolog"):
        ss.pop(k, None)
    st.rerun()
if ss.get("cmp_err_guardar"):
    st.error(ss.cmp_err_guardar)

df_precios_db = _get("cmp_precios", cc.SH_PRECIOS, cc.COLS_PRECIOS)
df_formatos_db = _get("cmp_formatos", cc.SH_FORMATOS, cc.COLS_FORMATOS)
df_provs_db = _get("cmp_provs", cc.SH_PROVEEDORES, cc.COLS_PROVEEDORES)
if "cmp_siempre" not in st.session_state:
    st.session_state["cmp_siempre"] = pd.DataFrame(columns=cc.COLS_SIEMPRE)     # temporal: vale sólo para el análisis en curso
df_recetas_db = _get("cmp_recetas", cc.SH_RECETAS, cc.COLS_RECETAS)
fmap = cc.build_formato_map(df_formatos_db)


def _cargar_pesolog():
    """PESO_LOGISTICO (kg por unidad) desde la hoja 'productos' del Google Sheet. No crea la hoja si no existe."""
    try:
        vals = spreadsheet.worksheet("productos").get_all_values()
    except Exception:
        return {}
    if len(vals) < 2:
        return {}
    return cc.mapa_peso_logistico(pd.DataFrame(vals[1:], columns=vals[0]))


if "cmp_pesolog" not in ss:
    ss.cmp_pesolog = _cargar_pesolog()

_SECC = ["📥 1. Datos y ventas", "📊 2. Cobertura", "🛒 3. Plan de compra", "🏷️ 4. Proveedores y precios", "📄 5. Orden de compra", "🚚 6. Logística", "🧪 7. Procesados"]
# Streamlit borra el estado de los widgets que no se dibujan; con esto se conservan al cambiar de sección
for _k in ("plan_tipos",):
    if _k in ss:
        ss[_k] = ss[_k]


def _keep(name, default):
    """Valor recordado de un filtro (sobrevive al cambio de sección sin usar la API de session_state de los widgets)."""
    return ss.get("_k_" + name, default)


def _aplicar_stock_rep():
    """Si hay un stock actualizado, reemplaza el Stock del reposition (por SKU; los que no vienen en el reporte quedan en 0)."""
    sn, man = ss.get("cmp_stock_new"), ss.get("cmp_stock_manual")
    if (not sn and not man) or "cmp_rep" not in ss:
        return
    r_ = ss.cmp_rep.copy()
    if sn:
        r_["Stock"] = [float(sn.get(str(x).strip(), 0.0)) for x in r_["SKU"]]
    if man:
        r_["Stock"] = [float(man.get(str(x).strip(), v_)) for x, v_ in zip(r_["SKU"], r_["Stock"])]
    ss.cmp_rep = r_


def _cb_aplicar_stock():
    d = ss.get("_stock_df")
    if d is None or "cmp_rep" not in ss:
        return
    fch = ss.get("stock_fecha_in") or date.today()
    r_ = ss.cmp_rep
    ss.cmp_stock_base = {str(a).strip(): float(b) for a, b in zip(r_["SKU"], r_["Stock"])}
    ss.cmp_stock_base_fecha = ss.get("cmp_stock_fecha") or ss.cmp_meta.get("fecha_fin") or fch
    ss.cmp_comprar_prev = dict(zip(ss.cmp_nec["Clave"], ss.cmp_nec["Comprar kg"])) if "cmp_nec" in ss else {}
    ss.cmp_stock_new = {str(a).strip(): float(b) for a, b in zip(d["SKU"], d["Stock"])}
    ss.cmp_stock_manual, ss.cmp_stock_orig = {}, {}      # un stock nuevo deja sin efecto los ajustes manuales anteriores
    ss.cmp_stock_fecha = fch
    _aplicar_stock_rep()
    ss._k_ini = fch + timedelta(days=1)          # se cubre desde el día siguiente al stock; la fecha final queda igual
    ss.pop("cmp_plan_cache_key", None)
    ss.cmp_msg_stock = (True, f"Stock actualizado al {fch:%d/%m/%Y}. La cobertura ahora parte el {fch + timedelta(days=1):%d/%m/%Y} "
                              f"y termina igual que antes. Revisa los 🔴 en el Plan de compra.")


def _nec_completa(g, peso_h, colchon):
    """Necesidad de compra con el método elegido (promedio simple o base ajustada sin ventas especiales puntuales).
    Deja en ss.cmp_esp la lista de ventas especiales detectadas."""
    mult, mn = float(_keep("esp_mult", 3.0)), float(_keep("esp_min", 5.0))
    fm = cc.build_formato_map(ss.cmp_formatos)
    ss.cmp_esp = cc.ventas_especiales(ss.get("cmp_ventas_lin"), g, ss.cmp_meta, fm, mult, mn)
    ga = cc.ajusta_por_especiales(g, ss.cmp_esp, ss.get("cmp_esp_repite", set()), float(ss.get("cmp_peso_hist", 0.0)),
                                  int(ss.cmp_meta.get("n_sem", 1)))
    ga["Venta diaria simple"] = ga["Venta diaria kg"]
    if _keep("metodo", "simple") == "ajustada":
        ga["Venta diaria kg"] = ga["Venta diaria ajustada"]
        ga["Venta semanal kg"] = ga["Venta semanal ajustada"]
    nec = cc.calcular_necesidad(ga, peso_h, colchon)
    # ---- compra de temporada: productos elegidos a mano se compran para repetir lo que se vendió el año anterior ----
    nec["Comprar normal kg"] = nec["Comprar kg"]
    nec["Temporada"] = False
    nec["Previo kg"] = 0.0
    nec["Venta diaria normal kg"] = nec["Venta diaria kg"]
    temp = ss.get("cmp_temp") or {}
    if ss.get("cmp_ly_lin") is not None and ss.get("_k_ini"):
        ini_t = ss._k_ini
        fin_t = _keep("temp_fin", ss.get("_k_fin") or ini_t)
        sh = timedelta(days=364)
        des_t = max(ini_t, min(_keep("temp_ini", ini_t), fin_t))
        ly, _cob = cc.ventas_ly_rango(ss.cmp_ly_lin, ss.cmp_ly_cal, fm, des_t - sh, fin_t - sh)
        ly = dict(zip(ly["Clave"], ly["Kg"]))
        cal_t = ss.get("cmp_cal_fut")
        pw = {pd.Timestamp(r["Fecha"]).date(): (float(r["Peso"]) if r["Incluir"] else 0.0) for r in cal_t.to_dict("records")} \
            if cal_t is not None and len(cal_t) else {}
        _pd = lambda a_, b_: sum(pw[d_] if d_ in pw else cc.peso_base(d_, ss.cmp_feriados)
                                 for d_ in (a_ + timedelta(days=i) for i in range((b_ - a_).days + 1)))
        peso_e = _pd(ini_t, fin_t)
        peso_pre = _pd(ini_t, des_t - timedelta(days=1)) if des_t > ini_t else 0.0
        nec["Previo kg"] = nec["Venta diaria kg"] * peso_pre * (1 + colchon / 100.0)
        for clave, extra in temp.items():
            m = nec["Clave"] == clave
            if not m.any():
                continue
            pre = float(nec.loc[m, "Venta diaria kg"].iloc[0]) * peso_pre * (1 + colchon / 100.0)
            dem = pre + float(ly.get(clave, 0.0)) * (1 + float(extra) / 100.0)
            nec.loc[m, "Previo kg"] = pre
            stock = float(nec.loc[m, "Stock kg"].iloc[0])
            daily = dem / peso_e if peso_e > 0 else 0.0
            nec.loc[m, "Demanda kg"] = dem
            nec.loc[m, "Comprar kg"] = max(0.0, dem - stock)
            nec.loc[m, "Venta diaria kg"] = daily
            nec.loc[m, "Cobertura días hábiles"] = (stock / daily) if daily > 0 else float("inf")
            nec.loc[m, "Temporada"] = True
    # ---- procesados (mixes, granolas…): su necesidad pasa a los ingredientes según la receta ----
    nec = cc.aplicar_recetas(nec, ss.get("cmp_recetas"))
    # ---- alerta: ¿el stock real bajó más de lo previsto desde el stock anterior? ----
    nec["Stock previsto kg"] = float("nan")
    nec["Desvío stock kg"] = float("nan")
    nec["Alerta stock"] = ""
    sb, f1_, f0_ = ss.get("cmp_stock_base"), ss.get("cmp_stock_fecha"), ss.get("cmp_stock_base_fecha")
    cons = sbk = None
    if sb and f1_ and f0_ and f1_ > f0_ and "cmp_rep" in ss:
        rb = ss.cmp_rep.copy()
        rb["Stock"] = [float(sb.get(str(x).strip(), a_)) for x, a_ in zip(rb["SKU"], rb["Stock"])]
        gb, _ = cc.demanda_por_producto(rb, ss.cmp_meta, fm, float(ss.get("cmp_peso_hist", 1.0)) or 1.0)
        sbk = nec["Clave"].map(dict(zip(gb["Clave"], gb["Stock kg"])))
        w_ = sum(cc.peso_base(f0_ + timedelta(days=i), ss.cmp_feriados) for i in range(1, (f1_ - f0_).days + 1))
        cons = nec["Venta diaria normal kg"] * w_
        prev = (sbk - cons).clip(lower=0)
        des = nec["Stock kg"] - prev
        tol = (cons * 0.25).clip(lower=1.0)
        nec["Stock previsto kg"] = prev
        nec["Desvío stock kg"] = des
        nec["Alerta stock"] = ["🔴" if (pd.notna(d_) and d_ < -t_) else ("🟢" if (pd.notna(d_) and d_ > t_) else "")
                               for d_, t_ in zip(des, tol)]
    # ---- productos agregados a mano a este análisis: compra mínima aunque no haya venta ----
    nec["Comprar base kg"] = nec["Comprar kg"]
    nec["Siempre"] = False
    sg_ = ss.get("cmp_siempre")
    if sg_ is not None and len(sg_):
        for cl_, ca_ in zip(sg_["Clave"], pd.to_numeric(sg_["Cantidad"], errors="coerce").fillna(0.0)):
            m_ = nec["Clave"] == cl_
            if m_.any() and ca_ > 0:
                nec.loc[m_, "Comprar kg"] = max(float(nec.loc[m_, "Comprar kg"].iloc[0]), float(ca_))
                nec.loc[m_, "Siempre"] = True
    # ---- cambios de venta: tendencia de las semanas del reporte y venta real desde el stock anterior ----
    wk_ = [c_ for c_ in nec.columns if re.fullmatch(r"W\d+ kg", c_)]
    nec["Tendencia %"] = float("nan")
    nec["Tend. kg/sem"] = 0.0
    if len(wk_) >= 2:
        prev_m = nec[wk_[:-1]].mean(axis=1)
        last_ = nec[wk_[-1]]
        nec["Tend. kg/sem"] = last_ - prev_m
        nec["Tendencia %"] = np.where(prev_m > 0, (last_ - prev_m) / prev_m * 100.0, np.where(last_ > 0, 999.0, np.nan))
    nec["Venta esperada kg"] = float("nan")
    nec["Venta real kg"] = float("nan")
    nec["Cambio venta %"] = float("nan")
    nec["Otras diferencias kg"] = float("nan")
    ss.cmp_vreal_cob = None
    if sb and f1_ and f0_ and f1_ > f0_ and ss.get("cmp_ventas_lin") is not None and ss.get("cmp_ventas") is not None:
        lr_, cob_ = cc.ventas_ly_rango(ss.cmp_ventas_lin, ss.cmp_ventas, fm, f0_ + timedelta(days=1), f1_)
        _dmn, _dmx = min(ss.cmp_ventas["Fecha"]), max(ss.cmp_ventas["Fecha"])
        _ok_cob = _dmn <= f0_ + timedelta(days=3) and _dmx >= f1_ - timedelta(days=2)     # tolera domingos/feriados sin ventas en los bordes
        ss.cmp_vreal_cob = 1.0 if _ok_cob else cob_
        if _ok_cob:
            real_ = nec["Clave"].map(dict(zip(lr_["Clave"], lr_["Kg"]))).fillna(0.0)
            nec["Venta esperada kg"] = cons
            nec["Venta real kg"] = real_
            nec["Cambio venta %"] = np.where(cons > 0, (real_ - cons) / cons * 100.0, np.where(real_ > 0, 999.0, np.nan))
            nec["Otras diferencias kg"] = nec["Stock kg"] - (sbk - real_)
    cp_ = ss.get("cmp_comprar_prev") or {}
    nec["Comprar antes kg"] = nec["Clave"].map(cp_) if cp_ else float("nan")
    return nec
_subs = _NAVMAP[_g]
sc1, sc2 = st.columns([3.6, 3.0], vertical_alignment="center")
if len(_subs) > 1:
    _sl = sc1.radio("Pestaña", [x[0] for x in _subs], horizontal=True, key=f"cmp_sub_{list(_NAVMAP).index(_g)}", label_visibility="collapsed")
    _idx = dict(_subs)[_sl]
else:
    _idx = _subs[0][1]
_nav = _SECC[_idx] if isinstance(_idx, int) else _idx
if "cmp_nec" in ss:
    _n_ = ss.cmp_nec
    _p_ = [f"Cobertura {ss.get('_k_ini'):%d/%m} al {ss.get('_k_fin'):%d/%m}" if ss.get("_k_ini") and ss.get("_k_fin") else None,
           f"Stock al {ss.cmp_stock_fecha:%d/%m/%Y}" if ss.get("cmp_stock_fecha") else None,
           f"{int((_n_['Comprar kg'] > 0).sum())} productos a comprar",
           (f"{int((_n_['Alerta stock'] == '🔴').sum())} con 🔴 stock bajo lo previsto" if "Alerta stock" in _n_ and (_n_["Alerta stock"] == "🔴").any() else None)]
    if ss.get("cmp_ultimo_guardado"):
        _p_.append(f"guardado {ss.cmp_ultimo_guardado}")
    sc2.caption(" · ".join(x for x in _p_ if x))
elif ss.get("cmp_ultimo_guardado"):
    sc2.caption(f"Último guardado: {ss.cmp_ultimo_guardado}")

def _agregar_oferta(producto, proveedor, formato, precio, disponible="Sí"):
    """Agrega una oferta a la base en memoria (valida y devuelve (ok, mensaje))."""
    producto, proveedor = str(producto or "").strip(), str(proveedor or "").strip()
    if not producto:
        return False, "Falta el producto."
    if not proveedor:
        return False, "Falta el proveedor."
    if not precio or float(precio) <= 0:
        return False, "El precio debe ser mayor que 0."
    db = ss.cmp_precios.copy().astype(object)
    dup = (db["Producto"].map(cc.norm) == cc.norm(producto)) & (db["Proveedor"].map(cc.norm) == cc.norm(proveedor)) \
          & (db["Formato"].astype(str).map(cc.norm) == cc.norm(formato or ""))
    fila = {"Producto": producto, "Formato": formato or "", "Precio Neto Kg": float(precio), "Proveedor": proveedor,
            "Disponible": disponible, "Actualizado": date.today().isoformat(), "Nota": "Agregado en la app"}
    if dup.any():                      # misma oferta: se actualiza el precio en vez de duplicar
        for k, v in fila.items():
            db.loc[dup, k] = v
        msg = "Oferta existente actualizada."
    else:
        db = pd.concat([db, pd.DataFrame([fila])], ignore_index=True)
        msg = "Oferta agregada."
    ss.cmp_precios = db[cc.COLS_PRECIOS].reset_index(drop=True)
    ss.cmp_dirty = True
    return True, msg


# ===================== Condiciones de proveedores (factura / pago) =====================
COND_OPTS = ["Transferencia", "Efectivo", "Cheque al día", "Cheque 15 días", "Cheque 30 días", "Cheque 45 días", "Cheque N# días"]


def _cond_label(cond, dias=None):
    cond = str(cond or "").strip()
    if cond == "Cheque N# días":
        d = cc.to_float(dias, 0.0)
        return f"Cheque {int(d)} días" if d > 0 else "Cheque N# días"
    return cond


def _prov_map():
    """{proveedor normalizado: {factura, cond, entrega, retiro}}; sin datos = factura Sí, sin viajes."""
    mp = {}
    for r in ss.cmp_provs.to_dict("records"):
        k = cc.norm(r.get("Proveedor", ""))
        if k:
            ent = str(r.get("Entrega", "") or "").strip()
            mp[k] = {"factura": cc.norm(r.get("Factura", "Sí")) not in ("no", "n", "false", "0"),
                     "cond": _cond_label(r.get("Condición de pago", ""), r.get("Días cheque", "")),
                     "entrega": ent, "retiro": cc.norm(ent).startswith("retiro"),
                     "min_kg": float(pd.to_numeric(r.get("Mín. kg compra", ""), errors="coerce") or 0.0),
                     "desp": cc.parse_dias(r.get("Días despacho", "")), "lim": cc.parse_dias(r.get("Pedir hasta", "")),
                     "fechas": cc.norm(r.get("Frecuencia", "")).startswith("fecha"),
                     "desp_txt": str(r.get("Días despacho", "") or ""), "lim_txt": str(r.get("Pedir hasta", "") or "")}
    return mp


def _prov_bajo_minimo(df):
    """[(proveedor, kg comprados, mínimo)] de los proveedores cuyo total en `df` queda bajo su mínimo de compra."""
    if df is None or df.empty:
        return []
    pm = _prov_map()
    kg = df.groupby("Proveedor")["Kg total"].sum()
    out = []
    for p_, k_ in kg.items():
        m_ = pm.get(cc.norm(p_), {}).get("min_kg", 0.0)
        if m_ > 0 and k_ + 1e-6 < m_:
            out.append((p_, float(k_), float(m_)))
    return out


# =====================================================================================
# TAB 1 - DATOS
# =====================================================================================
if _nav == _SECC[0]:
    if ss.get("cmp_msg_ses"):
        (st.success if ss.cmp_msg_ses[0] else st.error)(ss.cmp_msg_ses[1])
        ss.cmp_msg_ses = None
    if "cmp_ses_list" not in ss:
        try:
            ss.cmp_ses_list = _sesiones_listar()
            ss.pop("cmp_ses_err", None)
        except Exception as e:
            ss.cmp_ses_list = pd.DataFrame(columns=["ID", "Fecha", "Detalle"])
            ss.cmp_ses_err = str(e)
    _ses = ss.cmp_ses_list
    with st.container(border=True):
        st.markdown("#### 📂 Continuar un análisis guardado")
        if _ses.empty:
            st.caption("Aún no hay análisis guardados. Cuando avances, usa **💾 Guardar análisis** (arriba a la derecha) "
                       "y podrás retomarlo otro día desde aquí.")
            if ss.get("cmp_ses_err"):
                st.error(f"No pude leer la hoja de análisis guardados ({cc.SH_SESION}): {ss.cmp_ses_err}")
            else:
                st.caption(f"Hoja «{cc.SH_SESION}»: sin registros.")
            st.button("🔄 Buscar de nuevo", on_click=lambda: ss.pop("cmp_ses_list", None), key="ses_refresh")
        else:
            u = _ses.iloc[0]
            b1, b2 = st.columns([2, 3])
            b1.button(f"📂 Cargar último registro ({u['Fecha']})", on_click=_cb_cargar_sesion, args=(u["ID"],), type="primary",
                      use_container_width=True, help=u["Detalle"])
            b2.caption(f"Último: {u['Detalle']}")
            with st.expander("Elegir otro registro"):
                ops_s = {f"{r['Fecha']} · {r['Detalle']}": r["ID"] for _, r in _ses.iterrows()}
                el_s = st.selectbox("Registro", list(ops_s), key="ses_sel")
                st.button("Cargar este registro", on_click=lambda: _cb_cargar_sesion(ops_s[st.session_state.get("ses_sel", el_s)]))
            st.caption("Al cargar, se reemplaza el análisis que tengas abierto. Si subes archivos nuevos abajo, esos datos reemplazan a los cargados.")
    st.markdown("##### " + "Datos de la última semana de análisis")
    c1, c2, c3 = st.columns(3)
    with c1:
        f_rep = st.file_uploader("Reposition Stock (4 semanas + stock actual)", type=["xlsx", "xls"], key="up_rep")
    with c2:
        f_ven = st.file_uploader("Detalle de ventas (opcional: feriados, días flojos y ventas especiales)", type=["xlsx", "xls"], key="up_ven")
    with c3:
        f_ly = st.file_uploader("Detalle de ventas del año anterior (opcional: temporada)", type=["xlsx", "xls"], key="up_ly",
                                help="El mismo mes del año pasado + el mes que vas a comprar (p. ej. 01/10/2025 al 30/11/2025).")
    if f_ly is not None and ss.get("_sig_ly") != (f_ly.name, f_ly.size):
        try:
            _cal_ly, _lin_ly = cc.read_ventas_detalle(f_ly)
            ss.cmp_ly_cal = _cal_ly[["Fecha", "Documentos"]]
            ss.cmp_ly_lin = _lin_ly[["Fecha", "Doc", "Producto", "Cantidad"]]
            ss._sig_ly = (f_ly.name, f_ly.size)
        except Exception as e:
            st.error(f"No pude leer el detalle del año anterior: {e}")
    if ss.get("cmp_ly_cal") is not None and len(ss.cmp_ly_cal):
        st.caption(f"📅 Año anterior cargado: {min(ss.cmp_ly_cal['Fecha']):%d/%m/%Y} al {max(ss.cmp_ly_cal['Fecha']):%d/%m/%Y} · "
                   f"{int(ss.cmp_ly_cal['Documentos'].sum()):,} documentos".replace(",", "."))

    if f_rep is not None:
        try:
            ss.cmp_rep, ss.cmp_meta = cc.read_reposicion(f_rep)
            if ss.get("_sig_rep") != (f_rep.name, f_rep.size):          # Reposition nuevo = análisis general nuevo: se limpian los ajustes temporales
                ss._sig_rep = (f_rep.name, f_rep.size)
                ss.cmp_siempre = pd.DataFrame(columns=cc.COLS_SIEMPRE)
                for k_ in ("cmp_stock_new", "cmp_stock_base", "cmp_stock_fecha", "cmp_stock_base_fecha", "cmp_comprar_prev"):
                    ss.pop(k_, None)
                ss.cmp_stock_manual, ss.cmp_stock_orig = {}, {}
            ss.pop("cmp_cal_hist", None)
        except Exception as e:
            st.error(f"No pude leer el reposition: {e}")
    _aplicar_stock_rep()

    if "cmp_rep" in ss:
        with st.container(border=True):
            st.markdown("#### 🔄 Actualizar el stock (sin cambiar las ventas ni las fechas a cubrir)")
            if ss.get("cmp_msg_stock"):
                st.success(ss.cmp_msg_stock[1])
                ss.cmp_msg_stock = None
            s1_, s2_ = st.columns([2, 1])
            f_st = s1_.file_uploader("Stock actual (reporte «Stock actual» de Casa Matriz)", type=["xlsx", "xls"], key="up_stock",
                                     help="Reemplaza el stock del análisis. Las ventas y la fecha final de cobertura se mantienen; "
                                          "la cobertura pasa a partir del día siguiente a la fecha del stock.")
            if f_st is not None and ss.get("_sig_stock") != (f_st.name, f_st.size):
                try:
                    ss._stock_df = cc.leer_stock_actual(f_st)
                    ss._sig_stock = (f_st.name, f_st.size)
                    _m = re.search(r"(\d{2})-(\d{2})-(\d{4})", f_st.name)
                    ss.stock_fecha_in = date(int(_m.group(3)), int(_m.group(2)), int(_m.group(1))) if _m else date.today()
                except Exception as e:
                    ss.pop("_stock_df", None)
                    st.error(f"No pude leer el stock: {e}")
            if ss.get("_stock_df") is not None:
                _d = ss._stock_df
                _skus = set(ss.cmp_rep["SKU"].astype(str).str.strip())
                s2_.date_input("Stock al", key="stock_fecha_in", format="DD/MM/YYYY",
                               help="Fecha de ese stock (se toma del nombre del archivo si trae fecha).")
                st.caption(f"{len(_d)} productos en el reporte · {int(_d['SKU'].isin(_skus).sum())} coinciden con el análisis · "
                           f"{len(_skus - set(_d['SKU']))} del análisis no vienen en el reporte (se toman con stock 0, porque el reporte es «Con stock»).")
                st.button("✅ Aplicar stock actualizado", on_click=_cb_aplicar_stock, type="primary", key="btn_aplicar_stock")
            if ss.get("cmp_stock_fecha"):
                st.caption(f"Stock vigente: al {ss.cmp_stock_fecha:%d/%m/%Y}. En el Plan de compra, 🔴 = el stock bajó más de lo previsto "
                           f"(venta, merma, vencimiento…) y 🟢 = quedó más stock del previsto.")

        with st.container(border=True):
            st.markdown("#### ✏️ Ajustar el stock de un producto a mano")
            st.caption("Para corregir un producto puntual (conteo físico, merma, vencimiento, ingreso no registrado…). "
                       "El stock va en las unidades del reporte (sacos, bolsas, envases…). Los ajustes se guardan con el análisis "
                       "y se borran si aplicas un stock nuevo desde archivo.")
            _man = ss.get("cmp_stock_manual") or {}
            _q = st.text_input("🔎 Buscar producto", key="man_busca", placeholder="Escribe parte del nombre")
            _r0 = ss.cmp_rep[["Producto", "SKU", "Stock"]].copy()
            _r0["SKU"] = _r0["SKU"].astype(str).str.strip()
            if _q.strip():
                _v0 = _r0[_r0["Producto"].str.contains(_q.strip(), case=False, na=False)].head(40).reset_index(drop=True)
                if _v0.empty:
                    st.info("No encontré productos con ese nombre.")
                else:
                    _v0["Ajustado"] = ["✏️" if k in _man else "" for k in _v0["SKU"]]
                    with st.form("form_stock_manual"):
                        _ed = st.data_editor(_v0[["Ajustado", "Producto", "SKU", "Stock"]], hide_index=True, use_container_width=True,
                                             key="ed_stock_manual", disabled=["Ajustado", "Producto", "SKU"],
                                             column_config={"Stock": st.column_config.NumberColumn("Stock (editable)", min_value=0.0, step=1.0, format="%.2f"),
                                                            "SKU": None})
                        _ok = st.form_submit_button("Aplicar ajustes", type="primary")
                    if _ok:
                        _m2, _o2 = dict(_man), dict(ss.get("cmp_stock_orig") or {})
                        for k_, ant_, nuevo_ in zip(_v0["SKU"], _v0["Stock"], _ed["Stock"]):
                            if pd.notna(nuevo_) and abs(float(nuevo_) - float(ant_)) > 1e-9:
                                _o2.setdefault(k_, float(ant_))
                                _m2[k_] = float(nuevo_)
                        ss.cmp_stock_manual, ss.cmp_stock_orig = _m2, _o2
                        _aplicar_stock_rep()
                        ss.pop("cmp_plan_cache_key", None)
                        st.rerun()
            if _man:
                _nm = dict(zip(_r0["SKU"], _r0["Producto"]))
                _lst = pd.DataFrame({"Producto": [_nm.get(k, k) for k in _man], "Stock ajustado": list(_man.values()),
                                     "Stock original": [ss.get("cmp_stock_orig", {}).get(k, float("nan")) for k in _man]})
                st.markdown(f"**Ajustes activos ({len(_man)})**")
                st.dataframe(_lst, hide_index=True, use_container_width=True)

                def _cb_quitar_man():
                    sn_, orig_ = ss.get("cmp_stock_new"), ss.get("cmp_stock_orig") or {}
                    r_ = ss.cmp_rep.copy()
                    r_["Stock"] = [float(sn_.get(str(x).strip(), 0.0)) if (sn_ and str(x).strip() in ss.cmp_stock_manual)
                                   else (float(orig_[str(x).strip()]) if str(x).strip() in ss.cmp_stock_manual and str(x).strip() in orig_ else v_)
                                   for x, v_ in zip(r_["SKU"], r_["Stock"])]
                    ss.cmp_rep = r_
                    ss.cmp_stock_manual, ss.cmp_stock_orig = {}, {}
                    ss.pop("cmp_plan_cache_key", None)
                st.button("↩️ Quitar todos los ajustes manuales", on_click=_cb_quitar_man, key="btn_quitar_man")
    if f_ven is not None and ss.get("_sig_ven") != (f_ven.name, f_ven.size):
        try:
            ss.cmp_ventas, ss.cmp_ventas_lin = cc.read_ventas_detalle(f_ven)
            ss._sig_ven = (f_ven.name, f_ven.size)
            ss.pop("cmp_cal_hist", None)
        except Exception as e:
            st.error(f"No pude leer el detalle de ventas: {e}")

    if "cmp_rep" not in ss:
        st.info("Sube el *Reposition Stock* para comenzar.")
    else:
        rep, meta = ss.cmp_rep, ss.cmp_meta
        d_ini, d_fin = meta["semanas"][0][1], meta["semanas"][-1][2]
        st.success(f"✅ {len(rep)} productos · {meta['sucursal']} · {d_ini:%d/%m/%Y} al {d_fin:%d/%m/%Y} ({meta['n_sem']} semanas)")

        st.markdown("#### Días operativos de la ventana analizada")
        st.caption("Lun–Vie = 1 · Sáb = 0,5 · Dom/feriado = 0. Con el detalle de ventas, los días de muy baja actividad "
                   "se ajustan solos (puedes corregir el peso). La venta diaria se calcula dividiendo por la suma de pesos.")
        if "cmp_cal_hist" not in ss:
            ss.cmp_cal_hist = cc.calendario_historico(ss.get("cmp_ventas"), d_ini, d_fin, ss.cmp_feriados)
        cal_ed = st.data_editor(
            ss.cmp_cal_hist, hide_index=True, use_container_width=True, key="ed_cal_hist",
            disabled=["Fecha", "Día", "Docs", "Nota"],
            column_config={"Peso": st.column_config.NumberColumn("Peso", min_value=0.0, max_value=1.0, step=0.25, format="%.2f")})
        ss.cmp_cal_hist = cal_ed
        ss.cmp_peso_hist = float(cal_ed["Peso"].sum())
        st.metric("Días hábiles equivalentes de la ventana", f"{ss.cmp_peso_hist:.2f}")

        g, sin_fmt = cc.demanda_por_producto(rep, meta, fmap, ss.cmp_peso_hist)
        ss.cmp_base = g
        if not sin_fmt.empty:
            with st.expander(f"⚠️ {len(sin_fmt)} productos con formato desconocido (no se pueden pasar a kg)"):
                st.dataframe(sin_fmt, hide_index=True, use_container_width=True)
                st.caption("Agrega el formato en la ⚙️ Configuración → Proveedores y precios → Formatos (ej: 'Env. 250gr' = 0.25).")

# =====================================================================================
# TAB 2 - COBERTURA
# =====================================================================================
if _nav == _SECC[1]:
    st.markdown("##### " + "¿Cuántos días hábiles quieres cubrir?")
    if "cmp_base" not in ss:
        st.info("Primero carga los datos en 📥 Datos.")
    else:
        a, b, c = st.columns([1, 1, 1])
        ini = a.date_input("Desde", value=_keep("ini", date.today() + timedelta(days=1)))
        fin = b.date_input("Hasta", value=_keep("fin", date.today() + timedelta(days=14)))
        colchon = c.number_input("Colchón de seguridad %", 0.0, 100.0, float(_keep("col", 10.0)), 5.0)
        ss._k_ini, ss._k_fin, ss._k_col = ini, fin, colchon
        if fin < ini:
            st.error("La fecha final debe ser posterior a la inicial.")
            st.stop()
        sig = f"{ini}|{fin}"
        if ss.get("cmp_cal_fut_sig") != sig:
            ss.cmp_cal_fut = cc.calendario_futuro(ini, fin, ss.cmp_feriados)
            ss.cmp_cal_fut_sig = sig
        with st.expander("📅 Días incluidos (desmarca feriados o días sin operación)", expanded=False):
            cal_f = st.data_editor(
                ss.cmp_cal_fut, hide_index=True, use_container_width=True, key=f"ed_cal_fut_{sig}",
                disabled=["Fecha", "Día", "Nota"],
                column_config={"Peso": st.column_config.NumberColumn("Peso", min_value=0.0, max_value=1.0, step=0.25)})
        try:
            cal_f
        except NameError:
            cal_f = ss.cmp_cal_fut
        ss.cmp_cal_fut = cal_f
        peso_h = float(cal_f.loc[cal_f["Incluir"], "Peso"].sum())
        m1, m2 = st.columns(2)
        m1.metric("Días hábiles equivalentes a cubrir", f"{peso_h:.2f}")
        m2.metric("Días calendario", f"{(fin - ini).days + 1}")

        # ---- método de cálculo y ventas especiales ----
        hay_det = ss.get("cmp_ventas_lin") is not None and len(ss.get("cmp_ventas_lin"))
        m1, m2, m3 = st.columns([2.2, 1.2, 1.2])
        _opt_m = ["Promedio simple", "Base ajustada (sin ventas especiales puntuales)"]
        _m = m1.radio("Calcular la venta diaria con", _opt_m, index=1 if _keep("metodo", "simple") == "ajustada" else 0,
                      help="Base ajustada descarta del promedio las ventas especiales que NO marques como «Se repite» (más abajo). "
                           "Requiere subir el Detalle de ventas en 📥 Datos.")
        ss._k_metodo = "ajustada" if _m == _opt_m[1] else "simple"
        ss._k_esp_mult = m2.number_input("Es especial desde (× día normal)", 1.0, 50.0, float(_keep("esp_mult", 3.0)), 0.5,
                                         help="Una venta de un documento que equivale a esta cantidad de días de venta normal del producto.")
        ss._k_esp_min = m3.number_input("y al menos (kg/uni.)", 0.0, 1000.0, float(_keep("esp_min", 5.0)), 1.0,
                                        help="Evita marcar como especiales ventas de productos de muy poco movimiento.")
        _hay_ly = ss.get("cmp_ly_lin") is not None
        t0_, t1_, t2_ = st.columns([1.2, 1.2, 1.2])
        _ini_def = _keep("temp_ini", ini)
        ss._k_temp_ini = t0_.date_input("🍂 Temporada: desde", value=min(max(_ini_def, ini), fin), min_value=ini,
                                        disabled=not _hay_ly, format="DD/MM/YYYY",
                                        help="Primer día del período de temporada. Antes de esta fecha el producto se compra con la cobertura normal; "
                                             "desde aquí hasta la fecha final se compra para repetir lo vendido el año anterior.")
        _fin_def = _keep("temp_fin", fin)
        ss._k_temp_fin = t1_.date_input("🍂 Temporada: hasta", value=max(_fin_def, ss._k_temp_ini),
                                        min_value=ss._k_temp_ini, disabled=not _hay_ly, format="DD/MM/YYYY",
                                        help="Fecha hasta la que quieres cubrir los productos que elijas para compra de temporada "
                                             "(puede ser posterior a la fecha de cobertura general). Se calcula con lo que se vendió el año "
                                             "anterior en esas mismas fechas. Sube el detalle del año anterior en 📥 Datos.")
        ss._k_temp_n = t2_.number_input("Mostrar los N más vendidos", 5, 200, int(_keep("temp_n", 25)), 5, disabled=not _hay_ly)
        if ss._k_metodo == "ajustada" and not hay_det:
            st.warning("Sube el Detalle de ventas en 📥 Datos para usar la base ajustada; por ahora se usa el promedio simple.")
        nec = _nec_completa(ss.cmp_base, peso_h, colchon)
        ss.cmp_nec = nec
        esp = ss.get("cmp_esp")
        if esp is not None and len(esp):
            rep_set = ss.get("cmp_esp_repite", set())
            n_rep = int(esp["Llave"].isin(rep_set).sum())
            with st.expander(f"🔎 Ventas especiales detectadas: {len(esp)} · {esp['Clave'].nunique()} productos · "
                             f"{n_rep} marcadas «Se repite»", expanded=False):
                st.caption("Ventas de un documento que equivalen a varios días de venta normal. Marca **Se repite** en las que esperas "
                           "que vuelvan a ocurrir (quedan en la base); las demás se descuentan de la base ajustada.")
                tipo_c = esp["Tipo doc"].map(lambda t: "Factura" if "FACTURA" in str(t).upper() else ("Boleta" if "BOLETA" in str(t).upper() else str(t)))
                view_e = pd.DataFrame({
                    "Se repite": esp["Llave"].isin(rep_set).values,
                    "Fecha": esp["Fecha"].map(lambda d: f"{d:%d/%m}").values,
                    "Cliente": esp["Cliente"].values, "RUT": esp["RUT"].values, "Doc": tipo_c.values,
                    "Producto": esp["Producto"].values, "Cantidad": esp["Detalle"].values, "Kg/uni.": esp["Kg"].values,
                    "× día normal": esp["x día normal"].values,
                    "Compras del cliente": esp["Docs cliente"].values, "Veces especial": esp["Veces especial"].values,
                    "Llave": esp["Llave"].values})
                with st.form("form_esp"):
                    ed_e = st.data_editor(
                        view_e, hide_index=True, use_container_width=True, key="ed_esp", height=min(420, 60 + 36 * len(view_e)),
                        disabled=[c for c in view_e.columns if c != "Se repite"],
                        column_config={"Llave": None,
                                       "Se repite": st.column_config.CheckboxColumn("Se repite", width="small"),
                                       "Kg/uni.": st.column_config.NumberColumn(format="%.1f"),
                                       "× día normal": st.column_config.NumberColumn(format="%.1f"),
                                       "Compras del cliente": st.column_config.NumberColumn(help="Documentos del cliente en el período del detalle."),
                                       "Veces especial": st.column_config.NumberColumn(help="Cuántas ventas especiales tuvo este cliente.")})
                    ap_e = st.form_submit_button("Aplicar decisiones", type="primary")
                if ap_e:
                    ss.cmp_esp_repite = set(ed_e.loc[ed_e["Se repite"], "Llave"])
                    st.rerun()
        elif hay_det:
            st.caption("No se detectaron ventas especiales con estos criterios.")

        if _hay_ly:
            ini_p, fin_p, sh_p = ss._k_temp_ini, ss._k_temp_fin, timedelta(days=364)
            d0_p, d1_p = ini_p - sh_p, fin_p - sh_p
            ly_df, cob_p = cc.ventas_ly_rango(ss.cmp_ly_lin, ss.cmp_ly_cal, cc.build_formato_map(ss.cmp_formatos), d0_p, d1_p)
            temp = ss.get("cmp_temp", {})
            n_act = int(nec["Temporada"].sum())
            extra_kg = float((nec.loc[nec["Temporada"], "Comprar kg"] - nec.loc[nec["Temporada"], "Comprar normal kg"]).sum())
            with st.expander(f"🍂 Compra de temporada (año anterior): {n_act} producto(s) activados"
                             + (f" · {extra_kg:+,.0f} kg/uni. sobre la compra normal".replace(",", ".") if n_act else ""), expanded=False):
                st.caption(f"Se compara con las ventas del año anterior del **{d0_p:%d/%m/%Y} al {d1_p:%d/%m/%Y}** "
                           f"(las mismas fechas que {ini_p:%d/%m/%Y} al {fin_p:%d/%m/%Y} de este año). "
                           "Los productos que actives se compran para repetir esa venta (más el % extra que elijas) hasta la fecha de arriba; "
                           "todos los demás siguen con la cobertura general. Para los activados no se usa el colchón general: el % extra lo reemplaza.")
                if cob_p < 0.999:
                    dmax_ly = max(ss.cmp_ly_cal["Fecha"])
                    st.warning(f"El detalle del año anterior llega hasta el {dmax_ly:%d/%m/%Y}: cubre el {cob_p:.0%} de esas fechas. "
                               "Lo vendido después no está incluido, así que el «Necesario» queda por debajo de lo real.")
                txt_t = st.text_input("🔎 Buscar producto (muestra también los que no están entre los más vendidos)", key="temp_busca")
                d_t = ly_df.merge(nec[["Clave", "Base", "Tipo", "Stock kg", "Comprar normal kg", "Previo kg", "Venta diaria normal kg"]].rename(columns={"Base": "Prod"}), on="Clave")
                _pw_v = sum(cc.peso_base(ini_p + timedelta(days=i), ss.cmp_feriados) for i in range((fin_p - ini_p).days + 1))
                d_t["Normal"] = d_t["Venta diaria normal kg"] * _pw_v
                d_t["Aumento"] = d_t["Kg"] - d_t["Normal"]
                d_t["Veces"] = d_t["Kg"] / d_t["Normal"].where(d_t["Normal"] > 0)
                o1_, o2_ = st.columns([2, 1])
                orden_t = o1_.radio("Ordenar por", ["Mayor aumento vs. venta normal (kg)", "Mayor aumento (veces)", "Más vendidos año ant."],
                                    horizontal=True, key="temp_orden",
                                    help="«Venta normal» = lo que vendes hoy por día (últimas 4 semanas) llevado a esas fechas. "
                                         "El aumento es lo vendido el año pasado menos esa venta normal: así aparecen los productos que SÍ suben en temporada, "
                                         "aunque no sean los más vendidos.")
                min_t = o2_.number_input("Vendido año ant. mínimo (kg/uni.)", 0.0, 10000.0, 5.0, 1.0, key="temp_min",
                                         help="Oculta productos con muy poca venta, para que un salto de 1 a 3 kg no aparezca como «3 veces».")
                d_t = d_t[d_t["Kg"] >= min_t]
                if orden_t.startswith("Mayor aumento vs"):
                    d_t = d_t.sort_values("Aumento", ascending=False)
                elif orden_t.startswith("Mayor aumento ("):
                    d_t = d_t.sort_values("Veces", ascending=False, na_position="last")
                else:
                    d_t = d_t.sort_values("Kg", ascending=False)
                if txt_t.strip():
                    d_t = d_t[d_t["Prod"].str.contains(txt_t.strip(), case=False, na=False)].head(60)
                else:
                    d_t = d_t.head(int(ss._k_temp_n))
                if d_t.empty:
                    st.info("No hay productos para mostrar con estos filtros.")
                else:
                    extra_def = 10.0
                    ex_ = [float(temp.get(c, extra_def)) for c in d_t["Clave"]]
                    nec_t = [max(0.0, pr_ + k * (1 + e / 100.0) - s_) for k, e, s_, pr_ in zip(d_t["Kg"], ex_, d_t["Stock kg"], d_t["Previo kg"])]
                    view_t = pd.DataFrame({
                        "Usar": [c in temp for c in d_t["Clave"]],
                        "Producto": d_t["Prod"].values, "Tipo": d_t["Tipo"].values,
                        "Vendido año ant.": d_t["Kg"].values, "Venta normal": d_t["Normal"].values, "Veces": d_t["Veces"].values,
                        "Stock actual": d_t["Stock kg"].values,
                        "Solicitado hoy": d_t["Comprar normal kg"].values, "Extra %": ex_,
                        "Necesario temporada": nec_t,
                        "Diferencia": [n_ - sol for n_, sol in zip(nec_t, d_t["Comprar normal kg"])],
                        "Clave": d_t["Clave"].values})
                    with st.form("form_temp"):
                        ed_t = st.data_editor(
                            view_t, hide_index=True, use_container_width=True, key="ed_temp", height=min(460, 60 + 36 * len(view_t)),
                            disabled=[c for c in view_t.columns if c not in ("Usar", "Extra %")],
                            column_config={"Clave": None,
                                           "Usar": st.column_config.CheckboxColumn("Usar temporada", width="small"),
                                           "Vendido año ant.": st.column_config.NumberColumn("Vendido año ant. kg/uni.", format="%.1f",
                                                                                              help="Lo que se vendió en esas fechas el año anterior."),
                                           "Venta normal": st.column_config.NumberColumn("Venta normal esperada kg/uni.", format="%.1f",
                                                                                          help="Lo que venderías en esas fechas al ritmo de hoy."),
                                           "Veces": st.column_config.NumberColumn("Sube ×", format="%.1f×",
                                                                                   help="Vendido año ant. ÷ venta normal esperada. 3× = se vendió el triple que lo normal."),
                                           "Stock actual": st.column_config.NumberColumn("Stock actual kg/uni.", format="%.1f"),
                                           "Solicitado hoy": st.column_config.NumberColumn("Solicitado hoy kg/uni.", format="%.1f",
                                                                                            help="Lo que pide la compra normal (cobertura general)."),
                                           "Extra %": st.column_config.NumberColumn("Extra %", min_value=-50.0, max_value=300.0, step=5.0, format="%.0f",
                                                                                     help="Porcentaje que sumas a lo vendido el año anterior por si crees que subirá la venta."),
                                           "Necesario temporada": st.column_config.NumberColumn("Necesario kg/uni.", format="%.1f",
                                                                                                help="Vendido año ant. × (1 + extra) − stock actual."),
                                           "Diferencia": st.column_config.NumberColumn("Diferencia kg/uni.", format="%+.1f",
                                                                                       help="Necesario − Solicitado hoy.")})
                        ap_t = st.form_submit_button("Aplicar temporada", type="primary")
                    st.caption("Marca **Usar temporada** en los productos que quieres cubrir como el año pasado y ajusta su **Extra %**; luego pulsa Aplicar. "
                               "La compra de esos productos pasa a usar la columna *Necesario*.")
                    if ap_t:
                        nuevo = dict(temp)
                        for c_, u_, e_ in zip(ed_t["Clave"], ed_t["Usar"], ed_t["Extra %"]):
                            if u_:
                                nuevo[c_] = float(e_) if pd.notna(e_) else extra_def
                            else:
                                nuevo.pop(c_, None)
                        ss.cmp_temp = nuevo
                        st.rerun()

        f1, f2, f3 = st.columns([2, 2, 1])
        tipos = sorted(nec["Tipo"].dropna().unique())
        sel_t = f1.multiselect("Tipo de producto", tipos, default=[t for t in tipos if t.startswith("MP.") or t.startswith("ReV.")] or tipos)
        txt = f2.text_input("🔎 Buscar producto")
        solo = f3.toggle("Sólo a comprar", True)
        v = nec[nec["Tipo"].isin(sel_t)]
        if txt.strip():
            v = v[v["Base"].str.contains(txt.strip(), case=False, na=False)]
        if solo:
            v = v[v["Comprar kg"] > 0]
        show = v[["Base", "Tipo", "Temporada", "Stock kg", "Venta semanal kg", "Venta diaria simple", "Venta diaria ajustada",
                  "Especiales kg", "Demanda kg", "Comprar kg", "Cobertura días hábiles"]].copy()
        show["Temporada"] = show["Temporada"].map(lambda t: "🍂" if t else "")
        if not ss.get("cmp_temp"):
            show = show.drop(columns=["Temporada"])
        if ss.get("cmp_recetas") is not None and len(ss.cmp_recetas):
            show["Por procesados kg"] = v["Por procesados kg"]
        if ss.get("cmp_stock_base"):
            show.insert(2, "Alerta stock", v["Alerta stock"])
            for c_ in ("Stock previsto kg", "Desvío stock kg", "Comprar antes kg"):
                show[c_] = v[c_]
        show["Cobertura días hábiles"] = show["Cobertura días hábiles"].replace(float("inf"), None)
        st.dataframe(show.sort_values("Comprar kg", ascending=False), hide_index=True, use_container_width=True,
                     column_config={c: st.column_config.NumberColumn(
                         {"Stock kg": "Stock Actual kg/uni.", "Venta diaria simple": "Venta diaria simple kg/uni.",
                          "Venta diaria ajustada": "Venta diaria ajustada kg/uni.", "Especiales kg": "Especiales kg/uni.",
                          "Cobertura días hábiles": "Cobertura días hábiles"}.get(c, c.replace(" kg", " kg/uni.")), format="%.2f")
                                    for c in show.columns if c not in ("Base", "Tipo", "Temporada", "Alerta stock")})
        st.caption(f"{len(v)} productos · Demanda = venta diaria (**{'ajustada' if ss._k_metodo == 'ajustada' else 'simple'}**) × días hábiles × (1 + colchón). "
                   "Kg/uni.: los productos por kilo van en kg y los que se venden por unidad (ej. aceites, cajas de 6u) en unidades. "
                   "Los productos se agrupan por nombre base (sin formato), igual que tu hoja *Análisis Actual*.")

# =====================================================================================
# TAB 3 - PLAN DE COMPRA  (lista + panel de decisión, sólo se recalcula este bloque)
# =====================================================================================
_fragment = getattr(st, "fragment", None) or getattr(st, "experimental_fragment", None) or (lambda f: f)


def _color_precio(t):
    """0 = más barato (verde) ... 1 = más caro (rojo #fc0d1b)."""
    return "#fc0d1b" if t >= 0.999 else f"hsl({120 * (1 - t):.0f},{85 + 13 * t:.0f}%,{34 + 18 * t:.0f}%)"


def _num(x, dec=0):
    t = f"{float(x):,.{dec}f}"
    return t.replace(",", "X").replace(".", ",").replace("X", ".")


def _pn_cache():
    """Precios normalizados + agrupados por producto: se calculan una vez por cambio de la base."""
    key = (id(ss.cmp_precios), len(fmap))
    if ss.get("cmp_pn_key") != key:
        pn = cc.normaliza_precios(ss.cmp_precios, fmap)
        ss.cmp_pn, ss.cmp_gr, ss.cmp_pn_key = pn, cc.agrupa_precios(pn), key
        ss.pop("cmp_plan_cache", None)
    return ss.cmp_pn, ss.cmp_gr


def _plan_actual(crit):
    """Plan completo con las elecciones manuales. La base se calcula una vez; cada clic sólo recalcula esa fila."""
    pn, gr = _pn_cache()
    nec = ss.cmp_nec
    ck = (id(nec), ss.cmp_pn_key, crit)
    if ss.get("cmp_plan_cache_key") != ck or "cmp_plan_cache" not in ss:
        ss.cmp_plan_cache = cc.plan_de_compra(nec, pn, {}, criterio=crit, grupos=gr).set_index("Clave", drop=False)
        ss.cmp_plan_cache_key = ck
    plan = ss.cmp_plan_cache.copy()
    el = {k: v for k, v in ss.cmp_elegido.items() if k in plan.index}
    stale = set(el) | {k for k in ss.get("cmp_stale", set()) if k in plan.index}
    if stale:
        sub = cc.plan_de_compra(nec[nec["Clave"].isin(stale)], pn, el, criterio=crit, grupos=gr).set_index("Clave", drop=False)
        plan.loc[sub.index, sub.columns] = sub
    for k, (prov, fmt, b) in list(ss.cmp_bultos.items()):     # cantidades fijadas a mano
        if k in plan.index and plan.at[k, "Opciones"] > 0 and plan.at[k, "Proveedor"] == prov:
            kb = float(plan.at[k, "Kg bulto"])
            plan.at[k, "Bultos"] = int(b)
            plan.at[k, "Kg total"] = round(b * kb, 2)
            plan.at[k, "Total Neto"] = round(b * kb * float(plan.at[k, "Precio Neto Kg"]), 0)
    return plan.reset_index(drop=True)


def _cb_set_bultos(clave, prov, fmt, wkey):
    ss.cmp_bultos[clave] = (prov, fmt, int(ss[wkey]))


def _cb_reset_bultos(clave):
    ss.cmp_bultos.pop(clave, None)


def _cb_edit_oferta(clave, prov, fmt_old, pk, fk):
    precio, fmt_new = float(ss[pk]), ss[fk] or ""
    if precio <= 0:
        ss.cmp_msg = (False, "El precio debe ser mayor que 0.")
        return
    db = ss.cmp_precios.copy().astype(object)
    m = (db["Producto"].map(cc.norm) == clave) & (db["Proveedor"] == prov) & (db["Formato"].astype(str) == str(fmt_old))
    db.loc[m, "Precio Neto Kg"] = precio
    db.loc[m, "Formato"] = fmt_new
    db.loc[m, "Actualizado"] = date.today().isoformat()
    ss.cmp_precios = db
    if ss.cmp_elegido.get(clave) == (prov, fmt_old):
        ss.cmp_elegido[clave] = (prov, fmt_new)
    ss.cmp_bultos.pop(clave, None)
    ss.cmp_dirty = True
    ss.cmp_msg = (True, f"Oferta de {prov} actualizada.")


def _cb_elegir(clave, prov, fmt, sig):
    ss.cmp_bultos.pop(clave, None)
    ss.cmp_elegido[clave] = (prov, fmt)
    ss.cmp_rev.add(clave)
    ss.cmp_cur = sig


def _cb_ir(clave):
    ss.cmp_cur = clave


def _cb_omitir(clave, sig):
    ss[f"inc_{clave}"] = False
    ss.cmp_excl.add(clave)
    ss.cmp_rev.add(clave)
    ss.cmp_cur = sig


def _cb_quiebre(clave, prov, fmt):
    db = ss.cmp_precios.copy().astype(object)
    m = (db["Producto"].map(cc.norm) == clave) & (db["Proveedor"] == prov) & (db["Formato"].astype(str) == str(fmt))
    db.loc[m, "Disponible"] = "No"
    db.loc[m, "Actualizado"] = date.today().isoformat()
    ss.cmp_precios = db
    # actualización incremental (evita recalcular los ~500 productos)
    pn = ss.cmp_pn
    pn.loc[(pn["Clave"] == clave) & (pn["Proveedor"] == prov) & (pn["Formato"].astype(str) == str(fmt)), "Disponible"] = False
    nuevo = cc.agrupa_precios(pn[pn["Clave"] == clave])
    if clave in nuevo:
        ss.cmp_gr[clave] = nuevo[clave]
    else:
        ss.cmp_gr.pop(clave, None)
    ss.cmp_pn_key = (id(ss.cmp_precios), len(fmap))
    ss.setdefault("cmp_stale", set()).add(clave)
    ss.cmp_bultos.pop(clave, None)
    ss.cmp_elegido.pop(clave, None)
    ss.cmp_dirty = True


def _cb_pedir_del(clave, prov, fmt):
    ss.cmp_del = (clave, prov, fmt)


def _cb_cancel_del():
    ss.cmp_del = None


def _cb_confirm_del():
    clave, prov, fmt = ss.cmp_del
    db = ss.cmp_precios
    m = (db["Producto"].map(cc.norm) == clave) & (db["Proveedor"] == prov) & (db["Formato"].astype(str) == str(fmt))
    ss.cmp_precios = db[~m].reset_index(drop=True)      # objeto nuevo -> se reconstruye el cache de precios
    ss.cmp_elegido.pop(clave, None)
    ss.cmp_bultos.pop(clave, None)
    ss.cmp_dirty = True
    ss.cmp_del = None
    ss.cmp_msg = (True, f"Oferta de {prov} eliminada.")


def _cb_guardar_quiebres():
    _cmp_save(cc.SH_PRECIOS, ss.cmp_precios[cc.COLS_PRECIOS])
    ss.cmp_dirty = False
    st.toast("✅ Cambios guardados en Google Sheets", icon="✅")


def _cb_add_plan(producto):
    prov = ss.get("af_new", "").strip() or ss.get("af_sel") or ""
    ok, msg = _agregar_oferta(producto, prov, ss.get("af_fmt", ""), ss.get("af_precio", 0.0))
    ss.cmp_msg = (ok, msg)


def _cb_incluir(clave):
    if ss.get(f"inc_{clave}", True):
        ss.cmp_excl.discard(clave)
    else:
        ss.cmp_excl.add(clave)


@_fragment
def _recalc_nec():
    cf = ss.get("cmp_cal_fut")
    if "cmp_base" in ss and cf is not None and len(cf):
        ss.cmp_nec = _nec_completa(ss.cmp_base, float(cf.loc[cf["Incluir"], "Peso"].sum()), float(ss.get("_k_col", 10.0)))
        ss.pop("cmp_plan_cache_key", None)


def _panel_no_considerados(expandido=False):
    nec = ss.cmp_nec
    _, gr = _pn_cache()
    b = nec[(nec["Comprar base kg"] <= 0) | nec["Siempre"]].copy()
    vend, stk = b["Vendido kg"], b["Stock kg"]
    b["Motivo"] = np.where(b["Se produce"], "Se produce (receta)",
                           np.where(b["Comprar base kg"] > 0, "Ya se compra normal",
                                    np.where(vend <= 0, np.where(stk <= 0, "Sin venta y sin stock", "Sin venta (hay stock)"), "Stock suficiente")))
    b["Oferta"] = b["Clave"].isin(set(gr.keys()))
    n_sg = int(nec["Siempre"].sum())
    with st.expander(f"🔍 Productos no considerados en la compra ({int((nec['Comprar base kg'] <= 0).sum())}) · "
                     f"{n_sg} agregado(s) a este análisis", expanded=expandido):
        st.caption("Aquí están los productos que hoy NO entran al plan (sin venta, sin stock o con stock suficiente). "
                   "Marca **Considerar** para incluirlos en el plan de **este análisis**: la cantidad (parte en 1; cámbiala si necesitas otra) "
                   "se pide como mínimo, aunque no tengan venta (útil si el proveedor estaba en quiebre). "
                   "Es temporal: se guarda con 💾 Guardar análisis y desaparece cuando subes un Reposition nuevo.")
        k1, k2, k3 = st.columns([2, 2, 1.5])
        mots = ["Sin venta y sin stock", "Sin venta (hay stock)", "Stock suficiente", "Ya se compra normal", "Se produce (receta)"]
        sel_m = k1.multiselect("Motivo", mots, default=mots[:3], key="nc_mot")
        txt_n = k2.text_input("🔎 Buscar producto", key="nc_txt")
        solo_of = k3.toggle("Sólo con proveedor", True, key="nc_of", help="Oculta productos que no tienen ninguna oferta disponible en la base de precios.")
        v = b[b["Motivo"].isin(sel_m)]
        if solo_of:
            v = v[v["Oferta"] | v["Siempre"]]
        if txt_n.strip():
            v = v[v["Base"].str.contains(txt_n.strip(), case=False, na=False)]
        v = v.assign(_o=~v["Siempre"]).sort_values(["_o", "Motivo", "Base"]).head(300)
        if v.empty:
            st.info("No hay productos con estos filtros.")
            return
        sg = ss.get("cmp_siempre")
        cant_sg = dict(zip(sg["Clave"], pd.to_numeric(sg["Cantidad"], errors="coerce").fillna(0.0))) if sg is not None and len(sg) else {}
        view = pd.DataFrame({"Siempre": v["Siempre"].values, "Producto": v["Base"].values, "Tipo": v["Tipo"].values,
                             "Motivo": v["Motivo"].values, "Stock": v["Stock kg"].values, "Venta semanal": v["Venta semanal kg"].values,
                             "Proveedor": ["Sí" if o else "No" for o in v["Oferta"]],
                             "Cantidad": [float(cant_sg.get(c, 1.0)) for c in v["Clave"]], "Clave": v["Clave"].values})
        with st.form("form_siempre"):
            ed = st.data_editor(view, hide_index=True, use_container_width=True, key="ed_siempre", height=min(460, 60 + 36 * len(view)),
                                disabled=["Producto", "Tipo", "Motivo", "Stock", "Venta semanal", "Proveedor"],
                                column_config={"Clave": None, "Siempre": st.column_config.CheckboxColumn("Considerar", width="small"),
                                               "Stock": st.column_config.NumberColumn("Stock kg/uni.", format="%.1f"),
                                               "Venta semanal": st.column_config.NumberColumn("Venta semanal kg/uni.", format="%.1f"),
                                               "Cantidad": st.column_config.NumberColumn("Cantidad a pedir kg/uni.", min_value=0.0, step=1.0, format="%.1f",
                                                                                          help="Mínimo que se pedirá en cada análisis (en kg, o unidades si el producto es por unidad).")})
            ok = st.form_submit_button("Guardar y recalcular", type="primary")
        if ok:
            _ok_q = lambda q_: pd.notna(q_) and q_ > 0
            nuevo = {r_.Clave: (r_.Producto, float(r_.Cantidad) if _ok_q(r_.Cantidad) else 1.0) for r_ in ed.itertuples() if r_.Siempre}
            sin_cant = [r_.Producto for r_ in ed.itertuples() if r_.Siempre and not _ok_q(r_.Cantidad)]
            mostrados = set(ed["Clave"])
            filas = [{"Producto": nm, "Clave": k_, "Cantidad": q_, "Nota": ""} for k_, (nm, q_) in nuevo.items()]
            if sg is not None and len(sg):
                for r_ in sg.itertuples():
                    if r_.Clave not in mostrados:
                        filas.append({"Producto": r_.Producto, "Clave": r_.Clave, "Cantidad": r_.Cantidad, "Nota": getattr(r_, "Nota", "")})
            df_new = pd.DataFrame(filas, columns=cc.COLS_SIEMPRE)
            try:
                ss.cmp_siempre = df_new
                _recalc_nec()
                ss.cmp_msg_sg = (f"✅ Listo: {len(df_new)} producto(s) agregados a este análisis." +
                                 (f" Sin cantidad: se usó 1 para {', '.join(sin_cant[:5])}; ajústala en la tabla si necesitas otra." if sin_cant else ""))
            except Exception as e:
                ss.cmp_msg_sg = f"❌ No pude aplicar el cambio: {e}"
            st.rerun()


def _plan_panel():
    ss.setdefault("cmp_elegido", {})
    ss.setdefault("cmp_rev", set())
    ss.setdefault("cmp_excl", set())
    ss.setdefault("cmp_dirty", False)

    f1, f2, f3 = st.columns([3, 2, 2])
    crit_txt = f1.radio("Sugerir proveedor por", ["Precio por kg", "Costo total de la compra"], horizontal=True,
                        index=["Precio por kg", "Costo total de la compra"].index(_keep("crit", "Precio por kg")),
                        help="'Costo total' considera el redondeo a bultos completos: a veces el más barato por kg no es el más barato para la cantidad que necesitas.")
    crit = "total" if crit_txt.startswith("Costo") else "kg"
    ss._k_crit = crit_txt
    _vers = ["Todos", "Pendientes", "Sin proveedor"] + (["Stock bajo lo previsto"] if ss.get("cmp_stock_base") else [])
    ver = f2.radio("Mostrar", _vers, horizontal=True, index=_vers.index(_keep("ver", "Todos")) if _keep("ver", "Todos") in _vers else 0)
    if "plan_busca" not in ss:
        ss.plan_busca = _keep("busca", "")
    _tb, _bb = f3.columns([5, 1.4], vertical_alignment="bottom")
    txt = _tb.text_input("🔎 Buscar producto", key="plan_busca")
    _bb.button("✖", key="plan_busca_x", help="Limpiar la búsqueda", on_click=lambda: ss.__setitem__("plan_busca", ""))
    ss._k_ver, ss._k_busca = ver, txt

    plan_all = _plan_actual(crit)
    _nn = ss.cmp_nec.set_index("Clave")
    _alerta_map = _nn["Alerta stock"].to_dict() if "Alerta stock" in _nn else {}
    _antes_map = _nn["Comprar antes kg"].to_dict() if "Comprar antes kg" in _nn else {}
    _sg_set = set(_nn.index[_nn["Siempre"]]) if "Siempre" in _nn else set()
    _pp_map = _nn["Por procesados kg"].to_dict() if "Por procesados kg" in _nn else {}
    _tp_set = set(_nn.index[_nn["Temporada"]]) if "Temporada" in _nn else set()
    tipos = sorted(plan_all["Tipo"].unique())
    with st.expander("Filtrar por tipo de producto"):
        _def = None if "plan_tipos" in ss else ([t for t in tipos if t.startswith("MP.") or t.startswith("ReV.")] or tipos)
        sel_t = st.multiselect("Tipos", tipos, key="plan_tipos", default=_def)
    plan = plan_all[plan_all["Tipo"].isin(sel_t) | plan_all["Clave"].isin(_sg_set)]     # los agregados a mano se ven aunque su tipo esté filtrado
    if txt.strip():
        plan = plan[plan["Producto"].str.contains(txt.strip(), case=False, na=False)]
    if ver == "Pendientes":
        plan = plan[~plan["Clave"].isin(ss.cmp_rev)]
    elif ver == "Sin proveedor":
        plan = plan[plan["Estado"] == "Sin proveedor"]
    elif ver == "Stock bajo lo previsto":
        plan = plan[plan["Clave"].map(_alerta_map).eq("🔴")]
    plan = plan.reset_index(drop=True)

    if ss.cmp_dirty:
        w1, w2 = st.columns([4, 1])
        w1.warning("Hay cambios (quiebres / proveedores nuevos) sin guardar en Google Sheets. Ya se usan en este plan.")
        w2.button("💾 Guardar cambios", on_click=_cb_guardar_quiebres, use_container_width=True)

    if plan.empty:
        st.success("🎉 No hay productos para mostrar con estos filtros (o el stock cubre los días elegidos).")
        ss.cmp_plan_sel = plan
        return

    claves = plan["Clave"].tolist()
    cur = ss.get("cmp_cur")
    if cur not in claves:
        cur = claves[0]

    izq, der = st.columns([5, 4], gap="large")

    # ---------------- Lista (izquierda) ----------------
    with izq:
        vista = plan[["Producto", "Comprar kg", "Proveedor", "Formato", "Bultos", "Total Neto"]].copy()
        vista["Producto"] = [("🍂 " if c_ in _tp_set else "") + ("📌 " if c_ in _sg_set else "") + ("🧪 " if _pp_map.get(c_, 0) > 0 else "") + p_
                             for p_, c_ in zip(plan["Producto"], plan["Clave"])]
        vista.insert(0, "✓", ["✔" if c in ss.cmp_rev else "" for c in plan["Clave"]])
        vista.insert(1, "Incl.", ["✖" if c in ss.cmp_excl else "●" for c in plan["Clave"]])
        vista = vista.rename(columns={"Comprar kg": "Comprar kg/uni."})
        if ss.get("cmp_stock_base"):
            vista.insert(2, "⚠", [_alerta_map.get(c, "") for c in plan["Clave"]])
        pos_cur = claves.index(cur)

        def _est(df):
            e = pd.DataFrame("", index=df.index, columns=df.columns)
            e.loc[pos_cur, :] = "background-color: rgba(80,130,200,0.22)"
            sinp = plan["Estado"].eq("Sin proveedor").values
            for i in df.index[sinp]:
                e.at[i, "Proveedor"] = "color: #fc0d1b"
            return e
        sty = vista.style.apply(_est, axis=None).format(
            {"Comprar kg/uni.": lambda v: _num(v, 1), "Total Neto": lambda v: "$" + _num(v)})
        ev = st.dataframe(sty, hide_index=True, use_container_width=True, height=700, key="plan_tbl",
                          on_select="rerun", selection_mode="single-row",
                          column_config={"✓": st.column_config.TextColumn(width="small", help="Ya revisado"),
                                         "Incl.": st.column_config.TextColumn(width="small", help="● se incluye en la orden · ✖ omitido"),
                                         "⚠": st.column_config.TextColumn(width="small", help="🔴 el stock bajó más de lo previsto · 🟢 quedó más stock del previsto"),
                                         "Producto": st.column_config.TextColumn(width="large")})
        rows = list(ev.selection.rows) if ev is not None else []
        if rows and rows != ss.get("cmp_last_rows") and rows[0] < len(claves):
            cur = claves[rows[0]]
            ss.cmp_cur = cur
            pos_cur = rows[0]
        ss.cmp_last_rows = rows
        st.caption("Haz clic en una fila para revisarla a la derecha. Con **Elegir y seguir** avanzas al siguiente producto.")

    # ---------------- Panel de decisión (derecha) ----------------
    with der:
        r = plan[plan["Clave"] == cur].iloc[0]
        sig = claves[pos_cur + 1] if pos_cur + 1 < len(claves) else cur
        ant = claves[pos_cur - 1] if pos_cur > 0 else cur
        st.markdown(f"#### {r['Producto']}")
        st.caption(f"Producto {pos_cur + 1} de {len(claves)} · {r['Tipo']}")
        m1, m2, m3 = st.columns(3)
        m1.metric("Stock Actual kg/uni.", _num(r["Stock kg"], 1))
        m2.metric("Venta semanal kg/uni.", _num(r["Venta semanal kg"], 1))
        m3.metric("Comprar", _num(r["Comprar kg"], 1))
        if cur in _tp_set:
            _t1 = _nn.loc[cur]
            st.info(f"🍂 Compra de temporada: se calcula para repetir lo vendido el año anterior (+{float(ss.get('cmp_temp', {}).get(cur, 0)):.0f} %) "
                    f"hasta el {ss.get('_k_temp_fin'):%d/%m/%Y}. Con la compra normal se pedirían {_num(_t1['Comprar normal kg'], 1)} kg/uni.; "
                    f"con temporada, {_num(_t1['Comprar kg'], 1)}.")
        if _pp_map.get(cur, 0) > 0:
            st.info(f"🧪 Incluye {_num(_pp_map[cur], 1)} kg/uni. de demanda por productos procesados que usan este ingrediente (⚙️ Configuración → Procesados).")
        if cur in _sg_set:
            st.info("📌 Agregado a mano a este análisis: se pide como mínimo la cantidad que definiste en «Productos no considerados».")
        if _alerta_map.get(cur) in ("🔴", "🟢"):
            _n1 = _nn.loc[cur]
            _ant = _antes_map.get(cur)
            _tx = (f"Stock previsto {_num(_n1['Stock previsto kg'], 1)} · real {_num(_n1['Stock kg'], 1)} "
                   f"({_n1['Desvío stock kg']:+.1f}). " + (f"Antes se pedía {_num(_ant, 1)}, ahora {_num(_n1['Comprar kg'], 1)}." if pd.notna(_ant) else ""))
            (st.warning if _alerta_map[cur] == "🔴" else st.info)(("🔴 El stock bajó más de lo previsto. " if _alerta_map[cur] == "🔴"
                                                                    else "🟢 Quedó más stock del previsto. ") + _tx)

        if ss.get("cmp_msg"):
            (st.success if ss.cmp_msg[0] else st.error)(ss.cmp_msg[1])
            ss.cmp_msg = None
        ops = cc.opciones_producto(cur, ss.cmp_pn, r["Comprar kg"], ss.cmp_gr)
        pend = ss.get("cmp_del")
        if pend and pend[0] == cur:
            st.warning(f"¿Eliminar la oferta de **{pend[1]}** ({pend[2] or 'sin formato'}) para este producto? "
                       "Se borra de la base de proveedores.")
            d1, d2 = st.columns(2)
            d1.button("Sí, eliminar", on_click=_cb_confirm_del, type="primary", use_container_width=True)
            d2.button("Cancelar", on_click=_cb_cancel_del, use_container_width=True)
        if ops.empty:
            st.warning("Ningún proveedor con stock y precio para este producto. Agrégalo abajo, en ➕ Agregar proveedor para este producto.")
        else:
            pos = cc.posicion_elegida(ops, ss.cmp_elegido.get(cur), crit)
            pmin, pmax = ops["Precio Neto Kg"].min(), ops["Precio Neto Kg"].max()
            h = st.columns([3, 2, 2, 1.2, 2, 1.1, 1.5, 0.8])
            for c, t in zip(h, ["Proveedor", "Formato", "Precio kg/uni.", "Bultos", "Total neto", "Dif.", "", ""]):
                c.caption(t)
            for i, o in ops.iterrows():
                t_ = 0.0 if pmax == pmin else (o["Precio Neto Kg"] - pmin) / (pmax - pmin)
                es_sel = i == pos
                c = st.columns([3, 2, 2, 1.2, 2, 1.1, 1.5, 0.8])
                fmt_txt = o["Formato"] if not o["Sin formato"] else "(kg)"
                c[0].markdown(("✅ **" + o["Proveedor"] + "**") if es_sel else o["Proveedor"])
                c[1].write(fmt_txt)
                c[2].markdown(f"<span style='color:{_color_precio(t_)}'>${_num(o['Precio Neto Kg'], 2 if abs(o['Precio Neto Kg'] - round(o['Precio Neto Kg'])) > 0.005 else 0)}</span>",
                              unsafe_allow_html=True)
                c[3].write(int(o["Bultos"]))
                c[4].write("$" + _num(o["Total Neto"]))
                c[5].write("—" if o["Dif %"] < 0.5 else f"+{o['Dif %']:.0f}%")
                c[6].button("Elegir", key=f"el_{cur}_{i}", on_click=_cb_elegir, args=(cur, o["Proveedor"], o["Formato"], sig),
                            type="primary" if es_sel else "secondary", use_container_width=True,
                            help="Elegir este proveedor y pasar al siguiente producto")
                c[7].button("❌", key=f"qb_{cur}_{i}", on_click=_cb_quiebre, args=(cur, o["Proveedor"], o["Formato"]),
                            use_container_width=True,
                            help="Sin stock hoy: marca este proveedor en quiebre (temporal) y pasa al siguiente mejor")
            st.caption("❌ = sin stock hoy (temporal) · ✅ = opción vigente · Para borrar una oferta usa ✏️ Editar precio / formato.")

            # ---- cantidad a comprar (editable) ----
            o_sel = ops.iloc[pos]
            sug = int(o_sel["Bultos"])
            ov = ss.cmp_bultos.get(cur)
            hay = bool(ov and ov[0] == o_sel["Proveedor"] and ov[1] == o_sel["Formato"])
            val = int(ov[2]) if hay else sug
            wk = f"bl_{cur}_{pos}_{o_sel['Proveedor']}_{o_sel['Formato']}_{sug}_{val}"
            unidad = "kg" if o_sel["Sin formato"] else "bultos"
            q1, q2 = st.columns([1, 2])
            q1.number_input(f"Cantidad a comprar ({unidad})", min_value=0, step=1, value=val, key=wk,
                            on_change=_cb_set_bultos, args=(cur, o_sel["Proveedor"], o_sel["Formato"], wk))
            kg_v = val * float(o_sel["Kg bulto"])
            q2.markdown(f"<div style='padding-top:1.9rem'>Sugerido: <b>{sug}</b> · {_num(kg_v, 1)} kg/uni. · "
                        f"<b>${_num(kg_v * float(o_sel['Precio Neto Kg']))}</b> neto</div>", unsafe_allow_html=True)
            if hay:
                st.button("↺ Volver a la cantidad sugerida", on_click=_cb_reset_bultos, args=(cur,))

            # ---- editar precio / formato de una oferta ----
            with st.expander("✏️ Editar precio / formato de un proveedor"):
                idxs = list(range(len(ops)))

                def _lab(i):
                    f_ = "(sin formato)" if ops.iloc[i]["Sin formato"] else ops.iloc[i]["Formato"]
                    return f"{ops.iloc[i]['Proveedor']} · {f_}"
                e_i = st.selectbox("Oferta", idxs, index=min(pos, len(idxs) - 1), format_func=_lab, key=f"eo_{cur}")
                e_o = ops.iloc[e_i]
                fm_e = [""] + sorted(set(df_formatos_db["Formato"].astype(str)) - {""})
                cur_fmt = "" if e_o["Sin formato"] else str(e_o["Formato"])
                if cur_fmt not in fm_e:
                    fm_e.append(cur_fmt)
                pk = f"ep_{cur}_{e_i}_{e_o['Proveedor']}_{e_o['Precio Neto Kg']}"
                fk = f"ef_{cur}_{e_i}_{e_o['Proveedor']}_{cur_fmt}"
                e1, e2 = st.columns(2)
                e1.number_input("Precio neto por kg/uni.", min_value=0.0, step=10.0, format="%.2f",
                                value=float(e_o["Precio Neto Kg"]), key=pk)
                e2.selectbox("Formato", fm_e, index=fm_e.index(cur_fmt), key=fk)
                a1_, a2_ = st.columns(2)
                a1_.button("💾 Aplicar cambio", on_click=_cb_edit_oferta, args=(cur, e_o["Proveedor"], e_o["Formato"], pk, fk),
                           type="primary", use_container_width=True)
                a2_.button("🗑️ Eliminar esta oferta", key=f"dl_{cur}_{e_i}", on_click=_cb_pedir_del,
                           args=(cur, e_o["Proveedor"], e_o["Formato"]), use_container_width=True,
                           help="Borra esta oferta de la base de proveedores (pide confirmación)")

        with st.expander("➕ Agregar proveedor para este producto", expanded=ops.empty):
            provs_ex = sorted(ss.cmp_precios["Proveedor"].astype(str).str.strip().replace("", pd.NA).dropna().unique())
            fm_ops = [""] + sorted(set(df_formatos_db["Formato"].astype(str)) - {""})
            with st.form("form_add_plan", clear_on_submit=True):
                g1, g2 = st.columns(2)
                g1.selectbox("Proveedor existente", provs_ex, index=None, placeholder="Elige…", key="af_sel")
                g2.text_input("…o proveedor nuevo", key="af_new")
                g3, g4 = st.columns(2)
                g3.selectbox("Formato", fm_ops, key="af_fmt")
                g4.number_input("Precio neto por kg/uni.", min_value=0.0, step=10.0, format="%.2f", key="af_precio")
                st.form_submit_button("Agregar", type="primary", on_click=_cb_add_plan, args=(r["Producto"],))
        st.checkbox("Incluir en la orden", value=cur not in ss.cmp_excl, key=f"inc_{cur}", on_change=_cb_incluir, args=(cur,))
        b1, b2, b3, b4 = st.columns(4)
        b1.button("◀ Anterior", on_click=_cb_ir, args=(ant,), use_container_width=True)
        b2.button("✔ Aceptar y seguir", on_click=lambda: (ss.cmp_rev.add(cur), setattr(ss, "cmp_cur", sig)),
                  type="primary", use_container_width=True, help="Deja la opción vigente y pasa al siguiente")
        b3.button("Siguiente ▶", on_click=_cb_ir, args=(sig,), use_container_width=True)
        b4.button("🚫 Omitir", on_click=_cb_omitir, args=(cur, sig), use_container_width=True, help="No comprar este producto ahora")

    # ---------------- Totales ----------------
    inc = plan[~plan["Clave"].isin(ss.cmp_excl) & plan["Estado"].isin(["OK", "Falta formato (kg)"]) & (plan["Bultos"] > 0)]
    ss.cmp_plan_sel = inc.copy()
    neto = float(inc["Total Neto"].sum())
    k1, k2, k3, k4, k5 = st.columns(5)
    k1.metric("Líneas en la orden", len(inc))
    k2.metric("Revisadas", f"{len(set(claves) & ss.cmp_rev)}/{len(claves)}")
    k3.metric("Neto", "$" + _num(neto))
    _pm = _prov_map()
    iva = float(inc.loc[inc["Proveedor"].map(lambda p: _pm.get(cc.norm(p), {}).get("factura", True)), "Total Neto"].sum()) * 0.19
    iva = round(iva)
    k4.metric("IVA 19%", "$" + _num(iva), help="Sólo se suma IVA a los proveedores que facturan.")
    k5.metric("Total", "$" + _num(neto + iva))
    _bajo = _prov_bajo_minimo(inc)
    if _bajo:
        st.warning("⚠️ Proveedores bajo su mínimo de compra: " + " · ".join(f"**{p_}** {_num(k_, 0)} de {_num(m_, 0)} kg (faltan {_num(m_ - k_, 0)})"
                                                                         for p_, k_, m_ in _bajo))


if _nav == _SECC[2]:
    st.markdown("##### " + "Plan de compra: revisa y elige proveedor producto por producto")
    if "cmp_nec" not in ss:
        st.info("Define la cobertura en 📊 Análisis → Cobertura.")
    elif ss.cmp_precios.empty:
        st.warning("Aún no hay proveedores/precios. Impórtalos en ⚙️ Configuración → Proveedores y precios.")
    else:
        _plan_panel()

# =====================================================================================
# TAB 4 - PROVEEDORES Y PRECIOS
# =====================================================================================
if _nav == _SECC[3]:
    st.markdown("##### " + "Base de proveedores y precios (neto por kg / unidad)")
    sub1, sub_cond, sub2, sub3 = st.tabs(["Precios por proveedor", "Condiciones de Proveedor", "Formatos (kg por bulto)", "Importar desde tu Excel"])

    with sub1:
        st.caption("Una fila por producto + proveedor + formato. *Disponible = No* marca quiebre. "
                   "Producto debe llamarse igual que el nombre base (sin 'Env.', 'Saco', etc.).")
        txt = st.text_input("🔎 Buscar", value=_keep("bprecios", ""))
        ss._k_bprecios = txt
        d = ss.cmp_precios.copy()
        for c in cc.COLS_PRECIOS:
            if c not in d.columns:
                d[c] = ""
        d["Precio Neto Kg"] = d["Precio Neto Kg"].map(lambda x: cc.to_float(x, float("nan")))      # acepta "3656,3", "$3.656,30", etc.
        d["Disponible"] = d["Disponible"].replace("", "Sí")
        # orden: producto A-Z, con stock primero, del más barato al más caro
        d["_k"] = d["Producto"].map(cc.norm)
        d["_q"] = (d["Disponible"] == "No").astype(int)
        d = d.sort_values(["_k", "_q", "Precio Neto Kg"], na_position="last", kind="stable").reset_index(drop=True)
        mask = pd.Series(True, index=d.index)
        if txt.strip():
            mask = d["Producto"].astype(str).str.contains(txt.strip(), case=False, na=False) | \
                   d["Proveedor"].astype(str).str.contains(txt.strip(), case=False, na=False)

        # ---------- Vista con colores (sólo lectura) ----------
        def _clp(x):
            if pd.isna(x) or x == "":
                return ""
            x = float(x)
            t = f"{x:,.0f}" if abs(x - round(x)) < 0.005 else f"{x:,.2f}"
            return "$" + t.replace(",", "X").replace(".", ",").replace("X", ".")

        vista = d[mask].reset_index(drop=True)
        k_ser, q_ser = vista["_k"], vista["_q"]
        vista = vista[cc.COLS_PRECIOS]          # sólo columnas visibles; k/q se usan para los estilos

        def _estilos(df):
            st_df = pd.DataFrame("", index=df.index, columns=df.columns)
            grupo = (k_ser != k_ser.shift()).cumsum()
            for gi, idx in k_ser.groupby(grupo).groups.items():
                fondo = "background-color: rgba(80,130,200,0.13)" if gi % 2 else ""
                st_df.loc[idx, :] = fondo
                pr = df.loc[idx]
                ok = pr[(q_ser.loc[idx] == 0) & pr["Precio Neto Kg"].notna() & (pr["Precio Neto Kg"] > 0)]
                lo, hi = (ok["Precio Neto Kg"].min(), ok["Precio Neto Kg"].max()) if len(ok) else (0, 0)
                for i in idx:
                    if q_ser.loc[i] == 1:
                        st_df.at[i, "Disponible"] = fondo + "; color: #cc0000; font-weight: 600"
                        st_df.at[i, "Precio Neto Kg"] = fondo + "; color: #999999; text-decoration: line-through"
                    elif i in ok.index:
                        t = 0.0 if hi == lo else (df.at[i, "Precio Neto Kg"] - lo) / (hi - lo)
                        # texto verde (barato) -> rojo (caro); el fondo azul del grupo se mantiene
                        if t >= 0.999:
                            col = "#fc0d1b"      # rojo del precio más caro
                        else:
                            col = f"hsl({120 * (1 - t):.0f},{85 + 13 * t:.0f}%,{34 + 18 * t:.0f}%)"
                        st_df.at[i, "Precio Neto Kg"] = fondo + f"; color: {col}"
            return st_df

        st.caption(f"{k_ser.nunique()} productos · {len(vista)} ofertas · ordenado de más barato a más caro "
                   "(texto verde = más barato, rojo = más caro dentro de cada producto; tachado = quiebre).")
        sty = vista.style.apply(_estilos, axis=None).format({"Precio Neto Kg": _clp}, na_rep="")
        st.dataframe(sty, hide_index=True, use_container_width=True, height=420)

        # ---------- Alta de oferta ----------
        fm_opts = sorted(set(df_formatos_db["Formato"].astype(str)) - {""})
        with st.expander("➕ Agregar oferta de proveedor"):
            nombres = set(d["Producto"].astype(str))
            if "cmp_base" in ss:
                nombres |= set(ss.cmp_base["Base"].astype(str))
            provs_ex = sorted(d["Proveedor"].astype(str).str.strip().replace("", pd.NA).dropna().unique())
            with st.form("form_add_oferta", clear_on_submit=True):
                c1, c2 = st.columns(2)
                a_prod = c1.selectbox("Producto (nombre base)", sorted(nombres), index=None, placeholder="Escribe para buscar…")
                a_prov = c2.selectbox("Proveedor existente", provs_ex, index=None, placeholder="Elige…")
                c3, c4, c5, c6 = st.columns([2, 2, 2, 1])
                a_prov_new = c3.text_input("…o proveedor nuevo")
                a_fmt = c4.selectbox("Formato", [""] + fm_opts)
                a_precio = c5.number_input("Precio neto por kg/uni.", min_value=0.0, step=10.0, format="%.2f")
                a_disp = c6.selectbox("Disponible", ["Sí", "No"])
                a_ok = st.form_submit_button("Agregar", type="primary")
            if a_ok:
                ok, msg = _agregar_oferta(a_prod, (a_prov_new or "").strip() or a_prov, a_fmt, a_precio, a_disp)
                if ok:
                    _cmp_save(cc.SH_PRECIOS, ss.cmp_precios[cc.COLS_PRECIOS])
                    ss.cmp_dirty = False
                    st.toast("✅ " + msg, icon="✅")
                    st.rerun()
                else:
                    st.error(msg)

        # ---------- Eliminar proveedor ----------
        with st.expander("🗑️ Eliminar un proveedor completo"):
            provs_del = sorted(d["Proveedor"].astype(str).str.strip().replace("", pd.NA).dropna().unique())
            with st.form("form_del_prov"):
                p_del = st.selectbox("Proveedor", provs_del, index=None, placeholder="Elige el proveedor a eliminar…")
                n_of = int((d["Proveedor"] == p_del).sum()) if p_del else 0
                st.caption(f"Se eliminarán **{n_of}** ofertas de este proveedor de la base." if p_del else
                           "Borra todas las ofertas del proveedor elegido. Para quitar una sola oferta usa 🗑️ en el Plan de compra.")
                conf = st.checkbox("Confirmo que quiero eliminarlo")
                go = st.form_submit_button("Eliminar proveedor", type="primary")
            if go:
                if not p_del or not conf:
                    st.error("Elige un proveedor y marca la confirmación.")
                else:
                    nuevo = ss.cmp_precios[ss.cmp_precios["Proveedor"] != p_del].reset_index(drop=True)
                    _cmp_save(cc.SH_PRECIOS, nuevo[cc.COLS_PRECIOS], force=True)
                    ss.cmp_precios = nuevo
                    ss.cmp_dirty = False
                    st.toast(f"🗑️ {p_del} eliminado ({n_of} ofertas)", icon="✅")
                    st.rerun()

        # ---------- Edición ----------
        with st.expander("✏️ Editar precios (cambiar o borrar filas)"):
            with st.form("form_ed_precios"):
                edp = st.data_editor(
                    d[mask][cc.COLS_PRECIOS], num_rows="dynamic", use_container_width=True, hide_index=True, key="ed_precios",
                    column_config={
                        "Formato": st.column_config.SelectboxColumn("Formato", options=fm_opts),
                        "Disponible": st.column_config.SelectboxColumn("Disponible", options=["Sí", "No"]),
                        "Precio Neto Kg": st.column_config.NumberColumn("Precio Neto Kg/Un", format="%.2f")})
                guardar = st.form_submit_button("💾 Guardar precios", type="primary")
            if guardar:
                out = pd.concat([d[~mask][cc.COLS_PRECIOS], edp], ignore_index=True)
                out = out[(out["Producto"].astype(str).str.strip() != "") & (out["Proveedor"].astype(str).str.strip() != "")]
                _con0 = int((d["Precio Neto Kg"].fillna(0) > 0).sum())
                _con1 = int((pd.to_numeric(out["Precio Neto Kg"], errors="coerce").fillna(0) > 0).sum())
                if _con0 - _con1 > max(5, 0.05 * _con0):
                    st.error(f"⛔ No guardé: este cambio dejaría {_con0 - _con1} precios en blanco (de {_con0}). Si es lo que quieres, "
                             "borra esas filas a propósito; si no, recarga con 🔄 y revisa.")
                    st.stop()
                _cmp_save(cc.SH_PRECIOS, out[cc.COLS_PRECIOS])
                ss.cmp_precios = out[cc.COLS_PRECIOS].reset_index(drop=True)
                ss.cmp_dirty = False
                st.toast("✅ Precios guardados", icon="✅")
                st.rerun()

    with sub2:
        st.caption("Cuántos kg (o unidades) trae cada formato. Se usa para pasar ventas y compras a kg.")
        fd = df_formatos_db.copy()
        fd["Kg por bulto"] = fd["Kg por bulto"].map(lambda x: cc.to_float(x, float("nan")))
        edf = st.data_editor(fd, num_rows="dynamic", use_container_width=True, hide_index=True, key="ed_formatos")
        if st.button("💾 Guardar formatos", type="primary", key="save_formatos"):
            edf = edf.dropna(subset=["Formato"])
            _cmp_save(cc.SH_FORMATOS, edf[cc.COLS_FORMATOS])
            ss.cmp_formatos = edf[cc.COLS_FORMATOS].reset_index(drop=True)
            st.toast("✅ Formatos guardados", icon="✅")
            st.rerun()

    with sub3:
        st.caption("Sube tu *COMPRAS Análisis Reposición de Stock* (.xlsm). Se convierte la matriz Producto × Proveedor a la tabla "
                   "de abajo (1 = quiebre). El formato de cada producto se toma del más usado en tu hoja Cotizaciones.")
        up = st.file_uploader("Excel de compras", type=["xlsm", "xlsx"], key="up_migr")
        if up is not None:
            try:
                p_new, f_new, stt = cc.migrar_excel_compras(up)
                st.write(f"**{stt['productos']}** productos · **{stt['ofertas']}** ofertas · **{stt['quiebres']}** quiebres · "
                         f"**{stt['sin_formato']}** ofertas sin formato · **{len(f_new)}** formatos")
                modo = st.radio("Al guardar", ["Reemplazar lo existente", "Agregar sólo lo que no existe"], horizontal=True)
                if st.button("⬆️ Guardar en Google Sheets", type="primary", key="do_migr"):
                    if modo.startswith("Agregar") and not df_precios_db.empty:
                        key = lambda df: df["Producto"].map(cc.norm) + "|" + df["Proveedor"].map(cc.norm)
                        p_new = p_new[~key(p_new).isin(set(key(df_precios_db)))]
                        p_new = pd.concat([df_precios_db, p_new], ignore_index=True)
                        f_new = pd.concat([df_formatos_db, f_new], ignore_index=True).drop_duplicates("Formato")
                    _cmp_save(cc.SH_PRECIOS, p_new[cc.COLS_PRECIOS], force=True)
                    _cmp_save(cc.SH_FORMATOS, f_new[cc.COLS_FORMATOS])
                    ss.cmp_precios, ss.cmp_formatos = p_new[cc.COLS_PRECIOS].reset_index(drop=True), f_new[cc.COLS_FORMATOS].reset_index(drop=True)
                    st.success("✅ Importado")
                    st.rerun()
            except Exception as e:
                st.error(f"No pude importar: {e}")

    with sub_cond:
        st.markdown("#### Condiciones de Proveedor")
        st.caption("Indica si el proveedor **factura** (si no factura, en la orden de compra no se suma IVA), su **condición de pago** y cómo "
                   "**entrega**. Los productos de proveedores con *Retiro en bodega proveedor* se reparten juntos en los viajes del camión "
                   "(la capacidad se define en 📄 Pedidos → Logística), priorizando los que se agotan antes. **Máx. viajes** limita en cuántos viajes se puede "
                   "repartir la compra de cada proveedor (1 = todo en un solo viaje; vacío = sin límite).")
        base = ss.cmp_provs.copy()
        for c_ in cc.COLS_PROVEEDORES:
            if c_ not in base.columns:
                base[c_] = ""
        ya = {cc.norm(p) for p in base["Proveedor"].astype(str)}
        nuevos = sorted({str(p).strip() for p in ss.cmp_precios["Proveedor"].astype(str)} - {""})
        nuevos = [p for p in nuevos if cc.norm(p) not in ya]
        if nuevos:    # proveedores de la base de precios que aún no tienen condiciones
            base = pd.concat([base, pd.DataFrame({"Proveedor": nuevos, "Factura": "Sí", "Condición de pago": "",
                                                  "Días cheque": "", "Entrega": "", "Máx. viajes": "", "Mín. kg compra": "", "Frecuencia": "Semanal", "Días despacho": "", "Pedir hasta": "", "Nota": ""})], ignore_index=True)
        base["_k"] = base["Proveedor"].map(cc.norm)
        base = base.sort_values("_k").drop(columns="_k").reset_index(drop=True)
        base["Factura"] = base["Factura"].replace("", "Sí").fillna("Sí")
        base["Frecuencia"] = base["Frecuencia"].replace("", "Semanal").fillna("Semanal")
        base["Días cheque"] = pd.to_numeric(base["Días cheque"].replace("", None), errors="coerce")
        base["Entrega"] = base["Entrega"].fillna("").astype(str)
        base["Máx. viajes"] = pd.to_numeric(base["Máx. viajes"].replace("", np.nan), errors="coerce")
        base["Mín. kg compra"] = pd.to_numeric(base["Mín. kg compra"].replace("", np.nan), errors="coerce")
        sin_cond = int((base["Condición de pago"].astype(str).str.strip() == "").sum())
        k1, k2, k3, k4, k5 = st.columns(5)
        k1.metric("Proveedores", len(base))
        k2.metric("Sin factura", int((base["Factura"] == "No").sum()))
        k3.metric("Sin condición de pago", sin_cond)
        k4.metric("Retiro en bodega", sum(1 for e_ in base["Entrega"] if cc.norm(e_).startswith("retiro")))
        k5.metric("Sin entrega definida", int((base["Entrega"].str.strip() == "").sum()))

        vista = base.copy()
        vista["Condición de pago"] = [_cond_label(c_, d_) or "—" for c_, d_ in zip(base["Condición de pago"], base["Días cheque"])]
        vista["Entrega"] = [e_ or "—" for e_ in base["Entrega"]]
        vista["Máx. viajes"] = [f"{int(m_)}" if pd.notna(m_) else "sin límite" for m_ in base["Máx. viajes"]]
        vista["Mín. kg compra"] = [f"{_num(m_, 0)} kg" if pd.notna(m_) and m_ > 0 else "—" for m_ in base["Mín. kg compra"]]
        vista = vista[["Proveedor", "Factura", "Condición de pago", "Entrega", "Máx. viajes", "Mín. kg compra", "Frecuencia", "Días despacho", "Pedir hasta", "Nota"]]

        def _est_cond(df):
            e = pd.DataFrame("", index=df.index, columns=df.columns)
            for i, v in df["Factura"].items():
                if v == "No":
                    e.at[i, "Factura"] = "color: #fc0d1b"
            return e
        st.dataframe(vista.style.apply(_est_cond, axis=None), hide_index=True, use_container_width=True, height=420)

        with st.expander("✏️ Editar condiciones", expanded=ss.cmp_provs.empty):
            with st.form("form_provs"):
                edp_ = st.data_editor(
                    base[cc.COLS_PROVEEDORES], num_rows="dynamic", hide_index=True, use_container_width=True, key="ed_provs",
                    column_config={
                        "Factura": st.column_config.SelectboxColumn("Factura", options=["Sí", "No"], width="small"),
                        "Condición de pago": st.column_config.SelectboxColumn("Condición de pago", options=COND_OPTS),
                        "Días cheque": st.column_config.NumberColumn("Días cheque (si es N#)", min_value=1, step=1, format="%d",
                                                                     help="Sólo si la condición es 'Cheque N# días'"),
                        "Entrega": st.column_config.SelectboxColumn("Entrega", options=cc.ENTREGAS),
                        "Máx. viajes": st.column_config.NumberColumn("Máx. viajes", min_value=1, max_value=8, step=1, format="%d",
                                                                     help="En cuántos viajes del camión como máximo se puede repartir la compra de este proveedor "
                                                                          "(sólo si retiras en su bodega). 1 = todo junto en un solo viaje. Vacío = sin límite."),
                        "Mín. kg compra": st.column_config.NumberColumn("Mín. kg por compra", min_value=0, step=10, format="%d",
                                                                        help="Kilos mínimos que exige el proveedor en cada compra (ej. Alitrade 500). Vacío = sin mínimo. "
                                                                             "Si el plan lo deja bajo ese mínimo, la app avisa cuánto falta."),
                        "Frecuencia": st.column_config.SelectboxColumn("Frecuencia", options=cc.FRECUENCIAS, width="small",
                                                                       help="Semanal: se repite cada semana. Fechas específicas: tú anotas las fechas (ej. proveedor que despacha una vez al mes)."),
                        "Días despacho": st.column_config.TextColumn("Días / fechas de despacho", help="Semanal: Vie  ·  o varios: Mar, Vie.  Fechas específicas: 27/11, 18/12"),
                        "Pedir hasta": st.column_config.TextColumn("Pedir hasta", help="Semanal: día (Mar). Fechas específicas: fechas límite en el mismo orden (20/11, 11/12), "
                                                                   "o '5 días antes'. Vacío = el día anterior al despacho."),
                        "Nota": st.column_config.TextColumn("Nota", width="large")})
                guardar_p = st.form_submit_button("💾 Guardar condiciones", type="primary")
            if guardar_p:
                out = edp_.copy()
                out = out[out["Proveedor"].astype(str).str.strip() != ""]
                out["Factura"] = out["Factura"].replace("", "Sí").fillna("Sí")
                out = out[cc.COLS_PROVEEDORES].reset_index(drop=True)
                _cmp_save(cc.SH_PROVEEDORES, out)
                ss.cmp_provs = out
                st.toast("✅ Condiciones guardadas", icon="✅")
                st.rerun()

# =====================================================================================
# TAB 5 - ORDEN DE COMPRA  (viajes globales: un mismo camión retira de varios proveedores)
# =====================================================================================
def _fecha_es(d):
    return f"{cc.DIAS_ES[d.weekday()]} {d:%d/%m}"


def _peso_fn():
    """Peso de cada día futuro: el calendario de 📊 Análisis → Cobertura si existe; si no, la regla base."""
    cal = ss.get("cmp_cal_fut")
    mp = {}
    if cal is not None and len(cal):
        mp = {pd.Timestamp(r["Fecha"]).date(): (float(r["Peso"]) if r["Incluir"] else 0.0) for r in cal.to_dict("records")}
    fer = ss.cmp_feriados
    return lambda d: mp[d] if d in mp else cc.peso_base(d, fer)


def _lineas_viaje(grp):
    """Líneas de la orden con su peso real de camión (PESO_LOGISTICO)."""
    out = []
    for r in grp.to_dict("records"):
        f_, origen = cc.factor_peso(r["Clave"], r["Producto"], bool(r.get("Es unidad", False)), ss.cmp_pesolog)
        kgb = float(r["Kg bulto"]) if float(r["Kg bulto"]) > 0 else 1.0
        out.append({"Clave": r["Clave"], "Proveedor": r["Proveedor"], "Producto": r["Producto"], "Formato": r["Formato"],
                    "Bultos": int(r["Bultos"]), "Kg bulto": kgb, "Peso bulto": kgb * f_, "Origen": origen,
                    "Stock": float(r["Stock kg"]), "Diaria": float(r["Venta diaria kg"]), "Precio": float(r["Precio Neto Kg"])})
    return out


def _max_viajes_map():
    """{proveedor: máx. de viajes} según ⚙️ Configuración → Condiciones de Proveedor (vacío = sin límite)."""
    d = ss.get("cmp_provs")
    if d is None or d.empty or "Máx. viajes" not in d.columns:
        return {}
    v = pd.to_numeric(d["Máx. viajes"].replace("", np.nan), errors="coerce")
    return {str(p).strip(): int(m) for p, m in zip(d["Proveedor"], v) if pd.notna(m) and m >= 1}


ss.setdefault("cmp_vjg_cfg", None)     # {n, f1, paso}
ss.setdefault("cmp_vjg_asig", None)    # {sig, df}


def _params_viaje():
    cap = float(_keep("vj_cap", cc.CAP_VIAJE_DEF))
    return {"cap": cap, "max": max(float(_keep("vj_max", cc.MAX_VIAJE_DEF)), cap), "buf": int(_keep("vj_buf", 2)),
            "paso": int(_keep("vj_paso", 7)), "llenar": bool(_keep("vj_llenar", True))}


def _viajes_global(lineas):
    """Configuración (n° de viajes y fecha de cada uno) + distribución vigente de TODOS los productos de retiro en bodega.
    Lee los widgets de la 📄 Pedidos → Logística si existen; si no, la última configuración guardada."""
    P = _params_viaje()
    ini_ = ss.get("_k_ini", date.today() + timedelta(days=1))
    col_ = float(ss.get("_k_col", 10.0))
    pf = _peso_fn()
    f1_def = ini_
    while pf(f1_def) <= 0:
        f1_def += timedelta(days=1)
    cfg = ss.cmp_vjg_cfg
    if cfg is None:
        sg = cc.sugerir_viajes(lineas, f1_def, P["paso"], ini_, pf, P["buf"], col_, P["cap"], P["llenar"], max_viajes=_max_viajes_map())
        n0, p0 = (sg if sg else (1, P["paso"]))
        cfg = {"n": n0, "fechas": [f1_def + timedelta(days=p0 * i) for i in range(n0)]}
        ss.cmp_vjg_cfg = cfg
    n = int(ss.get("vjn", cfg["n"]))
    raw = []
    for i in range(n):
        if f"vjd_{i}" in ss:
            raw.append(ss[f"vjd_{i}"])
        elif i < len(cfg["fechas"]):
            raw.append(cfg["fechas"][i])
        else:
            raw.append((raw[-1] if raw else f1_def) + timedelta(days=P["paso"]))
    fechas = sorted(raw)
    sig = (tuple((l["Clave"], l["Proveedor"], l["Bultos"], l["Formato"], round(l["Peso bulto"], 4)) for l in lineas),
           tuple(fechas), P["buf"], P["llenar"], ini_, col_, P["cap"], tuple(sorted(_max_viajes_map().items())))
    st_ = ss.cmp_vjg_asig
    if st_ is None or st_["sig"] != sig:
        st_ = {"sig": sig, "df": cc.distribuir_viajes(lineas, fechas, ini_, pf, P["buf"], col_, P["cap"], P["llenar"], _max_viajes_map())}
        ss.cmp_vjg_asig = st_
    pb = {l["Clave"]: l["Peso bulto"] for l in lineas}
    return {"n": n, "fechas": fechas, "fechas_raw": raw, "asig": st_["df"], "sig": sig, "pb": pb,
            "pesos": cc.pesos_por_viaje(st_["df"], pb, n), "ini": ini_, "col": col_, "pf": pf, "P": P, "f1_def": f1_def}


def _estado_carga(w, cap, mx):
    return "🟢" if w <= cap + 1e-6 else ("🟡" if w <= mx + 1e-6 else "🔴")


def _cb_redistribuir():
    ss.cmp_vjg_asig = None


def _cb_sugerir(lineas):
    """Pone el n° de viajes y las fechas que sugiere el algoritmo (descarta lo escrito a mano)."""
    P = _params_viaje()
    pf = _peso_fn()
    ini_ = ss.get("_k_ini", date.today() + timedelta(days=1))
    f1 = ini_
    while pf(f1) <= 0:
        f1 += timedelta(days=1)
    sg = cc.sugerir_viajes(lineas, f1, P["paso"], ini_, pf, P["buf"], float(ss.get("_k_col", 10.0)), P["cap"], P["llenar"], max_viajes=_max_viajes_map())
    n0, p0 = (sg if sg else (8, 1))
    fechas = [f1 + timedelta(days=p0 * i) for i in range(n0)]
    ss.cmp_vjg_cfg = {"n": n0, "fechas": fechas}
    ss["vjn"] = n0
    for i in range(8):
        ss.pop(f"vjd_{i}", None)
    for i, f_ in enumerate(fechas):
        ss[f"vjd_{i}"] = f_
    ss.cmp_vjg_asig = None


def _cb_n_viajes():
    """Al cambiar el n° de viajes, los nuevos viajes toman una fecha por defecto (último + días sugeridos)."""
    P = _params_viaje()
    n = int(ss["vjn"])
    prev = None
    for i in range(n):
        if f"vjd_{i}" not in ss:
            base_ = prev or ss.get("cmp_vjg_cfg", {}).get("fechas", [date.today()])[-1]
            ss[f"vjd_{i}"] = base_ + timedelta(days=P["paso"])
        prev = ss[f"vjd_{i}"]


def _pedidos_despacho(grp, inf):
    """Para un proveedor que despacha en tienda: reparte la compra en sus despachos (sólo lo necesario hasta el siguiente).
    Devuelve {'cal': [(límite, despacho)], 'lins': líneas, 'asig': DataFrame V1..Vn} o None."""
    if inf.get("retiro") or not (inf.get("desp") or inf.get("fechas")):
        return None
    ini_ = ss.get("_k_ini", date.today() + timedelta(days=1))
    fin_ = ss.get("_k_fin", ini_ + timedelta(days=14))
    if inf.get("fechas"):
        cal = cc.calendario_fechas(date.today(), inf.get("desp_txt", ""), inf.get("lim_txt", ""), fin_)
    else:
        cal = cc.calendario_despacho(date.today(), inf["desp"], inf.get("lim", []), fin_)
    if not cal:
        return None
    lins = _lineas_viaje(grp)
    P = _params_viaje()
    asig = cc.distribuir_viajes(lins, [d for _, d in cal], ini_, _peso_fn(), P["buf"], float(ss.get("_k_col", 10.0)), 1e12, False, None)
    mk = float(inf.get("min_kg", 0.0) or 0.0)
    if mk > 0 and len(cal) > 1 and len(asig):
        asig = _aplicar_minimo(asig, lins, len(cal), mk)
    return {"cal": cal, "lins": lins, "asig": asig}


def _aplicar_minimo(asig, lins, n, mk):
    """Cada pedido debe llegar al mínimo de kg del proveedor: primero se adelantan bultos de pedidos siguientes (los que se agotan antes);
    si ya no queda de dónde, el pedido se junta con el anterior."""
    kgb = {l["Clave"]: l["Kg bulto"] for l in lins}
    rows = [dict(r) for r in asig.to_dict("records")]
    kg = lambda j: sum(int(r[f"V{j + 1}"]) * kgb[r["Clave"]] for r in rows)
    for j in range(n):
        if kg(j) <= 0 or kg(j) >= mk:
            continue
        for jj in range(j + 1, n):
            for r in rows:
                while int(r[f"V{jj + 1}"]) > 0 and kg(j) < mk:
                    r[f"V{jj + 1}"] = int(r[f"V{jj + 1}"]) - 1
                    r[f"V{j + 1}"] = int(r[f"V{j + 1}"]) + 1
                if kg(j) >= mk:
                    break
            if kg(j) >= mk:
                break
        if 0 < kg(j) < mk:
            prev = next((i for i in range(j - 1, -1, -1) if kg(i) > 0), None)
            if prev is not None:
                for r in rows:
                    r[f"V{prev + 1}"] = int(r[f"V{prev + 1}"]) + int(r[f"V{j + 1}"])
                    r[f"V{j + 1}"] = 0
    return pd.DataFrame(rows)


def _info_stock(lineas, asig, j, fecha):
    """{Clave: (stock hoy, stock al llegar, dura hasta)} para las líneas del pedido/viaje j (1..n), incluyendo lo pedido hasta j."""
    ini_ = ss.get("_k_ini", date.today() + timedelta(days=1))
    pre = cc._prefijo_pesos(ini_, _peso_fn())
    ln = {l["Clave"]: l for l in lineas}
    out = {}
    for r in asig.to_dict("records"):
        l = ln[r["Clave"]]
        cum = sum(int(r[f"V{i}"]) for i in range(1, j + 1)) * l["Kg bulto"]
        prev = sum(int(r[f"V{i}"]) for i in range(1, j)) * l["Kg bulto"]     # lo ya pedido en los pedidos/viajes anteriores
        al_llegar = max(0.0, l["Stock"] + prev - cc._acum(pre, ini_, fecha) * l["Diaria"])
        q = cc.fecha_quiebre(l["Stock"] + cum, l["Diaria"], ini_, pre)
        dura = f"{cc.DIAS_ES[q.weekday()]} {q:%d/%m}" if q else "sin quiebre"
        out[r["Clave"]] = (l["Stock"], al_llegar, dura, prev)
    return out


def _lineas_de_viaje(lineas, asig, j, prov=None):
    """Líneas (dict) con bultos en el viaje j (1..n); opcionalmente sólo de un proveedor."""
    ln = {l["Clave"]: l for l in lineas}
    out = []
    for r in asig.to_dict("records"):
        b = int(r[f"V{j}"])
        l = ln[r["Clave"]]
        if b > 0 and (prov is None or l["Proveedor"] == prov):
            out.append({**l, "BultosV": b, "KgV": b * l["Kg bulto"], "PesoV": b * l["Peso bulto"],
                        "NetoV": b * l["Kg bulto"] * l["Precio"]})
    return out


def _linea_msg(l, b=None):
    fm = str(l["Formato"] or "").strip()
    b = int(l["Bultos"] if b is None else b)
    cant = f"{b} kg" if fm in ("", "(kg)") else f"{b} x {fm}"
    return f"• {cant} – {l['Producto']}"


def _render_viajes(lineas, vj):
    """Bloque de viajes: configuración, carga del camión, reparto editable y plan de ruta por viaje."""
    P = vj["P"]
    ln = {l["Clave"]: l for l in lineas}
    c1, c2, c3 = st.columns([1, 1.6, 1.6])
    c1.number_input("Cantidad de viajes", 1, 8, step=1, key="vjn", on_change=_cb_n_viajes,
                    **({} if "vjn" in ss else {"value": int(vj["n"])}))
    c2.button("✨ Sugerir viajes y fechas", on_click=_cb_sugerir, args=(lineas,), use_container_width=True,
              help="Calcula cuántos viajes y qué días hacen falta para que todo quepa sin dejar productos sin stock.")
    c3.button("🔄 Redistribuir productos", on_click=_cb_redistribuir, use_container_width=True,
              help="Vuelve a repartir los productos entre tus viajes según cuándo se agotan (descarta tus cambios manuales de bultos).")
    st.caption("Define el día de cada viaje:")
    cols_f = st.columns(min(4, vj["n"]))
    for i in range(vj["n"]):
        cols_f[i % len(cols_f)].date_input(f"Viaje {i + 1}", key=f"vjd_{i}", format="DD/MM/YYYY",
                                           **({} if f"vjd_{i}" in ss else {"value": vj["fechas_raw"][i]}))
    ss.cmp_vjg_cfg = {"n": int(ss.get("vjn", vj["n"])), "fechas": [ss.get(f"vjd_{i}", vj["fechas_raw"][i]) for i in range(vj["n"])]}
    if vj["fechas_raw"] != sorted(vj["fechas_raw"]):
        st.warning("Las fechas de los viajes no van en orden; se ordenaron de la más próxima a la más lejana para el cálculo.")

    n, fechas, pesos, asig = vj["n"], vj["fechas"], vj["pesos"], vj["asig"]
    cap, mx = P["cap"], P["max"]
    for ini_j in range(0, n, 4):
        cols = st.columns(min(4, n - ini_j))
        for off, c in enumerate(cols):
            j = ini_j + off
            w = pesos[j]
            c.markdown(f"**Viaje {j + 1}** · {_fecha_es(fechas[j])}")
            c.progress(min(w / mx, 1.0) if mx > 0 else 0.0)
            c.caption(f"{_estado_carga(w, cap, mx)} {_num(w)} kg de {_num(cap)} (máx {_num(mx)})")
    if max(pesos) > cap + 1e-6:
        st.warning(f"Con tus viajes, alguno pasa de {_num(cap)} kg. Prueba **✨ Sugerir viajes y fechas**, agrega un viaje o cambia sus fechas.")
    for j, w in enumerate(pesos):
        if w > mx + 1e-6:
            st.error(f"Viaje {j + 1}: {_num(w)} kg supera el máximo del camión ({_num(mx)} kg). "
                     "Agrega un viaje, adelanta la fecha o sube la capacidad.")
        elif w > cap + 1e-6:
            st.warning(f"Viaje {j + 1}: {_num(w)} kg, sobre la capacidad normal ({_num(cap)} kg) pero dentro del máximo ({_num(mx)} kg).")
    _mvm = _max_viajes_map()
    if _mvm:
        _usados = {}
        for l_ in lineas:
            if l_["Proveedor"] in _mvm:
                fila_ = asig[asig["Clave"] == l_["Clave"]]
                if len(fila_):
                    _usados.setdefault(l_["Proveedor"], set()).update(j + 1 for j in range(vj["n"]) if int(fila_.iloc[0][f"V{j + 1}"]) > 0)
        _lim = [f"{p_} (máx. {m_})" for p_, m_ in _mvm.items() if p_ in _usados]
        if _lim:
            st.caption("Proveedores con límite de viajes: " + " · ".join(_lim) + ". Se editan en ⚙️ Configuración → Condiciones de Proveedor.")
        for p_, m_ in _mvm.items():
            if p_ in _usados and len(_usados[p_]) > m_:
                st.warning(f"{p_}: se reparte en {len(_usados[p_])} viajes (viajes {', '.join(map(str, sorted(_usados[p_])))}) y su límite es {m_}. "
                           "Pulsa **🔄 Redistribuir productos** para volver a respetarlo.")
    riesgo = asig[asig["Prioridad"] == "🔴 Riesgo"]
    if len(riesgo):
        st.error(f"🔴 {len(riesgo)} producto(s) se agotarían antes del viaje 1 (con {P['buf']} días de margen): "
                 + ", ".join(ln[c]["Producto"] for c in riesgo["Clave"].head(6)) + ("…" if len(riesgo) > 6 else "")
                 + ". Adelanta el viaje 1 o consíguelos por otra vía.")
    dud = [l["Producto"] for l in lineas if l["Origen"] in ("Estimado", "Sin dato")]
    if dud:
        st.caption(f"⚠️ {len(dud)} producto(s) sin PESO_LOGISTICO en la hoja *productos* (peso estimado con el nombre o 1 kg): "
                   + ", ".join(dud[:5]) + ("…" if len(dud) > 5 else ""))

    # ---- reparto editable: bultos por viaje, todos los proveedores juntos, los más urgentes arriba ----
    view = pd.DataFrame({
        "Prioridad": asig["Prioridad"],
        "Producto": [ln[c]["Producto"] for c in asig["Clave"]],
        "Proveedor": [ln[c]["Proveedor"] for c in asig["Clave"]],
        "Formato": [ln[c]["Formato"] or "(kg)" for c in asig["Clave"]],
        "Peso bulto": [f"{_num(ln[c]['Peso bulto'], 2)} kg" + (" ⚠" if ln[c]["Origen"] in ("Estimado", "Sin dato") else "")
                       for c in asig["Clave"]],
        "Se agota": [_fecha_es(q) if q else "—" for q in asig["Quiebre"]],
        "Total": [ln[c]["Bultos"] for c in asig["Clave"]],
    })
    for j in range(n):
        view[f"V{j + 1}"] = asig[f"V{j + 1}"].astype(int).values
    cfgcols = {f"V{j + 1}": st.column_config.NumberColumn(f"Viaje {j + 1} · {fechas[j]:%d/%m}", min_value=0, step=1, format="%d")
               for j in range(n)}
    cfgcols["Total"] = st.column_config.NumberColumn("Total bultos", format="%d")
    with st.form("form_vj"):
        st.caption("Ordenado por urgencia (los que se agotan antes, arriba). Escribe cuántos bultos van en cada viaje; "
                   "la suma de los viajes debe ser igual al total del plan.")
        edv = st.data_editor(view, hide_index=True, use_container_width=True, key="edvj", height=420,
                             disabled=["Prioridad", "Producto", "Proveedor", "Formato", "Peso bulto", "Se agota", "Total"],
                             column_config=cfgcols)
        ok_btn = st.form_submit_button("💾 Aplicar distribución", type="primary")
    if ok_btn:
        vcols = [f"V{j + 1}" for j in range(n)]
        suma = edv[vcols].fillna(0).astype(int).sum(axis=1)
        mal = edv[suma != edv["Total"]]
        if len(mal):
            st.error("La suma de los viajes no coincide con el total en: "
                     + ", ".join(f"{r.Producto} ({int(suma[i])} de {int(r.Total)})" for i, r in mal.head(6).iterrows()))
        else:
            nuevo = asig.copy()
            for c_ in vcols:
                nuevo[c_] = edv[c_].fillna(0).astype(int).values
            ss.cmp_vjg_asig = {"sig": vj["sig"], "df": nuevo}
            st.rerun()


def _render_plan_ruta(lineas, vj):
    """Una pestaña por viaje: a qué proveedores ir y qué retirar en cada uno."""
    tabs = st.tabs([f"Viaje {j + 1} · {_fecha_es(vj['fechas'][j])} · {_num(vj['pesos'][j])} kg" for j in range(vj["n"])])
    for j, tb in enumerate(tabs, 1):
        with tb:
            fl = _lineas_de_viaje(lineas, vj["asig"], j)
            if not fl:
                st.info("Este viaje no tiene productos.")
                continue
            det = pd.DataFrame([{"Proveedor": l["Proveedor"], "Producto": l["Producto"], "Formato": l["Formato"] or "(kg)",
                                 "Bultos": l["BultosV"], "Peso camión kg": l["PesoV"], "Total Neto": l["NetoV"]} for l in fl])
            resp = det.groupby("Proveedor").agg(Líneas=("Producto", "count"), **{"Peso camión kg": ("Peso camión kg", "sum"),
                                                                                  "Total Neto": ("Total Neto", "sum")}).reset_index()
            st.markdown("**Proveedores a visitar**")
            st.dataframe(resp.style.format({"Peso camión kg": lambda v: _num(v, 1), "Total Neto": lambda v: "$" + _num(v)}),
                         hide_index=True, use_container_width=True)
            det = det.sort_values(["Proveedor", "Producto"])
            st.markdown("**Detalle**")
            st.dataframe(det.style.format({"Peso camión kg": lambda v: _num(v, 1), "Total Neto": lambda v: "$" + _num(v)}),
                         hide_index=True, use_container_width=True)


if _nav == _SECC[4]:
    st.markdown("##### " + "Órdenes por proveedor")
    sel = ss.get("cmp_plan_sel")
    if sel is None or sel.empty:
        st.info("Arma y marca las líneas en el plan de compra (🛒 Plan de compra).")
    else:
        ok = sel[(sel["Proveedor"] != "— Sin proveedor —") & (sel["Bultos"] > 0)].copy()
        sinp = sel[sel["Proveedor"] == "— Sin proveedor —"]
        if len(sinp):
            st.warning(f"{len(sinp)} productos sin proveedor: " + ", ".join(sinp["Producto"].head(8)) + ("…" if len(sinp) > 8 else ""))
        pm = _prov_map()
        _info = lambda p: pm.get(cc.norm(p), {"factura": True, "cond": "", "entrega": "", "retiro": False})

        # --- líneas de todos los proveedores y distribución global de los viajes (antes del resumen) ---
        lins_all = _lineas_viaje(ok) if len(ok) else []
        lins_ret = [l for l in lins_all if _info(l["Proveedor"])["retiro"]]
        pes_prov = {}
        for l in lins_all:
            pes_prov[l["Proveedor"]] = pes_prov.get(l["Proveedor"], 0.0) + l["Bultos"] * l["Peso bulto"]
        vj = _viajes_global(lins_ret) if lins_ret else None
        viajes_de = {}
        if vj:
            for j in range(1, vj["n"] + 1):
                for l in _lineas_de_viaje(lins_ret, vj["asig"], j):
                    viajes_de.setdefault(l["Proveedor"], set()).add(j)

        desp_de = {}
        for p_, g_ in ok.groupby("Proveedor"):
            r_ = _pedidos_despacho(g_, _info(p_))
            if r_:
                desp_de[p_] = r_

        for p_ in sorted(ok["Proveedor"].unique()):
            i_ = _info(p_)
            if i_.get("fechas") and p_ not in desp_de:
                st.error(f"⚠️ {p_}: no tiene una fecha de despacho vigente. Anota su próxima fecha en ⚙️ Configuración → Condiciones de Proveedor.")

        res = ok.groupby("Proveedor").agg(Neto=("Total Neto", "sum")).reset_index()
        res["Factura"] = res["Proveedor"].map(lambda p: "Sí" if _info(p)["factura"] else "No")
        res["Condición de pago"] = res["Proveedor"].map(lambda p: _info(p)["cond"] or "—")
        res["Entrega"] = res["Proveedor"].map(lambda p: _info(p)["entrega"] or "—")
        res["Peso kg"] = res["Proveedor"].map(pes_prov)
        n_vj = vj["n"] if vj else 0
        for j in range(1, n_vj + 1):       # total (con IVA si factura) de lo que se retira en cada viaje
            def _tot_viaje(p, j=j):
                if p not in viajes_de or j not in viajes_de[p]:
                    return float("nan")
                neto_j = sum(l["NetoV"] for l in _lineas_de_viaje(lins_ret, vj["asig"], j, p))
                return round(neto_j * (1.19 if _info(p)["factura"] else 1.0))
            res[f"Viaje {j} Total"] = res["Proveedor"].map(_tot_viaje)
        def _tot_ahora(p):
            d_ = desp_de.get(p)
            if not d_:
                return float("nan")
            pri = _lineas_de_viaje(d_["lins"], d_["asig"], 1, p)
            return round(sum(l["NetoV"] for l in pri) * (1.19 if _info(p)["factura"] else 1.0))
        res["Pedir ahora"] = res["Proveedor"].map(_tot_ahora)
        res["IVA"] = (res["Neto"].where(res["Factura"] == "Sí", 0.0) * 0.19).round(0)      # sin factura: no se suma IVA
        res["Total"] = res["Neto"] + res["IVA"]
        res = res.sort_values("Total", ascending=False).reset_index(drop=True)
        res = res[["Proveedor", "Factura", "Condición de pago", "Entrega", "Peso kg"] + [f"Viaje {j} Total" for j in range(1, n_vj + 1)]
                  + (["Pedir ahora"] if res["Pedir ahora"].notna().any() else []) + ["Neto", "IVA", "Total"]]
        _cols_vj = [f"Viaje {j} Total" for j in range(1, n_vj + 1)]

        def _nf(x, signo=""):
            """Formato chileno: miles con punto, decimales con coma (sólo si hacen falta)."""
            if pd.isna(x):
                return ""
            x = float(x)
            return signo + (_num(x, 0) if abs(x - round(x)) < 0.005 else _num(x, 2))
        def _est_res(df):
            e = pd.DataFrame("", index=df.index, columns=df.columns)
            for i, v in df["Factura"].items():
                if v == "No":
                    e.at[i, "Factura"] = "color: #fc0d1b"
            return e
        _vista_res = res.drop(columns=["Neto", "IVA"])
        st.dataframe(_vista_res.style.apply(_est_res, axis=None).format(
                        {**{c: (lambda v: _nf(v, "$")) for c in _cols_vj + ["Total", "Pedir ahora"] if c in _vista_res.columns}, "Peso kg": lambda v: _num(v, 0)}, na_rep="—"),
                     hide_index=True, use_container_width=True)
        t1, t2, t3 = st.columns(3)
        t1.metric("Neto total", "$" + _nf(res["Neto"].sum()))
        t2.metric("IVA total", "$" + _nf(res["IVA"].sum()), help="Sólo proveedores que facturan.")
        t3.metric("Total a pagar", "$" + _nf(res["Total"].sum()))
        if (res["Factura"] == "No").any():
            st.caption("🔴 Proveedores sin factura: se considera sólo el valor neto (no se suma IVA).")
        _bajo5 = [b_ for b_ in _prov_bajo_minimo(ok) if b_[0] not in viajes_de]      # los que van por viajes se revisan viaje por viaje
        for p_, k_, m_ in _bajo5:
            st.warning(f"⚠️ {p_}: la compra es de {_num(k_, 0)} kg y su mínimo es {_num(m_, 0)} kg (faltan {_num(m_ - k_, 0)} kg).")
        for p_ in sorted(viajes_de):
            m_ = _info(p_).get("min_kg", 0.0)
            for j_ in sorted(viajes_de[p_]):
                k_ = sum(l["KgV"] for l in _lineas_de_viaje(lins_ret, vj["asig"], j_, p_))
                if m_ > 0 and k_ + 1e-6 < m_:
                    st.warning(f"⚠️ {p_} · viaje {j_}: {_num(k_, 0)} kg y su mínimo por compra es {_num(m_, 0)} kg (faltan {_num(m_ - k_, 0)} kg). "
                               "Junta sus compras en menos viajes (Máx. viajes) o súmale productos.")
        if (res["Condición de pago"] == "—").any() or (res["Entrega"] == "—").any():
            st.caption("⚠️ Hay proveedores sin condición de pago o sin forma de entrega: defínelos en 4 → Condiciones de Proveedor.")

        if vj:
            st.info(f"🚚 {len(viajes_de)} proveedor(es) se retiran con el camión en {vj['n']} viaje(s). "
                    "Define los viajes y sus fechas en 📄 Pedidos → Logística y viajes.")

        _cal = []
        _hoy = date.today()
        for p_, d_ in sorted(desp_de.items()):
            fl_ = None
            for j_, (C_, D_) in enumerate(d_["cal"], start=1):
                if _lineas_de_viaje(d_["lins"], d_["asig"], j_, p_):
                    fl_ = (C_, D_)
                    break
            if fl_ is None:
                continue
            C_, D_ = fl_
            fal = (C_ - _hoy).days
            _cal.append({"Proveedor": p_, "Pedir antes de": f"{cc.DIAS_ES[C_.weekday()]} {C_:%d/%m}", "Llega / despacha": f"{cc.DIAS_ES[D_.weekday()]} {D_:%d/%m}",
                         "Días": fal, "Estado": "🔴 ¡Hoy es el último día!" if fal == 0 else ("🟡 Pedir pronto" if fal <= 2 else "🟢 A tiempo")})
        if _cal:
            st.markdown("#### 📅 Próximo pedido por proveedor (despachan en tienda)")
            _cal.sort(key=lambda x: x["Días"])
            st.dataframe(pd.DataFrame(_cal).drop(columns="Días"), hide_index=True, use_container_width=True)
            for c_ in _cal:
                if c_["Días"] == 0:
                    st.error(f"🔴 {c_['Proveedor']}: hoy es el último día para pedir (despacha {c_['Llega / despacha']}).")
                elif c_["Días"] <= 2:
                    st.warning(f"🟡 {c_['Proveedor']}: pedir antes de {c_['Pedir antes de']} para que llegue el {c_['Llega / despacha']}.")
            st.caption("Cada pedido incluye sólo lo necesario para llegar al despacho siguiente; el resto del plan se pide en los despachos posteriores.")

        st.markdown("### Órdenes por proveedor")
        prov_sel = st.selectbox("Ver / copiar orden de", ["(todos)"] + res["Proveedor"].tolist())
        for prov, grp in ok.groupby("Proveedor"):
            if prov_sel not in ("(todos)", prov):
                continue
            inf = _info(prov)
            _tag = ("Factura" if inf["factura"] else "Sin factura") + " · " + (inf["cond"] or "sin condición de pago") \
                   + " · " + (inf["entrega"] or "entrega sin definir")
            with st.expander(f"{prov} · {clp(grp['Total Neto'].sum())} neto · {_tag}", expanded=prov_sel == prov):
                if prov in viajes_de:
                    for j in sorted(viajes_de[prov]):
                        fl = _lineas_de_viaje(lins_ret, vj["asig"], j, prov)
                        st.markdown(f"**Viaje {j} · {_fecha_es(vj['fechas'][j - 1])}** · {_num(sum(l['PesoV'] for l in fl), 1)} kg · "
                                    + (f"{clp(round(sum(l['NetoV'] for l in fl) * 1.19))} bruto (IVA incl.)" if inf["factura"]
                                       else f"{clp(round(sum(l['NetoV'] for l in fl)))} (sin factura)"))
                        _is = _info_stock(lins_ret, vj["asig"], j, vj["fechas"][j - 1])
                        det = pd.DataFrame([{"Producto": l["Producto"], "Formato": l["Formato"], "Bultos": l["BultosV"], "Kg total": l["KgV"],
                                             "Stock hoy": _is[l["Clave"]][0], "Ya pedido antes": _is[l["Clave"]][3], "Stock al llegar": _is[l["Clave"]][1], "Durará hasta": _is[l["Clave"]][2],
                                             "Precio Neto Kg": l["Precio"], "Total Neto": l["NetoV"]} for l in fl])
                        st.dataframe(det.style.format({"Kg total": _nf, "Stock hoy": _nf, "Ya pedido antes": _nf, "Stock al llegar": _nf, "Precio Neto Kg": lambda v: _nf(v, "$"),
                                                       "Total Neto": lambda v: _nf(v, "$")}),
                                     hide_index=True, use_container_width=True)
                        f_ = vj["fechas"][j - 1]
                        msg = (f"Hola, quisiera pedir para retirar el {cc.DIAS_ES[f_.weekday()]} {f_:%d/%m}:\n"
                               + "\n".join(_linea_msg(l, l["BultosV"]) for l in sorted(fl, key=lambda x: cc.norm(x["Producto"]))) + "\n¿Tienen stock? Gracias.")
                        st.code(msg, language=None)
                elif prov in desp_de:
                    d_ = desp_de[prov]
                    for j, (C_, D_) in enumerate(d_["cal"], start=1):
                        fl = _lineas_de_viaje(d_["lins"], d_["asig"], j, prov)
                        if not fl:
                            continue
                        neto_ = sum(l["NetoV"] for l in fl)
                        st.markdown(f"**Pedido {j} · pedir antes de {cc.DIAS_ES[C_.weekday()]} {C_:%d/%m} · llega {cc.DIAS_ES[D_.weekday()]} {D_:%d/%m}** "
                                    f"· {clp(round(neto_ * (1.19 if inf['factura'] else 1.0)))} " + ("bruto (IVA incl.)" if inf["factura"] else "(sin factura)"))
                        _is = _info_stock(d_["lins"], d_["asig"], j, D_)
                        det = pd.DataFrame([{"Producto": l["Producto"], "Formato": l["Formato"], "Bultos": l["BultosV"], "Kg total": l["KgV"],
                                             "Stock hoy": _is[l["Clave"]][0], "Ya pedido antes": _is[l["Clave"]][3], "Stock al llegar": _is[l["Clave"]][1], "Durará hasta": _is[l["Clave"]][2],
                                             "Precio Neto Kg": l["Precio"], "Total Neto": l["NetoV"]} for l in fl])
                        st.dataframe(det.style.format({"Kg total": _nf, "Stock hoy": _nf, "Ya pedido antes": _nf, "Stock al llegar": _nf, "Precio Neto Kg": lambda v: _nf(v, "$"),
                                                       "Total Neto": lambda v: _nf(v, "$")}),
                                     hide_index=True, use_container_width=True)
                        msg = (f"Hola, quisiera pedir para el despacho del {cc.DIAS_ES[D_.weekday()]} {D_:%d/%m}:\n"
                               + "\n".join(_linea_msg(l, l["BultosV"]) for l in sorted(fl, key=lambda x: cc.norm(x["Producto"]))) + "\n¿Tienen stock? Gracias.")
                        st.code(msg, language=None)
                else:
                    det = grp[["Producto", "Formato", "Bultos", "Kg total", "Precio Neto Kg", "Total Neto"]]
                    st.dataframe(det.style.format({"Kg total": _nf, "Precio Neto Kg": lambda v: _nf(v, "$"),
                                                   "Total Neto": lambda v: _nf(v, "$")}),
                                 hide_index=True, use_container_width=True)
                    lins = [l for l in lins_all if l["Proveedor"] == prov]
                    msg = "Hola, quisiera pedir / cotizar:\n" + "\n".join(_linea_msg(l) for l in sorted(lins, key=lambda x: cc.norm(x["Producto"]))) + "\n¿Tienen stock? Gracias."
                    st.code(msg, language=None)

        # Excel
        buf = BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as xw:
            res.to_excel(xw, sheet_name="Resumen", index=False)
            if vj:
                filas_v = []
                for j in range(1, vj["n"] + 1):
                    for l in _lineas_de_viaje(lins_ret, vj["asig"], j):
                        filas_v.append({"Viaje": j, "Fecha retiro": vj["fechas"][j - 1].strftime("%d/%m/%Y"), "Proveedor": l["Proveedor"],
                                        "Producto": l["Producto"], "Formato": l["Formato"], "Bultos": l["BultosV"],
                                        "Peso camión kg": l["PesoV"], "Total Neto": l["NetoV"]})
                pd.DataFrame(filas_v).sort_values(["Viaje", "Proveedor", "Producto"]).to_excel(xw, sheet_name="Viajes", index=False)
            for prov, grp in ok.groupby("Proveedor"):
                nm = re.sub(r"[\\/*?:\[\]]", "", str(prov))[:24] or "Prov"
                iva_r = 0.19 if _info(prov)["factura"] else 0.0
                if prov in viajes_de:
                    for j in sorted(viajes_de[prov]):
                        fl = _lineas_de_viaje(lins_ret, vj["asig"], j, prov)
                        g2 = pd.DataFrame([{"Producto": l["Producto"], "Formato": l["Formato"], "Bultos": l["BultosV"], "Kg total": l["KgV"],
                                            "Peso camión kg": l["PesoV"], "Precio Neto Kg": l["Precio"], "Total Neto": l["NetoV"],
                                            "Fecha retiro": vj["fechas"][j - 1].strftime("%d/%m/%Y")} for l in fl])
                        g2["IVA"] = g2["Total Neto"] * iva_r
                        g2["Total"] = g2["Total Neto"] + g2["IVA"]
                        g2.to_excel(xw, sheet_name=f"{nm} V{j}", index=False)
                else:
                    g2 = grp[["Producto", "Formato", "Bultos", "Kg total", "Precio Neto Kg", "Total Neto"]].copy()
                    g2["IVA"] = g2["Total Neto"] * iva_r
                    g2["Total"] = g2["Total Neto"] + g2["IVA"]
                    g2.to_excel(xw, sheet_name=nm, index=False)
            if len(sinp):
                sinp[["Producto", "Comprar kg"]].to_excel(xw, sheet_name="Sin proveedor", index=False)
        st.download_button("📥 Descargar Excel de órdenes", buf.getvalue(), file_name=f"Ordenes_Compra_{date.today():%Y%m%d}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        estado = st.selectbox("Estado al guardar", ["Cotizada", "Pedida"])
        if st.button("💾 Guardar en historial de órdenes (Google Sheets)", type="primary"):
            stamp = datetime.now().strftime("%Y%m%d-%H%M")
            provs = {p: i for i, p in enumerate(ok["Proveedor"].unique(), 1)}
            filas = []
            for prov, grp in ok.groupby("Proveedor", sort=False):
                base_no = f"OC-{stamp}-{provs[prov]:02d}"
                if prov in viajes_de:
                    for j in sorted(viajes_de[prov]):
                        for l in _lineas_de_viaje(lins_ret, vj["asig"], j, prov):
                            filas.append({"Fecha": date.today().isoformat(), "N° Orden": f"{base_no}-V{j}", "Proveedor": prov,
                                          "Producto": l["Producto"], "Formato": l["Formato"], "Bultos": l["BultosV"], "Kg": l["KgV"],
                                          "Precio Neto Kg": l["Precio"], "Total Neto": round(l["NetoV"], 0), "Estado": estado,
                                          "Viaje": j, "Fecha Viaje": vj["fechas"][j - 1].isoformat()})
                else:
                    for r in grp.to_dict("records"):
                        filas.append({"Fecha": date.today().isoformat(), "N° Orden": base_no, "Proveedor": prov,
                                      "Producto": r["Producto"], "Formato": r["Formato"], "Bultos": r["Bultos"], "Kg": r["Kg total"],
                                      "Precio Neto Kg": r["Precio Neto Kg"], "Total Neto": r["Total Neto"], "Estado": estado,
                                      "Viaje": "", "Fecha Viaje": ""})
            h = pd.DataFrame(filas, columns=cc.COLS_ORDENES)
            _cmp_append(cc.SH_ORDENES, h, cc.COLS_ORDENES)
            st.success(f"✅ {len(h)} líneas guardadas en '{cc.SH_ORDENES}'")


# =====================================================================================
# TAB 6 - LOGÍSTICA  (camión y viajes)
# =====================================================================================
if _nav == _SECC[5]:
    st.markdown("##### " + "Logística: camión y viajes")
    sel = ss.get("cmp_plan_sel")
    if sel is None or sel.empty:
        st.info("Arma y marca las líneas en el plan de compra (🛒 Plan de compra).")
    else:
        ok = sel[(sel["Proveedor"] != "— Sin proveedor —") & (sel["Bultos"] > 0)].copy()
        pm = _prov_map()
        _info = lambda p: pm.get(cc.norm(p), {"factura": True, "cond": "", "entrega": "", "retiro": False})
        lins_all = _lineas_viaje(ok) if len(ok) else []
        lins_ret = [l for l in lins_all if _info(l["Proveedor"])["retiro"]]

        with st.expander("⚙️ Parámetros del camión", expanded=True):
            p1, p2, p3, p4, p5 = st.columns(5)
            ss._k_vj_cap = p1.number_input("Capacidad del camión (kg)", 100, 10000, int(_keep("vj_cap", cc.CAP_VIAJE_DEF)), 50)
            ss._k_vj_max = p2.number_input("Máximo permitido (kg)", 100, 10000, int(_keep("vj_max", cc.MAX_VIAJE_DEF)), 50,
                                           help="Tope absoluto; entre la capacidad y este máximo el viaje queda en amarillo.")
            ss._k_vj_buf = p3.number_input("Margen de seguridad (días)", 0, 15, int(_keep("vj_buf", 2)), 1,
                                           help="Días de stock que deben quedar al llegar el siguiente viaje.")
            ss._k_vj_paso = p4.number_input("Días entre viajes (sugerencia)", 1, 30, int(_keep("vj_paso", 7)), 1,
                                            help="Sólo se usa para sugerir fechas; las fechas de cada viaje las defines tú abajo.")
            ss._k_vj_llenar = p5.toggle("Llenar el camión", bool(_keep("vj_llenar", True)),
                                        help="Si sobra capacidad en un viaje, se adelantan productos del siguiente (los que se agotan antes primero).")

        if not lins_ret:
            st.info("Ningún proveedor del plan tiene *Retiro en bodega proveedor*. Márcalo en 4 → Condiciones de Proveedor "
                    "para repartir sus productos en viajes del camión.")
        else:
            vj = _viajes_global(lins_ret)
            st.caption(f"{len(set(l['Proveedor'] for l in lins_ret))} proveedores con retiro · "
                       f"{_num(sum(l['Bultos'] * l['Peso bulto'] for l in lins_ret))} kg en total. "
                       "Los productos se reparten según cuándo se agotan, mezclando proveedores.")
            _render_viajes(lins_ret, vj)
            vj = _viajes_global(lins_ret)            # ya con la distribución aplicada
            st.markdown("#### Plan de ruta")
            _render_plan_ruta(lins_ret, vj)


# =====================================================================================
# TAB 7 - PROCESADOS (recetas: mix de frutos secos, granolas, etc.)
# =====================================================================================
if _nav == _SECC[6]:
    st.markdown("##### " + "Productos procesados: de qué están hechos")
    st.caption("Para los productos que tú preparas (Mix Frutos Secos, Granola Mix…): define qué ingredientes lleva y qué % de cada uno. "
               "El procesado ya no se compra: lo que haya que fabricar se suma a la compra de sus ingredientes en el Plan de compra. "
               "Los porcentajes son sobre el peso del producto terminado (10 kg de mix con 10 % de almendras = 1 kg de almendras).")
    if "cmp_nec" not in ss:
        st.info("Primero define la cobertura en 📊 Análisis → Cobertura.")
        st.stop()
    _nec7 = ss.cmp_nec
    _bases = sorted(_nec7["Base"].unique())
    _b2c = dict(zip(_nec7["Base"], _nec7["Clave"]))
    _rec = ss.get("cmp_recetas")
    _rec = _rec if _rec is not None else pd.DataFrame(columns=cc.COLS_RECETAS)
    if ss.get("cmp_msg_rec"):
        (st.success if ss.cmp_msg_rec.startswith("✅") else st.error)(ss.cmp_msg_rec)
        ss.cmp_msg_rec = None

    _con_rec = sorted(_rec["Producto"].unique())
    pr_sel = st.selectbox("Producto procesado (el que preparas)", _bases, index=None, placeholder="Escribe para buscar… (los que ya tienen receta: " +
                          (", ".join(_con_rec[:4]) + ("…" if len(_con_rec) > 4 else "") if _con_rec else "ninguno todavía") + ")", key="rec_prod")
    if pr_sel:
        _ex = _rec[_rec["Producto"] == pr_sel]
        _ini = pd.DataFrame({"Ingrediente": _ex["Ingrediente"].tolist(), "Porcentaje": pd.to_numeric(_ex["Porcentaje"], errors="coerce").fillna(0.0).tolist()})
        if _ini.empty:
            _ini = pd.DataFrame({"Ingrediente": pd.Series(dtype=str), "Porcentaje": pd.Series(dtype=float)})
        with st.form(f"form_rec_{_b2c[pr_sel]}"):
            st.markdown(f"**Ingredientes de {pr_sel}** (agrega filas con el + de la tabla)")
            ed_r = st.data_editor(_ini, num_rows="dynamic", hide_index=True, use_container_width=True, key=f"ed_rec_{_b2c[pr_sel]}",
                                  column_config={"Ingrediente": st.column_config.SelectboxColumn("Ingrediente", options=[b for b in _bases if b != pr_sel], required=True,
                                                                                                  help="Producto base (sin formato) que se usa como materia prima."),
                                                 "Porcentaje": st.column_config.NumberColumn("% del producto terminado", min_value=0.0, max_value=100.0,
                                                                                              step=1.0, format="%.1f", required=True)})
            g1, g2 = st.columns(2)
            ok_r = g1.form_submit_button("💾 Guardar receta", type="primary", use_container_width=True)
            del_r = g2.form_submit_button("🗑️ Eliminar receta", use_container_width=True)
        _tot = float(pd.to_numeric(ed_r["Porcentaje"], errors="coerce").fillna(0).sum()) if len(ed_r) else 0.0
        if _tot > 100.0001:
            st.warning(f"Los porcentajes suman {_tot:.1f} %: es más de 100 %.")
        else:
            st.caption(f"Suma actual: {_tot:.1f} % (lo que falte hasta 100 % no se compra, por ejemplo envase o merma).")
        if ok_r or del_r:
            _resto = _rec[_rec["Producto"] != pr_sel]
            _nuevas = []
            if ok_r:
                for r_ in ed_r.itertuples():
                    if r_.Ingrediente and pd.notna(r_.Porcentaje) and r_.Porcentaje > 0 and r_.Ingrediente in _b2c:
                        _nuevas.append({"Producto": pr_sel, "Clave": _b2c[pr_sel], "Ingrediente": r_.Ingrediente,
                                        "Clave ing": _b2c[r_.Ingrediente], "Porcentaje": float(r_.Porcentaje), "Nota": ""})
            _all = pd.concat([_resto, pd.DataFrame(_nuevas, columns=cc.COLS_RECETAS)], ignore_index=True)[cc.COLS_RECETAS]
            try:
                _cmp_save(cc.SH_RECETAS, _all)
                ss.cmp_recetas = _all
                _recalc_nec()
                ss.cmp_msg_rec = (f"✅ Receta de {pr_sel} guardada ({len(_nuevas)} ingredientes)." if ok_r and _nuevas
                                  else f"✅ Receta de {pr_sel} eliminada.")
            except Exception as e:
                ss.cmp_msg_rec = f"❌ No pude guardar en Google Sheets: {e}"
            st.rerun()

    if len(_rec):
        st.markdown("#### Recetas guardadas")
        _res = _rec.assign(P=pd.to_numeric(_rec["Porcentaje"], errors="coerce")).groupby("Producto").agg(
            Ingredientes=("Ingrediente", "count"), Suma=("P", "sum")).reset_index()
        _prod = _nec7.set_index("Base")["Producción kg"].to_dict()
        _res["A producir (kg)"] = _res["Producto"].map(_prod).fillna(0.0)
        _res["Detalle"] = _res["Producto"].map(lambda p: " · ".join(f"{i} {float(q):g}%" for i, q in
                                                                      zip(_rec.loc[_rec["Producto"] == p, "Ingrediente"], _rec.loc[_rec["Producto"] == p, "Porcentaje"])))
        st.dataframe(_res.rename(columns={"Suma": "Suma %"}), hide_index=True, use_container_width=True,
                     column_config={"Suma %": st.column_config.NumberColumn(format="%.1f"), "A producir (kg)": st.column_config.NumberColumn(format="%.1f")})
        _ing = _nec7[_nec7["Por procesados kg"] > 0][["Base", "Tipo", "Stock kg", "Por procesados kg", "Comprar kg"]]
        if len(_ing):
            st.markdown("#### Efecto en la compra de ingredientes")
            st.dataframe(_ing.sort_values("Por procesados kg", ascending=False).rename(columns={
                "Base": "Ingrediente", "Stock kg": "Stock kg/uni.", "Por procesados kg": "Demanda por procesados kg", "Comprar kg": "Comprar total kg"}),
                hide_index=True, use_container_width=True)
            st.caption("«Comprar total» ya incluye la venta propia del ingrediente más lo que consumen los procesados, descontando su stock.")
    else:
        st.info("Aún no hay recetas. Elige un producto procesado arriba y agrega sus ingredientes.")


# =====================================================================================
# PRODUCTOS NO CONSIDERADOS (pestaña de Análisis)
# =====================================================================================
if _nav == "nc":
    st.markdown("##### " + "Productos que hoy no entran al plan de compra")
    if "cmp_nec" not in ss:
        st.info("Primero define la cobertura en 📊 Análisis → Cobertura.")
    else:
        if ss.get("cmp_msg_sg"):
            (st.success if ss.cmp_msg_sg.startswith("✅") else st.error)(ss.cmp_msg_sg)
            ss.cmp_msg_sg = None
        _panel_no_considerados(expandido=True)


# =====================================================================================
# CAMBIOS DE VENTA (pestaña de Análisis)
# =====================================================================================
if _nav == "cv":
    st.markdown("##### " + "Productos cuya venta cambió")
    if "cmp_nec" not in ss:
        st.info("Primero define la cobertura en 📊 Análisis → Cobertura.")
    else:
        n_ = ss.cmp_nec
        c1_, c2_, c3_ = st.columns(3)
        pmin = c1_.number_input("Cambio mínimo %", 5.0, 500.0, float(_keep("cv_pct", 30.0)), 5.0,
                                help="Sólo avisa si la venta subió o bajó al menos este porcentaje.")
        kmin = c2_.number_input("y al menos (kg/uni.)", 0.0, 1000.0, float(_keep("cv_kg", 3.0)), 1.0,
                                help="Evita avisar por productos de muy poca venta (de 1 a 3 kg no es una tendencia).")
        solo_c = c3_.toggle("Sólo los que afectan la compra", True, key="cv_solo",
                            help="Muestra productos con compra pendiente o con aviso de stock; oculta los que tienen stock de sobra.")
        ss._k_cv_pct, ss._k_cv_kg = pmin, kmin
        sem_ = max(1, int(ss.cmp_meta.get("n_sem", 4)))

        # 1) tendencia con las semanas del reporte
        t_ = n_.copy()
        t_["Prom. previas"] = t_[[c for c in t_.columns if re.fullmatch(r"W\d+ kg", c)][:-1]].mean(axis=1) if sem_ >= 2 else 0.0
        t_ = t_[(t_["Tendencia %"].abs() >= pmin) & (t_["Tend. kg/sem"].abs() >= kmin)]
        if solo_c:
            t_ = t_[(t_["Comprar kg"] > 0) | (t_["Cobertura días hábiles"] < 15) | (t_["Alerta stock"] != "")]
        wlast = [c for c in n_.columns if re.fullmatch(r"W\d+ kg", c)]
        st.markdown("#### 1) Tendencia: última semana vs las anteriores del reporte")
        if len(wlast) < 2 or t_.empty:
            st.info("No hay productos con cambios de tendencia sobre estos umbrales.")
        else:
            v1 = pd.DataFrame({"Aviso": np.where(t_["Tend. kg/sem"] > 0, "⬆️ Sube", "⬇️ Baja"), "Producto": t_["Base"],
                               "Tipo": t_["Tipo"], "Semanas previas": t_["Prom. previas"], "Última semana": t_[wlast[-1]],
                               "Cambio %": t_["Tendencia %"].clip(upper=999), "Cambio kg/sem": t_["Tend. kg/sem"],
                               "Stock": t_["Stock kg"], "Comprar": t_["Comprar kg"]})
            v1 = v1.reindex(v1["Cambio kg/sem"].abs().sort_values(ascending=False).index)
            st.dataframe(v1, hide_index=True, use_container_width=True, height=min(420, 60 + 36 * len(v1)),
                         column_config={"Semanas previas": st.column_config.NumberColumn("Promedio semanas previas kg/uni.", format="%.1f"),
                                        "Última semana": st.column_config.NumberColumn("Última semana kg/uni.", format="%.1f"),
                                        "Cambio %": st.column_config.NumberColumn(format="%+.0f%%"),
                                        "Cambio kg/sem": st.column_config.NumberColumn("Cambio kg/uni. por semana", format="%+.1f"),
                                        "Stock": st.column_config.NumberColumn("Stock kg/uni.", format="%.1f"),
                                        "Comprar": st.column_config.NumberColumn("Comprar kg/uni.", format="%.1f")})
            st.caption("Compara la última semana del Reposition con el promedio de las semanas anteriores. Un feriado o una venta especial "
                       "pueden mover una semana: revisa antes de cambiar una compra.")

        # 2) venta real desde el stock anterior
        st.markdown("#### 2) Venta real desde el stock anterior")
        if not (ss.get("cmp_stock_base") and ss.get("cmp_stock_fecha") and ss.get("cmp_stock_base_fecha")):
            st.info("Se activa cuando actualizas el stock en 📥 Datos (compara contra el stock anterior).")
        elif ss.get("cmp_ventas_lin") is None:
            st.info("Sube el **Detalle de ventas** en 📥 Datos, con las ventas hasta la fecha del stock nuevo, para ver cuánto se vendió realmente.")
        elif ss.get("cmp_vreal_cob") is None or ss.cmp_vreal_cob < 1.0:
            _dm = max(ss.cmp_ventas["Fecha"]) if ss.get("cmp_ventas") is not None and len(ss.cmp_ventas) else None
            st.warning(f"El Detalle de ventas cargado cubre sólo el {0 if ss.cmp_vreal_cob is None else ss.cmp_vreal_cob:.0%} del período "
                       f"{ss.cmp_stock_base_fecha + timedelta(days=1):%d/%m} al {ss.cmp_stock_fecha:%d/%m}"
                       + (f" (llega hasta el {_dm:%d/%m/%Y})" if _dm else "") + ". Sube uno que llegue hasta la fecha del stock nuevo.")
        else:
            r_ = n_[n_["Venta real kg"].notna()].copy()
            r_["Dif"] = r_["Venta real kg"] - r_["Venta esperada kg"]
            ok_v = (r_["Cambio venta %"].abs() >= pmin) & (r_["Dif"].abs() >= kmin)
            ok_o = r_["Otras diferencias kg"].abs() >= max(kmin, 1.0)
            r_ = r_[ok_v | ok_o]
            if solo_c:
                r_ = r_[(r_["Comprar kg"] > 0) | (r_["Cobertura días hábiles"] < 15) | (r_["Alerta stock"] != "")]
            if r_.empty:
                st.info("No hay productos con diferencias sobre estos umbrales.")
            else:
                def _av(x):
                    a = []
                    if abs(x["Cambio venta %"]) >= pmin and abs(x["Dif"]) >= kmin:
                        a.append("⬆️ Vendió más" if x["Dif"] > 0 else "⬇️ Vendió menos")
                    if abs(x["Otras diferencias kg"]) >= max(kmin, 1.0):
                        a.append("⚠️ Faltó stock sin venta que lo explique" if x["Otras diferencias kg"] < 0 else "➕ Sobra stock (¿ingreso o devolución?)")
                    return " · ".join(a)
                v2 = pd.DataFrame({"Aviso": r_.apply(_av, axis=1), "Producto": r_["Base"], "Tipo": r_["Tipo"],
                                   "Esperada": r_["Venta esperada kg"], "Real": r_["Venta real kg"], "Cambio %": r_["Cambio venta %"].clip(upper=999),
                                   "Otras": r_["Otras diferencias kg"], "Stock": r_["Stock kg"], "Comprar": r_["Comprar kg"]})
                v2 = v2.reindex(pd.concat([r_["Dif"].abs(), r_["Otras diferencias kg"].abs()], axis=1).max(axis=1).sort_values(ascending=False).index)
                st.dataframe(v2, hide_index=True, use_container_width=True, height=min(460, 60 + 36 * len(v2)),
                             column_config={"Esperada": st.column_config.NumberColumn("Venta esperada kg/uni.", format="%.1f"),
                                            "Real": st.column_config.NumberColumn("Venta real kg/uni.", format="%.1f"),
                                            "Cambio %": st.column_config.NumberColumn(format="%+.0f%%"),
                                            "Otras": st.column_config.NumberColumn("Otras diferencias kg/uni.", format="%+.1f",
                                                                                    help="Stock real − (stock anterior − venta real). Negativo: bajó más de lo que explican las ventas (merma, vencimiento, ajuste). "
                                                                                         "Positivo: hay más de lo esperado (mercadería recibida, devolución o conteo distinto)."),
                                            "Stock": st.column_config.NumberColumn("Stock kg/uni.", format="%.1f"),
                                            "Comprar": st.column_config.NumberColumn("Comprar kg/uni.", format="%.1f")})
                st.caption(f"Período: {ss.cmp_stock_base_fecha + timedelta(days=1):%d/%m/%Y} al {ss.cmp_stock_fecha:%d/%m/%Y}. "
                           "«Esperada» = lo que se habría vendido al ritmo normal del producto; «Real» sale del Detalle de ventas. "
                           "Si recibiste mercadería en ese período, el stock sube y aparecerá en «Otras diferencias».")
