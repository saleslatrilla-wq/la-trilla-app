"""
compras_core.py  -  Lógica pura (sin Streamlit) del módulo de Compras de La Trilla.

Replica y mejora el flujo del Excel "COMPRAS Analisis Reposicion de Stock":
  Reposición de Stock  ->  Análisis Actual (kg)  ->  Precios (proveedores)  ->  Cotizaciones / Orden de compra
"""
from __future__ import annotations

import math
import re
import unicodedata
from datetime import date, datetime, timedelta
from io import BytesIO

import pandas as pd

# ---------------------------------------------------------------------------
# Nombres de pestañas en Google Sheets (todas con prefijo compras_)
# ---------------------------------------------------------------------------
SH_PRECIOS = "compras_precios"
SH_FORMATOS = "compras_formatos"
SH_CALENDARIO = "compras_calendario"
SH_ORDENES = "compras_ordenes"
SH_PROVEEDORES = "compras_proveedores"
SH_SESION = "compras_sesion"
SH_SIEMPRE = "compras_siempre"
COLS_SIEMPRE = ["Producto", "Clave", "Cantidad", "Nota"]
SH_RECETAS = "compras_recetas"
COLS_RECETAS = ["Producto", "Clave", "Ingrediente", "Clave ing", "Porcentaje", "Nota"]

COLS_PRECIOS = ["Producto", "Formato", "Precio Neto Kg", "Proveedor", "Disponible", "Actualizado", "Nota"]
COLS_FORMATOS = ["Formato", "Kg por bulto"]
COLS_SESION = ["ID", "Fecha", "Detalle", "Parte", "Partes", "Datos"]
COLS_CALENDARIO = ["Fecha", "Peso", "Nota"]
ENTREGAS = ["Despacha en tienda", "Retiro en bodega proveedor"]
CAP_VIAJE_DEF, MAX_VIAJE_DEF = 1500.0, 1750.0
COLS_PROVEEDORES = ["Proveedor", "Factura", "Condición de pago", "Días cheque", "Entrega", "Máx. viajes", "Mín. kg compra", "Días despacho", "Pedir hasta", "Nota"]
COLS_ORDENES = ["Fecha", "N° Orden", "Proveedor", "Producto", "Formato", "Bultos", "Kg",
                "Precio Neto Kg", "Total Neto", "Estado", "Viaje", "Fecha Viaje"]

# Orden de prioridad EXACTO de la fórmula de Excel (Análisis Actual!A)
FORMAT_TOKENS = ["Env.", "Bidon", "Caja", "Bolsa", "Saco", "Granel"]

# Feriados por defecto (editables en la app). Sólo los que se conocen con certeza.
FERIADOS_DEFAULT = {
    "2026-10-12": "Encuentro de Dos Mundos",
    "2026-10-31": "Día de las Iglesias Evangélicas",
    "2026-11-01": "Todos los Santos",
    "2026-12-08": "Inmaculada Concepción",
    "2026-12-25": "Navidad",
    "2027-01-01": "Año Nuevo",
}


# ---------------------------------------------------------------------------
# Utilidades de texto
# ---------------------------------------------------------------------------
def norm(s) -> str:
    """Clave de comparación: sin tildes, minúsculas, espacios colapsados."""
    if s is None or (isinstance(s, float) and math.isnan(s)):
        return ""
    s = unicodedata.normalize("NFKD", str(s))
    s = "".join(c for c in s if not unicodedata.combining(c))
    return re.sub(r"\s+", " ", s).strip().casefold()


def to_float(v, default=0.0) -> float:
    if v is None:
        return default
    if isinstance(v, (int, float)):
        return default if (isinstance(v, float) and math.isnan(v)) else float(v)
    s = str(v).strip().replace("$", "").replace(" ", "")
    if s in ("", "nan", "None", "-"):
        return default
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".")
    elif "," in s:
        s = s.replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return default


# ---------------------------------------------------------------------------
# Producto -> (base, formato) y kg por unidad
# ---------------------------------------------------------------------------
def split_producto(nombre: str) -> tuple[str, str]:
    """'Almendra Laminada Env. 100gr' -> ('Almendra Laminada', 'Env. 100gr').
    Sin token de formato -> (nombre, 'Unidad'). Misma regla que el Excel."""
    n = str(nombre).strip()
    cut = len(n)
    for tok in FORMAT_TOKENS:
        i = n.find(tok)
        if i >= 0:
            cut = i
            break
    base = n[:cut].strip()
    fmt = n[cut:].strip() if cut < len(n) else ""
    return base, (fmt if fmt else "Unidad")


_RX_KG = re.compile(r"(\d+(?:[.,]\d+)?)\s*(kgs?|kilos?)\b", re.I)
_RX_GR = re.compile(r"(\d+(?:[.,]\d+)?)\s*(grs?|g|gramos?)\b", re.I)
_RX_UN = re.compile(r"(\d+)\s*(u|un|unid\w*|unidades)\b", re.I)


def kg_from_text(fmt: str):
    """Intenta deducir kg (o unidades) desde el texto del formato. None si no puede."""
    f = str(fmt).replace(",", ".")
    f_low = norm(f)
    if f_low in ("unidad", "granel", "env. unidad", "granel unidad") or f_low.startswith("granel"):
        return 1.0
    m = _RX_KG.search(f)
    if m:
        return float(m.group(1))
    m = _RX_GR.search(f)
    if m:
        return float(m.group(1)) / 1000.0
    m = _RX_UN.search(f)
    if m:
        return float(m.group(1))
    return None


def build_formato_map(df_formatos: pd.DataFrame | None) -> dict:
    mp = {}
    if df_formatos is not None and not df_formatos.empty:
        for _, r in df_formatos.iterrows():
            k = norm(r.get("Formato", ""))
            v = to_float(r.get("Kg por bulto", None), default=float("nan"))
            if k and not math.isnan(v) and v > 0:
                mp[k] = v
    return mp


def kg_por_formato(fmt: str, fmap: dict):
    k = norm(fmt)
    if k in fmap:
        return fmap[k]
    return kg_from_text(fmt)


# ---------------------------------------------------------------------------
# Lectura de archivos (Bsale/Reposition y Detalle de ventas)
# ---------------------------------------------------------------------------
def _bytes(f) -> bytes:
    if hasattr(f, "getvalue"):
        return f.getvalue()
    if hasattr(f, "read"):
        return f.read()
    with open(f, "rb") as fh:
        return fh.read()


def read_excel_robust(f, **kw) -> dict:
    """Lee todas las hojas probando openpyxl -> calamine -> html. Devuelve {hoja: df}."""
    data = _bytes(f)
    last = None
    for engine in (None, "calamine"):
        try:
            return pd.read_excel(BytesIO(data), sheet_name=None, engine=engine, **kw)
        except Exception as e:  # noqa: BLE001
            last = e
    try:
        return {"html": pd.read_html(BytesIO(data), encoding="utf-8")[0]}
    except Exception:  # noqa: BLE001
        raise ValueError(
            "No pude leer el archivo. Si es un .xlsx de Bsale con estilos dañados, agrega "
            "'python-calamine' a requirements.txt. Detalle: %s" % last)


def read_reposicion(f) -> tuple[pd.DataFrame, dict]:
    """Devuelve (df, meta). df: Tipo, Producto, SKU, S1..Sn (unidades), Stock, Valor.
    meta: sucursal, fecha_fin, semanas [(label, d_ini, d_fin)]"""
    sheets = read_excel_robust(f, header=None)
    raw = next((d for d in sheets.values() if d.shape[0] > 6), None)
    if raw is None:
        raise ValueError("El archivo de reposición está vacío.")
    hdr = raw.index[raw.iloc[:, 0].astype(str).str.strip() == "Tipo de Producto"]
    if len(hdr) == 0:
        raise ValueError("No encontré la fila 'Tipo de Producto' (¿es el reporte Reposition Stock?).")
    h = hdr[0]
    cols = [str(c).strip() for c in raw.iloc[h]]
    df = raw.iloc[h + 1:].copy()
    df.columns = cols
    df = df.dropna(subset=["Producto"]).reset_index(drop=True)

    meta = {"sucursal": "", "fecha_fin": None, "semanas": []}
    for i in range(h):
        k = str(raw.iloc[i, 0]).strip()
        if k == "Sucursal":
            meta["sucursal"] = str(raw.iloc[i, 1])
        if k == "Fecha Termino":
            try:
                meta["fecha_fin"] = datetime.strptime(str(raw.iloc[i, 1])[:10], "%d/%m/%Y").date()
            except ValueError:
                pass

    wk_cols = []
    for c in cols:
        m = re.search(r"(\d{2}/\d{2}/\d{4})\s+al\s+(\d{2}/\d{2}/\d{4})", c)
        if c.startswith("Unidades Vendida") and m:
            d1 = datetime.strptime(m.group(1), "%d/%m/%Y").date()
            d2 = datetime.strptime(m.group(2), "%d/%m/%Y").date()
            wk_cols.append(c)
            meta["semanas"].append((c, d1, d2))
    out = pd.DataFrame({
        "Tipo": df["Tipo de Producto"].astype(str).str.strip(),
        "Producto": df["Producto"].astype(str).str.strip(),
        "SKU": df.get("SKU", ""),
    })
    for i, c in enumerate(wk_cols, 1):
        out[f"S{i}"] = pd.to_numeric(df[c], errors="coerce").fillna(0.0)
    out["Stock"] = pd.to_numeric(df["Stock Actual"], errors="coerce").fillna(0.0)
    out["Valor Existencia"] = pd.to_numeric(df.get("Valor Total Existencia", 0), errors="coerce").fillna(0.0)
    meta["n_sem"] = len(wk_cols)
    return out, meta


def read_ventas_detalle(f) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Detalle de ventas -> (resumen diario, líneas por documento y producto).
    Resumen: Fecha, Documentos, Lineas, Venta. Líneas: Fecha, Doc, Tipo doc, Cliente, RUT, Producto, Cantidad."""
    sheets = read_excel_robust(f)
    v = next(iter(sheets.values()))
    v = v[v["Tipo Movimiento"].astype(str).str.lower() == "venta"].copy()
    v["Fecha"] = pd.to_datetime(v["Fecha Venta"], format="%d/%m/%Y", errors="coerce").dt.date
    cal = v.groupby("Fecha").agg(
        Documentos=("Numero del documento", "nunique"),
        Lineas=("SKU", "size"),
        Venta=("Venta Total Bruta", "sum"),
    ).reset_index()
    col = lambda n: v[n] if n in v.columns else ""
    lin = pd.DataFrame({
        "Fecha": v["Fecha"],
        "Doc": col("Numero del documento").map(lambda x: str(int(x)) if isinstance(x, (int, float)) and not pd.isna(x) else str(x)),
        "Tipo doc": col("Tipo de Documento").astype(str),
        "Cliente": col("Nombre Cliente").fillna("Sin cliente").astype(str),
        "RUT": col("Cliente RUT").fillna("").astype(str),
        "Producto": col("Producto / Servicio + Variante").astype(str).str.strip(),
        "Cantidad": pd.to_numeric(col("Cantidad"), errors="coerce").fillna(0.0),
    }).dropna(subset=["Fecha"])
    lin = lin.groupby(["Fecha", "Doc", "Tipo doc", "Cliente", "RUT", "Producto"], as_index=False, sort=False)["Cantidad"].sum()
    return cal, lin


def read_ventas_calendario(f) -> pd.DataFrame:
    """Resumen diario del Detalle de ventas: Fecha, Documentos, Lineas, Venta Bruta."""
    return read_ventas_detalle(f)[0]


COLS_ESPECIALES = ["Fecha", "Cliente", "RUT", "Doc", "Tipo doc", "Producto", "Clave", "Detalle", "Kg", "x día normal",
                   "Docs cliente", "Veces especial", "Llave"]


def ventas_especiales(lin: pd.DataFrame | None, base: pd.DataFrame, meta: dict, fmap: dict,
                      mult: float = 3.0, min_kg: float = 5.0) -> pd.DataFrame:
    """Ventas de un documento que equivalen a `mult` o más días de venta normal de ese producto
    (y al menos `min_kg` kg/uni.), dentro de la ventana analizada. Kg en las mismas unidades que el plan."""
    if lin is None or lin.empty or base is None or base.empty:
        return pd.DataFrame(columns=COLS_ESPECIALES)
    d_ini, d_fin = meta["semanas"][0][1], meta["semanas"][-1][2]
    w = lin[(lin["Fecha"] >= d_ini) & (lin["Fecha"] <= d_fin)].copy()
    if w.empty:
        return pd.DataFrame(columns=COLS_ESPECIALES)
    prods = w["Producto"].unique()
    info = {}
    for p in prods:
        b, fm = split_producto(p)
        info[p] = (b, norm(b), fm, kg_por_formato(fm, fmap))
    w["Base"] = w["Producto"].map(lambda p: info[p][0])
    w["Clave"] = w["Producto"].map(lambda p: info[p][1])
    w["Fmt"] = w["Producto"].map(lambda p: info[p][2])
    w["kgu"] = w["Producto"].map(lambda p: info[p][3])
    w = w[w["kgu"].notna()]
    w["Kg"] = w["Cantidad"] * w["kgu"]
    # un renglón por documento y producto base (la línea de mayor cantidad describe el formato)
    w = w.sort_values("Kg", ascending=False)
    g = w.groupby(["Fecha", "Doc", "Tipo doc", "Cliente", "RUT", "Clave"], as_index=False, sort=False).agg(
        Producto=("Base", "first"), Kg=("Kg", "sum"), Cant=("Cantidad", "first"), Fmt=("Fmt", "first"))
    diaria = dict(zip(base["Clave"], base["Venta diaria kg"]))
    g["_d"] = g["Clave"].map(diaria)
    g = g[(g["_d"] > 0) & (g["Kg"] >= min_kg)]
    g["x día normal"] = g["Kg"] / g["_d"]
    g = g[g["x día normal"] >= mult].copy()
    if g.empty:
        return pd.DataFrame(columns=COLS_ESPECIALES)
    g["Detalle"] = [f"{c:g} × {f}" if f != "Unidad" else f"{c:g} u" for c, f in zip(g["Cant"], g["Fmt"])]
    docs_cli = w[w["RUT"] != ""].groupby("RUT")["Doc"].nunique().to_dict()
    veces = g[g["RUT"] != ""].groupby("RUT")["Doc"].count().to_dict()
    g["Docs cliente"] = g["RUT"].map(lambda r: docs_cli.get(r, 0) if r else 0)
    g["Veces especial"] = g["RUT"].map(lambda r: veces.get(r, 0) if r else 0)
    g["Llave"] = g["Doc"].astype(str) + "|" + g["Clave"]
    return g[COLS_ESPECIALES].sort_values("x día normal", ascending=False).reset_index(drop=True)


def ajusta_por_especiales(g: pd.DataFrame, esp: pd.DataFrame, repite: set, peso_ventana: float, n_sem: int) -> pd.DataFrame:
    """Agrega a la base: 'Especiales kg' (todas las ventas especiales), 'Venta diaria ajustada' y 'Venta semanal ajustada'
    (sin las especiales que NO están marcadas como 'Se repite')."""
    out = g.copy()
    tot, excl = {}, {}
    if esp is not None and not esp.empty:
        for r in esp.to_dict("records"):
            tot[r["Clave"]] = tot.get(r["Clave"], 0.0) + r["Kg"]
            if r["Llave"] not in repite:
                excl[r["Clave"]] = excl.get(r["Clave"], 0.0) + r["Kg"]
    out["Especiales kg"] = out["Clave"].map(tot).fillna(0.0)
    ex = out["Clave"].map(excl).fillna(0.0)
    vend = (out["Vendido kg"] - ex).clip(lower=0)
    out["Venta diaria ajustada"] = vend / peso_ventana if peso_ventana > 0 else 0.0
    out["Venta semanal ajustada"] = vend / max(1, n_sem)
    return out


# ---------------------------------------------------------------------------
# Calendario y pesos de días
# ---------------------------------------------------------------------------
def peso_base(d: date, feriados: dict) -> float:
    """Lun-Vie = 1, Sáb = 0.5, Dom = 0, feriado = 0 (el sábado feriado también 0)."""
    if d.isoformat() in feriados:
        return 0.0
    wd = d.weekday()
    return 1.0 if wd < 5 else (0.5 if wd == 5 else 0.0)


def calendario_historico(ventas_dia: pd.DataFrame | None, d_ini: date, d_fin: date,
                         feriados: dict) -> pd.DataFrame:
    """Un día por fila entre d_ini y d_fin con peso sugerido.
    Con el detalle de ventas, un día de baja actividad (docs < 60% de la mediana de su tipo de día)
    recibe un peso proporcional (múltiplos de 0.25) aunque no esté en la lista de feriados."""
    days = [d_ini + timedelta(days=i) for i in range((d_fin - d_ini).days + 1)]
    rows = []
    docs = {}
    if ventas_dia is not None and not ventas_dia.empty:
        docs = dict(zip(ventas_dia["Fecha"], ventas_dia["Documentos"]))
    # medianas por tipo de día
    med = {}
    for tipo, cond in (("lv", lambda d: d.weekday() < 5), ("sab", lambda d: d.weekday() == 5)):
        vals = [docs[d] for d in days if d in docs and cond(d) and d.isoformat() not in feriados]
        med[tipo] = float(pd.Series(vals).median()) if vals else None
    for d in days:
        base = 1.0 if d.weekday() < 5 else (0.5 if d.weekday() == 5 else 0.0)
        nota = ""
        w = base
        es_fer = d.isoformat() in feriados
        if es_fer:
            w, nota = 0.0, "Feriado: " + feriados[d.isoformat()]
        if d in docs and base > 0:
            tipo = "lv" if d.weekday() < 5 else "sab"
            m = med[tipo]
            if m and docs[d] < 0.6 * m:
                ratio = docs[d] / m
                w = base * max(0.0, round(ratio * 4) / 4)
                nota = (nota + " · " if nota else "") + f"Baja venta ({docs[d]} docs vs mediana {m:.0f})"
            elif es_fer:  # feriado pero con venta normal: manda el dato real
                w = base
                nota += " · abrió con venta normal"
        if d.weekday() == 6 and d not in docs:
            nota = nota or "Domingo"
        rows.append({"Fecha": d, "Día": ["Lun", "Mar", "Mié", "Jue", "Vie", "Sáb", "Dom"][d.weekday()],
                     "Docs": int(docs[d]) if d in docs else None, "Peso": w, "Nota": nota})
    return pd.DataFrame(rows)


def calendario_futuro(d_ini: date, d_fin: date, feriados: dict) -> pd.DataFrame:
    rows = []
    for i in range((d_fin - d_ini).days + 1):
        d = d_ini + timedelta(days=i)
        w = peso_base(d, feriados)
        nota = ("Feriado: " + feriados[d.isoformat()]) if d.isoformat() in feriados else ("Domingo" if d.weekday() == 6 else "")
        rows.append({"Incluir": w > 0, "Fecha": d,
                     "Día": ["Lun", "Mar", "Mié", "Jue", "Vie", "Sáb", "Dom"][d.weekday()],
                     "Peso": w, "Nota": nota})
    return pd.DataFrame(rows)


# ---------------------------------------------------------------------------
# Demanda diaria por producto base (kg)
# ---------------------------------------------------------------------------
def demanda_por_producto(rep: pd.DataFrame, meta: dict, fmap: dict, peso_ventana: float) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Agrupa el reporte por producto base y convierte todo a kg.
    peso_ventana = suma de pesos de los días de la ventana analizada (p.ej. 4 sem ≈ 22).
    Devuelve (df_base, df_lineas_sin_formato)."""
    sem_cols = [c for c in rep.columns if re.fullmatch(r"S\d+", c)]
    rows, sin_fmt = [], []
    for _, r in rep.iterrows():
        base, fmt = split_producto(r["Producto"])
        kgu = kg_por_formato(fmt, fmap)
        vendidas = float(sum(r[c] for c in sem_cols))
        if kgu is None:
            if vendidas > 0 or r["Stock"] > 0:
                sin_fmt.append({"Producto": r["Producto"], "Base": base, "Formato": fmt,
                                "Unid. vendidas": vendidas, "Stock": r["Stock"]})
            continue
        rows.append({"Base": base, "Clave": norm(base), "Tipo": r["Tipo"],
                     "Stock kg": r["Stock"] * kgu, "Vendido kg": vendidas * kgu,
                     "Venta semanal kg": 0.0, "Unidad": fmt == "Unidad",
                     **{f"W{i} kg": float(r[col]) * kgu for i, col in enumerate(sem_cols, 1)}})
    d = pd.DataFrame(rows)
    if d.empty:
        return d, pd.DataFrame(sin_fmt)
    g = d.groupby("Clave", as_index=False).agg(
        Base=("Base", "first"), Tipo=("Tipo", "first"),
        **{"Stock kg": ("Stock kg", "sum"), "Vendido kg": ("Vendido kg", "sum"), "Unidad": ("Unidad", "max")},
        **{f"W{i} kg": (f"W{i} kg", "sum") for i in range(1, len(sem_cols) + 1)})
    n_sem = max(1, meta.get("n_sem", len(sem_cols)))
    g["Venta semanal kg"] = g["Vendido kg"] / n_sem
    g["Venta diaria kg"] = g["Vendido kg"] / peso_ventana if peso_ventana > 0 else 0.0
    return g, pd.DataFrame(sin_fmt)


def calcular_necesidad(g: pd.DataFrame, peso_horizonte: float, colchon_pct: float = 0.0) -> pd.DataFrame:
    out = g.copy()
    out["Demanda kg"] = out["Venta diaria kg"] * peso_horizonte * (1 + colchon_pct / 100.0)
    out["Comprar kg"] = (out["Demanda kg"] - out["Stock kg"]).clip(lower=0)
    out["Cobertura días hábiles"] = out.apply(
        lambda r: (r["Stock kg"] / r["Venta diaria kg"]) if r["Venta diaria kg"] > 0 else float("inf"), axis=1)
    return out


# ---------------------------------------------------------------------------
# Proveedores: ranking por precio/kg y elección de formato
# ---------------------------------------------------------------------------
def normaliza_precios(df: pd.DataFrame | None, fmap: dict) -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame(columns=COLS_PRECIOS + ["Clave", "Kg bulto"])
    d = df.copy()
    for c in COLS_PRECIOS:
        if c not in d.columns:
            d[c] = ""
    d["Clave"] = d["Producto"].map(norm)
    d["Precio Neto Kg"] = d["Precio Neto Kg"].map(lambda x: to_float(x, 0.0))
    d["Disponible"] = d["Disponible"].map(lambda x: norm(x) not in ("no", "n", "false", "0", "sin stock", "quiebre"))
    d["Kg bulto"] = d["Formato"].map(lambda f: kg_por_formato(f, fmap) if str(f).strip() else None)
    d = d[(d["Clave"] != "") & (d["Proveedor"].astype(str).str.strip() != "")]
    return d.reset_index(drop=True)


def agrupa_precios(precios_n: pd.DataFrame) -> dict:
    """Pre-agrupa (una sola vez) las ofertas vigentes por producto, ordenadas por precio/kg."""
    if precios_n is None or precios_n.empty:
        return {}
    ok = precios_n[precios_n["Disponible"] & (precios_n["Precio Neto Kg"] > 0)]
    return {k: g.sort_values(["Precio Neto Kg", "Proveedor"]).reset_index(drop=True) for k, g in ok.groupby("Clave")}


def opciones_producto(clave: str, precios_n: pd.DataFrame, comprar_kg: float, grupos: dict | None = None) -> pd.DataFrame:
    """Opciones ordenadas por precio/kg (sólo con stock y precio). Calcula bultos, costo real y diferencia
    porcentual contra la opción de menor costo total para comprar_kg."""
    if grupos is not None:
        o = grupos.get(clave)
        if o is None:
            return pd.DataFrame()
        o = o.copy()
    else:
        if precios_n.empty:
            return pd.DataFrame()
        o = precios_n[(precios_n["Clave"] == clave) & precios_n["Disponible"] & (precios_n["Precio Neto Kg"] > 0)].copy()
        if o.empty:
            return o
        o = o.sort_values(["Precio Neto Kg", "Proveedor"]).reset_index(drop=True)
    o["Opción"] = range(1, len(o) + 1)
    o["Sin formato"] = o["Kg bulto"].isna() | (o["Kg bulto"] <= 0)
    o["Kg bulto"] = o["Kg bulto"].where(~o["Sin formato"], 1.0)  # sin formato: se compra por kg
    if comprar_kg > 0:
        o["Bultos"] = (comprar_kg / o["Kg bulto"] - 1e-9).map(math.ceil).astype(int)
    else:
        o["Bultos"] = 0
    o["Kg total"] = o["Bultos"] * o["Kg bulto"]
    o["Total Neto"] = o["Kg total"] * o["Precio Neto Kg"]
    mn = o["Total Neto"].min()
    o["Dif %"] = ((o["Total Neto"] / mn - 1) * 100) if mn and mn > 0 else 0.0
    return o


def posicion_elegida(ops: pd.DataFrame, elegido=None, criterio: str = "kg") -> int:
    """Posición (0..n-1) de la opción vigente: la elegida a mano si sigue disponible;
    si no, la sugerida (menor precio/kg o menor costo total de la compra)."""
    if elegido:
        m = ops[(ops["Proveedor"] == elegido[0]) & (ops["Formato"] == elegido[1])]
        if len(m):
            return int(m.index[0])
    if criterio == "total":
        return int(ops["Total Neto"].idxmin())
    return 0


def plan_de_compra(necesidad: pd.DataFrame, precios_n: pd.DataFrame, elegido: dict | None = None,
                   solo_comprar: bool = True, criterio: str = "kg", grupos: dict | None = None) -> pd.DataFrame:
    """Para cada producto con necesidad: toma la opción elegida o la sugerida."""
    elegido = elegido or {}
    if grupos is None:
        grupos = agrupa_precios(precios_n)
    rows = []
    for d in necesidad.to_dict("records"):
        comprar = d["Comprar kg"]
        if solo_comprar and comprar <= 0:
            continue
        ops = opciones_producto(d["Clave"], precios_n, comprar, grupos)
        n_ops = len(ops)
        base = {"Producto": d["Base"], "Clave": d["Clave"], "Tipo": d["Tipo"],
                "Stock kg": round(d["Stock kg"], 2), "Venta diaria kg": round(d["Venta diaria kg"], 3),
                "Venta semanal kg": round(d["Venta semanal kg"], 2),
                "Comprar kg": round(comprar, 2), "Opciones": n_ops, "Es unidad": bool(d.get("Unidad", False))}
        if n_ops == 0:
            base.update({"Opción": 0, "Proveedor": "— Sin proveedor —", "Formato": "", "Bultos": 0,
                         "Kg total": 0.0, "Kg bulto": 0.0, "Precio Neto Kg": 0.0, "Total Neto": 0.0, "Estado": "Sin proveedor"})
        else:
            p = posicion_elegida(ops, elegido.get(d["Clave"]), criterio)
            s = ops.iloc[p]
            base.update({"Opción": int(s["Opción"]), "Proveedor": s["Proveedor"],
                         "Formato": s["Formato"] if not s["Sin formato"] else "(kg)",
                         "Bultos": int(s["Bultos"]), "Kg total": round(float(s["Kg total"]), 2), "Kg bulto": float(s["Kg bulto"]),
                         "Precio Neto Kg": float(s["Precio Neto Kg"]), "Total Neto": round(float(s["Total Neto"]), 0),
                         "Estado": "Falta formato (kg)" if s["Sin formato"] else "OK"})
        rows.append(base)
    return pd.DataFrame(rows)


# ---------------------------------------------------------------------------
# Migración desde el Excel actual (hoja Precios + Cotizaciones + FORMATOS)
# ---------------------------------------------------------------------------
def migrar_excel_compras(f) -> tuple[pd.DataFrame, pd.DataFrame, dict]:
    """Convierte la matriz 'Precios' (producto x proveedor) a formato largo.
    - Celda con 1  -> proveedor SIN STOCK (se importa Disponible=No, sin precio).
    - Celda vacía/0 -> proveedor no cotiza ese producto (se omite).
    Formato por producto: el más usado en Cotizaciones (puede ajustarse luego en la app)."""
    sheets = read_excel_robust(f, header=None)
    p = sheets["Precios"]
    hdr = p.iloc[1].tolist()
    # proveedores empiezan en 'NAMA...' (col L) - detectamos tras 'Proveedor 4'
    start = next(i for i, v in enumerate(hdr) if str(v).strip() == "Proveedor 4") + 1
    provs = {i: str(hdr[i]).strip() for i in range(start, len(hdr)) if str(hdr[i]).strip() not in ("", "nan")}
    body = p.iloc[2:]

    # Formato más frecuente por producto desde Cotizaciones
    c = sheets.get("Cotizaciones")
    fmt_prod = {}
    if c is not None:
        cc = c.iloc[2:, [0, 1]].dropna()
        cc.columns = ["Producto", "Formato"]
        for k, grp in cc.groupby(cc["Producto"].map(norm)):
            fmt_prod[k] = grp["Formato"].astype(str).str.strip().mode().iloc[0]
    rows, today = [], date.today().isoformat()
    for _, r in body.iterrows():
        nombre = str(r.iloc[0]).strip()
        if nombre in ("", "nan"):
            continue
        for i, prov in provs.items():
            v = to_float(r.iloc[i], 0.0) if i < len(r) else 0.0
            if v == 0:
                continue
            disponible = "No" if v == 1 else "Sí"
            rows.append({"Producto": nombre, "Proveedor": prov, "Formato": fmt_prod.get(norm(nombre), ""),
                         "Precio Neto Kg": "" if v == 1 else v, "Disponible": disponible,
                         "Actualizado": today, "Nota": "Importado de Excel" + (" (quiebre)" if v == 1 else "")})
    precios = pd.DataFrame(rows, columns=COLS_PRECIOS)

    # Formatos: unión de FORMATOS!A:B y FORMATOS!D:E
    fm = sheets["FORMATOS"].iloc[2:]
    f1 = fm.iloc[:, [0, 1]].dropna(); f1.columns = COLS_FORMATOS
    f2 = fm.iloc[:, [3, 4]].dropna(); f2.columns = COLS_FORMATOS
    formatos = pd.concat([f1, f2], ignore_index=True)
    formatos["Formato"] = formatos["Formato"].astype(str).str.strip()
    formatos["Kg por bulto"] = formatos["Kg por bulto"].map(lambda x: to_float(x, 0.0))
    formatos = formatos[formatos["Kg por bulto"] > 0].drop_duplicates("Formato").reset_index(drop=True)
    stats = {"productos": precios["Producto"].nunique(), "ofertas": int((precios["Disponible"] == "Sí").sum()),
             "quiebres": int((precios["Disponible"] == "No").sum()),
             "sin_formato": int((precios["Formato"] == "").sum())}
    return precios, formatos, stats



# ---------------------------------------------------------------------------
# Peso logístico (kg reales que ocupa en el camión)
# ---------------------------------------------------------------------------
_RX_VOL_L = re.compile(r"(\d+(?:[.,]\d+)?)\s*(lt|lts|l|litros?)\b", re.I)
_RX_VOL_ML = re.compile(r"(\d+(?:[.,]\d+)?)\s*(ml|cc)\b", re.I)


def peso_desde_nombre(nombre: str):
    """Peso aproximado de una unidad leyendo el nombre ('500ml' -> 0.5, '250gr' -> 0.25). None si no puede."""
    f = str(nombre).replace(",", ".")
    for rx, div in ((_RX_KG, 1.0), (_RX_GR, 1000.0), (_RX_VOL_ML, 1000.0), (_RX_VOL_L, 1.0)):
        m = rx.search(f)
        if m:
            return float(m.group(1)) / div
    return None


def mapa_peso_logistico(df: pd.DataFrame | None) -> dict:
    """{producto normalizado: kg por unidad} desde la hoja productos (columna PESO_LOGISTICO)."""
    mp = {}
    if df is not None and not df.empty and "PESO_LOGISTICO" in df.columns and "PRODUCTO" in df.columns:
        for p, v in zip(df["PRODUCTO"], df["PESO_LOGISTICO"]):
            x = to_float(v, default=float("nan"))
            if norm(p) and not math.isnan(x) and x > 0:
                mp[norm(p)] = x
    return mp


def factor_peso(clave: str, nombre: str, es_unidad: bool, mapa: dict) -> tuple[float, str]:
    """kg de camión por cada 'kg/uni.' del plan. Productos por kilo = 1.
    Productos por unidad: PESO_LOGISTICO; sin dato se estima con el nombre y, si no se puede, 1 kg."""
    if clave in mapa:
        return mapa[clave], "Dato"
    if not es_unidad:
        return 1.0, "kg"
    est = peso_desde_nombre(nombre)
    if est and est > 0:
        return est, "Estimado"
    return 1.0, "Sin dato"


# ---------------------------------------------------------------------------
# Viajes: cuándo se agota cada producto y cómo repartir la compra entre viajes
# ---------------------------------------------------------------------------
DIAS_ES = ["Lun", "Mar", "Mié", "Jue", "Vie", "Sáb", "Dom"]


def _prefijo_pesos(ini: date, peso_fn, dias: int = 400) -> list:
    pre, acc = [0.0], 0.0
    for i in range(dias):
        acc += peso_fn(ini + timedelta(days=i))
        pre.append(acc)
    return pre


def _acum(pre: list, ini: date, hasta: date) -> float:
    k = (hasta - ini).days
    if k <= 0:
        return 0.0
    return pre[min(k, len(pre) - 1)]


def parse_dias(txt) -> list:
    """'Vie' / 'Mar, Vie' / 'lunes y jueves' -> [weekday,...] (0=lunes)."""
    out = []
    for t in re.split(r"[,;/ y]+", norm(str(txt or ""))):
        for i, d in enumerate(DIAS_ES):
            if t and norm(d).startswith(t[:3]) and i not in out:
                out.append(i)
    return out


def proximo_despacho(hoy: date, desp: list, lim: list):
    """Próximo despacho alcanzable desde `hoy`: devuelve (fecha_limite_pedido, fecha_despacho) o None.
    `lim` se empareja con `desp` por posición (si hay uno solo, vale para todos; si no hay, el límite es el día anterior)."""
    if not desp:
        return None
    mejor = None
    for k, wd in enumerate(desp):
        lw = lim[k] if k < len(lim) else (lim[0] if len(lim) == 1 else None)
        for n in range(0, 15):
            D = hoy + timedelta(days=n)
            if D.weekday() != wd:
                continue
            if lw is None:
                C = D - timedelta(days=1)
            else:
                C = D - timedelta(days=(D.weekday() - lw) % 7)
            if C >= hoy:
                if mejor is None or D < mejor[1]:
                    mejor = (C, D)
                break
    return mejor


def fecha_quiebre(stock: float, diaria: float, ini: date, pre: list):
    """Primer día en que se agota el stock (el stock es el de la víspera de `ini`). None si no se agota."""
    if diaria <= 0:
        return None
    if stock <= 0:
        return ini
    for k in range(1, len(pre)):
        if pre[k] * diaria >= stock:
            return ini + timedelta(days=k - 1)
    return None


def distribuir_viajes(lineas: list, fechas: list, ini: date, peso_fn, buffer: int = 2, colchon: float = 0.0,
                      cap: float = CAP_VIAJE_DEF, llenar: bool = True, max_viajes: dict | None = None) -> pd.DataFrame:
    """Reparte los bultos de cada línea entre los viajes sin dejar sin stock a ningún producto.
    lineas: dicts con Clave, Bultos, Kg bulto, Peso bulto, Stock, Diaria (stock/diaria en kg/uni. del plan).
    1) Mínimo en cada viaje: lo necesario para que el stock llegue hasta el viaje siguiente (+ margen de días).
    2) Si `llenar`, el camión de cada viaje se completa (hasta `cap`) adelantando compras de viajes siguientes,
       empezando por los productos que se agotan antes.
    max_viajes: {proveedor: n°} = en cuántos viajes como máximo se puede repartir la compra de ese proveedor
    (1 = todo junto en un solo viaje). Los proveedores sin límite se reparten libremente.
    Devuelve una fila por línea con V1..Vn (bultos), Quiebre, Prioridad y Orden."""
    n = len(fechas)
    pre = _prefijo_pesos(ini, peso_fn)
    f = 1 + colchon / 100.0
    filas = []
    for ln in lineas:
        B, kgb = int(ln["Bultos"]), max(float(ln["Kg bulto"]), 1e-9)
        stock, diaria = float(ln["Stock"]), float(ln["Diaria"])
        q = fecha_quiebre(stock, diaria, ini, pre)
        lim = q - timedelta(days=buffer) if q else None
        k = n if lim is None else sum(1 for t in fechas if t <= lim)
        r, prev = [], 0
        for j in range(n):
            if j == n - 1:
                rj = B
            else:
                cons = _acum(pre, ini, fechas[j + 1] + timedelta(days=buffer)) * diaria * f
                rj = min(B, max(0, math.ceil((cons - stock) / kgb - 1e-9)))
            rj = max(rj, prev)
            r.append(rj)
            prev = rj
        a = [r[0]] + [r[j] - r[j - 1] for j in range(1, n)]
        if n == 1:
            prio = "🔴 Riesgo" if k == 0 else "🟢 OK"
        else:
            prio = "🔴 Riesgo" if k == 0 else ("🔴 Crítico" if k == 1 else ("🟡 Normal" if k == 2 else "🟢 Puede esperar"))
        filas.append({"Clave": ln["Clave"], "pb": float(ln["Peso bulto"]), "a": a, "Quiebre": q,
                      "Prioridad": prio, "k": k, "lim": lim or date.max, "prov": ln.get("Proveedor"), "B": B, "kgb": kgb,
                      "stock": stock, "diaria": diaria, "perm": None})
    # ---- proveedores con límite de viajes: se compra sólo en M viajes (el primero, el más temprano en que algo se necesita) ----
    mv = {norm(k_): int(v_) for k_, v_ in (max_viajes or {}).items() if v_ and int(v_) >= 1}
    por_prov = {}
    for x in filas:
        por_prov.setdefault(norm(x["prov"] or ""), []).append(x)
    for pk, xs in por_prov.items():
        M = mv.get(pk)
        if not M or M >= n:
            continue
        j0 = min((next((j for j, v in enumerate(x["a"]) if v > 0), n - 1) for x in xs), default=n - 1)
        if M == 1 or j0 >= n - 1:
            S = [j0] if M == 1 else [n - 1]
        else:
            S = sorted({j0} | {j0 + int(round((n - 1 - j0) * i / (M - 1))) for i in range(1, M)})
        for x in xs:
            r = []
            for idx, s_ in enumerate(S):
                if idx == len(S) - 1:
                    rj = x["B"]
                else:
                    cons = _acum(pre, ini, fechas[S[idx + 1]] + timedelta(days=buffer)) * x["diaria"] * f
                    rj = min(x["B"], max(0, math.ceil((cons - x["stock"]) / x["kgb"] - 1e-9)))
                r.append(max(rj, r[-1] if r else 0))
            a = [0] * n
            for idx, s_ in enumerate(S):
                a[s_] = r[idx] - (r[idx - 1] if idx else 0)
            x["a"], x["perm"] = a, set(S)
    if llenar and n > 1:
        orden = sorted(range(len(filas)), key=lambda i: (filas[i]["k"], filas[i]["lim"]))
        for j in range(n - 1):
            carga = sum(x["a"][j] * x["pb"] for x in filas)
            for i in orden:
                libre = cap - carga
                if libre <= 0:
                    break
                x = filas[i]
                if x["perm"] is not None and j not in x["perm"]:
                    continue      # este proveedor no puede sumar un viaje más
                for jj in range(j + 1, n):
                    if x["a"][jj] <= 0 or x["pb"] <= 0:
                        continue
                    mover = min(x["a"][jj], int(libre // x["pb"]))
                    if mover > 0:
                        x["a"][jj] -= mover
                        x["a"][j] += mover
                        carga += mover * x["pb"]
                        libre = cap - carga
    out = []
    for x in filas:
        d = {"Clave": x["Clave"], "Prioridad": x["Prioridad"], "Quiebre": x["Quiebre"], "_k": x["k"], "_lim": x["lim"]}
        for j in range(n):
            d[f"V{j + 1}"] = int(x["a"][j])
        out.append(d)
    df = pd.DataFrame(out)
    if not df.empty:
        df = df.sort_values(["_k", "_lim"], kind="stable").drop(columns=["_k", "_lim"]).reset_index(drop=True)
    return df


def pesos_por_viaje(asig: pd.DataFrame, pesos_bulto: dict, n: int) -> list:
    """kg de camión de cada viaje. pesos_bulto: {Clave: kg por bulto}."""
    return [float(sum(int(r[f"V{j + 1}"]) * pesos_bulto.get(r["Clave"], 0.0) for r in asig.to_dict("records")))
            for j in range(n)]


def sugerir_viajes(lineas: list, f1: date, paso: int, ini: date, peso_fn, buffer: int = 2, colchon: float = 0.0,
                   cap: float = CAP_VIAJE_DEF, llenar: bool = True, maximo: int = 8, max_viajes: dict | None = None):
    """(n viajes, días entre viajes) con el menor número de viajes en que todo cabe bajo `cap`.
    Prueba primero con `paso` días entre viajes y, si no alcanza, con intervalos más cortos. None si no cabe nunca."""
    pb = {l["Clave"]: float(l["Peso bulto"]) for l in lineas}
    for n in range(1, maximo + 1):
        for p in range(max(1, paso), 0, -1):
            fechas = [f1 + timedelta(days=p * i) for i in range(n)]
            w = pesos_por_viaje(distribuir_viajes(lineas, fechas, ini, peso_fn, buffer, colchon, cap, llenar, max_viajes), pb, n)
            if max(w) <= cap + 1e-6:
                return n, p
            if n == 1:
                break
    return None


# ---------------------------------------------------------------------------
# Guardar / cargar el análisis (estado de la sesión) en una hoja de Google Sheets
# ---------------------------------------------------------------------------
def _enc(o):
    import numpy as np
    if o is None or isinstance(o, (bool, str)):
        return o
    if isinstance(o, np.bool_):
        return bool(o)
    if isinstance(o, (int, np.integer)):
        return int(o)
    if isinstance(o, (float, np.floating)):
        return None if math.isnan(o) else float(o)
    if o is pd.NaT or o is pd.NA:
        return None
    if isinstance(o, datetime):
        return {"__dt": o.isoformat()}
    if isinstance(o, date):
        return {"__d": o.isoformat()}
    if isinstance(o, tuple):
        return {"__t": [_enc(x) for x in o]}
    if isinstance(o, (set, frozenset)):
        return {"__s": [_enc(x) for x in sorted(o, key=str)]}
    if isinstance(o, list):
        return [_enc(x) for x in o]
    if isinstance(o, dict):
        return {"__m": [[_enc(k), _enc(v)] for k, v in o.items()]}
    if isinstance(o, pd.DataFrame):
        return {"__df": {"c": [str(c) for c in o.columns],
                         "r": [[_enc(v) for v in row] for row in o.itertuples(index=False, name=None)]}}
    raise TypeError(f"No se puede guardar un valor de tipo {type(o).__name__}")


def _dec(o):
    if isinstance(o, list):
        return [_dec(x) for x in o]
    if isinstance(o, dict):
        if "__d" in o:
            return date.fromisoformat(o["__d"])
        if "__dt" in o:
            return datetime.fromisoformat(o["__dt"])
        if "__t" in o:
            return tuple(_dec(x) for x in o["__t"])
        if "__s" in o:
            return set(_dec(x) for x in o["__s"])
        if "__m" in o:
            return {_dec(k): _dec(v) for k, v in o["__m"]}
        if "__df" in o:
            d = o["__df"]
            return pd.DataFrame([[_dec(v) for v in row] for row in d["r"]], columns=d["c"])
        return {k: _dec(v) for k, v in o.items()}
    return o


def empaquetar(estado: dict) -> str:
    """Estado (dict con DataFrames, fechas, sets, tuplas...) -> texto comprimido (base64) para guardar en una hoja."""
    import base64, json, zlib
    raw = json.dumps(_enc(estado), separators=(",", ":"), ensure_ascii=False).encode("utf-8")
    return base64.b64encode(zlib.compress(raw, 9)).decode("ascii")


def desempaquetar(texto: str) -> dict:
    import base64, json, zlib
    return _dec(json.loads(zlib.decompress(base64.b64decode(texto.encode("ascii"))).decode("utf-8")))


def partir(texto: str, n: int = 40000) -> list:
    """Google Sheets admite hasta 50.000 caracteres por celda."""
    return [texto[i:i + n] for i in range(0, len(texto), n)] or [""]


# ---------------------------------------------------------------------------
# Estacionalidad: índice de temporada por producto con las ventas del año anterior
# ---------------------------------------------------------------------------
K_CREDIBILIDAD = 30.0       # kg/uni. del año anterior a partir de los cuales el índice propio pesa más que el de su categoría


def pesos_dias_ly(ly_cal: pd.DataFrame, d0: date, d1: date) -> dict:
    """Peso de cada día del año anterior. Día sin ventas dentro del rango = tienda cerrada (peso 0);
    días de baja actividad se ajustan igual que en el calendario histórico."""
    cal = calendario_historico(ly_cal, d0, d1, {})
    con_ventas = set(ly_cal["Fecha"])
    return {r["Fecha"]: (float(r["Peso"]) if r["Fecha"] in con_ventas else 0.0) for r in cal.to_dict("records")}


def indice_temporada(ly_lin: pd.DataFrame | None, ly_cal: pd.DataFrame | None, base: pd.DataFrame, meta: dict, fmap: dict,
                     ini: date, fin: date, mult: float = 3.0, min_kg: float = 5.0, lo: float = 0.5, hi: float = 3.0,
                     dias: int = 364) -> dict:
    """Índice de temporada por producto = (venta diaria del año anterior en las fechas a cubrir) /
    (venta diaria del año anterior en las fechas equivalentes a la ventana analizada). Ambos con días hábiles equivalentes.
    Productos de poco volumen se apoyan en el índice de su categoría (y éste en el de toda la tienda).
    Las ventas especiales del año anterior se excluyen con el mismo criterio que este año.
    Devuelve {'df': DataFrame por Clave | None, 'aviso': str, 'ref': (d0,d1), 'tgt': (d0,d1), 'global': float}."""
    if ly_lin is None or ly_cal is None or ly_lin.empty or ly_cal.empty or base is None or base.empty:
        return {"df": None, "aviso": "Sin ventas del año anterior.", "global": 1.0}
    sh = timedelta(days=dias)
    r0, r1 = meta["semanas"][0][1] - sh, meta["semanas"][-1][2] - sh
    t0, t1 = ini - sh, fin - sh
    dmin, dmax = min(ly_cal["Fecha"]), max(ly_cal["Fecha"])
    w = pesos_dias_ly(ly_cal, dmin, dmax)
    def wsum(a, b):
        return sum(v for d, v in w.items() if a <= d <= b)
    def cobertura(a, b):
        n = (b - a).days + 1
        return sum(1 for i in range(n) if dmin <= a + timedelta(days=i) <= dmax) / n
    c_ref, c_tgt = cobertura(r0, r1), cobertura(t0, t1)
    w_ref, w_tgt = wsum(r0, r1), wsum(t0, t1)
    base_info = {"ref": (r0, r1), "tgt": (t0, t1), "cob_ref": c_ref, "cob_tgt": c_tgt, "global": 1.0}
    if c_tgt < 0.7 or c_ref < 0.7 or w_ref <= 0 or w_tgt <= 0:
        return {**base_info, "df": None,
                "aviso": (f"El archivo del año anterior ({dmin:%d/%m/%Y} al {dmax:%d/%m/%Y}) no cubre bien las fechas equivalentes: "
                          f"referencia {r0:%d/%m/%Y}–{r1:%d/%m/%Y} ({c_ref:.0%}), horizonte {t0:%d/%m/%Y}–{t1:%d/%m/%Y} ({c_tgt:.0%}).")}
    # kg por documento y producto base
    prods = ly_lin["Producto"].unique()
    info = {}
    for p in prods:
        b, fm = split_producto(p)
        info[p] = (norm(b), kg_por_formato(fm, fmap))
    x = ly_lin[["Fecha", "Doc", "Producto", "Cantidad"]].copy()
    x["Clave"] = x["Producto"].map(lambda p: info[p][0])
    x["kgu"] = x["Producto"].map(lambda p: info[p][1])
    x = x[x["kgu"].notna()]
    x["Kg"] = x["Cantidad"] * x["kgu"]
    x = x.groupby(["Fecha", "Doc", "Clave"], as_index=False, sort=False)["Kg"].sum()
    w_all = sum(w.values())
    diaria_all = (x.groupby("Clave")["Kg"].sum() / w_all) if w_all > 0 else pd.Series(dtype=float)
    x["_d"] = x["Clave"].map(diaria_all)
    x = x[~((x["_d"] > 0) & (x["Kg"] >= min_kg) & (x["Kg"] >= mult * x["_d"]))]          # sin ventas especiales
    in_ref = (x["Fecha"] >= r0) & (x["Fecha"] <= r1)
    in_tgt = (x["Fecha"] >= t0) & (x["Fecha"] <= t1)
    f = pd.DataFrame({"ref_kg": x[in_ref].groupby("Clave")["Kg"].sum(), "tgt_kg": x[in_tgt].groupby("Clave")["Kg"].sum()}).fillna(0.0)
    tipo = dict(zip(base["Clave"], base["Tipo"]))
    f["Tipo"] = f.index.map(lambda c: tipo.get(c))
    def _idx(tgt, ref):
        return (tgt / w_tgt) / (ref / w_ref) if ref > 0 else float("nan")
    g_idx = _idx(f["tgt_kg"].sum(), f["ref_kg"].sum())
    g_idx = 1.0 if (isinstance(g_idx, float) and math.isnan(g_idx)) else g_idx
    cat = {}
    for t, gg in f.dropna(subset=["Tipo"]).groupby("Tipo"):
        v = _idx(gg["tgt_kg"].sum(), gg["ref_kg"].sum())
        cat[t] = g_idx if (isinstance(v, float) and math.isnan(v)) else v
    rows = []
    ly_ref_d = (f["ref_kg"] / w_ref).to_dict()
    for r in base.to_dict("records"):
        c = r["Clave"]
        ref = float(f["ref_kg"].get(c, 0.0))
        tgt = float(f["tgt_kg"].get(c, 0.0))
        c_idx = cat.get(r["Tipo"], g_idx)
        raw = _idx(tgt, ref) if ref > 0 else float("nan")
        n = min(ref, tgt * w_ref / w_tgt)          # el lado con menos volumen limita la confianza en el cociente
        z = n / (n + K_CREDIBILIDAD)
        idx = c_idx if (isinstance(raw, float) and math.isnan(raw)) else z * raw + (1 - z) * c_idx
        idx = min(max(idx, lo), hi)
        diaria_ly = ref / w_ref if ref > 0 else float("nan")
        delta = (r["Venta diaria kg"] / diaria_ly - 1.0) if (ref >= 10 and diaria_ly > 0) else float("nan")
        rows.append({"Clave": c, "Índice calculado": idx, "Índice propio": raw, "Índice categoría": c_idx,
                     "Kg año ant. referencia": ref, "Kg año ant. horizonte": tgt, "Δ año anterior": delta})
    return {**base_info, "df": pd.DataFrame(rows), "aviso": "", "global": g_idx}


def ventas_ly_rango(ly_lin: pd.DataFrame | None, ly_cal: pd.DataFrame | None, fmap: dict, d0: date, d1: date):
    """Kg (o unidades) vendidos por producto base en el año anterior entre d0 y d1.
    Devuelve (DataFrame Clave/Base/Kg, fracción de esos días que cubre el archivo)."""
    vacio = pd.DataFrame(columns=["Clave", "Base", "Kg"])
    if ly_lin is None or ly_cal is None or ly_lin.empty or ly_cal.empty:
        return vacio, 0.0
    dmin, dmax = min(ly_cal["Fecha"]), max(ly_cal["Fecha"])
    n = (d1 - d0).days + 1
    cob = sum(1 for i in range(n) if dmin <= d0 + timedelta(days=i) <= dmax) / max(n, 1)
    w = ly_lin[(ly_lin["Fecha"] >= d0) & (ly_lin["Fecha"] <= d1)]
    if w.empty:
        return vacio, cob
    info = {}
    for p in w["Producto"].unique():
        b, fm = split_producto(p)
        info[p] = (norm(b), b, kg_por_formato(fm, fmap))
    x = pd.DataFrame({"Clave": w["Producto"].map(lambda p: info[p][0]), "Base": w["Producto"].map(lambda p: info[p][1]),
                      "kgu": w["Producto"].map(lambda p: info[p][2]), "Cantidad": w["Cantidad"]})
    x = x[x["kgu"].notna()]
    x["Kg"] = x["Cantidad"] * x["kgu"]
    g = x.groupby("Clave", as_index=False).agg(Base=("Base", "first"), Kg=("Kg", "sum"))
    return g, cob


def leer_stock_actual(f) -> pd.DataFrame:
    """Reporte 'Stock actual' (Casa Matriz): devuelve SKU, Producto, Stock. Los productos sin stock no vienen en el reporte."""
    sheets = read_excel_robust(f, header=None)
    raw = next((d for d in sheets.values() if d.shape[0] > 6), None)
    if raw is None:
        raise ValueError("El archivo de stock está vacío.")
    hdr = raw.index[raw.iloc[:, 0].astype(str).str.strip() == "Tipo de Producto"]
    if len(hdr) == 0:
        raise ValueError("No encontré la fila 'Tipo de Producto' (¿es el reporte Stock actual?).")
    h = hdr[0]
    cols = [str(c).strip() for c in raw.iloc[h]]
    df = raw.iloc[h + 1:].copy()
    df.columns = cols
    for c in ("SKU", "Producto", "Stock"):
        if c not in df.columns:
            raise ValueError(f"Falta la columna «{c}» en el reporte de stock.")
    df = df.dropna(subset=["Producto"])
    out = pd.DataFrame({"SKU": df["SKU"].astype(str).str.strip(), "Producto": df["Producto"].astype(str).str.strip(),
                        "Stock": pd.to_numeric(df["Stock"], errors="coerce").fillna(0.0)})
    return out[out["SKU"].ne("") & out["SKU"].ne("nan")].reset_index(drop=True)


def aplicar_recetas(nec: pd.DataFrame, rec: pd.DataFrame | None) -> pd.DataFrame:
    """Productos procesados (mix, granolas…): su necesidad de compra se traslada a los ingredientes según el % de la receta.
    - El procesado deja de comprarse (Comprar kg = 0, 'Se produce' = True) y queda 'Producción kg' = lo que hay que fabricar.
    - Cada ingrediente suma a su demanda lo que consumen los procesados (Demanda, Venta diaria y semanal, Comprar, Cobertura).
    Admite recetas encadenadas (un mix que usa otro mix)."""
    out = nec.copy()
    for col_ in ("Demanda kg", "Venta diaria kg", "Venta semanal kg", "Comprar kg", "Cobertura días hábiles", "Stock kg"):
        out[col_] = out[col_].astype(float)
    out["Producción kg"] = 0.0
    out["Por procesados kg"] = 0.0
    out["Se produce"] = False
    if rec is None or len(rec) == 0 or out.empty:
        return out
    r = rec.copy()
    r["Porcentaje"] = pd.to_numeric(r["Porcentaje"], errors="coerce").fillna(0.0)
    r = r[(r["Porcentaje"] > 0) & r["Clave"].isin(out["Clave"]) & r["Clave ing"].isin(out["Clave"])]
    if r.empty:
        return out
    ix = out.set_index("Clave")
    dem0, dia0, sem0, stk = ix["Demanda kg"], ix["Venta diaria kg"], ix["Venta semanal kg"], ix["Stock kg"]
    procs = set(r["Clave"])
    add_d = {k: 0.0 for k in ix.index}
    add_dia, add_sem = dict(add_d), dict(add_d)
    comp = {}
    for _ in range(6):                       # punto fijo (recetas encadenadas)
        comp = {k: max(0.0, float(dem0[k]) + add_d[k] - float(stk[k])) for k in ix.index}
        n_d, n_dia, n_sem = {k: 0.0 for k in ix.index}, {k: 0.0 for k in ix.index}, {k: 0.0 for k in ix.index}
        for p_, i_, pc_ in zip(r["Clave"], r["Clave ing"], r["Porcentaje"]):
            f = float(pc_) / 100.0
            n_d[i_] += comp[p_] * f
            n_dia[i_] += (float(dia0[p_]) + add_dia[p_]) * f
            n_sem[i_] += (float(sem0[p_]) + add_sem[p_]) * f
        add_d, add_dia, add_sem = n_d, n_dia, n_sem
    comp = {k: max(0.0, float(dem0[k]) + add_d[k] - float(stk[k])) for k in ix.index}
    for k in ix.index:
        m = out["Clave"] == k
        if k in procs:
            out.loc[m, "Producción kg"] = comp[k]
            out.loc[m, "Se produce"] = True
            out.loc[m, "Comprar kg"] = 0.0
        else:
            out.loc[m, "Comprar kg"] = comp[k]
        out.loc[m, "Por procesados kg"] = add_d[k]
        out.loc[m, "Demanda kg"] = float(dem0[k]) + add_d[k]
        out.loc[m, "Venta diaria kg"] = float(dia0[k]) + add_dia[k]
        out.loc[m, "Venta semanal kg"] = float(sem0[k]) + add_sem[k]
        dd = float(dia0[k]) + add_dia[k]
        out.loc[m, "Cobertura días hábiles"] = float(stk[k]) / dd if dd > 0 else float("inf")
    return out
