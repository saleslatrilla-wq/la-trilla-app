"""
Módulo de Pedidos – La Trilla
Pestaña 1: Requerimiento (Sala de Ventas)
Pestaña 2: Producción (Bodega / Envasado)

Canal de comunicación: Google Sheets
  - Hoja "pedidos_activos"  → pedidos con estado
  - Hoja "pedidos_chat"     → mensajes de chat por pedido
"""

import streamlit as st
import pandas as pd
from datetime import datetime

# ── utilidades ──────────────────────────────────────────────────────────────

spreadsheet = st.session_state.google_spreadsheet

SHEET_PEDIDOS = "pedidos_activos"
SHEET_CHAT    = "pedidos_chat"

COLS_PEDIDOS = ["id", "timestamp", "productos", "cantidades", "unidades",
                "notas_req", "estado", "respuesta_prod"]
COLS_CHAT    = ["pedido_id", "timestamp", "origen", "mensaje"]

def _ws(name, rows=2000, cols=20):
    try:
        return spreadsheet.worksheet(name)
    except Exception:
        return spreadsheet.add_worksheet(title=name, rows=str(rows), cols=str(cols))

def _load(sheet_name, cols):
    ws = _ws(sheet_name)
    vals = ws.get_all_values()
    if not vals or len(vals) < 2:
        return pd.DataFrame(columns=cols)
    df = pd.DataFrame(vals[1:], columns=vals[0])
    for c in cols:
        if c not in df.columns:
            df[c] = ""
    return df[cols]

def _append_row(sheet_name, row_dict, cols):
    ws = _ws(sheet_name)
    vals = ws.get_all_values()
    if not vals:
        ws.append_row(cols)
    row = [str(row_dict.get(c, "")) for c in cols]
    ws.append_row(row)

def _update_cell_by_id(sheet_name, pid, col_name, value, cols):
    """Actualiza una celda buscando por columna 'id'."""
    ws = _ws(sheet_name)
    vals = ws.get_all_values()
    if not vals:
        return
    header = vals[0]
    try:
        id_col   = header.index("id") + 1
        tgt_col  = header.index(col_name) + 1
    except ValueError:
        return
    for i, row in enumerate(vals[1:], start=2):
        if len(row) >= id_col and row[id_col - 1] == str(pid):
            ws.update_cell(i, tgt_col, str(value))
            return

def ts_now():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

def nuevo_id(df_pedidos):
    if df_pedidos.empty:
        return "P001"
    ids = df_pedidos["id"].tolist()
    nums = []
    for i in ids:
        try:
            nums.append(int(str(i).replace("P", "")))
        except Exception:
            pass
    return f"P{(max(nums) + 1):03d}" if nums else "P001"

# ── sonido de alarma (data-URI, tono repetitivo) ─────────────────────────────
ALARM_JS = """
<script>
(function() {
  if (window._alarmActive) return;
  window._alarmActive = true;

  const AudioContext = window.AudioContext || window.webkitAudioContext;
  if (!AudioContext) return;

  const ctx = new AudioContext();

  function beep(freq, start, dur) {
    const osc  = ctx.createOscillator();
    const gain = ctx.createGain();
    osc.connect(gain);
    gain.connect(ctx.destination);
    osc.type = 'square';
    osc.frequency.value = freq;
    gain.gain.setValueAtTime(0.4, ctx.currentTime + start);
    gain.gain.exponentialRampToValueAtTime(0.001, ctx.currentTime + start + dur);
    osc.start(ctx.currentTime + start);
    osc.stop(ctx.currentTime + start + dur + 0.01);
  }

  function playPattern() {
    if (!window._alarmActive) return;
    beep(880, 0.0,  0.12);
    beep(660, 0.15, 0.12);
    beep(880, 0.30, 0.12);
    beep(660, 0.45, 0.12);
    setTimeout(playPattern, 1400);
  }

  playPattern();

  window._stopAlarm = function() {
    window._alarmActive = false;
    ctx.close();
  };
})();
</script>
"""

STOP_ALARM_JS = """
<script>
if (window._stopAlarm) { window._stopAlarm(); }
window._alarmActive = false;
</script>
"""

# ── CSS compartido ────────────────────────────────────────────────────────────
st.markdown("""
<style>
/* burbuja chat Requerimiento */
.burbuja-req {
    background: #1a73e8;
    color: white;
    border-radius: 18px 18px 4px 18px;
    padding: 10px 16px;
    margin: 6px 0;
    max-width: 80%;
    font-size: 0.95rem;
    align-self: flex-end;
    margin-left: auto;
    word-wrap: break-word;
}
/* burbuja chat Producción */
.burbuja-prod {
    background: #e8f5e9;
    color: #1b5e20;
    border-radius: 18px 18px 18px 4px;
    padding: 10px 16px;
    margin: 6px 0;
    max-width: 80%;
    font-size: 0.95rem;
    align-self: flex-start;
    word-wrap: break-word;
}
.chat-ts { font-size: 0.72rem; color: #888; margin-top: 2px; }

/* tarjeta pedido */
.pedido-card {
    border: 2px solid #1a73e8;
    border-radius: 12px;
    padding: 16px;
    margin: 10px 0;
    background: #f8faff;
}

/* ALARMA */
.alarm-box {
    background: #ff0000;
    color: white;
    font-size: 2.2rem;
    font-weight: 900;
    text-align: center;
    border-radius: 16px;
    padding: 28px 20px;
    animation: pulse 0.6s ease-in-out infinite alternate;
    letter-spacing: 2px;
    margin-bottom: 20px;
}
@keyframes pulse {
    from { background: #ff0000; transform: scale(1);   box-shadow: 0 0 0px red; }
    to   { background: #cc0000; transform: scale(1.03); box-shadow: 0 0 32px red; }
}

/* badge estado */
.badge-pend  { background:#ff9800; color:white; border-radius:8px; padding:3px 10px; font-size:0.82rem; font-weight:700; }
.badge-verif { background:#2196f3; color:white; border-radius:8px; padding:3px 10px; font-size:0.82rem; font-weight:700; }
.badge-done  { background:#4caf50; color:white; border-radius:8px; padding:3px 10px; font-size:0.82rem; font-weight:700; }
.badge-nook  { background:#f44336; color:white; border-radius:8px; padding:3px 10px; font-size:0.82rem; font-weight:700; }
</style>
""", unsafe_allow_html=True)

# ── TABS principales ──────────────────────────────────────────────────────────
tab_req, tab_prod = st.tabs(["🛒 Requerimiento  (Sala de Ventas)",
                             "🏭 Producción  (Bodega / Envasado)"])

# ═══════════════════════════════════════════════════════════════════════════════
# PESTAÑA 1 – REQUERIMIENTO
# ═══════════════════════════════════════════════════════════════════════════════
with tab_req:

    st.subheader("🛒 Nuevo Requerimiento a Producción")

    # ── cargar productos desde sheet "productos" ──────────────────────────────
    @st.cache_data(ttl=120)
    def cargar_subproductos():
        try:
            ws   = spreadsheet.worksheet("productos")
            vals = ws.get_all_values()
            if not vals or len(vals) < 2:
                return []
            header = vals[0]
            sub_cols = [c for c in header if c.upper().startswith("SUB") and
                        not c.upper().endswith("_FACTOR")]
            items = []
            for row in vals[1:]:
                row_dict = dict(zip(header, row))
                for sc in sub_cols:
                    val = str(row_dict.get(sc, "")).strip()
                    if val and val not in ("", "nan", "None"):
                        items.append(val)
            return sorted(list(set(items)))
        except Exception as e:
            st.warning(f"No se pudo cargar la hoja 'productos': {e}")
            return []

    subproductos = cargar_subproductos()

    if not subproductos:
        st.error("⚠️ No se encontraron subproductos en la hoja 'productos'. "
                 "Asegúrate de que la hoja exista y tenga columnas SUB2–SUB6.")
        st.stop()

    # ── formulario de pedido ─────────────────────────────────────────────────
    with st.form("form_requerimiento", clear_on_submit=True):
        st.markdown("**Selecciona los productos que necesitas:**")

        n_items = st.number_input("¿Cuántos productos distintos?",
                                  min_value=1, max_value=10, value=1, step=1)

        seleccion_productos = []
        seleccion_cantidades = []
        seleccion_unidades   = []

        cols_form = st.columns([3, 1, 1])
        cols_form[0].markdown("**Producto**")
        cols_form[1].markdown("**Cantidad**")
        cols_form[2].markdown("**Unidad**")

        for i in range(int(n_items)):
            c1, c2, c3 = st.columns([3, 1, 1])
            prod = c1.selectbox(f"Producto {i+1}", subproductos,
                                key=f"prod_{i}", label_visibility="collapsed")
            cant = c2.number_input("Cant.", min_value=0.5, value=1.0, step=0.5,
                                   key=f"cant_{i}", label_visibility="collapsed")
            unid = c3.selectbox("Unid.", ["Unidad", "Kg", "Bolsa", "Bandeja", "Caja"],
                                key=f"unid_{i}", label_visibility="collapsed")
            seleccion_productos.append(prod)
            seleccion_cantidades.append(str(cant))
            seleccion_unidades.append(unid)

        notas = st.text_area("📝 Nota para producción (ej: ¿puede ser en una sola bolsa?)",
                             placeholder="Escribe aquí cualquier indicación especial...",
                             max_chars=300)

        enviado = st.form_submit_button("🚀 Enviar Requerimiento a Producción",
                                        type="primary", use_container_width=True)

    if enviado:
        df_ped = _load(SHEET_PEDIDOS, COLS_PEDIDOS)
        pid    = nuevo_id(df_ped)
        row    = {
            "id":           pid,
            "timestamp":    ts_now(),
            "productos":    " | ".join(seleccion_productos),
            "cantidades":   " | ".join(seleccion_cantidades),
            "unidades":     " | ".join(seleccion_unidades),
            "notas_req":    notas.strip(),
            "estado":       "PENDIENTE",
            "respuesta_prod": ""
        }
        _append_row(SHEET_PEDIDOS, row, COLS_PEDIDOS)
        st.success(f"✅ Requerimiento **{pid}** enviado a Producción.")
        st.balloons()

    # ── chat de seguimiento ───────────────────────────────────────────────────
    st.markdown("---")
    st.subheader("💬 Chat con Producción")

    df_ped_chat = _load(SHEET_PEDIDOS, COLS_PEDIDOS)
    if df_ped_chat.empty:
        st.info("Aún no hay requerimientos enviados.")
    else:
        pedidos_activos = df_ped_chat[df_ped_chat["estado"] != "ENTREGADO"]
        if pedidos_activos.empty:
            pedidos_activos = df_ped_chat

        pid_sel = st.selectbox(
            "Selecciona el pedido:",
            pedidos_activos["id"].tolist(),
            format_func=lambda x: (
                f"{x} – " +
                df_ped_chat.loc[df_ped_chat['id']==x, 'productos'].values[0][:60] + "..."
                if len(df_ped_chat.loc[df_ped_chat['id']==x, 'productos'].values[0]) > 60
                else f"{x} – " + df_ped_chat.loc[df_ped_chat['id']==x, 'productos'].values[0]
            )
        )

        # info del pedido seleccionado
        row_sel = df_ped_chat[df_ped_chat["id"] == pid_sel].iloc[0]
        estado  = row_sel["estado"]
        badge_map = {
            "PENDIENTE":         "badge-pend",
            "RECIBIDO":          "badge-verif",
            "VERIFICANDO STOCK": "badge-verif",
            "EN PREPARACIÓN":    "badge-verif",
            "LISTO":             "badge-done",
            "SIN STOCK":         "badge-nook",
            "ENTREGADO":         "badge-done",
        }
        badge_cls = badge_map.get(estado, "badge-pend")

        st.markdown(f"""
        <div class="pedido-card">
          <b>Pedido {pid_sel}</b> &nbsp; <span class="{badge_cls}">{estado}</span><br>
          <b>Productos:</b> {row_sel['productos']}<br>
          <b>Cantidades:</b> {row_sel['cantidades']} {row_sel['unidades']}<br>
          {"<b>Nota:</b> " + row_sel['notas_req'] if row_sel['notas_req'] else ""}
          {"<br><b>Respuesta producción:</b> " + row_sel['respuesta_prod'] if row_sel['respuesta_prod'] else ""}
        </div>
        """, unsafe_allow_html=True)

        # mensajes del chat
        df_chat = _load(SHEET_CHAT, COLS_CHAT)
        msgs    = df_chat[df_chat["pedido_id"] == pid_sel].reset_index(drop=True)

        chat_container = st.container()
        with chat_container:
            if msgs.empty:
                st.caption("Sin mensajes aún.")
            for _, m in msgs.iterrows():
                if m["origen"] == "REQ":
                    st.markdown(
                        f'<div class="burbuja-req">🛒 {m["mensaje"]}'
                        f'<div class="chat-ts">{m["timestamp"]}</div></div>',
                        unsafe_allow_html=True)
                else:
                    st.markdown(
                        f'<div class="burbuja-prod">🏭 {m["mensaje"]}'
                        f'<div class="chat-ts">{m["timestamp"]}</div></div>',
                        unsafe_allow_html=True)

        with st.form("chat_req_form", clear_on_submit=True):
            msg_req = st.text_input("Mensaje a Producción",
                                    placeholder="Ej: ¿pueden ser en una sola bolsa?",
                                    label_visibility="collapsed")
            if st.form_submit_button("Enviar 📤", use_container_width=True):
                if msg_req.strip():
                    _append_row(SHEET_CHAT, {
                        "pedido_id": pid_sel,
                        "timestamp": ts_now(),
                        "origen":    "REQ",
                        "mensaje":   msg_req.strip()
                    }, COLS_CHAT)
                    st.rerun()

        if st.button("🔄 Actualizar estado", use_container_width=True):
            st.rerun()

        # marcar como entregado
        if estado not in ("ENTREGADO",):
            if st.button("✅ Marcar como ENTREGADO", key="marcar_entregado",
                         use_container_width=True):
                _update_cell_by_id(SHEET_PEDIDOS, pid_sel, "estado",
                                   "ENTREGADO", COLS_PEDIDOS)
                st.success("Pedido marcado como Entregado.")
                st.rerun()


# ═══════════════════════════════════════════════════════════════════════════════
# PESTAÑA 2 – PRODUCCIÓN
# ═══════════════════════════════════════════════════════════════════════════════
with tab_prod:

    st.subheader("🏭 Panel de Producción / Bodega")

    # ── cargar pedidos ────────────────────────────────────────────────────────
    df_ped = _load(SHEET_PEDIDOS, COLS_PEDIDOS)

    pendientes = df_ped[df_ped["estado"].isin(
        ["PENDIENTE", "RECIBIDO", "VERIFICANDO STOCK", "EN PREPARACIÓN"])
    ].reset_index(drop=True)

    # ── ALARMA si hay pedidos PENDIENTE ──────────────────────────────────────
    hay_pendiente = not pendientes[pendientes["estado"] == "PENDIENTE"].empty

    if hay_pendiente:
        st.markdown(
            '<div class="alarm-box">🚨 ¡NUEVO PEDIDO! — REQUIERE ATENCIÓN 🚨</div>',
            unsafe_allow_html=True
        )
        st.components.v1.html(ALARM_JS, height=0)

    if not pendientes.empty:
        for _, ped in pendientes.iterrows():
            pid     = ped["id"]
            estado  = ped["estado"]
            is_pend = estado == "PENDIENTE"

            badge_map = {
                "PENDIENTE":         "badge-pend",
                "RECIBIDO":          "badge-verif",
                "VERIFICANDO STOCK": "badge-verif",
                "EN PREPARACIÓN":    "badge-verif",
            }
            badge_cls = badge_map.get(estado, "badge-pend")

            # ── tarjeta del pedido ────────────────────────────────────────────
            with st.container():
                st.markdown(f"""
                <div class="pedido-card" style="{'border-color:#ff0000;background:#fff5f5;' if is_pend else ''}">
                  <b>📦 Pedido {pid}</b> &nbsp; <span class="{badge_cls}">{estado}</span><br>
                  🕐 <i>{ped['timestamp']}</i><br><br>
                  <b>Productos solicitados:</b><br>
                """, unsafe_allow_html=True)

                prods = ped["productos"].split(" | ")
                cants = ped["cantidades"].split(" | ")
                unids = ped["unidades"].split(" | ")
                for p, c, u in zip(prods, cants, unids):
                    st.markdown(f"&nbsp;&nbsp;&nbsp;• **{p}** — {c} {u}")

                if ped["notas_req"]:
                    st.info(f"📝 Nota del solicitante: **{ped['notas_req']}**")

                st.markdown("</div>", unsafe_allow_html=True)

                # ── botones de respuesta rápida ───────────────────────────────
                st.markdown("**Respuesta rápida:**")
                b1, b2, b3, b4, b5 = st.columns(5)

                def cambiar_estado(pid_p, nuevo_estado, respuesta=""):
                    _update_cell_by_id(SHEET_PEDIDOS, pid_p, "estado",
                                       nuevo_estado, COLS_PEDIDOS)
                    if respuesta:
                        _update_cell_by_id(SHEET_PEDIDOS, pid_p, "respuesta_prod",
                                           respuesta, COLS_PEDIDOS)
                    _append_row(SHEET_CHAT, {
                        "pedido_id": pid_p,
                        "timestamp": ts_now(),
                        "origen":    "PROD",
                        "mensaje":   f"[{nuevo_estado}] {respuesta}".strip(" []")
                    }, COLS_CHAT)
                    st.components.v1.html(STOP_ALARM_JS, height=0)

                if b1.button("✅ RECIBIDO", key=f"rec_{pid}", use_container_width=True):
                    cambiar_estado(pid, "RECIBIDO", "Pedido recibido, estamos en ello.")
                    st.rerun()

                if b2.button("🔍 VERIFICANDO STOCK", key=f"ver_{pid}", use_container_width=True):
                    cambiar_estado(pid, "VERIFICANDO STOCK", "Verificando disponibilidad de stock.")
                    st.rerun()

                if b3.button("⚙️ EN PREPARACIÓN", key=f"prep_{pid}", use_container_width=True):
                    cambiar_estado(pid, "EN PREPARACIÓN", "Producto en proceso de preparación/envasado.")
                    st.rerun()

                if b4.button("🟢 LISTO", key=f"listo_{pid}", use_container_width=True):
                    cambiar_estado(pid, "LISTO", "¡Pedido listo para retirar!")
                    st.rerun()

                if b5.button("❌ SIN STOCK", key=f"stock_{pid}", use_container_width=True):
                    cambiar_estado(pid, "SIN STOCK", "No hay stock disponible.")
                    st.rerun()

                # ── chat producción ────────────────────────────────────────────
                with st.expander(f"💬 Chat del pedido {pid}", expanded=is_pend):
                    df_chat = _load(SHEET_CHAT, COLS_CHAT)
                    msgs    = df_chat[df_chat["pedido_id"] == pid].reset_index(drop=True)

                    if msgs.empty:
                        st.caption("Sin mensajes aún.")
                    for _, m in msgs.iterrows():
                        if m["origen"] == "REQ":
                            st.markdown(
                                f'<div class="burbuja-req">🛒 {m["mensaje"]}'
                                f'<div class="chat-ts">{m["timestamp"]}</div></div>',
                                unsafe_allow_html=True)
                        else:
                            st.markdown(
                                f'<div class="burbuja-prod">🏭 {m["mensaje"]}'
                                f'<div class="chat-ts">{m["timestamp"]}</div></div>',
                                unsafe_allow_html=True)

                    with st.form(f"chat_prod_{pid}", clear_on_submit=True):
                        msg_prod = st.text_input(
                            "Mensaje a Sala de Ventas",
                            placeholder="Ej: Vamos en 10 minutos...",
                            label_visibility="collapsed",
                            key=f"msg_prod_inp_{pid}"
                        )
                        if st.form_submit_button("Enviar 📤", use_container_width=True):
                            if msg_prod.strip():
                                _append_row(SHEET_CHAT, {
                                    "pedido_id": pid,
                                    "timestamp": ts_now(),
                                    "origen":    "PROD",
                                    "mensaje":   msg_prod.strip()
                                }, COLS_CHAT)
                                st.rerun()

                st.markdown("---")
    else:
        st.success("✅ Sin pedidos pendientes por ahora.")

    # ── historial de pedidos terminados ──────────────────────────────────────
    with st.expander("📜 Historial de pedidos (LISTO / SIN STOCK / ENTREGADO)"):
        terminados = df_ped[df_ped["estado"].isin(
            ["LISTO", "SIN STOCK", "ENTREGADO"])].reset_index(drop=True)
        if terminados.empty:
            st.info("No hay pedidos finalizados aún.")
        else:
            st.dataframe(
                terminados[["id", "timestamp", "productos", "cantidades",
                             "unidades", "estado", "respuesta_prod"]],
                use_container_width=True, hide_index=True
            )

    col_ref1, col_ref2 = st.columns(2)
    if col_ref1.button("🔄 Actualizar pedidos", use_container_width=True, key="ref_prod"):
        st.rerun()
    if col_ref2.button("🔕 Silenciar alarma", use_container_width=True, key="sil_alarm"):
        st.components.v1.html(STOP_ALARM_JS, height=0)
