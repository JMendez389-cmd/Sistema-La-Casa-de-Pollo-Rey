import base64
import html
import json
import os
import shutil
import time
from datetime import date, datetime, timedelta
import pandas as pd
import streamlit as st

ARCHIVO_DATOS = "inventario_tienda.json"
CARPETA_RESPALDOS = "respaldos"
CARPETA_MEDIA = "media_app"
DIAS_AVISO_VENCIMIENTO = 5
UMBRAL_STOCK_AMARILLO = 5.0   # 5 unidades -> amarillo
UMBRAL_STOCK_ROJO = 4.0       # <= 4 unidades -> rojo (>= 6 -> verde)
UNIDADES_MEDIDA = ["Unidad", "Libra", "Onza", "Kilogramo", "Litro", "Caja", "Bolsa"]
CATEGORIAS_INICIALES = [
    "Pollo fresco",
    "Congelados",
    "Insumos y adobos",
    "Bebidas y empaques",
    "General",
]
TITULOS_PESTANAS_BASE = [
    "➕ Registrar Producto",
    "📋 Inventario",
    "📦 Compra y Reposición",
    "🛒 Registrar Venta",
    "📈 Control de Ventas",
    "📉 Merma y Liquidación",
    "🔔 Notificaciones",
]
PLANTILLA_PRINCIPAL = {
    "nombre": "Plantilla Principal",
    "oscuro": {
        "bg_app": "linear-gradient(180deg, #0D1526 0%, #132038 55%, #172742 100%)",
        "txt_main": "#F1F5F9",
        "h2_color": "#FDE68A",
        "h3_color": "#F8FAFC",
        "strong_color": "#93C5FD",
        "cap_color": "#CBD5E1",
        "metric_bg": "linear-gradient(160deg, #1B2B48 0%, #142138 100%)",
        "metric_bdr": "rgba(125, 165, 224, 0.42)",
        "metric_top": "#FACC15",
        "metric_lbl": "#93C5FD",
        "metric_val": "#F8FAFC",
        "inp_bg": "#213354",
        "inp_bdr": "#38BDF8",
        "form_bg": "#142138",
        "form_bdr": "rgba(125, 165, 224, 0.40)",
        "btn_bg": "linear-gradient(180deg, #1E5FD6 0%, #123F94 100%)",
        "btn_bdr": "#F4B400",
        "banner_bg": "linear-gradient(95deg, #0B3A8D 0%, #1654BE 68%, #F4B400 100%)",
    },
    "claro": {
        "bg_app": "linear-gradient(180deg, #F1F5F9 0%, #FFFFFF 45%, #FFFBEB 100%)",
        "txt_main": "#111827",
        "h2_color": "#072A6E",
        "h3_color": "#111827",
        "strong_color": "#072A6E",
        "cap_color": "#1F2937",
        "metric_bg": "#FFFFFF",
        "metric_bdr": "#111827",
        "metric_top": "#0B3A8D",
        "metric_lbl": "#072A6E",
        "metric_val": "#111827",
        "inp_bg": "#FFFFFF",
        "inp_bdr": "#38BDF8",
        "form_bg": "#FFFFFF",
        "form_bdr": "#38BDF8",
        "btn_bg": "linear-gradient(180deg, #1E5FD6 0%, #123F94 100%)",
        "btn_bdr": "#F4B400",
        "banner_bg": "linear-gradient(95deg, #0B3A8D 0%, #1654BE 68%, #F4B400 100%)",
    },
}

# Contraseña maestra: cámbiala aquí en el código si alguna vez deseas modificarla
CLAVE_PRINCIPAL = "81754"
# Clave inicial para entrar al Modo Editor desde la nube o enlace compartido
CLAVE_EDITOR = "81754"

SECCIONES_CONTRASENA = [
    ("papelera", "Papelera (eliminación definitiva)"),
    ("inventario", "Inventario (eliminar registro)"),
    ("compras", "Compra y Reposición (eliminar compra)"),
    ("ventas", "Control de Ventas (eliminar venta)"),
    ("mermas", "Merma y Liquidación (eliminar registro)"),
    ("notificaciones", "Notificaciones (eliminar notificación)"),
]
OPCIONES_TIEMPO_PAPELERA = {
    "Desactivado (no borrar automáticamente)": None,
    "1 día (24 horas)": 24.0,
    "3 días (72 horas)": 72.0,
    "7 días": 168.0,
    "15 días": 360.0,
    "30 días": 720.0,
    "Personalizado (horas o días)": "custom",
}
SEGUNDOS_AVISO_RESPALDO = 4


# ==============================================================================
# CONEXIÓN A LA NUBE (GOOGLE SHEETS VÍA SECRETS DE STREAMLIT)
# ==============================================================================
def nube_configurada():
    try:
        nube_cfg = st.secrets.get("nube_sheets", {})
        return bool(
            nube_cfg.get("sheet_id")
            and nube_cfg.get("client_email")
            and nube_cfg.get("private_key")
        )
    except Exception:
        return False


def obtener_cliente_sheets():
    import gspread
    from google.oauth2.service_account import Credentials

    nube_cfg = st.secrets["nube_sheets"]
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    info_cuenta = {
        "type": "service_account",
        "project_id": nube_cfg.get("project_id", ""),
        "private_key_id": nube_cfg.get("private_key_id", ""),
        "private_key": str(nube_cfg["private_key"]).replace("\\n", "\n"),
        "client_email": nube_cfg["client_email"],
        "client_id": nube_cfg.get("client_id", ""),
        "auth_uri": "https://accounts.google.com/o/oauth2/auth",
        "token_uri": "https://oauth2.googleapis.com/token",
    }
    creds = Credentials.from_service_account_info(info_cuenta, scopes=scopes)
    cliente = gspread.authorize(creds)
    hoja = cliente.open_by_key(nube_cfg["sheet_id"]).sheet1
    return hoja


def leer_datos_nube():
    hoja = obtener_cliente_sheets()
    valores = hoja.get_all_values()
    if not valores:
        return None
    fragmentos = [fila[0] for fila in valores if fila and fila[0]]
    texto_json = "".join(fragmentos).strip()
    if not texto_json:
        return None
    return json.loads(texto_json)


def escribir_datos_nube(datos_dict):
    hoja = obtener_cliente_sheets()
    texto_json = json.dumps(datos_dict, ensure_ascii=False)
    tam_bloque = 40000
    bloques = [[texto_json[i:i + tam_bloque]] for i in range(0, len(texto_json), tam_bloque)] or [[""]]
    hoja.clear()
    hoja.update(range_name=f"A1:A{len(bloques)}", values=bloques)


def imagen_a_base64(ruta_imagen):
    if not ruta_imagen or not os.path.exists(ruta_imagen):
        return ""
    with open(ruta_imagen, "rb") as f:
        return base64.b64encode(f.read()).decode("utf-8")


def obtener_ruta_base():
    return os.path.dirname(os.path.abspath(__file__))


def obtener_carpeta_respaldos():
    ruta = os.path.join(obtener_ruta_base(), CARPETA_RESPALDOS)
    os.makedirs(ruta, exist_ok=True)
    return ruta


def obtener_carpeta_media():
    ruta = os.path.join(obtener_ruta_base(), CARPETA_MEDIA)
    os.makedirs(ruta, exist_ok=True)
    return ruta


def crear_respaldo_json():
    carpeta_resp = obtener_carpeta_respaldos()
    sello_tiempo = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    nombre_archivo = f"inventario_respaldo_{sello_tiempo}.json"
    ruta_destino = os.path.join(carpeta_resp, nombre_archivo)

    ruta_origen = os.path.join(obtener_ruta_base(), ARCHIVO_DATOS)
    if os.path.exists(ruta_origen):
        shutil.copy2(ruta_origen, ruta_destino)
    else:
        with open(ruta_destino, "w", encoding="utf-8") as f:
            json.dump(cargar_datos(), f, indent=4, ensure_ascii=False)

    return nombre_archivo, ruta_destino


def obtener_clave_editor(datos):
    val = str(datos.get("clave_editor") or "").strip()
    return val if val else CLAVE_EDITOR


def obtener_clave_seccion(datos, seccion_key):
    claves = datos.get("contrasenas_secciones") or {}
    val = str(claves.get(seccion_key, CLAVE_PRINCIPAL) or "").strip()
    return val if val else CLAVE_PRINCIPAL


def preparar_confirmacion_eliminacion(firma_key, sel_valor, conf_key, clave_key, flag_key):
    if st.session_state.get(flag_key):
        st.session_state[conf_key] = False
        st.session_state[clave_key] = ""
        st.session_state[flag_key] = False
        st.session_state[firma_key] = sel_valor
    elif st.session_state.get(firma_key) != sel_valor:
        st.session_state[firma_key] = sel_valor
        st.session_state[conf_key] = False
        st.session_state[clave_key] = ""


def solicitar_limpieza_eliminacion(flag_key):
    st.session_state[flag_key] = True


def mover_a_papelera(datos, tipo_item, item, extra_info=""):
    ahora = datetime.now()
    datos.setdefault("papelera", [])
    datos["papelera"].append({
        "id_papelera": f"pap_{ahora.strftime('%Y%m%d%H%M%S%f')}",
        "tipo": tipo_item,
        "eliminado_en": ahora.strftime("%Y-%m-%d %H:%M"),
        "eliminado_en_iso": ahora.isoformat(),
        "resumen": extra_info,
        "datos_item": item,
    })


def purgar_papelera_vencida(datos):
    cfg_pap = datos.get("papelera_auto_config") or {}
    if not cfg_pap.get("activo", False):
        return 0

    horas_limite = float(cfg_pap.get("horas_limite", 0.0) or 0.0)
    if horas_limite <= 0:
        return 0

    papelera_actual = datos.get("papelera", [])
    if not papelera_actual:
        return 0

    ahora = datetime.now()
    limite_dt = ahora - timedelta(hours=horas_limite)
    conservados = []
    eliminados = 0

    for item in papelera_actual:
        iso_str = item.get("eliminado_en_iso", "")
        dt_item = None
        if iso_str:
            try:
                dt_item = datetime.fromisoformat(str(iso_str))
            except ValueError:
                dt_item = None
        if dt_item is None and item.get("eliminado_en"):
            try:
                dt_item = datetime.strptime(str(item["eliminado_en"])[:16], "%Y-%m-%d %H:%M")
            except ValueError:
                dt_item = None

        if dt_item is not None and dt_item <= limite_dt:
            eliminados += 1
        else:
            conservados.append(item)

    if eliminados > 0:
        datos["papelera"] = conservados
        guardar_datos(datos)

    return eliminados


def normalizar_lotes_producto(prod):
    unidad = prod.get("unidad", "Unidad")
    cant_total = round(float(prod.get("cantidad", 0.0)), 2)
    fecha_venc = prod.get("fecha_vencimiento", "")
    fecha_reg = prod.get("fecha_registro", date.today().strftime("%Y-%m-%d"))
    estado_prod = prod.get("estado", "Buen estado")

    lotes = prod.get("lotes")
    if not isinstance(lotes, list) or not lotes:
        lotes = [{
            "id_lote": f"lote_ini_{prod.get('id', datetime.now().strftime('%H%M%S%f'))}",
            "cantidad": cant_total,
            "unidad": unidad,
            "fecha_vencimiento": fecha_venc,
            "fecha_ingreso": fecha_reg,
            "estado": estado_prod,
            "origen": "Registro inicial",
        }]
    else:
        for idx, l in enumerate(lotes):
            l.setdefault("id_lote", f"lote_{idx}_{datetime.now().strftime('%H%M%S%f')}")
            l["cantidad"] = round(float(l.get("cantidad", 0.0)), 2)
            l.setdefault("unidad", unidad)
            l.setdefault("fecha_vencimiento", fecha_venc)
            l.setdefault("fecha_ingreso", fecha_reg)
            l.setdefault("estado", estado_prod)
            l.setdefault("origen", "Compra / Reposición")

    prod["lotes"] = lotes
    return prod


def sincronizar_resumen_producto(prod):
    normalizar_lotes_producto(prod)
    lotes_activos = [l for l in prod["lotes"] if float(l.get("cantidad", 0.0)) > 0]
    total = round(sum(float(l.get("cantidad", 0.0)) for l in lotes_activos), 2)
    prod["cantidad"] = total

    fechas_validas = [
        l.get("fecha_vencimiento", "") for l in lotes_activos if l.get("fecha_vencimiento")
    ]
    if fechas_validas:
        prod["fecha_vencimiento"] = min(fechas_validas)

    if any(l.get("estado") == "Mal estado / Merma" for l in lotes_activos):
        prod["estado"] = "Mal estado / Merma"
    else:
        prod["estado"] = "Buen estado"


def normalizar_estructura_datos(datos):
    datos.setdefault("productos", [])
    datos.setdefault("compras", [])
    datos.setdefault("ventas", [])
    datos.setdefault("mermas", [])
    datos.setdefault("papelera", [])
    datos.setdefault("papelera_auto_config", {
        "activo": False,
        "modo_etiqueta": "Desactivado (no borrar automáticamente)",
        "horas_limite": 0.0,
    })
    if not str(datos.get("clave_editor") or "").strip():
        datos["clave_editor"] = CLAVE_EDITOR

    datos.setdefault("contrasenas_secciones", {})
    for k_sec, _ in SECCIONES_CONTRASENA:
        if not str(datos["contrasenas_secciones"].get(k_sec, "")).strip():
            datos["contrasenas_secciones"][k_sec] = CLAVE_PRINCIPAL

    datos.setdefault("categorias", list(CATEGORIAS_INICIALES))
    datos.setdefault("preferencia_tema", "oscuro")
    datos.setdefault("notificaciones_descartadas", [])
    datos.setdefault("config_app", {})

    cfg = datos["config_app"]
    cfg.setdefault("nombre_negocio", "La Casa del Pollo Rey")
    cfg.setdefault("subtitulo_negocio", "Sistema de Inventario · Compras · Ventas")
    cfg.setdefault("ruta_logo", "")
    titulos_cfg = cfg.get("titulos_pestanas")
    if not isinstance(titulos_cfg, list) or len(titulos_cfg) != 7:
        cfg["titulos_pestanas"] = list(TITULOS_PESTANAS_BASE)

    plantillas = cfg.get("plantillas")
    if not isinstance(plantillas, dict) or "Plantilla Principal" not in plantillas:
        plantillas = {"Plantilla Principal": PLANTILLA_PRINCIPAL}
        cfg["plantillas"] = plantillas
    if cfg.get("plantilla_activa") not in plantillas:
        cfg["plantilla_activa"] = "Plantilla Principal"

    if datos["preferencia_tema"] not in ("oscuro", "claro"):
        datos["preferencia_tema"] = "oscuro"

    if not datos["categorias"]:
        datos["categorias"] = ["General"]
    elif "General" not in datos["categorias"]:
        datos["categorias"].append("General")

    for idx, p in enumerate(datos["productos"]):
        if "id" not in p:
            p["id"] = f"prod_leg_{idx}_{datetime.now().strftime('%H%M%S%f')}"
        p.setdefault("unidad", "Unidad")
        p.setdefault("estado", "Buen estado")
        p.setdefault("categoria", "General")
        if p["categoria"] not in datos["categorias"]:
            datos["categorias"].append(p["categoria"])
        normalizar_lotes_producto(p)

    for idx, c in enumerate(datos["compras"]):
        if "id" not in c:
            c["id"] = f"comp_leg_{idx}_{datetime.now().strftime('%H%M%S%f')}"
        c.setdefault("unidad", "Unidad")
        c.setdefault("categoria", "General")

    for idx, v in enumerate(datos["ventas"]):
        if "id" not in v:
            v["id"] = f"vta_leg_{idx}_{datetime.now().strftime('%H%M%S%f')}"
        v.setdefault("unidad", "Unidad")
        v.setdefault("hora", "00:00")
        v.setdefault("caducidad_lote", "Sin registro")

    for idx, m in enumerate(datos["mermas"]):
        if "id" not in m:
            m["id"] = f"merma_leg_{idx}_{datetime.now().strftime('%H%M%S%f')}"
        m.setdefault("tipo_registro", "Merma / Vencido")
        m.setdefault("unidad", "Unidad")
        m.setdefault("hora", "00:00")
        m.setdefault("caducidad_lote", "Sin registro")

    return datos


def cargar_datos():
    datos_iniciales = normalizar_estructura_datos({
        "productos": [],
        "compras": [],
        "ventas": [],
        "mermas": [],
        "papelera": [],
        "papelera_auto_config": {
            "activo": False,
            "modo_etiqueta": "Desactivado (no borrar automáticamente)",
            "horas_limite": 0.0,
        },
        "clave_editor": CLAVE_EDITOR,
        "contrasenas_secciones": {k: CLAVE_PRINCIPAL for k, _ in SECCIONES_CONTRASENA},
        "categorias": list(CATEGORIAS_INICIALES),
        "preferencia_tema": "oscuro",
        "notificaciones_descartadas": [],
        "config_app": {
            "nombre_negocio": "La Casa del Pollo Rey",
            "subtitulo_negocio": "Sistema de Inventario · Compras · Ventas",
            "ruta_logo": "",
            "titulos_pestanas": list(TITULOS_PESTANAS_BASE),
            "plantilla_activa": "Plantilla Principal",
            "plantillas": {"Plantilla Principal": PLANTILLA_PRINCIPAL},
        },
    })

    if nube_configurada():
        try:
            datos_nube = leer_datos_nube()
            if isinstance(datos_nube, dict):
                return normalizar_estructura_datos(datos_nube)
            if os.path.exists(ARCHIVO_DATOS):
                with open(ARCHIVO_DATOS, "r", encoding="utf-8") as f:
                    datos_loc = normalizar_estructura_datos(json.load(f))
                escribir_datos_nube(datos_loc)
                return datos_loc
            escribir_datos_nube(datos_iniciales)
            return datos_iniciales
        except Exception as e:
            st.session_state["error_nube_msg"] = f"Aviso nube: {e}"

    if not os.path.exists(ARCHIVO_DATOS):
        guardar_datos(datos_iniciales)
        return datos_iniciales

    with open(ARCHIVO_DATOS, "r", encoding="utf-8") as f:
        datos = json.load(f)

    return normalizar_estructura_datos(datos)


def guardar_datos(datos):
    normalizar_estructura_datos(datos)
    if nube_configurada():
        try:
            escribir_datos_nube(datos)
            st.session_state["error_nube_msg"] = ""
        except Exception as e:
            st.session_state["error_nube_msg"] = f"Error al guardar en la nube: {e}"
    try:
        with open(ARCHIVO_DATOS, "w", encoding="utf-8") as f:
            json.dump(datos, f, indent=4, ensure_ascii=False)
    except Exception:
        pass


def calcular_dias_restantes(fecha_venc_str):
    if not fecha_venc_str:
        return None
    try:
        fecha_venc = datetime.strptime(str(fecha_venc_str)[:10], "%Y-%m-%d").date()
        return (fecha_venc - date.today()).days
    except ValueError:
        return None


def fmt_num(valor):
    num = round(float(valor), 2)
    if num.is_integer():
        return str(int(num))
    return f"{num:.2f}".rstrip("0").rstrip(".")


def fmt_quetzales(valor):
    return f"Q{float(valor or 0.0):.2f}"


def obtener_paleta_activa(cfg_app, modo_oscuro=True):
    plantillas = cfg_app.get("plantillas") or {"Plantilla Principal": PLANTILLA_PRINCIPAL}
    activa_nombre = cfg_app.get("plantilla_activa", "Plantilla Principal")
    plantilla = plantillas.get(activa_nombre) or plantillas.get("Plantilla Principal") or PLANTILLA_PRINCIPAL
    modo_key = "oscuro" if modo_oscuro else "claro"
    base_modo = PLANTILLA_PRINCIPAL[modo_key]
    merged = dict(base_modo)
    merged.update(plantilla.get(modo_key) or {})
    return merged


def tarjeta_metrica(etiqueta, valor, modo_oscuro=True, paleta=None):
    pal = paleta or obtener_paleta_activa({}, modo_oscuro=modo_oscuro)
    card_bg = pal["metric_bg"]
    card_bdr = pal["metric_bdr"]
    card_top = pal["metric_top"]
    lbl_col = pal["metric_lbl"]
    val_col = pal["metric_val"]

    html_card = f"""
    <div translate="no" class="notranslate" style="background:{card_bg}; border:1.5px solid {card_bdr}; border-top:3.5px solid {card_top}; border-radius:12px; padding:10px 14px; box-shadow:0 6px 16px rgba(0,0,0,0.14); margin-bottom:10px;">
        <div style="color:{lbl_col}; font-family:'Space Grotesk',sans-serif; font-weight:600; font-size:0.84rem; margin-bottom:4px;">{html.escape(str(etiqueta))}</div>
        <div translate="no" class="notranslate" style="color:{val_col}; font-family:'Montserrat',sans-serif; font-weight:700; font-size:1.35rem; line-height:1.2;">{html.escape(str(valor))}</div>
    </div>
    """
    st.markdown(html_card, unsafe_allow_html=True)


def formatear_cantidad(cantidad, unidad):
    return f"{fmt_num(cantidad)} {unidad}"


def clasificar_vencimiento_lote(lote, fallback_estado="Buen estado"):
    if lote.get("estado", fallback_estado) == "Mal estado / Merma":
        return "Vencido / Mal estado"
    dias = calcular_dias_restantes(lote.get("fecha_vencimiento", ""))
    if dias is None:
        return "Sin fecha"
    if dias < 0:
        return "Vencido"
    if dias <= DIAS_AVISO_VENCIMIENTO:
        return "Próximo a vencer"
    return "Buen estado"


def clasificar_vencimiento(prod):
    normalizar_lotes_producto(prod)
    lotes_activos = [l for l in prod["lotes"] if float(l.get("cantidad", 0.0)) > 0]
    if not lotes_activos:
        lotes_activos = prod["lotes"]

    situaciones = [clasificar_vencimiento_lote(l, prod.get("estado", "Buen estado")) for l in lotes_activos]
    if "Vencido / Mal estado" in situaciones:
        return "Vencido / Mal estado"
    if "Vencido" in situaciones:
        return "Vencido"
    if "Próximo a vencer" in situaciones:
        return "Próximo a vencer"
    if all(s == "Sin fecha" for s in situaciones):
        return "Sin fecha"
    return "Buen estado"


def clasificar_stock(cantidad):
    cant = float(cantidad or 0.0)
    if cant <= UMBRAL_STOCK_ROJO:
        return "Poco producto"
    if cant <= UMBRAL_STOCK_AMARILLO:
        return "Por surtir pronto"
    return "Stock suficiente"


def descontar_de_lotes(prod, cantidad_a_descontar):
    normalizar_lotes_producto(prod)
    unidad_p = prod.get("unidad", "Unidad")
    restante = round(float(cantidad_a_descontar), 2)
    fechas_usadas = []

    lotes_ordenados = sorted(
        prod["lotes"],
        key=lambda l: (
            l.get("fecha_vencimiento") or "9999-12-31",
            l.get("fecha_ingreso") or "",
        )
    )
    for lote in lotes_ordenados:
        if restante <= 0:
            break
        disp_lote = round(float(lote.get("cantidad", 0.0)), 2)
        if disp_lote > 0:
            tomado = min(disp_lote, restante)
            lote["cantidad"] = round(disp_lote - tomado, 2)
            restante = round(restante - tomado, 2)

            cad_lote = lote.get("fecha_vencimiento", "") or "Sin fecha"
            if len(lotes_ordenados) > 1 and tomado < cantidad_a_descontar:
                fechas_usadas.append(f"{cad_lote} ({formatear_cantidad(tomado, unidad_p)})")
            else:
                fechas_usadas.append(cad_lote)

    prod["lotes"] = [l for l in prod["lotes"] if float(l.get("cantidad", 0.0)) > 0]
    sincronizar_resumen_producto(prod)

    if fechas_usadas:
        unicas_sin_cant = [f.split(" (")[0] for f in fechas_usadas]
        if len(set(unicas_sin_cant)) == 1 and len(fechas_usadas) == 1:
            return unicas_sin_cant[0]
        return ", ".join(fechas_usadas)
    return prod.get("fecha_vencimiento", "") or "Sin fecha"


def estilo_resaltado_estado(situacion, modo_oscuro=True):
    if modo_oscuro:
        if situacion in ("Buen estado", "Stock suficiente", "Reciente", "Buen estado (otro precio / baja)"):
            return "rgba(22, 163, 74, 0.34)", "#F8FAFC"
        if situacion in ("Próximo a vencer", "Por surtir pronto"):
            return "rgba(217, 119, 6, 0.38)", "#FEF9C3"
        if situacion in ("Vencido", "Vencido / Mal estado", "Poco producto", "Merma / Vencido", "Nuevo ingreso en mal estado"):
            return "rgba(220, 38, 38, 0.38)", "#FEF2F2"
        return "rgba(56, 189, 248, 0.14)", "#F1F5F9"
    else:
        if situacion in ("Buen estado", "Stock suficiente", "Reciente", "Buen estado (otro precio / baja)"):
            return "#BBF7D0", "#111827"
        if situacion in ("Próximo a vencer", "Por surtir pronto"):
            return "#FEF08A", "#111827"
        if situacion in ("Vencido", "Vencido / Mal estado", "Poco producto", "Merma / Vencido", "Nuevo ingreso en mal estado"):
            return "#FECACA", "#111827"
        return "#F1F5F9", "#111827"


def preparar_df_enumerado(df_visible):
    df_num = df_visible.copy().reset_index(drop=True)
    df_num.insert(0, "#", [str(i + 1) for i in range(len(df_num))])
    return df_num


def renderizar_tabla_moderna(df_visible, modo_oscuro=True, estilos_celdas=None, enumerar=True):
    if df_visible.empty:
        return

    df_tabla = preparar_df_enumerado(df_visible) if enumerar else df_visible.reset_index(drop=True)

    if modo_oscuro:
        wrap_bg = "linear-gradient(180deg, #14233C 0%, #0F1A2E 100%)"
        hdr_bg = "linear-gradient(180deg, #1B2E4E 0%, #14233C 100%)"
        hdr_fg = "#E2E8F0"
        row_bg = "#132138"
        row_alt_bg = "#172742"
        cell_fg = "#F1F5F9"
        num_fg = "#93C5FD"
        bdr_color = "#38BDF8"
    else:
        wrap_bg = "#FFFFFF"
        hdr_bg = "#E2E8F0"
        hdr_fg = "#111827"
        row_bg = "#FFFFFF"
        row_alt_bg = "#F8FAFC"
        cell_fg = "#111827"
        num_fg = "#1E40AF"
        bdr_color = "#38BDF8"

    ths = [
        f"<th translate='no' class='notranslate' style='background:{hdr_bg}; color:{hdr_fg}; border:1px solid {bdr_color}; padding:6px 10px; text-align:left; font-weight:700; font-size:0.82rem; letter-spacing:0.01em; white-space:nowrap;'>{html.escape(str(col))}</th>"
        for col in df_tabla.columns
    ]

    trs = []
    for r_idx, (_, row) in enumerate(df_tabla.iterrows()):
        base_bg = row_bg if r_idx % 2 == 0 else row_alt_bg
        tds = []
        for col in df_tabla.columns:
            val = row[col]
            txt = "" if pd.isna(val) else str(val)
            bg_cell = base_bg
            fg_cell = num_fg if col == "#" else cell_fg
            fw_cell = "600" if col == "#" else "500"

            if (
                estilos_celdas is not None
                and col != "#"
                and r_idx in estilos_celdas.index
                and col in estilos_celdas.columns
            ):
                regla = str(estilos_celdas.loc[r_idx, col] or "")
                if regla:
                    for parte in regla.split(";"):
                        parte = parte.strip()
                        if parte.startswith("background-color:"):
                            bg_cell = parte.split(":", 1)[1].strip()
                        elif parte.startswith("color:"):
                            fg_cell = parte.split(":", 1)[1].strip()
                        elif parte.startswith("font-weight:"):
                            fw_cell = parte.split(":", 1)[1].strip()

            tds.append(
                f"<td translate='no' class='notranslate' style='background:{bg_cell}; color:{fg_cell}; border:1px solid {bdr_color}; padding:6px 10px; font-weight:{fw_cell}; font-size:0.84rem; line-height:1.25; white-space:nowrap;'>{html.escape(txt)}</td>"
            )
        trs.append(f"<tr>{''.join(tds)}</tr>")

    tabla_html = f"""
    <div translate="no" class="notranslate" style="width:100%; overflow-x:auto; background:{wrap_bg}; border:1.5px solid {bdr_color}; border-radius:10px; box-shadow:0 5px 14px rgba(0,0,0,0.14); margin:4px 0 10px 0;">
        <table style="width:100%; border-collapse:collapse; margin:0; font-family:'Inter',sans-serif;">
            <thead><tr>{''.join(ths)}</tr></thead>
            <tbody>{''.join(trs)}</tbody>
        </table>
    </div>
    """
    st.markdown(tabla_html, unsafe_allow_html=True)


def estilizar_celdas_inventario(df_visible, situaciones_cad, niveles_stock, modo_oscuro=True):
    estilos = pd.DataFrame("", index=df_visible.index, columns=df_visible.columns)
    if "Caducidad" in df_visible.columns:
        col_cad = df_visible.columns.get_loc("Caducidad")
        for i, sit in enumerate(situaciones_cad):
            bg, fg = estilo_resaltado_estado(sit, modo_oscuro=modo_oscuro)
            estilos.iloc[i, col_cad] = f"background-color: {bg}; color: {fg}; font-weight: 600;"
    if "Stock" in df_visible.columns:
        col_stk = df_visible.columns.get_loc("Stock")
        for i, niv in enumerate(niveles_stock):
            bg, fg = estilo_resaltado_estado(niv, modo_oscuro=modo_oscuro)
            estilos.iloc[i, col_stk] = f"background-color: {bg}; color: {fg}; font-weight: 600;"
    return estilos


def estilizar_celdas_estado(df_visible, situaciones_cad, niveles_stock, modo_oscuro=True):
    estilos = pd.DataFrame("", index=df_visible.index, columns=df_visible.columns)
    if "Estado" in df_visible.columns:
        col_est = df_visible.columns.get_loc("Estado")
        for i, sit in enumerate(situaciones_cad):
            bg, fg = estilo_resaltado_estado(sit, modo_oscuro=modo_oscuro)
            estilos.iloc[i, col_est] = f"background-color: {bg}; color: {fg}; font-weight: 600;"
    if "Nivel Stock" in df_visible.columns:
        col_ns = df_visible.columns.get_loc("Nivel Stock")
        for i, niv in enumerate(niveles_stock):
            bg, fg = estilo_resaltado_estado(niv, modo_oscuro=modo_oscuro)
            estilos.iloc[i, col_ns] = f"background-color: {bg}; color: {fg}; font-weight: 600;"
    return estilos


def estilizar_tabla_reposicion(df_visible, niveles_stock, modo_oscuro=True):
    estilos = pd.DataFrame("", index=df_visible.index, columns=df_visible.columns)
    for i, niv in enumerate(niveles_stock):
        bg, fg = estilo_resaltado_estado(niv, modo_oscuro=modo_oscuro)
        estilos.iloc[i, :] = f"background-color: {bg}; color: {fg}; font-weight: 600;"
    return estilos


def estilizar_notificaciones(df_visible, prioridades, modo_oscuro=True):
    estilos = pd.DataFrame("", index=df_visible.index, columns=df_visible.columns)
    for i, prio in enumerate(prioridades):
        bg, fg = estilo_resaltado_estado(prio, modo_oscuro=modo_oscuro)
        estilos.iloc[i, :] = f"background-color: {bg}; color: {fg}; font-weight: 600;"
    if "Tipo de alerta" in df_visible.columns:
        col_alerta = df_visible.columns.get_loc("Tipo de alerta")
        for i, prio in enumerate(prioridades):
            bg, fg = estilo_resaltado_estado(prio, modo_oscuro=modo_oscuro)
            estilos.iloc[i, col_alerta] = f"background-color: {bg}; color: {fg}; font-weight: 700;"
    return estilos


def estilizar_mermas(df_visible, tipos_reg, modo_oscuro=True):
    estilos = pd.DataFrame("", index=df_visible.index, columns=df_visible.columns)
    if "Tipo" in df_visible.columns:
        col_tipo = df_visible.columns.get_loc("Tipo")
        for i, t in enumerate(tipos_reg):
            bg, fg = estilo_resaltado_estado(t, modo_oscuro=modo_oscuro)
            estilos.iloc[i, col_tipo] = f"background-color: {bg}; color: {fg}; font-weight: 700;"
    return estilos


def aplicar_css_tema(modo_oscuro: bool, paleta: dict):
    bg_app = paleta["bg_app"]
    txt_main = paleta["txt_main"]
    h2_color = paleta["h2_color"]
    h3_color = paleta["h3_color"]
    strong_color = paleta["strong_color"]
    cap_color = paleta["cap_color"]
    metric_bg = paleta["metric_bg"]
    metric_bdr = paleta["metric_bdr"]
    metric_top = paleta["metric_top"]
    metric_lbl = paleta["metric_lbl"]
    metric_val = paleta["metric_val"]
    inp_bg = paleta["inp_bg"]
    inp_bdr = paleta["inp_bdr"]
    form_bg = paleta["form_bg"]
    form_bdr = paleta["form_bdr"]
    btn_bg = paleta["btn_bg"]
    btn_bdr = paleta["btn_bdr"]

    if modo_oscuro:
        chip_bg = "linear-gradient(135deg, rgba(30, 58, 138, 0.52), rgba(37, 99, 235, 0.28))"
        chip_txt = "#FDE68A"
        chip_bdr = "rgba(147, 197, 253, 0.55)"
        hueco_bg = "#142138"
        step_bg = "linear-gradient(180deg, #254E8A 0%, #183663 100%)"
        step_fg = "#F8FAFC"
        step_bdr = "#38BDF8"
        vacio_bg = "linear-gradient(160deg, #1A2B47 0%, #132036 100%)"
        vacio_bdr = "#38BDF8"
        vacio_txt = "#F8FAFC"
        vacio_sub = "#E2E8F0"
        sec_pap_bg = "linear-gradient(180deg, #162540 0%, #111D33 100%)"
        sec_pap_bdr = "rgba(56, 189, 248, 0.45)"
    else:
        chip_bg = "#FEF08A"
        chip_txt = "#111827"
        chip_bdr = "#111827"
        hueco_bg = "#FFFFFF"
        step_bg = "linear-gradient(180deg, #E0F2FE 0%, #BAE6FD 100%)"
        step_fg = "#111827"
        step_bdr = "#38BDF8"
        vacio_bg = "#FFFFFF"
        vacio_bdr = "#38BDF8"
        vacio_txt = "#072A6E"
        vacio_sub = "#111827"
        sec_pap_bg = "#F8FAFC"
        sec_pap_bdr = "#38BDF8"

    css = f"""
    <style>
    div[data-testid="InputInstructions"],
    div[data-testid="stNumberInput"] [data-testid="InputInstructions"],
    div[data-testid="stTextInput"] [data-testid="InputInstructions"],
    div[data-testid="stNumberInput"] small,
    div[data-testid="stTextInput"] small {{
        display: none !important;
        visibility: hidden !important;
        height: 0 !important;
        width: 0 !important;
        overflow: hidden !important;
    }}

    .stApp,
    div[data-testid="stAppViewContainer"],
    section.main,
    div.block-container,
    section[data-testid="stSidebar"] > div {{
        background: {bg_app} !important;
        color: {txt_main};
        font-family: 'Inter', sans-serif;
    }}
    h1 {{
        font-family: 'Montserrat', sans-serif !important;
        font-size: 1.55rem !important;
        font-weight: 700 !important;
        color: {metric_val};
        letter-spacing: -0.015em;
        margin-bottom: 0.25rem !important;
    }}
    h2 {{
        font-family: 'Outfit', sans-serif !important;
        font-size: 1.22rem !important;
        font-weight: 600 !important;
        color: {h2_color};
        letter-spacing: -0.01em;
    }}
    h3 {{
        font-family: 'Outfit', sans-serif !important;
        font-size: 1.06rem !important;
        font-weight: 600 !important;
        color: {h3_color};
        letter-spacing: -0.01em;
    }}
    h4, p, span, label, .stMarkdown {{
        font-family: 'Inter', sans-serif;
        color: {txt_main};
    }}
    .stMarkdown strong {{
        font-family: 'Outfit', sans-serif !important;
        font-size: 1.0rem !important;
        font-weight: 600 !important;
        color: {strong_color};
    }}
    .stCaption {{
        color: {cap_color} !important;
        font-size: 0.88rem !important;
        font-family: 'Inter', sans-serif !important;
    }}

    .header-simetrico {{
        display: flex;
        align-items: center;
        gap: 16px;
        min-height: 128px;
        width: 100%;
    }}
    .logo-marco {{
        height: 128px;
        width: 128px;
        min-width: 128px;
        padding: 6px;
        border-radius: 14px;
        border: 2.5px solid #F4B400;
        box-shadow: 0 8px 20px rgba(0, 0, 0, 0.28);
        display: flex;
        align-items: center;
        justify-content: center;
        background-color: #FFFFFF;
        box-sizing: border-box;
    }}
    .logo-marco img {{
        max-height: 100%;
        max-width: 100%;
        width: auto;
        height: auto;
        object-fit: contain;
        image-rendering: auto;
        display: block;
    }}
    .pollo-rey-banner {{
        background: {paleta["banner_bg"]};
        border-radius: 16px;
        padding: 12px 22px;
        border: 2.5px solid #F4B400;
        box-shadow: 0 8px 20px rgba(0, 0, 0, 0.28), inset 0 1px 0 rgba(255, 255, 255, 0.22);
        min-height: 128px;
        flex: 1;
        box-sizing: border-box;
        display: flex;
        flex-direction: column;
        justify-content: center;
    }}
    .pollo-rey-titulo {{
        font-family: 'Montserrat', sans-serif !important;
        font-size: 2.15rem !important;
        font-weight: 800 !important;
        color: #FFFFFF !important;
        margin: 0;
        line-height: 1.08;
        letter-spacing: -0.02em;
        text-shadow: 0 2px 6px rgba(0, 0, 0, 0.30);
    }}
    .pollo-rey-sub {{
        display: inline-block;
        width: fit-content;
        margin-top: 6px;
        background-color: #FFD633;
        color: #0B3A8D !important;
        font-family: 'Space Grotesk', sans-serif !important;
        font-weight: 700;
        font-size: 0.74rem;
        padding: 2px 10px;
        border-radius: 999px;
        border-left: 4px solid #D62828;
    }}

    div.st-key-btn_toggle_sidebar button,
    div.st-key-btn_badge_notif button,
    div.st-key-btn_respaldo_top button,
    div.st-key-btn_alternar_tema button,
    div.st-key-btn_actualizar_nube button,
    div.st-key-btn_modo_rol button {{
        min-height: 38px !important;
        border-radius: 12px !important;
        padding: 4px 10px !important;
    }}

    .cuadro-vacio {{
        margin-top: 12px;
        padding: 18px 22px;
        border-radius: 10px;
        background: {vacio_bg};
        border: 1.5px solid {vacio_bdr};
        box-shadow: 0 6px 16px rgba(0, 0, 0, 0.14);
    }}
    .cuadro-vacio-titulo {{
        font-family: 'Outfit', sans-serif !important;
        font-size: 1.05rem !important;
        font-weight: 600 !important;
        color: {vacio_txt} !important;
        margin: 0 0 6px 0;
    }}
    .cuadro-vacio-desc {{
        font-family: 'Inter', sans-serif !important;
        font-size: 0.9rem !important;
        font-weight: 400 !important;
        color: {vacio_sub} !important;
        margin: 0;
    }}

    .papelera-seccion-hdr {{
        background: {sec_pap_bg};
        border: 1.5px solid {sec_pap_bdr};
        border-left: 4px solid #F4B400;
        border-radius: 10px;
        padding: 8px 14px;
        margin-top: 14px;
        margin-bottom: 8px;
        font-family: 'Outfit', sans-serif;
        font-weight: 600;
        font-size: 1.0rem;
        color: {h2_color};
    }}

    .cat-chip {{
        display: inline-block;
        font-family: 'Outfit', sans-serif !important;
        font-weight: 600;
        font-size: 0.86rem;
        padding: 4px 12px;
        margin: 2px 5px 4px 0;
        border-radius: 999px;
        background: {chip_bg};
        color: {chip_txt} !important;
        border: 1.5px solid {chip_bdr};
    }}

    div[data-testid="stMetric"],
    div[data-testid="stMetric"] div[data-testid="stMetricValue"] {{
        translate: none !important;
    }}
    div[data-testid="stMetric"] {{
        background: {metric_bg};
        border: 1.5px solid {metric_bdr};
        border-top: 3.5px solid {metric_top};
        border-radius: 12px;
        padding: 10px 14px;
        box-shadow: 0 6px 16px rgba(0, 0, 0, 0.14);
    }}
    div[data-testid="stMetric"] label {{
        color: {metric_lbl} !important;
        font-family: 'Space Grotesk', sans-serif !important;
        font-weight: 600 !important;
        font-size: 0.84rem !important;
    }}
    div[data-testid="stMetric"] div[data-testid="stMetricValue"] {{
        color: {metric_val} !important;
        font-family: 'Montserrat', sans-serif !important;
        font-weight: 700 !important;
        font-size: 1.35rem !important;
    }}

    div[data-testid="stTextInput"],
    div[data-testid="stTextInput"] > div,
    div[data-testid="stDateInput"],
    div[data-testid="stDateInput"] > div,
    div[data-testid="stSelectbox"],
    div[data-testid="stSelectbox"] > div,
    div[data-testid="stMultiSelect"],
    div[data-testid="stMultiSelect"] > div,
    .stButton,
    .stButton > div,
    .stFormSubmitButton,
    .stFormSubmitButton > div {{
        background-color: transparent !important;
        background: transparent !important;
        border: none !important;
        box-shadow: none !important;
        outline: none !important;
    }}

    div[data-testid="stTextInput"] input,
    div[data-testid="stDateInput"] input,
    div[data-testid="stSelectbox"] > div > div,
    div[data-testid="stMultiSelect"] > div > div {{
        background-color: {inp_bg} !important;
        color: {txt_main} !important;
        -webkit-text-fill-color: {txt_main} !important;
        border: 1.5px solid {inp_bdr} !important;
        border-radius: 10px !important;
        font-family: 'Inter', sans-serif !important;
        font-size: 0.94rem !important;
        font-weight: 500 !important;
        min-height: 42px !important;
    }}

    div[data-testid="stNumberInput"],
    div[data-testid="stNumberInput"] > div,
    div[data-testid="stNumberInput"] > div > div,
    div[data-testid="stNumberInput"] [data-baseweb="input"],
    div[data-testid="stNumberInput"] [data-baseweb="base-input"] {{
        background-color: {hueco_bg} !important;
        background: {hueco_bg} !important;
        border: none !important;
        border-color: transparent !important;
        outline: none !important;
        box-shadow: none !important;
        overflow: visible !important;
        gap: 6px !important;
        align-items: center !important;
    }}

    div[data-testid="stNumberInput"] input {{
        background-color: {inp_bg} !important;
        color: {txt_main} !important;
        -webkit-text-fill-color: {txt_main} !important;
        border: 1.5px solid {inp_bdr} !important;
        border-radius: 10px !important;
        font-family: 'Inter', sans-serif !important;
        font-size: 0.94rem !important;
        font-weight: 500 !important;
        min-height: 42px !important;
        padding: 6px 10px !important;
        box-shadow: none !important;
        outline: none !important;
    }}

    div[data-testid="stNumberInput"] [data-baseweb="input"] > div:last-child,
    div[data-testid="stNumberInput"] div:has(> button) {{
        background-color: {hueco_bg} !important;
        background: {hueco_bg} !important;
        border: none !important;
        border-color: transparent !important;
        outline: none !important;
        box-shadow: none !important;
        gap: 6px !important;
        padding: 0 !important;
        align-items: center !important;
    }}

    div[data-testid="stNumberInput"] button {{
        background: {step_bg} !important;
        color: {step_fg} !important;
        border: 1.5px solid {step_bdr} !important;
        border-radius: 9px !important;
        box-shadow: none !important;
        outline: none !important;
        font-weight: 600 !important;
        min-height: 36px !important;
        min-width: 34px !important;
    }}
    div[data-testid="stNumberInput"] button svg {{
        color: {step_fg} !important;
        fill: {step_fg} !important;
    }}

    div[data-testid="stForm"],
    div[data-testid="stForm"] > div,
    div[data-testid="stForm"] div[data-testid="stHorizontalBlock"],
    div[data-testid="stForm"] div[data-testid="stColumn"],
    div[data-testid="stForm"] div[data-testid="stVerticalBlock"] {{
        background: {form_bg} !important;
        background-color: {form_bg} !important;
    }}
    div[data-testid="stForm"] {{
        border: 1.5px solid {form_bdr} !important;
        border-radius: 16px !important;
        padding: 18px !important;
        box-shadow: 0 8px 20px rgba(0, 0, 0, 0.16) !important;
    }}

    .stButton > button[kind="primary"] {{
        background: linear-gradient(180deg, #E13232 0%, #C41E1E 100%) !important;
        border: 1.5px solid #FCA5A5 !important;
        color: #FFFFFF !important;
        font-family: 'Space Grotesk', sans-serif !important;
        font-weight: 700 !important;
        border-radius: 12px !important;
        padding: 0.45rem 1.05rem !important;
        box-shadow: 0 4px 0 #7F1212, 0 6px 12px rgba(0, 0, 0, 0.22) !important;
        transition: all 0.12s ease !important;
    }}
    .stButton > button[kind="primary"] p,
    .stButton > button[kind="primary"] span {{
        color: #FFFFFF !important;
    }}
    .stButton > button[kind="primary"]:hover {{
        transform: translateY(-1px);
    }}
    .stButton > button[kind="primary"]:active {{
        transform: translateY(3px);
        box-shadow: 0 2px 0 #7F1212, 0 4px 8px rgba(0, 0, 0, 0.2) !important;
    }}

    .stButton > button[kind="secondary"],
    .stButton > button,
    .stFormSubmitButton > button {{
        background: {btn_bg} !important;
        border: 1.5px solid {btn_bdr} !important;
        color: #FFFFFF !important;
        font-family: 'Space Grotesk', sans-serif !important;
        font-weight: 700 !important;
        border-radius: 12px !important;
        padding: 0.45rem 1.05rem !important;
        box-shadow: 0 4px 0 #092254, 0 6px 12px rgba(0, 0, 0, 0.22) !important;
        transition: all 0.12s ease !important;
    }}
    .stButton > button[kind="secondary"] p,
    .stButton > button[kind="secondary"] span,
    .stButton > button p,
    .stButton > button span,
    .stFormSubmitButton > button p,
    .stFormSubmitButton > button span {{
        color: #FFFFFF !important;
    }}
    .stButton > button[kind="secondary"]:hover,
    .stButton > button:hover,
    .stFormSubmitButton > button:hover {{
        transform: translateY(-1px);
    }}
    .stButton > button[kind="secondary"]:active,
    .stButton > button:active,
    .stFormSubmitButton > button:active {{
        transform: translateY(3px);
        box-shadow: 0 2px 0 #092254, 0 4px 8px rgba(0, 0, 0, 0.2) !important;
    }}
    </style>
    <script>
    (function() {{
        const doc = window.parent ? window.parent.document : document;
        if (doc.__autoSelectNumBound) return;
        doc.__autoSelectNumBound = true;
        doc.addEventListener("focusin", function(e) {{
            const el = e.target;
            if (el && el.tagName === "INPUT" && el.closest('[data-testid="stNumberInput"]')) {{
                setTimeout(() => {{ try {{ el.select(); }} catch (_) {{}} }}, 0);
            }}
        }}, true);
        doc.addEventListener("click", function(e) {{
            const el = e.target;
            if (el && el.tagName === "INPUT" && el.closest('[data-testid="stNumberInput"]')) {{
                setTimeout(() => {{ try {{ el.select(); }} catch (_) {{}} }}, 0);
            }}
        }}, true);
    }})();
    </script>
    """
    st.markdown(css, unsafe_allow_html=True)


def construir_notificaciones_urgentes(productos, descartadas=None):
    descartadas_set = set(descartadas or [])
    alertas = []
    for p in productos:
        normalizar_lotes_producto(p)
        prod_id = p.get("id", "")
        unidad_p = p.get("unidad", "Unidad")
        cant_num = float(p.get("cantidad", 0.0))
        nivel_stk = clasificar_stock(cant_num)
        cad_str = p.get("fecha_vencimiento", "") or "Sin fecha"
        dias_rest = calcular_dias_restantes(p.get("fecha_vencimiento", ""))

        lotes_activos = [l for l in p["lotes"] if float(l.get("cantidad", 0.0)) > 0]
        for lote in lotes_activos:
            sit_lote = clasificar_vencimiento_lote(lote, p.get("estado", "Buen estado"))
            cad_lote = lote.get("fecha_vencimiento", "") or "Sin fecha"
            dias_lote = calcular_dias_restantes(lote.get("fecha_vencimiento", ""))
            cant_lote = float(lote.get("cantidad", 0.0))

            if sit_lote in ("Vencido", "Vencido / Mal estado", "Próximo a vencer"):
                clave_cad = f"{prod_id}|cad|{lote.get('id_lote', '')}|{sit_lote}|{cad_lote}"
                if clave_cad not in descartadas_set:
                    if sit_lote == "Próximo a vencer":
                        detalle = (
                            f"Lote vence en {dias_lote} día(s) ({cad_lote}). "
                            f"En este lote hay {formatear_cantidad(cant_lote, unidad_p)} "
                            f"(total producto: {formatear_cantidad(cant_num, unidad_p)})."
                        )
                        prioridad_estilo = "Próximo a vencer"
                    else:
                        detalle = (
                            f"Lote en estado '{sit_lote}' ({cad_lote}). "
                            f"En este lote hay {formatear_cantidad(cant_lote, unidad_p)} "
                            f"(total producto: {formatear_cantidad(cant_num, unidad_p)})."
                        )
                        prioridad_estilo = "Vencido"

                    alertas.append({
                        "clave_notif": clave_cad,
                        "Tipo de alerta": sit_lote,
                        "Producto": p["nombre"],
                        "Categoría": p.get("categoria", "General"),
                        "Cantidad actual": formatear_cantidad(cant_lote, unidad_p),
                        "Caducidad": cad_lote,
                        "Días Rest.": dias_lote,
                        "Detalle": detalle,
                        "_estilo": prioridad_estilo,
                        "_orden": 1 if prioridad_estilo == "Vencido" else 2,
                    })

        if nivel_stk in ("Poco producto", "Por surtir pronto"):
            clave_stk = f"{prod_id}|stock|{nivel_stk}|{fmt_num(cant_num)}"
            if clave_stk not in descartadas_set:
                if nivel_stk == "Poco producto":
                    detalle_stk = f"Quedan {formatear_cantidad(cant_num, unidad_p)} en total (4 o menos). Conviene reponer pronto."
                    estilo_stk = "Poco producto"
                    orden_stk = 1
                else:
                    detalle_stk = f"Quedan {formatear_cantidad(cant_num, unidad_p)} en total (5 unidades). Ir planeando reposición."
                    estilo_stk = "Por surtir pronto"
                    orden_stk = 2

                alertas.append({
                    "clave_notif": clave_stk,
                    "Tipo de alerta": nivel_stk,
                    "Producto": p["nombre"],
                    "Categoría": p.get("categoria", "General"),
                    "Cantidad actual": formatear_cantidad(cant_num, unidad_p),
                    "Caducidad": cad_str,
                    "Días Rest.": dias_rest,
                    "Detalle": detalle_stk,
                    "_estilo": estilo_stk,
                    "_orden": orden_stk,
                })

    alertas.sort(key=lambda x: (x["_orden"], x["Producto"].lower()))
    return alertas


# ==============================================================================
# INICIALIZACIÓN DE SESIÓN, ROLES (LECTOR / EDITOR) Y DATOS
# ==============================================================================
if "rol_acceso" not in st.session_state:
    st.session_state["rol_acceso"] = "lector"
if "mostrar_login_editor" not in st.session_state:
    st.session_state["mostrar_login_editor"] = False
if "tema_app" not in st.session_state:
    st.session_state["tema_app"] = "oscuro"
if "pestana_activa" not in st.session_state:
    st.session_state["pestana_activa"] = 0
if "barra_lateral_abierta" not in st.session_state:
    st.session_state["barra_lateral_abierta"] = False
if "claves_desbloqueadas" not in st.session_state:
    st.session_state["claves_desbloqueadas"] = False

datos = cargar_datos()
clave_editor_activa = obtener_clave_editor(datos)
es_editor = st.session_state.get("rol_acceso") == "editor"

purgados_inicio = purgar_papelera_vencida(datos) if es_editor else 0
cfg_app = datos["config_app"]

if "tema_inicializado" not in st.session_state:
    st.session_state["tema_app"] = datos.get("preferencia_tema", "oscuro")
    st.session_state["tema_inicializado"] = True

num_papelera = len(datos.get("papelera", []))
etiqueta_papelera_menu = f"🗑️ Papelera ({num_papelera})"
etiqueta_claves_menu = "🔐 Contraseñas"

if es_editor and st.session_state.get("forzar_papelera_lateral"):
    st.session_state["barra_lateral_abierta"] = True
    st.session_state["radio_vista_lateral"] = etiqueta_papelera_menu
    st.session_state["forzar_papelera_lateral"] = False

estado_barra = "expanded" if (es_editor and st.session_state["barra_lateral_abierta"]) else "collapsed"
st.set_page_config(
    page_title=cfg_app.get("nombre_negocio", "La Casa del Pollo Rey"),
    layout="wide",
    initial_sidebar_state=estado_barra,
)

modo_oscuro = st.session_state["tema_app"] == "oscuro"
paleta_activa = obtener_paleta_activa(cfg_app, modo_oscuro=modo_oscuro)
aplicar_css_tema(modo_oscuro=modo_oscuro, paleta=paleta_activa)

categorias_disponibles = list(datos["categorias"])
alertas_activas = construir_notificaciones_urgentes(
    datos["productos"],
    descartadas=datos.get("notificaciones_descartadas", []),
)
num_alertas = len(alertas_activas)

if es_editor and st.session_state["barra_lateral_abierta"]:
    with st.sidebar:
        st.markdown("### ☰ Menú lateral")
        if st.button("Cerrar barra lateral", key="btn_cerrar_sidebar", use_container_width=True):
            st.session_state["barra_lateral_abierta"] = False
            st.session_state["claves_desbloqueadas"] = False
            st.rerun()

        st.divider()
        opciones_radio_lat = [
            "Sistema principal",
            "⚙️ Configuración",
            etiqueta_papelera_menu,
            etiqueta_claves_menu,
        ]
        if st.session_state.get("radio_vista_lateral") not in opciones_radio_lat:
            st.session_state["radio_vista_lateral"] = "Sistema principal"
        vista_lateral = st.radio(
            "Ir a:",
            opciones_radio_lat,
            key="radio_vista_lateral",
        )

        st.divider()
        st.caption(f"Plantilla activa: **{cfg_app.get('plantilla_activa', 'Plantilla Principal')}**")
        st.caption(f"Elementos en papelera: **{num_papelera}**")
        cfg_pap_side = datos.get("papelera_auto_config") or {}
        estado_auto_side = cfg_pap_side.get("modo_etiqueta", "Desactivado") if cfg_pap_side.get("activo") else "Desactivado"
        st.caption(f"Autoborrado papelera: **{estado_auto_side}**")
else:
    vista_lateral = "Sistema principal"

if not es_editor or not st.session_state["barra_lateral_abierta"] or vista_lateral != etiqueta_claves_menu:
    st.session_state["claves_desbloqueadas"] = False

# Barra superior: cambia según si está en Modo Editor o Modo Lector
if es_editor:
    col_menu_izq, col_rol_top, col_badge_top, col_resp_top, col_tema_top = st.columns([1.6, 1.6, 2.2, 1.5, 1.4], gap="small")
    with col_menu_izq:
        etiqueta_menu_izq = "✕ Cerrar barra lateral" if st.session_state["barra_lateral_abierta"] else "☰ Abrir barra lateral"
        if st.button(etiqueta_menu_izq, key="btn_toggle_sidebar", use_container_width=True):
            st.session_state["barra_lateral_abierta"] = not st.session_state["barra_lateral_abierta"]
            st.session_state["claves_desbloqueadas"] = False
            st.rerun()
    with col_rol_top:
        if st.button("👁️ Pasar a Lector", key="btn_modo_rol", use_container_width=True):
            st.session_state["rol_acceso"] = "lector"
            st.session_state["barra_lateral_abierta"] = False
            st.session_state["claves_desbloqueadas"] = False
            st.session_state["mostrar_login_editor"] = False
            st.rerun()
    with col_badge_top:
        tipo_badge = "primary" if num_alertas > 0 else "secondary"
        etiqueta_badge = f"🔔 Pendientes  ● {num_alertas}" if num_alertas > 0 else "🔔 Pendientes  ○ 0"
        if st.button(etiqueta_badge, key="btn_badge_notif", type=tipo_badge, use_container_width=True):
            st.session_state["pestana_activa"] = 6
            st.rerun()
    with col_resp_top:
        if st.button("💾 Respaldo", key="btn_respaldo_top", use_container_width=True):
            nombre_resp, ruta_resp = crear_respaldo_json()
            st.session_state["ultimo_respaldo_msg"] = f"Respaldo guardado: {nombre_resp} (en {ruta_resp})"
            st.session_state["ultimo_respaldo_ts"] = time.time()
    with col_tema_top:
        texto_boton_tema = "☀️ Claro" if modo_oscuro else "🌙 Oscuro"
        if st.button(texto_boton_tema, key="btn_alternar_tema", use_container_width=True):
            nuevo_tema = "claro" if modo_oscuro else "oscuro"
            st.session_state["tema_app"] = nuevo_tema
            datos["preferencia_tema"] = nuevo_tema
            guardar_datos(datos)
            st.rerun()
else:
    col_lec_info, col_act_lec, col_entrar_ed, col_tema_lec = st.columns([4.5, 1.6, 1.9, 1.3], gap="small")
    with col_lec_info:
        estado_nube_txt = "Nube activa" if nube_configurada() else "Modo local"
        st.caption(f"👁️ **Modo Lector (Solo vista)** · {estado_nube_txt}")
    with col_act_lec:
        if st.button("🔄 Actualizar datos", key="btn_actualizar_nube", use_container_width=True):
            st.rerun()
    with col_entrar_ed:
        if st.button("✏️ Acceso Editor", key="btn_modo_rol", use_container_width=True):
            st.session_state["mostrar_login_editor"] = not st.session_state.get("mostrar_login_editor", False)
            st.rerun()
    with col_tema_lec:
        texto_boton_tema = "☀️ Claro" if modo_oscuro else "🌙 Oscuro"
        if st.button(texto_boton_tema, key="btn_alternar_tema", use_container_width=True):
            st.session_state["tema_app"] = "claro" if modo_oscuro else "oscuro"
            st.rerun()

    if st.session_state.get("mostrar_login_editor", False):
        with st.form("form_login_editor"):
            clave_ed_in = st.text_input("Contraseña de Editor *", type="password", key="input_clave_editor_acceso")
            c_ed1, c_ed2 = st.columns([1.2, 1.2])
            confirmar_ed = c_ed1.form_submit_button("Entrar como Editor")
            cancelar_ed = c_ed2.form_submit_button("Cancelar")
            if confirmar_ed:
                if clave_ed_in == clave_editor_activa:
                    st.session_state["rol_acceso"] = "editor"
                    st.session_state["mostrar_login_editor"] = False
                    st.rerun()
                else:
                    st.error("Contraseña de Editor incorrecta.")
            elif cancelar_ed:
                st.session_state["mostrar_login_editor"] = False
                st.rerun()

if st.session_state.get("error_nube_msg"):
    st.warning(st.session_state["error_nube_msg"])

if purgados_inicio > 0:
    st.info(f"Autolimpieza de papelera: se eliminaron automáticamente {purgados_inicio} elemento(s) que cumplieron el tiempo configurado.")

msg_resp = st.session_state.get("ultimo_respaldo_msg")
ts_resp = st.session_state.get("ultimo_respaldo_ts", 0.0)
if msg_resp:
    transcurrido = time.time() - float(ts_resp)
    if transcurrido < SEGUNDOS_AVISO_RESPALDO:
        st.success(msg_resp)
        time.sleep(min(1.0, SEGUNDOS_AVISO_RESPALDO - transcurrido))
        st.rerun()
    else:
        st.session_state["ultimo_respaldo_msg"] = ""
        st.session_state["ultimo_respaldo_ts"] = 0.0

carpeta_actual = obtener_ruta_base()
ruta_logo_cfg = cfg_app.get("ruta_logo", "")
ruta_logo = ruta_logo_cfg if (ruta_logo_cfg and os.path.exists(ruta_logo_cfg)) else os.path.join(carpeta_actual, "logo.jpeg")
logo_b64 = imagen_a_base64(ruta_logo)

nombre_negocio_actual = cfg_app.get("nombre_negocio", "La Casa del Pollo Rey")
subtitulo_negocio_actual = cfg_app.get("subtitulo_negocio", "Sistema de Inventario · Compras · Ventas")

logo_html = (
    f'<div class="logo-marco"><img src="data:image/jpeg;base64,{logo_b64}" alt="{html.escape(nombre_negocio_actual)}"></div>'
    if logo_b64 else ""
)
st.markdown(
    f"""
    <div class="header-simetrico">
        {logo_html}
        <div class="pollo-rey-banner">
            <p class="pollo-rey-titulo">{html.escape(nombre_negocio_actual)}</p>
            <span class="pollo-rey-sub">{html.escape(subtitulo_negocio_actual)}</span>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

st.markdown("<div style='margin-bottom: 8px;'></div>", unsafe_allow_html=True)

# ==============================================================================
# PANTALLA EXCLUSIVA DE MODO LECTOR (PRODUCTOS, VENTAS/BALANCE, MERMA Y ALERTAS)
# ==============================================================================
if not es_editor:
    st.subheader("📋 Estado actual de los productos")
    st.caption("Vista de solo lectura. Si se acaban de ingresar o modificar productos, haz clic en **Actualizar datos**.")

    total_prods_lec = len(datos["productos"])
    stock_bajo_lec = sum(1 for p in datos["productos"] if clasificar_stock(float(p.get("cantidad", 0.0))) in ("Poco producto", "Por surtir pronto"))
    urgentes_lec = sum(1 for a in alertas_activas if a.get("_orden") == 1)

    rm1, rm2, rm3 = st.columns(3)
    with rm1:
        tarjeta_metrica("Productos registrados", str(total_prods_lec), modo_oscuro=modo_oscuro, paleta=paleta_activa)
    with rm2:
        tarjeta_metrica("Por reponer / stock bajo", str(stock_bajo_lec), modo_oscuro=modo_oscuro, paleta=paleta_activa)
    with rm3:
        tarjeta_metrica("Alertas urgentes", str(urgentes_lec), modo_oscuro=modo_oscuro, paleta=paleta_activa)

    if datos["productos"]:
        filas_lec = []
        for p in datos["productos"]:
            normalizar_lotes_producto(p)
            unidad_p = p.get("unidad", "Unidad")
            cant_num = float(p.get("cantidad", 0.0))
            dias_rest = calcular_dias_restantes(p.get("fecha_vencimiento", ""))
            situacion = clasificar_vencimiento(p)
            nivel_stk = clasificar_stock(cant_num)
            filas_lec.append({
                "Producto": p["nombre"],
                "Categoría": p.get("categoria", "General"),
                "Disponible": formatear_cantidad(cant_num, unidad_p),
                "Nivel Stock": nivel_stk,
                "Caducidad": p.get("fecha_vencimiento", "") or "Sin fecha",
                "Días Rest.": dias_rest if dias_rest is not None else "-",
                "Estado": situacion,
                "_situacion": situacion,
                "_nivel_stock": nivel_stk,
            })

        df_lec = pd.DataFrame(filas_lec)

        lf1, lf2, lf3 = st.columns(3)
        sit_lec_sel = lf1.multiselect(
            "Filtrar por estado / caducidad",
            options=["Buen estado", "Próximo a vencer", "Vencido", "Vencido / Mal estado", "Sin fecha"],
            default=[],
            placeholder="Todos los estados",
            key="filtro_estado_lector",
        )
        stk_lec_sel = lf2.multiselect(
            "Filtrar por nivel de stock",
            options=["Stock suficiente", "Por surtir pronto", "Poco producto"],
            default=[],
            placeholder="Todos los niveles de stock",
            key="filtro_stock_lector",
        )
        cats_lec_lista = sorted(df_lec["Categoría"].unique().tolist())
        cat_lec_sel = lf3.multiselect(
            "Filtrar por categoría",
            options=cats_lec_lista,
            default=[],
            placeholder="Todas las categorías",
            key="filtro_cat_lector",
        )

        m_sit_lec = df_lec["_situacion"].isin(sit_lec_sel) if sit_lec_sel else pd.Series([True] * len(df_lec), index=df_lec.index)
        m_stk_lec = df_lec["_nivel_stock"].isin(stk_lec_sel) if stk_lec_sel else pd.Series([True] * len(df_lec), index=df_lec.index)
        m_cat_lec = df_lec["Categoría"].isin(cat_lec_sel) if cat_lec_sel else pd.Series([True] * len(df_lec), index=df_lec.index)

        df_lec_fil = df_lec[m_sit_lec & m_stk_lec & m_cat_lec].reset_index(drop=True)
        sits_lec = df_lec_fil["_situacion"].tolist()
        stks_lec = df_lec_fil["_nivel_stock"].tolist()
        df_lec_vis = df_lec_fil.drop(columns=["_situacion", "_nivel_stock"])

        if not df_lec_vis.empty:
            renderizar_tabla_moderna(
                df_lec_vis,
                modo_oscuro=modo_oscuro,
                estilos_celdas=estilizar_celdas_estado(df_lec_vis, sits_lec, stks_lec, modo_oscuro=modo_oscuro),
                enumerar=True,
            )
        else:
            st.markdown(
                """
                <div class="cuadro-vacio">
                    <p class="cuadro-vacio-titulo">No hay productos que coincidan con los filtros seleccionados</p>
                    <p class="cuadro-vacio-desc">Quita o cambia alguno de los filtros de estado, stock o categoría para ver todos los productos.</p>
                </div>
                """,
                unsafe_allow_html=True,
            )
    else:
        st.markdown(
            """
            <div class="cuadro-vacio">
                <p class="cuadro-vacio-titulo">Sin productos registrados por el momento</p>
                <p class="cuadro-vacio-desc">Cuando el editor registre productos y actualice los datos, aparecerán aquí.</p>
            </div>
            """,
            unsafe_allow_html=True,
        )

    st.divider()
    st.subheader("📈 Control de ventas, compras y balance")
    periodo_lec = st.selectbox(
        "Periodo",
        ["Todo", "Hoy", "Últimos 7 días", "Últimos 30 días"],
        key="filtro_periodo_lector",
    )
    hoy_lec = date.today()

    if datos["ventas"]:
        df_v_lec = pd.DataFrame(datos["ventas"]).copy()
        df_v_lec["fecha_dt"] = pd.to_datetime(df_v_lec["fecha"], errors="coerce")
        df_v_lec = df_v_lec.dropna(subset=["fecha_dt"])
        df_v_lec["total"] = df_v_lec["total"].astype(float)
        if periodo_lec == "Hoy":
            df_v_lec = df_v_lec[df_v_lec["fecha_dt"].dt.date == hoy_lec]
        elif periodo_lec == "Últimos 7 días":
            df_v_lec = df_v_lec[df_v_lec["fecha_dt"].dt.date >= (hoy_lec - timedelta(days=6))]
        elif periodo_lec == "Últimos 30 días":
            df_v_lec = df_v_lec[df_v_lec["fecha_dt"].dt.date >= (hoy_lec - timedelta(days=29))]
        total_v_lec = float(df_v_lec["total"].sum()) if not df_v_lec.empty else 0.0
    else:
        df_v_lec = pd.DataFrame()
        total_v_lec = 0.0

    if datos["compras"]:
        df_c_lec = pd.DataFrame(datos["compras"]).copy()
        df_c_lec["fecha_dt"] = pd.to_datetime(df_c_lec["fecha"], errors="coerce")
        df_c_lec = df_c_lec.dropna(subset=["fecha_dt"])
        df_c_lec["total_gastado"] = df_c_lec["total_gastado"].astype(float)
        if periodo_lec == "Hoy":
            df_c_lec = df_c_lec[df_c_lec["fecha_dt"].dt.date == hoy_lec]
        elif periodo_lec == "Últimos 7 días":
            df_c_lec = df_c_lec[df_c_lec["fecha_dt"].dt.date >= (hoy_lec - timedelta(days=6))]
        elif periodo_lec == "Últimos 30 días":
            df_c_lec = df_c_lec[df_c_lec["fecha_dt"].dt.date >= (hoy_lec - timedelta(days=29))]
        total_c_lec = float(df_c_lec["total_gastado"].sum()) if not df_c_lec.empty else 0.0
    else:
        total_c_lec = 0.0

    perdida_lec = sum(
        float(m.get("perdida_estimada", 0.0))
        for m in datos["mermas"]
        if m.get("tipo_registro") in ("Merma / Vencido", "Nuevo ingreso en mal estado")
    )
    cobrado_liq_lec = sum(
        float(m.get("monto_recuperado", 0.0))
        for m in datos["mermas"]
        if m.get("tipo_registro") == "Buen estado (otro precio / baja)"
    )
    dif_liq_lec = sum(
        float(m.get("perdida_estimada", 0.0))
        for m in datos["mermas"]
        if m.get("tipo_registro") == "Buen estado (otro precio / baja)"
    )
    balance_lec = round(total_v_lec - total_c_lec - perdida_lec, 2)

    vm1, vm2, vm3 = st.columns(3)
    with vm1:
        tarjeta_metrica("Total en ventas", fmt_quetzales(total_v_lec), modo_oscuro=modo_oscuro, paleta=paleta_activa)
    with vm2:
        tarjeta_metrica("Compras / reposición", fmt_quetzales(total_c_lec), modo_oscuro=modo_oscuro, paleta=paleta_activa)
    with vm3:
        tarjeta_metrica("Balance (Ventas - Compras - Merma)", fmt_quetzales(balance_lec), modo_oscuro=modo_oscuro, paleta=paleta_activa)

    if not df_v_lec.empty:
        filas_v_lec = [
            {
                "Fecha": r["fecha"],
                "Hora": r.get("hora", ""),
                "Producto": r["producto"],
                "Cantidad": formatear_cantidad(r["cantidad"], r.get("unidad", "Unidad")),
                "Total": fmt_quetzales(r.get("total", 0.0)),
            }
            for _, r in df_v_lec.sort_values("fecha_dt", ascending=False).iterrows()
        ]
        renderizar_tabla_moderna(pd.DataFrame(filas_v_lec), modo_oscuro=modo_oscuro, enumerar=True)
    else:
        st.markdown(
            """
            <div class="cuadro-vacio">
                <p class="cuadro-vacio-titulo">Sin ventas registradas en este periodo</p>
                <p class="cuadro-vacio-desc">Cuando se registren ventas, aparecerán aquí.</p>
            </div>
            """,
            unsafe_allow_html=True,
        )

    st.divider()
    st.subheader("📉 Merma, desperdicio y liquidación")
    lm1, lm2, lm3 = st.columns(3)
    with lm1:
        tarjeta_metrica("Desperdicio / pérdida (vencidos / mal estado)", fmt_quetzales(perdida_lec), modo_oscuro=modo_oscuro, paleta=paleta_activa)
    with lm2:
        tarjeta_metrica("Cobrado en rebaja / otro precio", fmt_quetzales(cobrado_liq_lec), modo_oscuro=modo_oscuro, paleta=paleta_activa)
    with lm3:
        tarjeta_metrica("Diferencia en rebaja / baja", fmt_quetzales(dif_liq_lec), modo_oscuro=modo_oscuro, paleta=paleta_activa)

    if datos["mermas"]:
        filas_m_lec = []
        tipos_m_lec = []
        for m in reversed(datos["mermas"]):
            t_reg = m.get("tipo_registro", "Merma / Vencido")
            tipos_m_lec.append(t_reg)
            filas_m_lec.append({
                "Fecha": m.get("fecha", ""),
                "Tipo": t_reg,
                "Producto": m.get("producto", ""),
                "Cantidad": formatear_cantidad(m.get("cantidad", 0), m.get("unidad", "Unidad")),
                "Cobrado": fmt_quetzales(m.get("monto_recuperado", 0.0)),
                "Pérdida / Dif.": fmt_quetzales(m.get("perdida_estimada", 0.0)),
            })
        df_m_lec_vis = pd.DataFrame(filas_m_lec)
        renderizar_tabla_moderna(
            df_m_lec_vis,
            modo_oscuro=modo_oscuro,
            estilos_celdas=estilizar_mermas(df_m_lec_vis, tipos_m_lec, modo_oscuro=modo_oscuro),
            enumerar=True,
        )
    else:
        st.markdown(
            """
            <div class="cuadro-vacio">
                <p class="cuadro-vacio-titulo">Sin registros de merma o liquidación</p>
                <p class="cuadro-vacio-desc">Aquí se mostrará lo desperdiciado y lo vendido en rebaja o liquidación.</p>
            </div>
            """,
            unsafe_allow_html=True,
        )

    st.divider()
    st.subheader(f"🔔 Alertas y notificaciones ({num_alertas})")
    if alertas_activas:
        df_n_lec = pd.DataFrame(alertas_activas).reset_index(drop=True)
        tipos_lec = sorted(df_n_lec["Tipo de alerta"].unique().tolist())
        tipos_sel_lec = st.multiselect(
            "Filtrar por tipo de alerta",
            options=tipos_lec,
            default=[],
            placeholder="Todos los tipos de alerta",
            key="filtro_alertas_lector",
        )
        if tipos_sel_lec:
            df_n_lec = df_n_lec[df_n_lec["Tipo de alerta"].isin(tipos_sel_lec)].reset_index(drop=True)

        est_n_lec = df_n_lec["_estilo"].tolist()
        df_n_lec_vis = df_n_lec[["Tipo de alerta", "Producto", "Categoría", "Cantidad actual", "Caducidad", "Días Rest.", "Detalle"]]
        if not df_n_lec_vis.empty:
            renderizar_tabla_moderna(
                df_n_lec_vis,
                modo_oscuro=modo_oscuro,
                estilos_celdas=estilizar_notificaciones(df_n_lec_vis, est_n_lec, modo_oscuro=modo_oscuro),
                enumerar=True,
            )
    else:
        st.markdown(
            """
            <div class="cuadro-vacio">
                <p class="cuadro-vacio-titulo">Sin alertas pendientes</p>
                <p class="cuadro-vacio-desc">No hay productos vencidos, próximos a vencer ni con stock bajo.</p>
            </div>
            """,
            unsafe_allow_html=True,
        )

# ==============================================================================
# VISTA 1: CONFIGURACIÓN (NOMBRE, LOGOTIPO, TÍTULOS Y PLANTILLAS DE COLOR)
# ==============================================================================
elif vista_lateral == "⚙️ Configuración":
    st.header("⚙️ Configuración y Personalización")

    st.subheader("Identidad, logotipo y nombres de pestañas")
    with st.form("form_identidad_app"):
        c_id1, c_id2 = st.columns(2)
        nuevo_nombre_neg = c_id1.text_input("Nombre del negocio / título principal *", value=nombre_negocio_actual)
        nuevo_sub_neg = c_id2.text_input("Subtítulo del banner", value=subtitulo_negocio_actual)

        archivo_logo = st.file_uploader("Cambiar logotipo (PNG / JPG / JPEG)", type=["png", "jpg", "jpeg"])

        st.markdown("**Nombres de pestañas**")
        titulos_actuales = cfg_app.get("titulos_pestanas", list(TITULOS_PESTANAS_BASE))
        nuevos_titulos = []
        p_cols = st.columns(2)
        for idx_t, t_act in enumerate(titulos_actuales):
            with p_cols[idx_t % 2]:
                nuevos_titulos.append(st.text_input(f"Pestaña {idx_t + 1}", value=t_act, key=f"cfg_tab_title_{idx_t}"))

        guardar_identidad = st.form_submit_button("Guardar identidad y pestañas")
        if guardar_identidad:
            if not nuevo_nombre_neg.strip():
                st.warning("El nombre del negocio / título principal es obligatorio.")
            else:
                cfg_app["nombre_negocio"] = nuevo_nombre_neg.strip()
                cfg_app["subtitulo_negocio"] = nuevo_sub_neg.strip()
                cfg_app["titulos_pestanas"] = [t.strip() or TITULOS_PESTANAS_BASE[i] for i, t in enumerate(nuevos_titulos)]

                if archivo_logo is not None:
                    ext = os.path.splitext(archivo_logo.name)[1].lower() or ".png"
                    ruta_guardada_logo = os.path.join(obtener_carpeta_media(), f"logo_personalizado{ext}")
                    with open(ruta_guardada_logo, "wb") as f_out:
                        f_out.write(archivo_logo.getbuffer())
                    cfg_app["ruta_logo"] = ruta_guardada_logo

                guardar_datos(datos)
                st.rerun()

    st.divider()
    st.subheader("Plantillas de colores")

    plantillas_dict = cfg_app.get("plantillas", {"Plantilla Principal": PLANTILLA_PRINCIPAL})
    nombres_plantillas = list(plantillas_dict.keys())
    activa_actual = cfg_app.get("plantilla_activa", "Plantilla Principal")
    idx_plat = nombres_plantillas.index(activa_actual) if activa_actual in nombres_plantillas else 0

    col_sel_p, col_btn_p = st.columns([2.5, 1.5])
    plantilla_elegida = col_sel_p.selectbox("Seleccionar plantilla activa", nombres_plantillas, index=idx_plat, key="sel_plantilla_activa")
    if col_btn_p.button("Aplicar plantilla", key="btn_aplicar_plantilla"):
        cfg_app["plantilla_activa"] = plantilla_elegida
        guardar_datos(datos)
        st.rerun()

    st.markdown("**Agregar nueva plantilla**")
    with st.form("form_nueva_plantilla", clear_on_submit=True):
        nombre_nueva_plat = st.text_input("Nombre de la nueva plantilla *")
        cp1, cp2, cp3, cp4 = st.columns(4)
        col_h2_osc = cp1.color_picker("Color títulos (oscuro)", value="#FDE68A")
        col_card_osc = cp2.color_picker("Color cuadros (oscuro)", value="#1B2B48")
        col_top_osc = cp3.color_picker("Borde superior cuadro", value="#FACC15")
        col_btn_osc = cp4.color_picker("Color botones", value="#1E5FD6")

        cp5, cp6, cp7, cp8 = st.columns(4)
        col_h2_cla = cp5.color_picker("Color títulos (claro)", value="#072A6E")
        col_card_cla = cp6.color_picker("Color cuadros (claro)", value="#FFFFFF")
        col_top_cla = cp7.color_picker("Borde superior (claro)", value="#0B3A8D")
        col_banner_1 = cp8.color_picker("Color banner principal", value="#0B3A8D")

        crear_plat = st.form_submit_button("Agregar plantilla")
        if crear_plat:
            if not nombre_nueva_plat.strip():
                st.warning("Escribe un nombre para la nueva plantilla.")
            else:
                nom_p = nombre_nueva_plat.strip()
                nueva_p = {
                    "nombre": nom_p,
                    "oscuro": {
                        "h2_color": col_h2_osc,
                        "metric_bg": col_card_osc,
                        "metric_top": col_top_osc,
                        "btn_bg": col_btn_osc,
                        "banner_bg": f"linear-gradient(95deg, {col_banner_1} 0%, #1654BE 68%, #F4B400 100%)",
                    },
                    "claro": {
                        "h2_color": col_h2_cla,
                        "metric_bg": col_card_cla,
                        "metric_top": col_top_cla,
                        "btn_bg": col_btn_osc,
                        "banner_bg": f"linear-gradient(95deg, {col_banner_1} 0%, #1654BE 68%, #F4B400 100%)",
                    },
                }
                plantillas_dict[nom_p] = nueva_p
                cfg_app["plantillas"] = plantillas_dict
                cfg_app["plantilla_activa"] = nom_p
                guardar_datos(datos)
                st.rerun()

# ==============================================================================
# VISTA 2: PAPELERA POR SECCIONES + CONTEO AUTOMÁTICO CONFIGURABLE
# ==============================================================================
elif vista_lateral.startswith("🗑️ Papelera"):
    st.header("🗑️ Papelera")
    st.caption("Organizada por sección según su origen. Selecciona uno o varios elementos en cada cuadro para restaurarlos o eliminarlos definitivamente.")

    cfg_pap = datos.get("papelera_auto_config") or {
        "activo": False,
        "modo_etiqueta": "Desactivado (no borrar automáticamente)",
        "horas_limite": 0.0,
    }
    modo_guardado = cfg_pap.get("modo_etiqueta", "Desactivado (no borrar automáticamente)")
    lista_modos_pap = list(OPCIONES_TIEMPO_PAPELERA.keys())
    idx_modo_pap = lista_modos_pap.index(modo_guardado) if modo_guardado in lista_modos_pap else 0

    with st.expander("⏱️ Tiempo para eliminar automáticamente de la papelera", expanded=True):
        st.caption("Elige cuánto tiempo quieres que pase para que los elementos de la papelera se borren solos.")
        col_t1, col_t2 = st.columns([2.2, 1.8])
        modo_tiempo_sel = col_t1.selectbox(
            "Tiempo de borrado automático",
            lista_modos_pap,
            index=idx_modo_pap,
            key="sel_tiempo_papelera",
        )
        valor_modo = OPCIONES_TIEMPO_PAPELERA[modo_tiempo_sel]
        horas_finales = 0.0
        activo_final = False
        etiqueta_final = modo_tiempo_sel

        if valor_modo == "custom":
            unidad_custom = col_t2.selectbox("Unidad de tiempo", ["Días", "Horas"], key="sel_unidad_custom_pap")
            horas_prev = float(cfg_pap.get("horas_limite", 24.0) or 24.0)
            val_def = round(horas_prev / 24.0, 2) if unidad_custom == "Días" else round(horas_prev, 2)
            cant_custom = st.number_input(
                f"Cantidad de {unidad_custom.lower()} para borrar automáticamente *",
                min_value=0.5,
                step=1.0,
                value=max(0.5, float(val_def)),
                format="%.1f",
                key="num_custom_tiempo_pap",
            )
            horas_finales = float(cant_custom) * (24.0 if unidad_custom == "Días" else 1.0)
            activo_final = horas_finales > 0
            etiqueta_final = f"Cada {fmt_num(cant_custom)} {unidad_custom.lower()}"
        elif isinstance(valor_modo, float):
            horas_finales = valor_modo
            activo_final = True

        estado_actual_txt = etiqueta_final if (cfg_pap.get("activo") and float(cfg_pap.get("horas_limite", 0.0)) > 0) else "Desactivado"
        st.caption(f"Configuración guardada actualmente: **{estado_actual_txt}**")

        if st.button("Guardar tiempo de borrado automático", key="btn_guardar_tiempo_papelera"):
            datos["papelera_auto_config"] = {
                "activo": activo_final,
                "modo_etiqueta": etiqueta_final,
                "horas_limite": horas_finales,
            }
            purgados_ahora = purgar_papelera_vencida(datos)
            guardar_datos(datos)
            st.session_state["forzar_papelera_lateral"] = True
            if purgados_ahora > 0:
                st.session_state["aviso_purga_pap"] = f"Se guardó el tiempo ({etiqueta_final}) y se eliminaron {purgados_ahora} elemento(s) vencido(s)."
            else:
                st.session_state["aviso_purga_pap"] = f"Tiempo de borrado automático guardado: {etiqueta_final}."
            st.rerun()

    if st.session_state.get("aviso_purga_pap"):
        st.success(st.session_state["aviso_purga_pap"])
        st.session_state["aviso_purga_pap"] = ""

    clave_papelera_req = obtener_clave_seccion(datos, "papelera")
    papelera_items = datos.get("papelera", [])

    if papelera_items:
        secciones_papelera = [
            ("Producto", "📋 Productos eliminados de Inventario", "prod"),
            ("Compra", "📦 Compras y reposiciones eliminadas", "comp"),
            ("Venta", "📈 Registros eliminados de Control de Ventas", "vta"),
            ("Merma", "📉 Mermas, mal estado o rebajas eliminadas", "merma"),
        ]

        alguna_seccion_con_datos = False

        for tipo_clave, titulo_sec, prefijo_key in secciones_papelera:
            items_sec = [it for it in reversed(papelera_items) if it.get("tipo") == tipo_clave]
            if not items_sec:
                continue

            alguna_seccion_con_datos = True
            st.markdown(f'<div class="papelera-seccion-hdr">{html.escape(titulo_sec)} ({len(items_sec)})</div>', unsafe_allow_html=True)

            filas_sec = [
                {
                    "Eliminado": it.get("eliminado_en", ""),
                    "Detalle": it.get("resumen", ""),
                }
                for it in items_sec
            ]
            renderizar_tabla_moderna(pd.DataFrame(filas_sec), modo_oscuro=modo_oscuro, enumerar=True)

            mapa_opc_sec = {}
            for idx_i, it in enumerate(items_sec, start=1):
                etiqueta_item = f"#{idx_i} · [{it.get('eliminado_en', '')}] {it.get('resumen', '')}"
                mapa_opc_sec[etiqueta_item] = it["id_papelera"]

            lista_opc_sec = list(mapa_opc_sec.keys())

            key_ms = f"pap_ms_{prefijo_key}"
            key_conf = f"pap_conf_{prefijo_key}"
            key_clave = f"pap_clave_{prefijo_key}"
            key_reset = f"pap_reset_{prefijo_key}"

            if st.session_state.get(key_reset):
                st.session_state[key_ms] = []
                st.session_state[key_conf] = False
                st.session_state[key_clave] = ""
                st.session_state[key_reset] = False
            elif key_ms in st.session_state:
                st.session_state[key_ms] = [x for x in st.session_state[key_ms] if x in mapa_opc_sec]

            seleccionados = st.multiselect(
                "Selecciona uno o varios elementos de esta sección",
                options=lista_opc_sec,
                key=key_ms,
                placeholder="Elige uno o más registros...",
            )

            col_s_rest, col_s_del = st.columns([1.4, 2.6])

            with col_s_rest:
                if st.button("♻️ Restaurar seleccionados", key=f"btn_rest_{prefijo_key}", disabled=not seleccionados):
                    ids_rest = {mapa_opc_sec[lbl] for lbl in seleccionados}
                    for it in papelera_items:
                        if it["id_papelera"] in ids_rest:
                            tipo_r = it.get("tipo")
                            obj_r = it.get("datos_item")
                            if tipo_r == "Producto" and isinstance(obj_r, dict):
                                datos["productos"].append(obj_r)
                            elif tipo_r == "Compra" and isinstance(obj_r, dict):
                                datos["compras"].append(obj_r)
                            elif tipo_r == "Venta" and isinstance(obj_r, dict):
                                datos["ventas"].append(obj_r)
                            elif tipo_r == "Merma" and isinstance(obj_r, dict):
                                datos["mermas"].append(obj_r)
                    datos["papelera"] = [x for x in papelera_items if x["id_papelera"] not in ids_rest]
                    guardar_datos(datos)
                    st.session_state[key_reset] = True
                    st.session_state["forzar_papelera_lateral"] = True
                    st.rerun()

            with col_s_del:
                conf_s = st.checkbox("Confirmar eliminación definitiva", key=key_conf)
                clave_s = st.text_input("Contraseña *", type="password", key=key_clave)
                puede_borrar = bool(seleccionados and conf_s)
                if st.button(
                    "Eliminar seleccionados definitivamente",
                    type="primary",
                    key=f"btn_del_def_{prefijo_key}",
                    disabled=not puede_borrar,
                ):
                    if clave_s == clave_papelera_req:
                        ids_del = {mapa_opc_sec[lbl] for lbl in seleccionados}
                        datos["papelera"] = [x for x in papelera_items if x["id_papelera"] not in ids_del]
                        guardar_datos(datos)
                        st.session_state[key_reset] = True
                        st.session_state["forzar_papelera_lateral"] = True
                        st.rerun()
                    else:
                        st.error("Contraseña incorrecta.")

        if not alguna_seccion_con_datos:
            st.markdown(
                """
                <div class="cuadro-vacio">
                    <p class="cuadro-vacio-titulo">La papelera está vacía</p>
                    <p class="cuadro-vacio-desc">Los productos o registros eliminados aparecerán aquí agrupados por sección.</p>
                </div>
                """,
                unsafe_allow_html=True,
            )
    else:
        st.markdown(
            """
            <div class="cuadro-vacio">
                <p class="cuadro-vacio-titulo">La papelera está vacía</p>
                <p class="cuadro-vacio-desc">Los productos o registros eliminados aparecerán aquí para poder restaurarlos si lo necesitas.</p>
            </div>
            """,
            unsafe_allow_html=True,
        )

# ==============================================================================
# VISTA 3: GESTIÓN DE CONTRASEÑAS POR SECCIÓN (PROTEGIDA CON CLAVE PRINCIPAL)
# ==============================================================================
elif vista_lateral == etiqueta_claves_menu:
    st.header("🔐 Contraseñas por sección")
    st.caption("Ingresa la contraseña principal para desbloquear y editar la contraseña de cada sección. Se bloqueará automáticamente al salir de esta vista.")

    if not st.session_state.get("claves_desbloqueadas", False):
        with st.form("form_desbloquear_claves"):
            clave_maestra_in = st.text_input("Contraseña principal *", type="password", key="input_clave_maestra_sec")
            entrar_claves = st.form_submit_button("Desbloquear contraseñas")
            if entrar_claves:
                if clave_maestra_in == CLAVE_PRINCIPAL:
                    st.session_state["claves_desbloqueadas"] = True
                    st.rerun()
                else:
                    st.error("Contraseña principal incorrecta.")
    else:
        with st.form("form_editar_claves_secciones"):
            st.markdown("**Acceso al Modo Editor (pantalla inicial / enlace compartido)**")
            nueva_clave_editor = st.text_input(
                "Contraseña para entrar a Modo Editor *",
                value=clave_editor_activa,
                key="edit_clave_modo_editor",
            )
            st.divider()

            nuevas_claves = {}
            for k_sec, nombre_sec in SECCIONES_CONTRASENA:
                valor_actual_sec = obtener_clave_seccion(datos, k_sec)
                st.markdown(f"**{nombre_sec}**")
                nuevas_claves[k_sec] = st.text_input(
                    f"Contraseña para {nombre_sec} *",
                    value=valor_actual_sec,
                    key=f"edit_clave_{k_sec}",
                )

            guardar_claves = st.form_submit_button("Guardar contraseñas")
            if guardar_claves:
                vacias = [nombre_sec for k_sec, nombre_sec in SECCIONES_CONTRASENA if not str(nuevas_claves.get(k_sec, "")).strip()]
                if not str(nueva_clave_editor or "").strip():
                    st.warning("La contraseña para entrar a Modo Editor no puede quedar vacía.")
                elif vacias:
                    st.warning(f"Ninguna contraseña puede quedar vacía: {', '.join(vacias)}.")
                else:
                    datos["clave_editor"] = str(nueva_clave_editor).strip()
                    datos["contrasenas_secciones"] = {
                        k_sec: str(nuevas_claves[k_sec]).strip()
                        for k_sec, _ in SECCIONES_CONTRASENA
                    }
                    guardar_datos(datos)
                    st.success("Contraseñas actualizadas correctamente para cada sección.")

# ==============================================================================
# VISTA PRINCIPAL DEL SISTEMA EDITOR (7 PESTAÑAS)
# ==============================================================================
else:
    titulos_cfg = cfg_app.get("titulos_pestanas", list(TITULOS_PESTANAS_BASE))
    etiqueta_notif = f"{titulos_cfg[6]} ({num_alertas})" if num_alertas > 0 else f"{titulos_cfg[6]} (0)"

    titulos_pestanas = [
        titulos_cfg[0],
        titulos_cfg[1],
        titulos_cfg[2],
        titulos_cfg[3],
        titulos_cfg[4],
        titulos_cfg[5],
        etiqueta_notif,
    ]

    colores_pestana = [
        ("linear-gradient(180deg, #1E63E9 0%, #0E3A94 100%)", "#FFD633", "#FFD633", "0 4px 0 #071E4D, 0 8px 16px rgba(30, 99, 233, 0.35)"),
        ("linear-gradient(180deg, #16A34A 0%, #0E6E32 100%)", "#FFD633", "#FFD633", "0 4px 0 #073B1A, 0 8px 16px rgba(22, 163, 74, 0.35)"),
        ("linear-gradient(180deg, #D97706 0%, #92400E 100%)", "#FFD633", "#FFD633", "0 4px 0 #451A03, 0 8px 16px rgba(217, 119, 6, 0.35)"),
        ("linear-gradient(180deg, #E11D48 0%, #9F1239 100%)", "#FFD633", "#FFD633", "0 4px 0 #58081E, 0 8px 16px rgba(225, 29, 72, 0.38)"),
        ("linear-gradient(180deg, #7C3AED 0%, #4C1D95 100%)", "#FFD633", "#FFD633", "0 4px 0 #2E1065, 0 8px 16px rgba(124, 58, 237, 0.38)"),
        ("linear-gradient(180deg, #0D9488 0%, #115E59 100%)", "#FFD633", "#FFD633", "0 4px 0 #042F2E, 0 8px 16px rgba(13, 148, 136, 0.38)"),
        ("linear-gradient(180deg, #DC2626 0%, #7F1D1D 100%)", "#FFD633", "#FFD633", "0 4px 0 #450A0A, 0 8px 16px rgba(220, 38, 38, 0.38)"),
    ]

    idx_act = st.session_state["pestana_activa"]
    if idx_act >= len(titulos_pestanas):
        idx_act = 0
        st.session_state["pestana_activa"] = 0

    bg_act, bdr_act, txt_act, shd_act = colores_pestana[idx_act]

    css_nav = f"""
    <style>
    div[class*="st-key-btn_nav_"],
    div[class*="st-key-btn_nav_"] > div,
    div[class*="st-key-btn_nav_"] button {{
        background-color: transparent !important;
        box-shadow: none !important;
    }}
    div[class*="st-key-btn_nav_"] button {{
        min-height: 42px !important;
        border-radius: 11px !important;
        padding: 0.42rem 0.45rem !important;
        white-space: nowrap !important;
        overflow: hidden !important;
        text-overflow: ellipsis !important;
        transition: transform 0.12s ease, box-shadow 0.12s ease !important;
    }}
    div[class*="st-key-btn_nav_"] button [data-testid="stMarkdownContainer"] p,
    div[class*="st-key-btn_nav_"] button p,
    div[class*="st-key-btn_nav_"] button span,
    div[class*="st-key-btn_nav_"] button div {{
        font-family: 'Outfit', sans-serif !important;
        font-size: 0.84rem !important;
        font-weight: 700 !important;
        line-height: 1.15 !important;
        white-space: nowrap !important;
    }}
    div.st-key-btn_nav_{idx_act} button {{
        background: {bg_act} !important;
        border: 2.5px solid {bdr_act} !important;
        box-shadow: {shd_act} !important;
        transform: translateY(-1px) !important;
    }}
    div.st-key-btn_nav_{idx_act} button [data-testid="stMarkdownContainer"] p,
    div.st-key-btn_nav_{idx_act} button p,
    div.st-key-btn_nav_{idx_act} button span,
    div.st-key-btn_nav_{idx_act} button div {{
        color: {txt_act} !important;
    }}
    </style>
    """
    st.markdown(css_nav, unsafe_allow_html=True)

    cols_nav = st.columns(7, gap="small")
    for i, titulo in enumerate(titulos_pestanas):
        with cols_nav[i]:
            if st.button(titulo, key=f"btn_nav_{i}", use_container_width=True):
                st.session_state["pestana_activa"] = i
                st.rerun()

    st.markdown("<div style='margin-bottom: 8px;'></div>", unsafe_allow_html=True)

    # ==========================================================================
    # SECCIÓN 1: ALTA DE PRODUCTO Y GESTIÓN / EDICIÓN DE CATEGORÍAS
    # ==========================================================================
    if idx_act == 0:
        st.header(titulos_cfg[0])

        with st.form("form_alta_producto", clear_on_submit=True):
            c1, c2 = st.columns([3, 2])
            nombre = c1.text_input("Nombre del producto *")
            idx_cat_def = categorias_disponibles.index("General") if "General" in categorias_disponibles else 0
            categoria = c2.selectbox("Categoría (si no eliges, queda General)", categorias_disponibles, index=idx_cat_def)

            c3, c4, c5 = st.columns([1.8, 1.5, 2])
            cantidad = c3.number_input("Cantidad inicial *", min_value=0.0, step=1.0, value=0.0, format="%.2f")
            unidad = c4.selectbox("Unidad de medida *", UNIDADES_MEDIDA)
            precio = c5.number_input("Precio unitario *", min_value=0.0, step=0.5, value=0.0, format="%.2f")

            fecha_venc = st.date_input("Fecha de caducidad / vencimiento", value=date.today() + timedelta(days=15))
            guardar_prod = st.form_submit_button("Guardar producto")

            if guardar_prod:
                if not nombre.strip():
                    st.warning("El nombre del producto es obligatorio.")
                elif float(cantidad) <= 0:
                    st.warning("La cantidad inicial debe ser mayor a 0.")
                elif float(precio) <= 0:
                    st.warning("El precio unitario es obligatorio y debe ser mayor a 0.")
                else:
                    cat_final = categoria if categoria in categorias_disponibles else "General"
                    cant_limpia = round(float(cantidad), 2)
                    precio_limpio = float(precio)
                    fecha_venc_str = fecha_venc.strftime("%Y-%m-%d")
                    fecha_hoy_str = date.today().strftime("%Y-%m-%d")
                    prod_id = datetime.now().strftime("%Y%m%d%H%M%S%f")

                    datos["productos"].append({
                        "id": prod_id,
                        "nombre": nombre.strip(),
                        "categoria": cat_final,
                        "cantidad": cant_limpia,
                        "unidad": unidad,
                        "precio_unitario": precio_limpio,
                        "fecha_vencimiento": fecha_venc_str,
                        "estado": "Buen estado",
                        "fecha_registro": fecha_hoy_str,
                        "lotes": [{
                            "id_lote": f"lote_{prod_id}",
                            "cantidad": cant_limpia,
                            "unidad": unidad,
                            "fecha_vencimiento": fecha_venc_str,
                            "fecha_ingreso": fecha_hoy_str,
                            "estado": "Buen estado",
                            "origen": "Registro inicial",
                        }],
                    })
                    guardar_datos(datos)
                    st.rerun()

        st.divider()
        st.subheader("Clasificación y edición de categorías")
        chips_html = "".join(f"<span class='cat-chip'>{cat}</span>" for cat in categorias_disponibles)
        st.markdown(f"<div>{chips_html}</div>", unsafe_allow_html=True)

        col_add, col_ren, col_del = st.columns(3)

        with col_add:
            st.markdown("**Agregar categoría**")
            nueva_cat = st.text_input("Nombre de la nueva categoría *", key="nueva_cat_input")
            if st.button("Agregar categoría", key="btn_agregar_cat"):
                nombre_cat = nueva_cat.strip()
                if not nombre_cat:
                    st.warning("Escribe el nombre de la categoría.")
                elif not any(c.lower() == nombre_cat.lower() for c in datos["categorias"]):
                    datos["categorias"].append(nombre_cat)
                    guardar_datos(datos)
                    st.rerun()

        with col_ren:
            st.markdown("**Editar / renombrar categoría**")
            cat_a_editar = st.selectbox("Categoría a editar", categorias_disponibles, key="sel_cat_editar")
            nuevo_nombre_cat = st.text_input("Nuevo nombre *", value=cat_a_editar, key="input_nuevo_nombre_cat")
            if st.button("Guardar cambio de nombre", key="btn_editar_cat"):
                limpio = nuevo_nombre_cat.strip()
                if not limpio:
                    st.warning("El nuevo nombre de la categoría es obligatorio.")
                elif not (limpio.lower() != cat_a_editar.lower() and any(c.lower() == limpio.lower() for c in datos["categorias"])):
                    datos["categorias"] = [limpio if c == cat_a_editar else c for c in datos["categorias"]]
                    for p in datos["productos"]:
                        if p.get("categoria") == cat_a_editar:
                            p["categoria"] = limpio
                    for comp in datos["compras"]:
                        if comp.get("categoria") == cat_a_editar:
                            comp["categoria"] = limpio
                    guardar_datos(datos)
                    st.rerun()

        with col_del:
            st.markdown("**Eliminar categoría**")
            cats_eliminables = [c for c in categorias_disponibles if c != "General"]
            if cats_eliminables:
                cat_a_borrar = st.selectbox("Categoría a eliminar", cats_eliminables, key="sel_cat_borrar")
                if st.button("Eliminar categoría", key="btn_borrar_cat"):
                    datos["categorias"] = [c for c in datos["categorias"] if c != cat_a_borrar]
                    for p in datos["productos"]:
                        if p.get("categoria") == cat_a_borrar:
                            p["categoria"] = "General"
                    for comp in datos["compras"]:
                        if comp.get("categoria") == cat_a_borrar:
                            comp["categoria"] = "General"
                    guardar_datos(datos)
                    st.rerun()

        if datos["productos"]:
            st.markdown("**Cambiar categoría a un producto existente**")
            opciones_prod_cat = {
                f"{p['nombre']} ({formatear_cantidad(p['cantidad'], p.get('unidad', 'Unidad'))}) — Actual: {p.get('categoria', 'General')}": p["id"]
                for p in datos["productos"]
            }
            prod_sel_cat_label = st.selectbox("Producto", list(opciones_prod_cat.keys()), key="sel_prod_cambiar_cat")
            prod_obj_cat = next(p for p in datos["productos"] if p["id"] == opciones_prod_cat[prod_sel_cat_label])
            cat_actual_prod = prod_obj_cat.get("categoria", "General")
            idx_cat = categorias_disponibles.index(cat_actual_prod) if cat_actual_prod in categorias_disponibles else 0
            nueva_cat_prod = st.selectbox("Nueva categoría del producto", categorias_disponibles, index=idx_cat, key="sel_nueva_cat_prod")
            if st.button("Asignar categoría", key="btn_asignar_cat_prod"):
                prod_obj_cat["categoria"] = nueva_cat_prod
                guardar_datos(datos)
                st.rerun()

    # ==========================================================================
    # SECCIÓN 2: LISTADO COMPLETO DE PRODUCTOS, SUBCUADRO POR PRODUCTO Y ESTADO
    # ==========================================================================
    elif idx_act == 1:
        st.header(titulos_cfg[1])

        if datos["productos"]:
            st.subheader("Listado de productos")

            buscar_prod = st.text_input(
                "🔍 Buscar producto",
                placeholder="Escribe el nombre o categoría del producto...",
                key="buscador_prod_inv",
            ).strip().lower()

            filas_detalle = []
            for p in datos["productos"]:
                normalizar_lotes_producto(p)
                unidad_p = p.get("unidad", "Unidad")
                cant_num = float(p.get("cantidad", 0.0))
                situacion = clasificar_vencimiento(p)
                nivel_stk = clasificar_stock(cant_num)
                filas_detalle.append({
                    "ID": p["id"],
                    "Producto": p["nombre"],
                    "Categoría": p.get("categoria", "General"),
                    "Cantidad": formatear_cantidad(cant_num, unidad_p),
                    "Medida": unidad_p,
                    "Precio": fmt_quetzales(p.get("precio_unitario", 0.0)),
                    "Stock": nivel_stk,
                    "Caducidad": p.get("fecha_vencimiento", "") or "Sin fecha",
                    "_situacion": situacion,
                    "_nivel_stock": nivel_stk,
                })

            df_det = pd.DataFrame(filas_detalle)

            f_col1, f_col2, f_col3 = st.columns(3)
            situaciones_filtro = f_col1.multiselect(
                "Filtrar por situación de caducidad",
                options=["Buen estado", "Próximo a vencer", "Vencido", "Vencido / Mal estado", "Sin fecha"],
                default=[],
                placeholder="Todos los estados de caducidad",
            )
            niveles_filtro = f_col2.multiselect(
                "Filtrar por nivel de stock",
                options=["Stock suficiente", "Por surtir pronto", "Poco producto"],
                default=[],
                placeholder="Todos los niveles de stock",
            )
            cats_en_uso = sorted(df_det["Categoría"].unique().tolist())
            categorias_filtro = f_col3.multiselect(
                "Filtrar por categoría",
                options=cats_en_uso,
                default=[],
                placeholder="Todas las categorías",
            )

            mascara_busqueda = (
                df_det["Producto"].str.lower().str.contains(buscar_prod, na=False)
                | df_det["Categoría"].str.lower().str.contains(buscar_prod, na=False)
            ) if buscar_prod else pd.Series([True] * len(df_det), index=df_det.index)

            mascara_cad = df_det["_situacion"].isin(situaciones_filtro) if situaciones_filtro else pd.Series([True] * len(df_det), index=df_det.index)
            mascara_stk = df_det["_nivel_stock"].isin(niveles_filtro) if niveles_filtro else pd.Series([True] * len(df_det), index=df_det.index)
            mascara_cat = df_det["Categoría"].isin(categorias_filtro) if categorias_filtro else pd.Series([True] * len(df_det), index=df_det.index)

            df_filtrado = df_det[
                mascara_cad
                & mascara_stk
                & mascara_cat
                & mascara_busqueda
            ].drop(columns=["ID"]).reset_index(drop=True)

            situaciones_det = df_filtrado["_situacion"].tolist()
            niveles_det = df_filtrado["_nivel_stock"].tolist()
            df_mostrar_visible = df_filtrado.drop(columns=["_situacion", "_nivel_stock"])

            renderizar_tabla_moderna(
                df_mostrar_visible,
                modo_oscuro=modo_oscuro,
                estilos_celdas=estilizar_celdas_inventario(
                    df_mostrar_visible, situaciones_det, niveles_det, modo_oscuro=modo_oscuro
                ),
                enumerar=True,
            )

            st.divider()
            st.subheader("Detalle de lotes y caducidad por producto")

            opciones_subcuadro = {
                f"{p['nombre']} ({formatear_cantidad(float(p.get('cantidad', 0.0)), p.get('unidad', 'Unidad'))})": p["id"]
                for p in datos["productos"]
            }
            sel_sub_label = st.selectbox(
                "Selecciona un producto para ver su desglose",
                list(opciones_subcuadro.keys()),
                key="sel_subcuadro_producto",
            )
            prod_sub = next(p for p in datos["productos"] if p["id"] == opciones_subcuadro[sel_sub_label])
            normalizar_lotes_producto(prod_sub)
            unidad_sub = prod_sub.get("unidad", "Unidad")
            lotes_sub = [l for l in prod_sub["lotes"] if float(l.get("cantidad", 0.0)) > 0] or prod_sub["lotes"]

            fechas_lotes = [l.get("fecha_vencimiento", "") for l in lotes_sub if l.get("fecha_vencimiento")]
            proxima_fecha = min(fechas_lotes) if fechas_lotes else "Sin fecha"
            ultimo_ingreso = max((l.get("fecha_ingreso", "") for l in lotes_sub if l.get("fecha_ingreso")), default="N/A")

            m_sub1, m_sub2, m_sub3 = st.columns(3)
            with m_sub1:
                tarjeta_metrica("Total acumulado", formatear_cantidad(float(prod_sub.get("cantidad", 0.0)), unidad_sub), modo_oscuro=modo_oscuro, paleta=paleta_activa)
            with m_sub2:
                tarjeta_metrica("Próxima caducidad", proxima_fecha, modo_oscuro=modo_oscuro, paleta=paleta_activa)
            with m_sub3:
                tarjeta_metrica("Última entrada / compra", ultimo_ingreso, modo_oscuro=modo_oscuro, paleta=paleta_activa)

            filas_lotes_sub = []
            for idx_l, l in enumerate(
                sorted(lotes_sub, key=lambda x: (x.get("fecha_vencimiento") or "9999-12-31", x.get("fecha_ingreso") or "")),
                start=1,
            ):
                sit_l = clasificar_vencimiento_lote(l, prod_sub.get("estado", "Buen estado"))
                dias_l = calcular_dias_restantes(l.get("fecha_vencimiento", ""))
                es_reciente = "Sí" if l.get("fecha_ingreso") == ultimo_ingreso else "No"
                filas_lotes_sub.append({
                    "Lote": f"Lote #{idx_l}",
                    "Cantidad": formatear_cantidad(float(l.get("cantidad", 0.0)), unidad_sub),
                    "Entrada": l.get("fecha_ingreso", ""),
                    "Caducidad": l.get("fecha_vencimiento", "") or "Sin fecha",
                    "Días Rest.": dias_l if dias_l is not None else "-",
                    "Situación": sit_l,
                    "Compra Reciente": es_reciente,
                })

            renderizar_tabla_moderna(pd.DataFrame(filas_lotes_sub), modo_oscuro=modo_oscuro, enumerar=True)

            st.divider()
            st.subheader("Estado de cada producto")

            filas_estado = []
            for p in datos["productos"]:
                normalizar_lotes_producto(p)
                unidad_p = p.get("unidad", "Unidad")
                cant_num = float(p.get("cantidad", 0.0))
                dias_rest = calcular_dias_restantes(p.get("fecha_vencimiento", ""))
                situacion = clasificar_vencimiento(p)
                nivel_stk = clasificar_stock(cant_num)
                filas_estado.append({
                    "Producto": p["nombre"],
                    "Categoría": p.get("categoria", "General"),
                    "Cantidad": formatear_cantidad(cant_num, unidad_p),
                    "Nivel Stock": nivel_stk,
                    "Caducidad": p.get("fecha_vencimiento", "") or "Sin fecha",
                    "Días Rest.": dias_rest if dias_rest is not None else "-",
                    "Estado": situacion,
                    "_situacion": situacion,
                    "_nivel_stock": nivel_stk,
                })

            df_estado = pd.DataFrame(filas_estado)
            if buscar_prod:
                df_estado = df_estado[
                    df_estado["Producto"].str.lower().str.contains(buscar_prod, na=False)
                    | df_estado["Categoría"].str.lower().str.contains(buscar_prod, na=False)
                ]
            mascara_cad_est = df_estado["_situacion"].isin(situaciones_filtro) if situaciones_filtro else pd.Series([True] * len(df_estado), index=df_estado.index)
            mascara_stk_est = df_estado["_nivel_stock"].isin(niveles_filtro) if niveles_filtro else pd.Series([True] * len(df_estado), index=df_estado.index)
            mascara_cat_est = df_estado["Categoría"].isin(categorias_filtro) if categorias_filtro else pd.Series([True] * len(df_estado), index=df_estado.index)

            df_estado = df_estado[mascara_cad_est & mascara_stk_est & mascara_cat_est].drop(columns=["Categoría"]).reset_index(drop=True)
            situaciones_est = df_estado["_situacion"].tolist()
            niveles_est = df_estado["_nivel_stock"].tolist()
            df_estado_visible = df_estado.drop(columns=["_situacion", "_nivel_stock"])

            renderizar_tabla_moderna(
                df_estado_visible,
                modo_oscuro=modo_oscuro,
                estilos_celdas=estilizar_celdas_estado(
                    df_estado_visible, situaciones_est, niveles_est, modo_oscuro=modo_oscuro
                ),
                enumerar=True,
            )

            st.divider()
            st.subheader("🔒 Eliminar un registro del inventario")
            clave_inv_req = obtener_clave_seccion(datos, "inventario")
            opciones_eliminar = {
                f"{p['nombre']} [{p.get('categoria', 'General')}] | {formatear_cantidad(p['cantidad'], p.get('unidad', 'Unidad'))} | Precio: {fmt_quetzales(p.get('precio_unitario', 0.0))} | Vence: {p.get('fecha_vencimiento', 'N/A')}": p["id"]
                for p in datos["productos"]
            }
            sel_elim_label = st.selectbox("Selecciona el producto o lote a eliminar", list(opciones_eliminar.keys()), key="sel_inv_borrar")
            preparar_confirmacion_eliminacion("inv_ultimo_sel_borrar", sel_elim_label, "conf_inv_borrar", "clave_elim_inv", "inv_limpiar_next")

            confirmar = st.checkbox("Confirmo que deseo eliminar este registro del inventario", key="conf_inv_borrar")
            clave_inv = st.text_input("Ingresa la contraseña para eliminar el registro *", type="password", key="clave_elim_inv")

            if st.button("Eliminar registro", type="primary", key="btn_elim_inv", disabled=not confirmar):
                if confirmar and clave_inv == clave_inv_req:
                    id_borrar = opciones_eliminar[sel_elim_label]
                    prod_obj = next((p for p in datos["productos"] if p["id"] == id_borrar), None)
                    if prod_obj:
                        mover_a_papelera(
                            datos,
                            "Producto",
                            prod_obj,
                            f"{prod_obj.get('nombre', '')} ({formatear_cantidad(prod_obj.get('cantidad', 0), prod_obj.get('unidad', 'Unidad'))})",
                        )
                    datos["productos"] = [p for p in datos["productos"] if p["id"] != id_borrar]
                    guardar_datos(datos)
                    solicitar_limpieza_eliminacion("inv_limpiar_next")
                    st.rerun()
                elif confirmar:
                    st.error("Contraseña incorrecta.")
        else:
            st.session_state["conf_inv_borrar"] = False
            st.session_state["clave_elim_inv"] = ""
            st.session_state["inv_ultimo_sel_borrar"] = ""
            st.session_state["inv_limpiar_next"] = False
            st.markdown(
                """
                <div class="cuadro-vacio">
                    <p class="cuadro-vacio-titulo">En este momento no hay productos en el inventario</p>
                    <p class="cuadro-vacio-desc">Registra un producto en la pestaña Registrar Producto o compra uno en Compra y Reposición para verlo aquí.</p>
                </div>
                """,
                unsafe_allow_html=True,
            )

    # ==========================================================================
    # SECCIÓN 3: COMPRAR PRODUCTO NUEVO, SUMARLO AL INVENTARIO Y REGISTRAR GASTO
    # ==========================================================================
    elif idx_act == 2:
        st.header(titulos_cfg[2])

        if st.session_state.get("comp_reset_todo", False):
            st.session_state["comp_sel_producto_existente"] = "— Escribir otro producto —"
            st.session_state["comp_input_nombre_nuevo"] = ""
            st.session_state["comp_sel_categoria"] = "General" if "General" in categorias_disponibles else categorias_disponibles[0]
            st.session_state["comp_input_cant"] = 0.0
            st.session_state["comp_sel_unidad"] = UNIDADES_MEDIDA[0]
            st.session_state["comp_input_precio"] = 0.0
            st.session_state["comp_input_gasto"] = 0.0
            st.session_state["comp_fecha_cad"] = date.today() + timedelta(days=15)
            st.session_state["comp_firma_seleccion"] = "— Escribir otro producto —"
            st.session_state["comp_reset_todo"] = False

        productos_por_reponer = []
        if datos["productos"]:
            filas_alerta_rep = []
            for p in datos["productos"]:
                normalizar_lotes_producto(p)
                unidad_p = p.get("unidad", "Unidad")
                cant_num = float(p.get("cantidad", 0.0))
                nivel_stk = clasificar_stock(cant_num)
                if nivel_stk in ("Poco producto", "Por surtir pronto"):
                    productos_por_reponer.append(p["nombre"])
                    filas_alerta_rep.append({
                        "Producto": p["nombre"],
                        "Categoría": p.get("categoria", "General"),
                        "Disponible": formatear_cantidad(cant_num, unidad_p),
                        "Estado de Stock": nivel_stk,
                        "Caducidad": p.get("fecha_vencimiento", "") or "Sin fecha",
                        "_cant_num": cant_num,
                        "_nivel_stock": nivel_stk,
                    })

            total_criticos = sum(1 for r in filas_alerta_rep if r["_nivel_stock"] == "Poco producto")
            total_preventivos = sum(1 for r in filas_alerta_rep if r["_nivel_stock"] == "Por surtir pronto")

            st.subheader("Productos por reponer")
            m_rep1, m_rep2 = st.columns(2)
            with m_rep1:
                tarjeta_metrica("Poco producto (≤ 4)", str(total_criticos), modo_oscuro=modo_oscuro, paleta=paleta_activa)
            with m_rep2:
                tarjeta_metrica("Por surtir pronto (5)", str(total_preventivos), modo_oscuro=modo_oscuro, paleta=paleta_activa)

            if filas_alerta_rep:
                df_rep = pd.DataFrame(filas_alerta_rep).sort_values("_cant_num", ascending=True).reset_index(drop=True)
                niveles_rep = df_rep["_nivel_stock"].tolist()
                df_rep_visible = df_rep.drop(columns=["_cant_num", "_nivel_stock"])
                renderizar_tabla_moderna(
                    df_rep_visible,
                    modo_oscuro=modo_oscuro,
                    estilos_celdas=estilizar_tabla_reposicion(df_rep_visible, niveles_rep, modo_oscuro=modo_oscuro),
                    enumerar=True,
                )

                col_rap1, col_rap2 = st.columns([2.5, 1.5])
                prod_rapido = col_rap1.selectbox(
                    "Reponer ahora",
                    sorted(set(productos_por_reponer)),
                    key="sel_reponer_ahora",
                )
                if col_rap2.button("Reponer ahora", key="btn_reponer_ahora"):
                    st.session_state["comp_sel_producto_existente"] = prod_rapido
                    st.session_state["comp_firma_seleccion"] = ""
                    st.rerun()
            st.divider()

        mapa_prod_existente = {}
        for p in datos["productos"]:
            clave_nom = p["nombre"].strip().lower()
            if clave_nom not in mapa_prod_existente:
                mapa_prod_existente[clave_nom] = p

        nombres_existentes = sorted({p["nombre"] for p in datos["productos"]})
        opciones_prod_comp = ["— Escribir otro producto —"] + nombres_existentes

        if st.session_state.get("comp_sel_producto_existente") not in opciones_prod_comp:
            st.session_state["comp_sel_producto_existente"] = "— Escribir otro producto —"

        c1, c2 = st.columns([2.2, 1.8])
        producto_elegido = c1.selectbox(
            "Producto existente",
            opciones_prod_comp,
            key="comp_sel_producto_existente",
        )
        nombre_nuevo = c2.text_input(
            "O escribe el nombre del producto *",
            key="comp_input_nombre_nuevo",
            disabled=(producto_elegido != "— Escribir otro producto —"),
        )

        nombre_previo = "" if producto_elegido == "— Escribir otro producto —" else producto_elegido
        prod_ref = mapa_prod_existente.get(nombre_previo.strip().lower()) if nombre_previo else None

        if st.session_state.get("comp_firma_seleccion") != producto_elegido:
            st.session_state["comp_firma_seleccion"] = producto_elegido
            if prod_ref:
                cat_def = prod_ref.get("categoria", "General")
                if cat_def not in categorias_disponibles:
                    cat_def = "General" if "General" in categorias_disponibles else categorias_disponibles[0]
                uni_def = prod_ref.get("unidad", "Unidad")
                if uni_def not in UNIDADES_MEDIDA:
                    uni_def = UNIDADES_MEDIDA[0]
                precio_def = float(prod_ref.get("precio_unitario", 0.0))
                st.session_state["comp_sel_categoria"] = cat_def
                st.session_state["comp_sel_unidad"] = uni_def
                st.session_state["comp_input_precio"] = precio_def
                st.session_state["comp_input_cant"] = 0.0
                st.session_state["comp_input_gasto"] = 0.0
            else:
                st.session_state["comp_sel_categoria"] = "General" if "General" in categorias_disponibles else categorias_disponibles[0]
                st.session_state["comp_sel_unidad"] = UNIDADES_MEDIDA[0]
                st.session_state["comp_input_precio"] = 0.0
                st.session_state["comp_input_cant"] = 0.0
                st.session_state["comp_input_gasto"] = 0.0

        if "comp_input_cant" not in st.session_state:
            st.session_state["comp_input_cant"] = 0.0
        if "comp_input_gasto" not in st.session_state:
            st.session_state["comp_input_gasto"] = 0.0
        if "comp_fecha_cad" not in st.session_state:
            st.session_state["comp_fecha_cad"] = date.today() + timedelta(days=15)

        with st.form("form_compra_reposicion", clear_on_submit=False):
            c3, c4, c5, c6 = st.columns([1.8, 1.7, 1.5, 2.0])
            categoria_compra = c3.selectbox("Categoría", categorias_disponibles, key="comp_sel_categoria")
            cant_comprada = c4.number_input("Cantidad comprada *", min_value=0.0, step=1.0, key="comp_input_cant", format="%.2f")
            unidad_compra = c5.selectbox("Unidad de medida *", UNIDADES_MEDIDA, key="comp_sel_unidad")
            precio_compra = c6.number_input("Precio unitario del lote *", min_value=0.0, step=0.5, key="comp_input_precio", format="%.2f")

            c7, c8 = st.columns([2.0, 2.5])
            gasto_total = c7.number_input("Total gastado en la compra *", min_value=0.0, step=1.0, key="comp_input_gasto", format="%.2f")
            fecha_cad_nueva = c8.date_input("Fecha de caducidad de este nuevo lote", key="comp_fecha_cad")

            guardar_compra = st.form_submit_button("Registrar compra y sumar al inventario")

            if guardar_compra:
                nombre_final = nombre_nuevo.strip() if producto_elegido == "— Escribir otro producto —" else producto_elegido
                cant_c = round(float(cant_comprada), 2)
                gasto_c = round(float(gasto_total), 2)
                precio_in = float(precio_compra)

                if not nombre_final:
                    st.warning("El nombre del producto es obligatorio.")
                elif cant_c <= 0:
                    st.warning("La cantidad comprada debe ser mayor a 0.")
                elif precio_in <= 0 and gasto_c <= 0:
                    st.warning("Ingresa el precio unitario o el total gastado en la compra (mayor a 0).")
                else:
                    precio_c = precio_in if precio_in > 0 else round(gasto_c / cant_c, 2)
                    if gasto_c == 0 and precio_c > 0:
                        gasto_c = round(cant_c * precio_c, 2)

                    cat_final_c = categoria_compra if categoria_compra in categorias_disponibles else "General"
                    fecha_cad_str = fecha_cad_nueva.strftime("%Y-%m-%d")
                    fecha_hoy_str = date.today().strftime("%Y-%m-%d")

                    lote_igual = next(
                        (p for p in datos["productos"]
                         if p["nombre"].lower() == nombre_final.lower()
                         and p.get("unidad", "Unidad").lower() == unidad_compra.lower()
                         and p.get("estado") == "Buen estado"),
                        None
                    )

                    if lote_igual:
                        normalizar_lotes_producto(lote_igual)
                        lote_igual["categoria"] = cat_final_c
                        if precio_c > 0:
                            lote_igual["precio_unitario"] = precio_c

                        lote_misma_fecha = next(
                            (l for l in lote_igual["lotes"]
                             if l.get("fecha_vencimiento") == fecha_cad_str
                             and l.get("estado", "Buen estado") == "Buen estado"),
                            None
                        )
                        if lote_misma_fecha:
                            lote_misma_fecha["cantidad"] = round(float(lote_misma_fecha["cantidad"]) + cant_c, 2)
                            lote_misma_fecha["fecha_ingreso"] = fecha_hoy_str
                        else:
                            lote_igual["lotes"].append({
                                "id_lote": f"lote_{datetime.now().strftime('%Y%m%d%H%M%S%f')}",
                                "cantidad": cant_c,
                                "unidad": unidad_compra,
                                "fecha_vencimiento": fecha_cad_str,
                                "fecha_ingreso": fecha_hoy_str,
                                "estado": "Buen estado",
                                "origen": "Compra / Reposición",
                            })
                        sincronizar_resumen_producto(lote_igual)
                    else:
                        prod_id = datetime.now().strftime("%Y%m%d%H%M%S%f")
                        datos["productos"].append({
                            "id": prod_id,
                            "nombre": nombre_final,
                            "categoria": cat_final_c,
                            "cantidad": cant_c,
                            "unidad": unidad_compra,
                            "precio_unitario": precio_c,
                            "fecha_vencimiento": fecha_cad_str,
                            "estado": "Buen estado",
                            "fecha_registro": fecha_hoy_str,
                            "lotes": [{
                                "id_lote": f"lote_{prod_id}",
                                "cantidad": cant_c,
                                "unidad": unidad_compra,
                                "fecha_vencimiento": fecha_cad_str,
                                "fecha_ingreso": fecha_hoy_str,
                                "estado": "Buen estado",
                                "origen": "Compra / Reposición",
                            }],
                        })

                    datos["compras"].append({
                        "id": datetime.now().strftime("%Y%m%d%H%M%S%f"),
                        "fecha": fecha_hoy_str,
                        "producto": nombre_final,
                        "categoria": cat_final_c,
                        "cantidad": cant_c,
                        "unidad": unidad_compra,
                        "caducidad": fecha_cad_str,
                        "total_gastado": gasto_c
                    })

                    guardar_datos(datos)
                    st.session_state["comp_reset_todo"] = True
                    st.rerun()

        st.divider()
        st.subheader("Historial de compras y gasto registrado")
        if datos["compras"]:
            total_invertido = sum(float(c.get("total_gastado", 0.0)) for c in datos["compras"])
            tarjeta_metrica("Total gastado en compras", fmt_quetzales(total_invertido), modo_oscuro=modo_oscuro, paleta=paleta_activa)

            filas_compras = [
                {
                    "Fecha": c.get("fecha", ""),
                    "Producto": c.get("producto", ""),
                    "Categoría": c.get("categoria", "General"),
                    "Cantidad Comprada": formatear_cantidad(c.get("cantidad", 0), c.get("unidad", "Unidad")),
                    "Caducidad Lote": c.get("caducidad", "") or "Sin fecha",
                    "Total Gastado": fmt_quetzales(c.get("total_gastado", 0.0))
                }
                for c in reversed(datos["compras"])
            ]
            renderizar_tabla_moderna(pd.DataFrame(filas_compras), modo_oscuro=modo_oscuro, enumerar=True)

            st.divider()
            st.subheader("🔒 Eliminar una compra o reposición")
            clave_comp_req = obtener_clave_seccion(datos, "compras")
            opciones_comp_elim = {
                f"[{c.get('fecha', '')}] {c.get('producto', '')} — {formatear_cantidad(c.get('cantidad', 0), c.get('unidad', 'Unidad'))} | Cad: {c.get('caducidad', 'N/A')} ({fmt_quetzales(c.get('total_gastado', 0.0))})": c["id"]
                for c in reversed(datos["compras"])
            }
            sel_comp_borrar = st.selectbox(
                "Selecciona la compra o reposición que deseas eliminar",
                list(opciones_comp_elim.keys()),
                key="sel_comp_borrar"
            )
            preparar_confirmacion_eliminacion("comp_ultimo_sel_borrar", sel_comp_borrar, "conf_comp_borrar", "clave_elim_comp", "comp_limpiar_next")

            restar_stock_comp = st.checkbox(
                "También descontar esta cantidad del lote correspondiente en el inventario",
                value=False,
                key="restar_stock_comp"
            )
            confirmar_comp = st.checkbox(
                "Confirmo que deseo eliminar este registro de compra/reposición",
                key="conf_comp_borrar"
            )
            clave_comp = st.text_input(
                "Ingresa la contraseña para eliminar la compra o reposición *",
                type="password",
                key="clave_elim_comp"
            )

            if st.button("Eliminar compra seleccionada", type="primary", key="btn_elim_comp", disabled=not confirmar_comp):
                if confirmar_comp and clave_comp == clave_comp_req:
                    id_comp_borrar = opciones_comp_elim[sel_comp_borrar]
                    compra_obj = next((c for c in datos["compras"] if c["id"] == id_comp_borrar), None)

                    if compra_obj:
                        mover_a_papelera(
                            datos,
                            "Compra",
                            compra_obj,
                            f"{compra_obj.get('producto', '')} — {formatear_cantidad(compra_obj.get('cantidad', 0), compra_obj.get('unidad', 'Unidad'))} ({fmt_quetzales(compra_obj.get('total_gastado', 0.0))})",
                        )
                        if restar_stock_comp:
                            cant_restar = float(compra_obj.get("cantidad", 0.0))
                            prod_nom = str(compra_obj.get("producto", "")).lower()
                            prod_uni = str(compra_obj.get("unidad", "Unidad")).lower()
                            prod_cad = str(compra_obj.get("caducidad", ""))

                            lote_obj = next(
                                (p for p in datos["productos"]
                                 if p["nombre"].lower() == prod_nom
                                 and p.get("unidad", "Unidad").lower() == prod_uni
                                 and p.get("estado") == "Buen estado"),
                                None
                            )
                            if lote_obj:
                                normalizar_lotes_producto(lote_obj)
                                lote_exacto = next(
                                    (l for l in lote_obj["lotes"]
                                     if l.get("fecha_vencimiento") == prod_cad
                                     and float(l.get("cantidad", 0.0)) >= cant_restar),
                                    None
                                )
                                if lote_exacto:
                                    lote_exacto["cantidad"] = round(float(lote_exacto["cantidad"]) - cant_restar, 2)
                                    lote_obj["lotes"] = [l for l in lote_obj["lotes"] if float(l.get("cantidad", 0.0)) > 0]
                                    sincronizar_resumen_producto(lote_obj)
                                else:
                                    descontar_de_lotes(lote_obj, cant_restar)

                    datos["compras"] = [c for c in datos["compras"] if c["id"] != id_comp_borrar]
                    guardar_datos(datos)
                    solicitar_limpieza_eliminacion("comp_limpiar_next")
                    st.rerun()
                elif confirmar_comp:
                    st.error("Contraseña incorrecta.")
        else:
            st.session_state["conf_comp_borrar"] = False
            st.session_state["clave_elim_comp"] = ""
            st.session_state["comp_ultimo_sel_borrar"] = ""
            st.session_state["comp_limpiar_next"] = False

    # ==========================================================================
    # SECCIÓN 4: REGISTRAR VENTAS CON AJUSTE UNO A UNO Y CÁLCULO DE INGRESO
    # ==========================================================================
    elif idx_act == 3:
        st.header(titulos_cfg[3])

        disponibles_venta = [
            p for p in datos["productos"]
            if float(p.get("cantidad", 0.0)) > 0 and p.get("estado") == "Buen estado"
        ]

        if disponibles_venta:
            if st.session_state.get("vta_reset_todo", False):
                st.session_state["sel_vta_label"] = "— Selecciona un producto —"
                st.session_state["vta_prod_id"] = ""
                st.session_state["vta_cant_acum"] = 0.0
                st.session_state["vta_input_cant"] = 0.0
                st.session_state["vta_precio_unit"] = 0.0
                st.session_state["vta_reset_todo"] = False

            opciones_vta = {
                f"{p['nombre']} [{p.get('categoria', 'General')}] | Disp: {formatear_cantidad(p['cantidad'], p.get('unidad', 'Unidad'))} | Vence: {p.get('fecha_vencimiento', 'N/A')} | Precio: {fmt_quetzales(p.get('precio_unitario', 0.0))}": p["id"]
                for p in disponibles_venta
            }
            lista_opciones_vta = ["— Selecciona un producto —"] + list(opciones_vta.keys())
            if st.session_state.get("sel_vta_label") not in lista_opciones_vta:
                st.session_state["sel_vta_label"] = "— Selecciona un producto —"

            sel_vta_label = st.selectbox("Selecciona el producto vendido *", lista_opciones_vta, key="sel_vta_label")

            if sel_vta_label == "— Selecciona un producto —":
                st.session_state["vta_prod_id"] = ""
                st.session_state["vta_input_cant"] = 0.0
                st.session_state["vta_precio_unit"] = 0.0

                with st.form("form_solo_venta_vacio", clear_on_submit=False):
                    v1, v2, v3 = st.columns(3)
                    v1.number_input("Cantidad vendida *", min_value=0.0, step=1.0, key="vta_input_cant", format="%.2f", disabled=True)
                    v2.number_input("Precio de venta unitario *", min_value=0.0, step=0.5, key="vta_precio_unit", format="%.2f", disabled=True)
                    with v3:
                        tarjeta_metrica("Total del ingreso por esta venta", fmt_quetzales(0.0), modo_oscuro=modo_oscuro, paleta=paleta_activa)
                    st.form_submit_button("Confirmar venta", disabled=True)
            else:
                prod_v = next(p for p in datos["productos"] if p["id"] == opciones_vta[sel_vta_label])
                normalizar_lotes_producto(prod_v)
                unidad_v = prod_v.get("unidad", "Unidad")
                max_v = float(prod_v["cantidad"])
                min_v = 0.0
                paso_v = 1.0 if max_v >= 1.0 else max(0.01, round(max_v / 4, 2))

                if st.session_state.get("vta_prod_id") != prod_v["id"]:
                    st.session_state["vta_prod_id"] = prod_v["id"]
                    st.session_state["vta_cant_acum"] = 0.0
                    st.session_state["vta_input_cant"] = 0.0
                    st.session_state["vta_precio_unit"] = float(prod_v["precio_unitario"])

                cant_actual_ses = min(max_v, max(0.0, round(float(st.session_state.get("vta_cant_acum", 0.0)), 2)))
                st.session_state["vta_cant_acum"] = cant_actual_ses
                if "vta_input_cant" not in st.session_state or float(st.session_state["vta_input_cant"]) > max_v:
                    st.session_state["vta_input_cant"] = cant_actual_ses

                col_r1, col_r2, col_r3 = st.columns([1.5, 1.5, 1.5])
                with col_r1:
                    if st.button(f"➕ +1 {unidad_v}", key="btn_vta_mas_uno"):
                        nueva_c = min(max_v, round(float(st.session_state["vta_input_cant"]) + 1.0, 2))
                        st.session_state["vta_cant_acum"] = nueva_c
                        st.session_state["vta_input_cant"] = nueva_c
                        st.rerun()
                with col_r2:
                    if st.button(f"➖ -1 {unidad_v}", key="btn_vta_menos_uno"):
                        nueva_c = max(0.0, round(float(st.session_state["vta_input_cant"]) - 1.0, 2))
                        st.session_state["vta_cant_acum"] = nueva_c
                        st.session_state["vta_input_cant"] = nueva_c
                        st.rerun()
                with col_r3:
                    if st.button("Reiniciar a 0", key="btn_vta_reset"):
                        st.session_state["vta_cant_acum"] = 0.0
                        st.session_state["vta_input_cant"] = 0.0
                        st.rerun()

                with st.form("form_solo_venta", clear_on_submit=False):
                    v1, v2, v3 = st.columns(3)
                    cant_v = v1.number_input(
                        f"Cantidad vendida ({unidad_v}) *",
                        min_value=min_v,
                        max_value=max_v,
                        step=paso_v,
                        key="vta_input_cant",
                        format="%.2f"
                    )
                    precio_v = v2.number_input(
                        "Precio de venta unitario *",
                        min_value=0.0,
                        step=0.5,
                        key="vta_precio_unit",
                        format="%.2f"
                    )
                    total_ingreso_venta = round(float(cant_v) * float(precio_v), 2)
                    with v3:
                        tarjeta_metrica("Total del ingreso por esta venta", fmt_quetzales(total_ingreso_venta), modo_oscuro=modo_oscuro, paleta=paleta_activa)

                    registrar_vta = st.form_submit_button("Confirmar venta")

                    if registrar_vta:
                        cant_limpia = round(float(cant_v), 2)
                        if cant_limpia <= 0:
                            st.warning("La cantidad vendida debe ser mayor a 0.")
                        elif float(precio_v) <= 0:
                            st.warning("El precio de venta unitario es obligatorio y debe ser mayor a 0.")
                        else:
                            total_limpio = round(cant_limpia * float(precio_v), 2)
                            caducidad_vendida = descontar_de_lotes(prod_v, cant_limpia)

                            datos["ventas"].append({
                                "id": datetime.now().strftime("%Y%m%d%H%M%S%f"),
                                "fecha": date.today().strftime("%Y-%m-%d"),
                                "hora": datetime.now().strftime("%H:%M"),
                                "producto": prod_v["nombre"],
                                "cantidad": cant_limpia,
                                "unidad": unidad_v,
                                "caducidad_lote": caducidad_vendida,
                                "precio_unitario": float(precio_v),
                                "total": total_limpio
                            })

                            guardar_datos(datos)
                            st.session_state["vta_reset_todo"] = True
                            st.rerun()
        else:
            st.markdown(
                """
                <div class="cuadro-vacio">
                    <p class="cuadro-vacio-titulo">No hay productos disponibles en buen estado para vender</p>
                    <p class="cuadro-vacio-desc">Registra o repone productos en inventario para poder realizar ventas.</p>
                </div>
                """,
                unsafe_allow_html=True,
            )

        total_ingresos_ventas = sum(float(v.get("total", 0.0)) for v in datos["ventas"])
        st.divider()
        tarjeta_metrica("Suma total de ingresos por ventas", fmt_quetzales(total_ingresos_ventas), modo_oscuro=modo_oscuro, paleta=paleta_activa)

    # ==========================================================================
    # SECCIÓN 5: HISTORIAL DE VENTAS, BALANCE, FILTROS Y GRÁFICAS DE LÍNEA
    # ==========================================================================
    elif idx_act == 4:
        st.header(titulos_cfg[4])

        if datos["ventas"]:
            df_ventas = pd.DataFrame(datos["ventas"]).copy()
            df_ventas["fecha_dt"] = pd.to_datetime(df_ventas["fecha"], errors="coerce")
            df_ventas = df_ventas.dropna(subset=["fecha_dt"])
            df_ventas["total"] = df_ventas["total"].astype(float)
            if "caducidad_lote" not in df_ventas.columns:
                df_ventas["caducidad_lote"] = "Sin registro"
            else:
                df_ventas["caducidad_lote"] = df_ventas["caducidad_lote"].fillna("Sin registro").replace("", "Sin registro")

            f_v1, f_v2, f_v3 = st.columns(3)
            rango_rapido = f_v1.selectbox(
                "Periodo",
                ["Todo", "Hoy", "Últimos 7 días", "Últimos 30 días", "Personalizado"],
                key="filtro_periodo_ventas",
            )
            prods_venta_lista = sorted(df_ventas["producto"].dropna().unique().tolist())
            prods_venta_sel = f_v2.multiselect(
                "Producto",
                options=prods_venta_lista,
                default=[],
                placeholder="Todos los productos",
                key="filtro_prod_ventas",
            )
            if rango_rapido == "Personalizado":
                fecha_min_d = df_ventas["fecha_dt"].min().date()
                fecha_max_d = max(df_ventas["fecha_dt"].max().date(), date.today())
                rango_fechas = f_v3.date_input(
                    "Rango de fechas",
                    value=(fecha_min_d, fecha_max_d),
                    key="filtro_rango_fechas_ventas",
                )
            else:
                rango_fechas = None
                f_v3.empty()

            df_vf = df_ventas.copy()
            hoy_d = date.today()
            if rango_rapido == "Hoy":
                df_vf = df_vf[df_vf["fecha_dt"].dt.date == hoy_d]
            elif rango_rapido == "Últimos 7 días":
                df_vf = df_vf[df_vf["fecha_dt"].dt.date >= (hoy_d - timedelta(days=6))]
            elif rango_rapido == "Últimos 30 días":
                df_vf = df_vf[df_vf["fecha_dt"].dt.date >= (hoy_d - timedelta(days=29))]
            elif rango_rapido == "Personalizado" and isinstance(rango_fechas, (tuple, list)) and len(rango_fechas) == 2:
                d_ini, d_fin = rango_fechas
                df_vf = df_vf[(df_vf["fecha_dt"].dt.date >= d_ini) & (df_vf["fecha_dt"].dt.date <= d_fin)]

            if prods_venta_sel:
                df_vf = df_vf[df_vf["producto"].isin(prods_venta_sel)]

            total_general = float(df_vf["total"].sum()) if not df_vf.empty else 0.0
            num_ventas = len(df_vf)
            promedio_venta = float(df_vf["total"].mean()) if not df_vf.empty else 0.0

            df_compras_all = pd.DataFrame(datos["compras"]) if datos["compras"] else pd.DataFrame(columns=["fecha", "producto", "total_gastado"])
            if not df_compras_all.empty:
                df_compras_all["fecha_dt"] = pd.to_datetime(df_compras_all["fecha"], errors="coerce")
                df_compras_all["total_gastado"] = df_compras_all["total_gastado"].astype(float)
                df_cf = df_compras_all.dropna(subset=["fecha_dt"]).copy()
                if rango_rapido == "Hoy":
                    df_cf = df_cf[df_cf["fecha_dt"].dt.date == hoy_d]
                elif rango_rapido == "Últimos 7 días":
                    df_cf = df_cf[df_cf["fecha_dt"].dt.date >= (hoy_d - timedelta(days=6))]
                elif rango_rapido == "Últimos 30 días":
                    df_cf = df_cf[df_cf["fecha_dt"].dt.date >= (hoy_d - timedelta(days=29))]
                elif rango_rapido == "Personalizado" and isinstance(rango_fechas, (tuple, list)) and len(rango_fechas) == 2:
                    d_ini, d_fin = rango_fechas
                    df_cf = df_cf[(df_cf["fecha_dt"].dt.date >= d_ini) & (df_cf["fecha_dt"].dt.date <= d_fin)]
                if prods_venta_sel and "producto" in df_cf.columns:
                    df_cf = df_cf[df_cf["producto"].isin(prods_venta_sel)]
                total_compras_periodo = float(df_cf["total_gastado"].sum()) if not df_cf.empty else 0.0
            else:
                total_compras_periodo = 0.0

            balance_neto = round(total_general - total_compras_periodo, 2)

            m1, m2, m3, m4 = st.columns(4)
            with m1:
                tarjeta_metrica("Total en ventas", fmt_quetzales(total_general), modo_oscuro=modo_oscuro, paleta=paleta_activa)
            with m2:
                tarjeta_metrica("Compras del periodo", fmt_quetzales(total_compras_periodo), modo_oscuro=modo_oscuro, paleta=paleta_activa)
            with m3:
                tarjeta_metrica("Balance (Ventas - Compras)", fmt_quetzales(balance_neto), modo_oscuro=modo_oscuro, paleta=paleta_activa)
            with m4:
                tarjeta_metrica("Ventas / Promedio", f"{num_ventas} · {fmt_quetzales(promedio_venta)}", modo_oscuro=modo_oscuro, paleta=paleta_activa)

            st.divider()
            st.subheader("Historial de ventas")
            if not df_vf.empty:
                filas_hist = [
                    {
                        "Fecha": row["fecha"],
                        "Hora": row.get("hora", ""),
                        "Producto": row["producto"],
                        "Cantidad": formatear_cantidad(row["cantidad"], row.get("unidad", "Unidad")),
                        "Caducidad Lote": row.get("caducidad_lote", "Sin registro"),
                        "Precio Unit.": fmt_quetzales(row.get("precio_unitario", 0.0)),
                        "Total": fmt_quetzales(row.get("total", 0.0))
                    }
                    for _, row in df_vf.sort_values("fecha_dt", ascending=False).iterrows()
                ]
                renderizar_tabla_moderna(pd.DataFrame(filas_hist), modo_oscuro=modo_oscuro, enumerar=True)
            else:
                st.markdown(
                    """
                    <div class="cuadro-vacio">
                        <p class="cuadro-vacio-titulo">No hay ventas con los filtros seleccionados</p>
                        <p class="cuadro-vacio-desc">Cambia el periodo o el producto para ver los registros correspondientes.</p>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )

            st.divider()
            st.subheader("🔒 Eliminar un registro de venta")
            clave_vta_req = obtener_clave_seccion(datos, "ventas")
            opciones_vta_elim = {
                f"[{v.get('fecha', '')} {v.get('hora', '')}] {v.get('producto', '')} — {formatear_cantidad(v.get('cantidad', 0), v.get('unidad', 'Unidad'))} | Cad: {v.get('caducidad_lote', 'Sin registro')} ({fmt_quetzales(v.get('total', 0.0))})": v["id"]
                for v in reversed(datos["ventas"])
            }
            sel_vta_borrar = st.selectbox("Selecciona la venta que deseas eliminar", list(opciones_vta_elim.keys()), key="sel_vta_borrar")
            preparar_confirmacion_eliminacion("vta_ultimo_sel_borrar", sel_vta_borrar, "conf_vta_borrar", "clave_elim_vta", "vta_limpiar_next")

            confirmar_vta = st.checkbox("Confirmo que deseo eliminar este registro de venta", key="conf_vta_borrar")
            clave_vta = st.text_input("Ingresa la contraseña para eliminar la venta *", type="password", key="clave_elim_vta")

            if st.button("Eliminar venta seleccionada", type="primary", key="btn_elim_vta", disabled=not confirmar_vta):
                if confirmar_vta and clave_vta == clave_vta_req:
                    id_vta_borrar = opciones_vta_elim[sel_vta_borrar]
                    vta_obj = next((v for v in datos["ventas"] if v["id"] == id_vta_borrar), None)
                    if vta_obj:
                        mover_a_papelera(
                            datos,
                            "Venta",
                            vta_obj,
                            f"{vta_obj.get('producto', '')} — {formatear_cantidad(vta_obj.get('cantidad', 0), vta_obj.get('unidad', 'Unidad'))} ({fmt_quetzales(vta_obj.get('total', 0.0))})",
                        )
                    datos["ventas"] = [v for v in datos["ventas"] if v["id"] != id_vta_borrar]
                    guardar_datos(datos)
                    solicitar_limpieza_eliminacion("vta_limpiar_next")
                    st.rerun()
                elif confirmar_vta:
                    st.error("Contraseña incorrecta.")

            if not df_vf.empty:
                st.divider()
                st.subheader("Gráficas de ventas (subidas y bajones)")

                st.subheader("Ventas por Día")
                serie_dia = (
                    df_vf.groupby(df_vf["fecha_dt"].dt.strftime("%Y-%m-%d"))["total"]
                    .sum()
                    .sort_index()
                )
                st.line_chart(serie_dia)

                st.subheader("Ventas por Semana")
                df_vf["inicio_semana"] = df_vf["fecha_dt"].dt.to_period("W-SUN").apply(lambda p: p.start_time.strftime("%Y-%m-%d"))
                serie_semana = (
                    df_vf.groupby("inicio_semana")["total"]
                    .sum()
                    .sort_index()
                )
                st.line_chart(serie_semana)

                st.subheader("Ventas por Mes")
                serie_mes = (
                    df_vf.groupby(df_vf["fecha_dt"].dt.to_period("M").astype(str))["total"]
                    .sum()
                    .sort_index()
                )
                st.line_chart(serie_mes)

                st.subheader("Ventas por Año")
                serie_anio = (
                    df_vf.groupby(df_vf["fecha_dt"].dt.year.astype(str))["total"]
                    .sum()
                    .sort_index()
                )
                st.line_chart(serie_anio)
        else:
            st.session_state["conf_vta_borrar"] = False
            st.session_state["clave_elim_vta"] = ""
            st.session_state["vta_ultimo_sel_borrar"] = ""
            st.session_state["vta_limpiar_next"] = False
            st.markdown(
                """
                <div class="cuadro-vacio">
                    <p class="cuadro-vacio-titulo">En este momento no hay ventas registradas</p>
                    <p class="cuadro-vacio-desc">Cuando registres una venta en la pestaña Registrar Venta, aquí aparecerá el historial y la evolución de tus ventas.</p>
                </div>
                """,
                unsafe_allow_html=True,
            )

    # ==========================================================================
    # SECCIÓN 6: MERMA / VENCIDO, NUEVO INGRESO EN MAL ESTADO Y BUEN ESTADO
    # ==========================================================================
    elif idx_act == 5:
        st.header(titulos_cfg[5])

        tipo_mov = st.radio(
            "Tipo de registro",
            [
                "Merma / Vencido (pérdida)",
                "Nuevo ingreso en mal estado",
                "Buen estado (otro precio / baja)",
            ],
            horizontal=True,
            key="tipo_mov_merma",
        )

        if tipo_mov == "Merma / Vencido (pérdida)":
            productos_venc_o_por_vencer = [
                p for p in datos["productos"]
                if float(p.get("cantidad", 0.0)) > 0
                and clasificar_vencimiento(p) in ("Vencido", "Próximo a vencer", "Vencido / Mal estado")
            ]

            if productos_venc_o_por_vencer:
                opciones_m = {
                    f"{p['nombre']} [{clasificar_vencimiento(p)}] | Disp: {formatear_cantidad(p['cantidad'], p.get('unidad', 'Unidad'))} | Vence: {p.get('fecha_vencimiento', 'N/A')} | Precio: {fmt_quetzales(p.get('precio_unitario', 0.0))}": p["id"]
                    for p in productos_venc_o_por_vencer
                }
                sel_m_label = st.selectbox(
                    "Selecciona el producto vencido o próximo a vencer *",
                    list(opciones_m.keys()),
                    key="sel_prod_merma_venc",
                )
                prod_m = next(p for p in datos["productos"] if p["id"] == opciones_m[sel_m_label])
                normalizar_lotes_producto(prod_m)
                unidad_m = prod_m.get("unidad", "Unidad")
                max_m = float(prod_m["cantidad"])
                precio_base_m = float(prod_m.get("precio_unitario", 0.0))
                paso_m = 1.0 if max_m >= 1.0 else max(0.01, round(max_m / 4, 2))

                with st.form("form_merma_perdida", clear_on_submit=True):
                    c_m1, c_m2, c_m3 = st.columns(3)
                    cant_m = c_m1.number_input(
                        f"Cantidad vencida o por vencer ({unidad_m}) *",
                        min_value=0.0,
                        max_value=max_m,
                        step=paso_m,
                        value=0.0,
                        format="%.2f",
                    )
                    costo_unit_m = c_m2.number_input(
                        "Valor / costo unitario de referencia *",
                        min_value=0.0,
                        step=0.5,
                        value=precio_base_m,
                        format="%.2f",
                    )
                    motivo_m = c_m3.selectbox(
                        "Motivo",
                        ["Producto vencido", "Próximo a vencer / descarte", "Descarte de lote"],
                    )
                    perdida_calc = round(float(cant_m) * float(costo_unit_m), 2)
                    tarjeta_metrica("Pérdida estimada por esta merma", fmt_quetzales(perdida_calc), modo_oscuro=modo_oscuro, paleta=paleta_activa)

                    guardar_m = st.form_submit_button("Registrar merma y descontar inventario")
                    if guardar_m:
                        cant_limpia = round(float(cant_m), 2)
                        if cant_limpia <= 0:
                            st.warning("La cantidad debe ser mayor a 0.")
                        elif float(costo_unit_m) <= 0:
                            st.warning("El valor / costo unitario es obligatorio y debe ser mayor a 0.")
                        else:
                            perdida_limpia = round(cant_limpia * float(costo_unit_m), 2)
                            cad_descontada = descontar_de_lotes(prod_m, cant_limpia)
                            datos["mermas"].append({
                                "id": datetime.now().strftime("%Y%m%d%H%M%S%f"),
                                "fecha": date.today().strftime("%Y-%m-%d"),
                                "hora": datetime.now().strftime("%H:%M"),
                                "tipo_registro": "Merma / Vencido",
                                "producto": prod_m["nombre"],
                                "cantidad": cant_limpia,
                                "unidad": unidad_m,
                                "caducidad_lote": cad_descontada,
                                "motivo": motivo_m,
                                "precio_original": float(costo_unit_m),
                                "precio_aplicado": 0.0,
                                "descuento_pct": 100.0,
                                "monto_recuperado": 0.0,
                                "perdida_estimada": perdida_limpia,
                            })
                            guardar_datos(datos)
                            st.rerun()
            else:
                st.markdown(
                    """
                    <div class="cuadro-vacio">
                        <p class="cuadro-vacio-titulo">No hay productos vencidos ni próximos a vencer en este momento</p>
                        <p class="cuadro-vacio-desc">Aquí solo aparecen los productos que ya están vencidos o dentro del rango de días por vencer.</p>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )

        elif tipo_mov == "Nuevo ingreso en mal estado":
            productos_disponibles_mal = [
                p for p in datos["productos"]
                if float(p.get("cantidad", 0.0)) > 0
            ]

            if productos_disponibles_mal:
                opciones_ni = {
                    f"{p['nombre']} [{p.get('categoria', 'General')}] | Disp: {formatear_cantidad(p['cantidad'], p.get('unidad', 'Unidad'))} | Ingreso: {p.get('fecha_registro', 'N/A')} | Vence: {p.get('fecha_vencimiento', 'N/A')}": p["id"]
                    for p in productos_disponibles_mal
                }
                sel_ni_label = st.selectbox(
                    "Selecciona el producto (nuevo ingreso en mal estado) *",
                    list(opciones_ni.keys()),
                    key="sel_prod_nuevo_mal",
                )
                prod_ni = next(p for p in datos["productos"] if p["id"] == opciones_ni[sel_ni_label])
                normalizar_lotes_producto(prod_ni)
                unidad_ni = prod_ni.get("unidad", "Unidad")
                max_ni = float(prod_ni["cantidad"])
                precio_base_ni = float(prod_ni.get("precio_unitario", 0.0))
                paso_ni = 1.0 if max_ni >= 1.0 else max(0.01, round(max_ni / 4, 2))

                with st.form("form_nuevo_ingreso_mal_estado", clear_on_submit=True):
                    c_n1, c_n2, c_n3 = st.columns(3)
                    cant_ni = c_n1.number_input(
                        f"Cantidad en mal estado ({unidad_ni}) *",
                        min_value=0.0,
                        max_value=max_ni,
                        step=paso_ni,
                        value=0.0,
                        format="%.2f",
                    )
                    costo_unit_ni = c_n2.number_input(
                        "Valor / costo unitario de referencia *",
                        min_value=0.0,
                        step=0.5,
                        value=precio_base_ni,
                        format="%.2f",
                    )
                    motivo_ni = c_n3.selectbox(
                        "Motivo del mal estado",
                        [
                            "Nuevo ingreso en mal estado",
                            "Mal estado / golpeado o dañado",
                            "Empaque o lote defectuoso",
                            "Descarte de lote",
                        ],
                    )
                    perdida_ni_calc = round(float(cant_ni) * float(costo_unit_ni), 2)
                    tarjeta_metrica("Pérdida estimada por mal estado", fmt_quetzales(perdida_ni_calc), modo_oscuro=modo_oscuro, paleta=paleta_activa)

                    guardar_ni = st.form_submit_button("Registrar ingreso en mal estado y descontar")
                    if guardar_ni:
                        cant_limpia = round(float(cant_ni), 2)
                        if cant_limpia <= 0:
                            st.warning("La cantidad en mal estado debe ser mayor a 0.")
                        elif float(costo_unit_ni) <= 0:
                            st.warning("El valor / costo unitario es obligatorio y debe ser mayor a 0.")
                        else:
                            perdida_ni_limpia = round(cant_limpia * float(costo_unit_ni), 2)
                            cad_descontada = descontar_de_lotes(prod_ni, cant_limpia)
                            datos["mermas"].append({
                                "id": datetime.now().strftime("%Y%m%d%H%M%S%f"),
                                "fecha": date.today().strftime("%Y-%m-%d"),
                                "hora": datetime.now().strftime("%H:%M"),
                                "tipo_registro": "Nuevo ingreso en mal estado",
                                "producto": prod_ni["nombre"],
                                "cantidad": cant_limpia,
                                "unidad": unidad_ni,
                                "caducidad_lote": cad_descontada,
                                "motivo": motivo_ni,
                                "precio_original": float(costo_unit_ni),
                                "precio_aplicado": 0.0,
                                "descuento_pct": 100.0,
                                "monto_recuperado": 0.0,
                                "perdida_estimada": perdida_ni_limpia,
                            })
                            guardar_datos(datos)
                            st.rerun()
            else:
                st.markdown(
                    """
                    <div class="cuadro-vacio">
                        <p class="cuadro-vacio-titulo">No hay productos en inventario para registrar</p>
                        <p class="cuadro-vacio-desc">Primero registra o compra productos en el sistema.</p>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )

        else:
            productos_buen_estado = [
                p for p in datos["productos"]
                if float(p.get("cantidad", 0.0)) > 0 and p.get("estado") == "Buen estado"
            ]

            if productos_buen_estado:
                opciones_be = {
                    f"{p['nombre']} [{p.get('categoria', 'General')}] | Disp: {formatear_cantidad(p['cantidad'], p.get('unidad', 'Unidad'))} | Vence: {p.get('fecha_vencimiento', 'N/A')} | Precio: {fmt_quetzales(p.get('precio_unitario', 0.0))}": p["id"]
                    for p in productos_buen_estado
                }
                sel_be_label = st.selectbox(
                    "Selecciona el producto en buen estado *",
                    list(opciones_be.keys()),
                    key="sel_prod_buen_estado",
                )
                prod_be = next(p for p in datos["productos"] if p["id"] == opciones_be[sel_be_label])
                normalizar_lotes_producto(prod_be)
                unidad_be = prod_be.get("unidad", "Unidad")
                max_be = float(prod_be["cantidad"])
                precio_base_be = float(prod_be.get("precio_unitario", 0.0))
                paso_be = 1.0 if max_be >= 1.0 else max(0.01, round(max_be / 4, 2))

                accion_be = st.radio(
                    "Cómo deseas dar salida al producto en buen estado",
                    [
                        "Vender a otro precio / con porcentaje de rebaja",
                        "Dar de baja / eliminar sin cobro (o a otro precio)",
                    ],
                    horizontal=True,
                    key="accion_buen_estado",
                )

                modo_rebaja_be = st.radio(
                    "Cómo definir el precio",
                    ["Por porcentaje de descuento (%)", "Escribir precio unitario directo"],
                    horizontal=True,
                    key="modo_rebaja_be",
                )

                with st.form("form_buen_estado_salida", clear_on_submit=True):
                    c_b1, c_b2, c_b3 = st.columns(3)
                    cant_be = c_b1.number_input(
                        f"Cantidad ({unidad_be}) *",
                        min_value=0.0,
                        max_value=max_be,
                        step=paso_be,
                        value=0.0,
                        format="%.2f",
                    )
                    precio_orig_be = c_b2.number_input(
                        "Precio normal unitario *",
                        min_value=0.0,
                        step=0.5,
                        value=precio_base_be,
                        format="%.2f",
                    )
                    if modo_rebaja_be == "Por porcentaje de descuento (%)":
                        pct_be = c_b3.number_input(
                            "Porcentaje quitado / descuento (%)",
                            min_value=0.0,
                            max_value=100.0,
                            step=5.0,
                            value=15.0 if accion_be.startswith("Vender") else 100.0,
                            format="%.1f",
                        )
                        precio_final_be = round(float(precio_orig_be) * (1.0 - float(pct_be) / 100.0), 2)
                    else:
                        precio_final_be = c_b3.number_input(
                            "Precio unitario aplicado *",
                            min_value=0.0,
                            step=0.5,
                            value=max(0.0, round(float(precio_orig_be) * 0.85, 2)) if accion_be.startswith("Vender") else 0.0,
                            format="%.2f",
                        )
                        pct_be = round(((float(precio_orig_be) - float(precio_final_be)) / float(precio_orig_be)) * 100.0, 1) if float(precio_orig_be) > 0 else 0.0

                    cobrado_be = round(float(cant_be) * float(precio_final_be), 2)
                    dif_be = round(max(0.0, float(cant_be) * (float(precio_orig_be) - float(precio_final_be))), 2)

                    bm1, bm2 = st.columns(2)
                    with bm1:
                        tarjeta_metrica(f"Monto a cobrar ({fmt_num(pct_be)}% desc.)", fmt_quetzales(cobrado_be), modo_oscuro=modo_oscuro, paleta=paleta_activa)
                    with bm2:
                        tarjeta_metrica("Diferencia frente al precio normal", fmt_quetzales(dif_be), modo_oscuro=modo_oscuro, paleta=paleta_activa)

                    guardar_be = st.form_submit_button("Registrar salida de producto en buen estado")
                    if guardar_be:
                        cant_limpia = round(float(cant_be), 2)
                        if cant_limpia <= 0:
                            st.warning("La cantidad debe ser mayor a 0.")
                        elif float(precio_orig_be) <= 0:
                            st.warning("El precio normal unitario es obligatorio y debe ser mayor a 0.")
                        elif accion_be.startswith("Vender") and float(precio_final_be) <= 0:
                            st.warning("Para vender a otro precio, el precio aplicado debe ser mayor a 0.")
                        else:
                            cobrado_limpio = round(cant_limpia * float(precio_final_be), 2)
                            dif_limpia = round(max(0.0, cant_limpia * (float(precio_orig_be) - float(precio_final_be))), 2)
                            cad_descontada = descontar_de_lotes(prod_be, cant_limpia)
                            fecha_hoy = date.today().strftime("%Y-%m-%d")
                            hora_hoy = datetime.now().strftime("%H:%M")

                            motivo_be = (
                                f"Buen estado — Rebaja / otro precio ({fmt_num(pct_be)}%)"
                                if accion_be.startswith("Vender")
                                else f"Buen estado — Baja / ajuste ({fmt_num(pct_be)}%)"
                            )

                            datos["mermas"].append({
                                "id": datetime.now().strftime("%Y%m%d%H%M%S%f"),
                                "fecha": fecha_hoy,
                                "hora": hora_hoy,
                                "tipo_registro": "Buen estado (otro precio / baja)",
                                "producto": prod_be["nombre"],
                                "cantidad": cant_limpia,
                                "unidad": unidad_be,
                                "caducidad_lote": cad_descontada,
                                "motivo": motivo_be,
                                "precio_original": float(precio_orig_be),
                                "precio_aplicado": float(precio_final_be),
                                "descuento_pct": float(pct_be),
                                "monto_recuperado": cobrado_limpio,
                                "perdida_estimada": dif_limpia,
                            })

                            if cobrado_limpio > 0:
                                datos["ventas"].append({
                                    "id": datetime.now().strftime("%Y%m%d%H%M%S%f"),
                                    "fecha": fecha_hoy,
                                    "hora": hora_hoy,
                                    "producto": f"{prod_be['nombre']} (Otro precio -{fmt_num(pct_be)}%)",
                                    "cantidad": cant_limpia,
                                    "unidad": unidad_be,
                                    "caducidad_lote": cad_descontada,
                                    "precio_unitario": float(precio_final_be),
                                    "total": cobrado_limpio,
                                })

                            guardar_datos(datos)
                            st.rerun()
            else:
                st.markdown(
                    """
                    <div class="cuadro-vacio">
                        <p class="cuadro-vacio-titulo">No hay productos en buen estado disponibles</p>
                        <p class="cuadro-vacio-desc">Registra o compra productos en buen estado para darles salida a otro precio o de baja.</p>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )

        st.divider()
        st.subheader("Historial de mermas, nuevo ingreso en mal estado y buen estado")
        if datos["mermas"]:
            total_perdida_merma = sum(
                float(m.get("perdida_estimada", 0.0))
                for m in datos["mermas"]
                if m.get("tipo_registro") in ("Merma / Vencido", "Nuevo ingreso en mal estado")
            )
            total_dif_be = sum(
                float(m.get("perdida_estimada", 0.0))
                for m in datos["mermas"]
                if m.get("tipo_registro") == "Buen estado (otro precio / baja)"
            )
            total_cobrado_be = sum(
                float(m.get("monto_recuperado", 0.0))
                for m in datos["mermas"]
                if m.get("tipo_registro") == "Buen estado (otro precio / baja)"
            )

            mm1, mm2, mm3 = st.columns(3)
            with mm1:
                tarjeta_metrica("Pérdida (vencidos / mal estado)", fmt_quetzales(total_perdida_merma), modo_oscuro=modo_oscuro, paleta=paleta_activa)
            with mm2:
                tarjeta_metrica("Cobrado (otro precio)", fmt_quetzales(total_cobrado_be), modo_oscuro=modo_oscuro, paleta=paleta_activa)
            with mm3:
                tarjeta_metrica("Diferencia en otro precio / baja", fmt_quetzales(total_dif_be), modo_oscuro=modo_oscuro, paleta=paleta_activa)

            filas_m = []
            tipos_m = []
            for m in reversed(datos["mermas"]):
                t_reg = m.get("tipo_registro", "Merma / Vencido")
                tipos_m.append(t_reg)
                filas_m.append({
                    "Fecha": m.get("fecha", ""),
                    "Hora": m.get("hora", ""),
                    "Tipo": t_reg,
                    "Producto": m.get("producto", ""),
                    "Cantidad": formatear_cantidad(m.get("cantidad", 0), m.get("unidad", "Unidad")),
                    "Caducidad Lote": m.get("caducidad_lote", "Sin registro"),
                    "Precio Normal": fmt_quetzales(m.get("precio_original", 0.0)),
                    "Precio Aplicado": fmt_quetzales(m.get("precio_aplicado", 0.0)),
                    "Cobrado": fmt_quetzales(m.get("monto_recuperado", 0.0)),
                    "Pérdida / Dif.": fmt_quetzales(m.get("perdida_estimada", 0.0)),
                })
            df_m_vis = pd.DataFrame(filas_m)
            renderizar_tabla_moderna(
                df_m_vis,
                modo_oscuro=modo_oscuro,
                estilos_celdas=estilizar_mermas(df_m_vis, tipos_m, modo_oscuro=modo_oscuro),
                enumerar=True,
            )

            st.divider()
            st.subheader("🔒 Eliminar un registro de merma, mal estado o rebaja")
            clave_merma_req = obtener_clave_seccion(datos, "mermas")
            opciones_merma_elim = {
                f"[{m.get('fecha', '')} {m.get('hora', '')}] {m.get('tipo_registro', 'Merma / Vencido')} — {m.get('producto', '')} ({formatear_cantidad(m.get('cantidad', 0), m.get('unidad', 'Unidad'))} | {fmt_quetzales(m.get('perdida_estimada', 0.0))})": m["id"]
                for m in reversed(datos["mermas"])
            }
            sel_merma_borrar = st.selectbox(
                "Selecciona el registro que deseas eliminar",
                list(opciones_merma_elim.keys()),
                key="sel_merma_borrar",
            )
            preparar_confirmacion_eliminacion("merma_ultimo_sel_borrar", sel_merma_borrar, "conf_merma_borrar", "clave_elim_merma", "merma_limpiar_next")

            confirmar_merma = st.checkbox(
                "Confirmo que deseo eliminar este registro del historial",
                key="conf_merma_borrar",
            )
            clave_merma = st.text_input(
                "Ingresa la contraseña para eliminar el registro *",
                type="password",
                key="clave_elim_merma",
            )

            if st.button("Eliminar registro seleccionado", type="primary", key="btn_elim_merma", disabled=not confirmar_merma):
                if confirmar_merma and clave_merma == clave_merma_req:
                    id_merma_borrar = opciones_merma_elim[sel_merma_borrar]
                    merma_obj = next((m for m in datos["mermas"] if m["id"] == id_merma_borrar), None)
                    if merma_obj:
                        mover_a_papelera(
                            datos,
                            "Merma",
                            merma_obj,
                            f"{merma_obj.get('producto', '')} — {formatear_cantidad(merma_obj.get('cantidad', 0), merma_obj.get('unidad', 'Unidad'))} ({fmt_quetzales(merma_obj.get('perdida_estimada', 0.0))})",
                        )
                    datos["mermas"] = [m for m in datos["mermas"] if m["id"] != id_merma_borrar]
                    guardar_datos(datos)
                    solicitar_limpieza_eliminacion("merma_limpiar_next")
                    st.rerun()
                elif confirmar_merma:
                    st.error("Contraseña incorrecta.")
        else:
            st.session_state["conf_merma_borrar"] = False
            st.session_state["clave_elim_merma"] = ""
            st.session_state["merma_ultimo_sel_borrar"] = ""
            st.session_state["merma_limpiar_next"] = False
            st.markdown(
                """
                <div class="cuadro-vacio">
                    <p class="cuadro-vacio-titulo">Sin registros en este historial todavía</p>
                    <p class="cuadro-vacio-desc">Aquí aparecerán los registros de vencidos/por vencer, nuevo ingreso en mal estado y salidas de buen estado.</p>
                </div>
                """,
                unsafe_allow_html=True,
            )

    # ==========================================================================
    # SECCIÓN 7: NOTIFICACIONES DE VENCIMIENTO Y STOCK BAJO / POR SURTIR
    # ==========================================================================
    elif idx_act == 6:
        st.header(titulos_cfg[6])

        if alertas_activas:
            df_notif = pd.DataFrame(alertas_activas)

            n1, n2, n3 = st.columns(3)
            with n1:
                tarjeta_metrica("Total de notificaciones", str(len(df_notif)), modo_oscuro=modo_oscuro, paleta=paleta_activa)
            with n2:
                tarjeta_metrica(
                    "Urgentes (Vencido / ≤ 4 stock)",
                    str(int(df_notif["Tipo de alerta"].isin(["Vencido", "Vencido / Mal estado", "Poco producto"]).sum())),
                    modo_oscuro=modo_oscuro,
                    paleta=paleta_activa,
                )
            with n3:
                tarjeta_metrica(
                    "Por atender pronto (Por vencer / 5 stock)",
                    str(int(df_notif["Tipo de alerta"].isin(["Próximo a vencer", "Por surtir pronto"]).sum())),
                    modo_oscuro=modo_oscuro,
                    paleta=paleta_activa,
                )

            tipos_disponibles = sorted(df_notif["Tipo de alerta"].unique().tolist())
            tipos_elegidos = st.multiselect(
                "Filtrar notificaciones",
                options=tipos_disponibles,
                default=[],
                placeholder="Todos los tipos de alerta",
                key="filtro_tipos_notif",
            )

            mascara_notif = (
                df_notif["Tipo de alerta"].isin(tipos_elegidos)
                if tipos_elegidos
                else pd.Series([True] * len(df_notif), index=df_notif.index)
            )
            df_notif_filtrada = df_notif[mascara_notif].reset_index(drop=True)
            estilos_notif = df_notif_filtrada["_estilo"].tolist()
            df_notif_visible = df_notif_filtrada.drop(columns=["clave_notif", "_estilo", "_orden"])

            st.subheader("Detalle de alertas activas")
            renderizar_tabla_moderna(
                df_notif_visible,
                modo_oscuro=modo_oscuro,
                estilos_celdas=estilizar_notificaciones(df_notif_visible, estilos_notif, modo_oscuro=modo_oscuro),
                enumerar=True,
            )

            prods_alerta_stock = sorted({
                row["Producto"]
                for _, row in df_notif_filtrada.iterrows()
                if row["Tipo de alerta"] in ("Poco producto", "Por surtir pronto")
            })
            if prods_alerta_stock:
                col_na, col_nb = st.columns([2.5, 1.5])
                sel_na = col_na.selectbox("Reponer producto con alerta", prods_alerta_stock, key="sel_notif_reponer")
                if col_nb.button("Reponer ahora", key="btn_notif_reponer"):
                    st.session_state["comp_sel_producto_existente"] = sel_na
                    st.session_state["comp_firma_seleccion"] = ""
                    st.session_state["pestana_activa"] = 2
                    st.rerun()

            st.divider()
            st.subheader("Lectura detallada por producto")
            for _, item in df_notif_filtrada.iterrows():
                alerta_txt = (
                    f"**{item['Producto']}** ({item['Categoría']}) — **{item['Tipo de alerta']}**: "
                    f"{item['Detalle']}"
                )
                if item["_estilo"] in ("Vencido", "Poco producto"):
                    st.error(alerta_txt)
                else:
                    st.warning(alerta_txt)

            st.divider()
            st.subheader("🔒 Eliminar una notificación manualmente")
            clave_notif_req = obtener_clave_seccion(datos, "notificaciones")

            opciones_notif_elim = {
                f"[{row['Tipo de alerta']}] {row['Producto']} ({row['Cantidad actual']} | Cad: {row['Caducidad']})": row["clave_notif"]
                for _, row in df_notif.iterrows()
            }
            sel_notif_borrar = st.selectbox(
                "Selecciona la notificación que deseas eliminar",
                list(opciones_notif_elim.keys()),
                key="sel_notif_borrar",
            )
            preparar_confirmacion_eliminacion("notif_ultimo_sel_borrar", sel_notif_borrar, "conf_notif_borrar", "clave_elim_notif", "notif_limpiar_next")

            confirmar_notif = st.checkbox(
                "Confirmo que deseo eliminar esta notificación",
                key="conf_notif_borrar",
            )
            clave_notif = st.text_input(
                "Ingresa la contraseña para eliminar la notificación *",
                type="password",
                key="clave_elim_notif",
            )

            if st.button("Eliminar notificación seleccionada", type="primary", key="btn_elim_notif", disabled=not confirmar_notif):
                if confirmar_notif and clave_notif == clave_notif_req:
                    clave_seleccionada = opciones_notif_elim[sel_notif_borrar]
                    if clave_seleccionada not in datos["notificaciones_descartadas"]:
                        datos["notificaciones_descartadas"].append(clave_seleccionada)
                        guardar_datos(datos)
                    solicitar_limpieza_eliminacion("notif_limpiar_next")
                    st.rerun()
                elif confirmar_notif:
                    st.error("Contraseña incorrecta.")
        else:
            st.session_state["conf_notif_borrar"] = False
            st.session_state["clave_elim_notif"] = ""
            st.session_state["notif_ultimo_sel_borrar"] = ""
            st.session_state["notif_limpiar_next"] = False
            st.markdown(
                """
                <div class="cuadro-vacio">
                    <p class="cuadro-vacio-titulo">En este momento no tienes notificaciones pendientes</p>
                    <p class="cuadro-vacio-desc">No hay productos vencidos, próximos a vencer ni con stock bajo por reponer.</p>
                </div>
                """,
                unsafe_allow_html=True,
            )
