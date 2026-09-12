"""
Medcell Despacho - app de Streamlit
=====================================================================
App independiente (repo propio) para planificar el despacho semanal de OC
en camiones de 13/16 pallets, a partir de la pestaña "SB" del Refresh.

Que hace:
- Lee la pestana "SB" del Refresh (Excel de compras/paletizado).
- Filtra por la semana que elijas (ej: 38, 39, 40...).
- Separa PRIMERO las lineas con nombre en "Directos": esas NO ocupan pallets
  ni ventanas de camion, solo quedan como tabla informativa.
- Con lo que queda, clasifica cada OC (columna "Pedido") en 4 categorias de
  prioridad de despacho.
- Arma camiones (13 o 16 pallets, elegidos libremente) sin partir ninguna OC,
  y SIN mezclar nunca Farma con Consumo Masivo en el mismo camion.
- Asigna cada camion a un dia x ventana de la semana.
- Devuelve 3 tablas (resumen de camiones, detalle de OC y Directos informativo)
  listas para mostrar en Streamlit y descargar como Excel.

Supuestos configurables (todos ajustables desde la UI, no hardcodeados):
- Columna base para pallets: "Pallets Pos." o "Pallets posibles" (sin redondear).
- Capacidades de camion disponibles (por defecto 13 y 16).
- Dias habiles de la semana (por defecto Lunes a Viernes).
- Ventanas de despacho por dia (por defecto 4).
- Orden de prioridad de carga (por defecto 1,2,5,3 - ver logica abajo).
"""

import datetime
from datetime import date
import math
import os
from io import BytesIO
import io

import pandas as pd
import streamlit as st
import plotly.graph_objects as go
import plotly.express as px
from plotly.colors import sample_colorscale
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# Colores usados solo por el panel de Stock y Caducidad (BBD STOCK)
COLOR_ROJO = "#FB7185"          # crítico (>90%)
COLOR_AMARILLO = "#FBBF24"      # atención (70-90%)
COLOR_VERDE = "#34D399"         # saludable (<70%)
COLOR_ACENTO_1 = "#818CF8"      # índigo claro
COLOR_ACENTO_2 = "#22C55E"      # verde
COLOR_NEUTRO = "#334155"        # slate oscuro de fondo para escalas
COLOR_CARD_BG = "#161B2C"
COLOR_CARD_BORDER = "#2A2F45"
COLOR_TEXT_MUTED = "#94A3B8"
COLOR_GRID = "rgba(148, 163, 184, 0.12)"   # grilla sutil

st.set_page_config(
    page_title="Medcell Despacho",
    page_icon="🚚",
    layout="wide",
)

# --------------------------------------------------------------------------
# 0. IDENTIDAD VISUAL - Medcell Despacho
# --------------------------------------------------------------------------
# Paleta: azul profundo (confianza, logistica) + verde solo reservado para el
# estado "Facturado" (para no chocar semanticamente con el resto de la app).
# El rojo queda reservado para alertas (sin ventanas suficientes, etc).

_BRAND_CSS = """
<style>
/* Fuerza tema oscuro siempre, sin importar la preferencia del navegador o
   el toggle de tema de quien abre la app (Streamlit permite claro/oscuro
   por sesion; esto lo anula visualmente). */
html, body, [data-testid="stAppViewContainer"], [data-testid="stApp"],
[data-testid="stHeader"], [data-testid="stSidebar"], [data-testid="stMain"],
.main, .block-container {
    background-color: #0B1120 !important;
    color: #F5F7FA !important;
}
[data-testid="stSidebar"] { background-color: #0F1626 !important; }
[data-testid="stHeader"] { background: transparent !important; }
p, span, label, li, div, h1, h2, h3, h4, h5, h6 { color: #F5F7FA; }
.stCaption, [data-testid="stCaptionContainer"] { color: #93A2B8 !important; }
[data-testid="stExpander"] {
    background-color: #141B2D !important;
    border: 1px solid #232E45 !important;
    border-radius: 10px !important;
}
[data-testid="stDataFrame"], [data-testid="stTable"] {
    background-color: #141B2D !important;
}

.medcell-header {
    background: linear-gradient(135deg, #0B4F86 0%, #12294A 100%);
    padding: 1.4rem 1.8rem;
    border-radius: 10px;
    margin-bottom: 1.4rem;
}
.medcell-header h1 {
    color: #FFFFFF;
    font-size: 1.9rem;
    font-weight: 800;
    margin: 0;
    letter-spacing: -0.01em;
    text-transform: uppercase;
}
.medcell-header h1 .brand-accent {
    color: #8ECFFF;
}
.medcell-header .dev {
    color: #7A93AC;
    font-size: 0.78rem;
    margin: 0.15rem 0 0 0;
}
.medcell-header p {
    color: #BFD9F2;
    margin: 0.25rem 0 0 0;
    font-size: 0.95rem;
}
.medcell-header .tag {
    display: inline-block;
    background: #1DB980;
    color: #06251A;
    font-size: 0.72rem;
    font-weight: 700;
    padding: 0.15rem 0.55rem;
    border-radius: 999px;
    margin-left: 0.6rem;
    vertical-align: middle;
}

/* Tarjetas KPI oscuras, estilo Medcell Almacenamiento */
.kpi-card {
    background: #141B2D;
    border: 1px solid #232E45;
    border-radius: 12px;
    padding: 1.1rem 1.2rem;
    text-align: center;
    height: 132px;
    display: flex;
    flex-direction: column;
    align-items: center;
    justify-content: center;
    box-sizing: border-box;
}
.kpi-card .kpi-value {
    font-size: 1.9rem;
    font-weight: 800;
    color: #FFFFFF;
    line-height: 1.1;
}
.kpi-card .kpi-label {
    color: #93A2B8;
    font-size: 0.82rem;
    margin-top: 0.35rem;
}
.kpi-card .kpi-badge {
    display: inline-block;
    margin-top: 0.5rem;
    font-size: 0.68rem;
    font-weight: 700;
    padding: 0.15rem 0.6rem;
    border-radius: 999px;
}
</style>
"""


def _header():
    st.markdown(_BRAND_CSS, unsafe_allow_html=True)
    st.markdown(
        """
        <div class="medcell-header">
            <h1>🚚 MEDCELL <span class="brand-accent">DESPACHO</span></h1>
            <div class="dev">Desarrollado por Sebastián Alexis Pérez López</div>
            <p>Plan de despacho semanal: distribuye las OC en camiones de 13/16 pallets,
            respetando Farma / Consumo Masivo, Directos y Facturados.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_kpi_cards(cards: list[dict]):
    """Fila de tarjetas KPI oscuras, todas del mismo tamaño. Cada card:
    {value, label, badge_text (opcional), badge_color (opcional, hex)}."""
    cols = st.columns(len(cards))
    for col, card in zip(cols, cards):
        badge_html = ""
        if card.get("badge_text"):
            color = card.get("badge_color", "#3B9EFF")
            badge_html = (
                f"<div class='kpi-badge' style='background:{color}26;color:{color};'>"
                f"{card['badge_text']}</div>"
            )
        html = (
            f"<div class='kpi-card'>"
            f"<div class='kpi-value'>{card['value']}</div>"
            f"<div class='kpi-label'>{card['label']}</div>"
            f"{badge_html}"
            f"</div>"
        )
        with col:
            st.markdown(html, unsafe_allow_html=True)

# --------------------------------------------------------------------------
# 1. LECTURA DE DATOS
# --------------------------------------------------------------------------

# La fila 4 del Excel (indice 3, 0-based) trae los encabezados reales.
HEADER_ROW = 3

COLS_NECESARIAS = [
    "Semana", "OC", "Pedido", "Fecha vence", "Descripción",
    "Solicitado", "1 Posible", "Pronto-vence",
    "Pallets Pos.", "Pallets posibles", "Directos",
]


@st.cache_data(show_spinner="Leyendo pestaña SB...")
def leer_sb(archivo) -> pd.DataFrame:
    """Lee la pestaña SB del Refresh y devuelve un DataFrame limpio."""
    df = pd.read_excel(archivo, sheet_name="SB", header=HEADER_ROW)
    # Normaliza nombres de columnas (a veces vienen con espacios extra)
    df.columns = [str(c).strip() for c in df.columns]
    return df


# --------------------------------------------------------------------------
# 1b. FACTURADOS
# --------------------------------------------------------------------------
# Se detecta automaticamente desde la pestana "OC" del MISMO Refresh:
#   "Pedido de Venta" = mismo numero que "Pedido" en SB.
#   "Pendiente" = 0 (sumado por Pedido de Venta) -> ya se despacho/facturo.
# No requiere mantener ningun archivo aparte.


@st.cache_data(show_spinner="Revisando pestaña OC del Refresh...")
def cargar_facturados_desde_refresh(archivo) -> set:
    """Detecta automaticamente los Pedido ya 100% despachados/facturados
    usando la pestana 'OC' del mismo Refresh (columnas 'Pedido de Venta' y
    'Pendiente')."""
    try:
        archivo.seek(0)
    except Exception:
        pass
    try:
        df_oc = pd.read_excel(archivo, sheet_name="OC")
    except Exception:
        return set()
    finally:
        try:
            archivo.seek(0)
        except Exception:
            pass

    df_oc.columns = [str(c).strip() for c in df_oc.columns]
    if "Pedido de Venta" not in df_oc.columns or "Pendiente" not in df_oc.columns:
        return set()

    pendiente_total = df_oc.groupby("Pedido de Venta")["Pendiente"].sum()
    ya_facturados = pendiente_total[pendiente_total <= 0].index
    return {str(int(x)) for x in ya_facturados if pd.notna(x)}


# --------------------------------------------------------------------------
# 2. CLASIFICACION DE PRIORIDAD POR OC
# --------------------------------------------------------------------------

def _clasificar(row) -> tuple[int, str]:
    # Los Directos ya fueron separados antes de llegar aca: aqui solo quedan
    # las lineas que SI van por camion.
    sol, pos1, pv = row["sol"], row["pos1"], row["pv"]
    if pos1 == sol:
        return 1, "1 - Solicitado = 1er Posible (completo)"
    if (sol - pos1) <= pv:
        return 2, "2 - Diferencia cubierta por Pronto-vence"
    if pos1 == 0:
        return 3, "3 - Sin 1er Posible (sin stock disponible)"
    return 5, "5 - Otros / parcial sin cobertura"


def separar_directos(df: pd.DataFrame, semana: int, pallet_col: str):
    """Separa las lineas con nombre en 'Directos' ANTES de armar camiones:
    esas lineas no ocupan espacio en ningun camion, solo se listan como info.
    Devuelve (df_camion, df_directos_info)."""
    d = df[df["Semana"] == semana].copy()
    d["directos_flag"] = d["Directos"].apply(
        lambda x: str(x).strip() not in ("", "-", "nan", "None")
    )
    df_directos = d[d["directos_flag"]].copy()
    df_camion = d[~d["directos_flag"]].copy()
    return df_camion, df_directos


def resumen_directos(df_directos: pd.DataFrame) -> pd.DataFrame:
    """Tabla informativa de lineas Directas (no van en camion)."""
    if df_directos.empty:
        return df_directos
    cols = ["Pedido", "OC", "Fecha vence", "SKU SB", "Descripción",
            "Directos", "Solicitado"]
    cols = [c for c in cols if c in df_directos.columns]
    return df_directos[cols].rename(columns={"Directos": "Proveedor directo"})


def agrupar_por_oc(df_camion: pd.DataFrame, pallet_col: str) -> pd.DataFrame:
    """Agrupa por Pedido (=OC) SOLO las lineas que van por camion (sin Directos)
    y clasifica prioridad."""
    if df_camion.empty:
        return df_camion

    agg = df_camion.groupby("Pedido").agg(
        oc=("OC", "first"),
        fecha_vence=("Fecha vence", "first"),
        division=("Division", "first"),
        sol=("Solicitado", "sum"),
        pos1=("1 Posible", "sum"),
        pv=("Pronto-vence", "sum"),
        pallets=(pallet_col, "sum"),
        # "Posible actual $" se pone en $0 apenas la OC queda 100% despachada
        # (ya no hay nada "posible" pendiente), asi que para reflejar el
        # valor real de la OC (y cuanto se despacho de verdad) usamos
        # "Solicitado $", que no se resetea a 0.
        monto=("Solicitado $", "sum"),
        n_sku=("Pedido", "count"),
    ).reset_index()

    pr = agg.apply(_clasificar, axis=1, result_type="expand")
    agg["prioridad"], agg["prioridad_label"] = pr[0], pr[1]
    # Los pallets de una OC siempre se redondean HACIA ARRIBA (nunca decimales):
    # una OC que ocupa 5.4 o 5.6 reserva igual 6 posiciones de pallet.
    agg["pallets"] = agg["pallets"].apply(lambda v: math.ceil(round(v, 6)))
    agg["monto"] = agg["monto"].round(0)
    # OC con 0 pallets (linea unica que quedo en 0 tras excluir directos) no
    # necesita camion; se deja fuera del bin-packing.
    agg = agg[agg["pallets"] > 0].reset_index(drop=True)
    return agg


# --------------------------------------------------------------------------
# 3. ARMADO DE CAMIONES (bin packing, OC nunca se parte)
# --------------------------------------------------------------------------

def armar_camiones(agg: pd.DataFrame, capacidades: list[int],
                    orden_prioridad: list[int]) -> list[dict]:
    """Arma camiones respetando 2 reglas duras:
    1) una OC nunca se parte entre camiones.
    2) Farma y Consumo Masivo NUNCA van en el mismo camion -> se empacan por
       separado (una division no le "presta" espacio a la otra) y despues se
       intercalan segun prioridad para asignar las ventanas."""
    cap_max = max(capacidades)

    def cerrar(b):
        for cap in sorted(capacidades):
            if b["total"] <= cap:
                b["camion"] = cap
                return b
        b["camion"] = cap_max
        return b

    def empacar(items):
        items = sorted(items, key=lambda x: (orden_prioridad.index(x["prioridad"]), -x["pallets"]))
        bins_local, actual = [], {"items": [], "total": 0.0}
        for it in items:
            if it["pallets"] > cap_max:
                if actual["items"]:
                    bins_local.append(cerrar(actual))
                    actual = {"items": [], "total": 0.0}
                bins_local.append(cerrar({"items": [it], "total": it["pallets"]}))
                continue
            if actual["total"] + it["pallets"] <= cap_max:
                actual["items"].append(it)
                actual["total"] += it["pallets"]
            else:
                bins_local.append(cerrar(actual))
                actual = {"items": [it], "total": it["pallets"]}
        if actual["items"]:
            bins_local.append(cerrar(actual))
        return bins_local

    agg = agg.copy()
    agg["division"] = agg["division"].fillna("Sin división")
    bins_all = []
    for division, sub in agg.groupby("division"):
        items = sub.to_dict("records")
        for it in items:
            it["division"] = division
        bins_dv = empacar(items)
        for b in bins_dv:
            b["division"] = division
        bins_all.extend(bins_dv)

    # Se intercalan los camiones de todas las divisiones segun prioridad, para
    # que las ventanas mas tempranas de la semana las tomen las OC mas
    # urgentes sin importar de que division sean.
    bins_all.sort(key=lambda b: min(orden_prioridad.index(it["prioridad"]) for it in b["items"]))
    return bins_all


def asignar_ventanas(bins: list[dict], semana: int, anio: int,
                      dias: list[str], ventanas_por_dia: int):
    lunes = datetime.date.fromisocalendar(anio, semana, 1)
    dia_offset = {"Lunes": 0, "Martes": 1, "Miercoles": 2, "Miércoles": 2,
                  "Jueves": 3, "Viernes": 4, "Sabado": 5, "Sábado": 5, "Domingo": 6}
    slots = []
    for d in dias:
        fecha = lunes + datetime.timedelta(days=dia_offset.get(d, 0))
        for v in range(1, ventanas_por_dia + 1):
            slots.append({"dia": d, "fecha": fecha, "ventana": v})

    overflow = len(bins) > len(slots)
    for i, b in enumerate(bins):
        slot = slots[i] if i < len(slots) else {"dia": "SIN VENTANA", "fecha": None, "ventana": "-"}
        b.update(slot)
        b["camion_num"] = i + 1
        for it in b["items"]:
            it["camion_num"] = i + 1
            it["dia"] = slot["dia"]
            it["ventana"] = slot["ventana"]
            it["fecha"] = slot["fecha"]
    return bins, slots, overflow


# --------------------------------------------------------------------------
# 4. ORQUESTADOR: de DataFrame crudo a las 2 tablas finales
# --------------------------------------------------------------------------

def generar_plan(df: pd.DataFrame, semana: int, anio: int, pallet_col: str,
                  capacidades: list[int], dias: list[str], ventanas_por_dia: int,
                  orden_prioridad: list[int], facturados: set | None = None):
    facturados = facturados or set()
    df_camion, df_directos = separar_directos(df, semana, pallet_col)
    tabla_directos = resumen_directos(df_directos)
    agg = agrupar_por_oc(df_camion, pallet_col)
    if agg.empty:
        return None, None, None, tabla_directos

    bins = armar_camiones(agg, capacidades, orden_prioridad)
    bins, slots, overflow = asignar_ventanas(bins, semana, anio, dias, ventanas_por_dia)

    resumen = pd.DataFrame([{
        "Camión #": b["camion_num"], "Día": b["dia"], "Fecha": b["fecha"],
        "Ventana": b["ventana"], "División": b["division"],
        "Tipo camión (pallets)": b["camion"],
        "Pallets cargados": round(b["total"], 2), "Capacidad": b["camion"],
        "Utilización %": round(b["total"] / b["camion"] * 100, 1),
        "# OCs": len(b["items"]),
        "Pedidos incluidos": ", ".join(str(it["Pedido"]) for it in b["items"]),
    } for b in bins])

    detalle_rows = []
    for b in bins:
        for it in b["items"]:
            es_facturado = str(it["Pedido"]).strip() in facturados
            detalle_rows.append({
                "Pedido (OC)": it["Pedido"], "OC": it["oc"], "Fecha vence": it["fecha_vence"],
                "División": it["division"],
                "Prioridad": it["prioridad"], "Descripción prioridad": it["prioridad_label"],
                "Pallets": it["pallets"], "Monto": it.get("monto", 0), "# SKUs": it["n_sku"],
                "Camión #": it["camion_num"], "Día": it["dia"], "Fecha": it.get("fecha"),
                "Ventana": it["ventana"], "Facturado": "Sí" if es_facturado else "No",
            })
    detalle = pd.DataFrame(detalle_rows)

    info = {"camiones": len(bins), "ventanas_disponibles": len(slots), "overflow": overflow,
            "oc_directos": tabla_directos["Pedido"].nunique() if not tabla_directos.empty else 0,
            "lineas_directos": len(tabla_directos),
            "oc_facturadas": int((detalle["Facturado"] == "Sí").sum())}
    return resumen, detalle, info, tabla_directos


def _resaltar_facturado(row):
    style = "background-color: #C6EFCE; color: #0b3d24" if row.get("Facturado") == "Sí" else ""
    return [style] * len(row)


def formato_clp(valor) -> str:
    try:
        return "$" + f"{int(round(float(valor))):,}".replace(",", ".")
    except Exception:
        return "$0"


_COLOR_PRIORIDAD = {1: "#1DB980", 2: "#0B4F86", 5: "#F2994A", 3: "#E4572E"}
_LABEL_PRIORIDAD = {
    1: "Completo (Solicitado = 1er Posible)",
    2: "Cubierto con Pronto-vence",
    5: "Parcial, sin cobertura",
    3: "Sin 1er Posible (sin stock)",
}
_ORDEN_DIAS = ["Lunes", "Martes", "Miercoles", "Miércoles", "Jueves", "Viernes", "Sabado", "Sábado"]


def _semaforo_utilizacion(pct: float) -> str:
    if pct >= 90:
        return "#1DB980"  # verde
    if pct >= 70:
        return "#F2994A"  # amarillo/naranjo
    return "#E4572E"      # rojo


def render_leyenda_calendario():
    """Leyenda de colores de las tarjetas del calendario (semaforo de prioridad
    + chip de Facturado), para que se entienda de un vistazo que significa
    cada color."""
    items_html = "".join(
        f"<div style='display:flex;align-items:center;gap:0.4rem;margin-right:1.1rem;'>"
        f"<span style='width:10px;height:10px;border-radius:3px;background:{color};"
        f"display:inline-block;'></span>"
        f"<span style='font-size:0.74rem;color:#C7D2E0;'>{_LABEL_PRIORIDAD[p]}</span></div>"
        for p, color in _COLOR_PRIORIDAD.items()
    )
    leyenda_html = (
        "<div style='display:flex;flex-wrap:wrap;align-items:center;"
        "background:#141B2D;border:1px solid #232E45;border-radius:8px;"
        "padding:0.55rem 0.8rem;margin-bottom:0.9rem;'>"
        "<span style='font-size:0.74rem;color:#8494AC;font-weight:600;"
        "margin-right:1rem;'>Colores de las tarjetas:</span>"
        f"{items_html}"
        "<div style='display:flex;align-items:center;gap:0.4rem;'>"
        "<span style='background:#C6EFCE;color:#0b3d24;font-size:0.6rem;"
        "font-weight:700;padding:0.05rem 0.4rem;border-radius:999px;'>FACTURADO</span>"
        "<span style='font-size:0.74rem;color:#C7D2E0;'>= ya despachado según el Refresh</span>"
        "</div>"
        "</div>"
    )
    st.markdown(leyenda_html, unsafe_allow_html=True)


def render_calendario(detalle: pd.DataFrame, resumen: pd.DataFrame | None = None):
    """Vista tipo calendario/kanban: una columna por dia, con un KPI de
    despacho arriba (camiones, utilizacion y Facturados vs No) y tarjetas por
    OC con Pedido, OC, Monto y Pallets."""
    if detalle.empty:
        st.info("No hay OC para mostrar en el calendario.")
        return

    render_leyenda_calendario()

    dias_presentes = [d for d in _ORDEN_DIAS if d in detalle["Día"].unique()]
    cols = st.columns(len(dias_presentes)) if dias_presentes else []

    for col, dia in zip(cols, dias_presentes):
        sub = detalle[detalle["Día"] == dia]
        fecha = sub["Fecha"].iloc[0] if "Fecha" in sub.columns and len(sub) else None
        fecha_str = fecha.strftime("%d-%b") if hasattr(fecha, "strftime") else ""

        # KPI de despacho del dia: camiones y utilizacion promedio (semaforo)
        n_camiones_dia, util_prom, color_kpi = 0, 0.0, "#9CA3AF"
        if resumen is not None and not resumen.empty and "Día" in resumen.columns:
            r_dia = resumen[resumen["Día"] == dia]
            if not r_dia.empty:
                n_camiones_dia = len(r_dia)
                util_prom = r_dia["Utilización %"].mean()
                color_kpi = _semaforo_utilizacion(util_prom)

        # KPI Facturados vs No Facturados del dia
        n_facturados = int((sub["Facturado"] == "Sí").sum())
        n_no_facturados = len(sub) - n_facturados
        pct_facturado = (n_facturados / len(sub) * 100) if len(sub) else 0

        with col:
            st.markdown(
                f"<div style='font-weight:700;font-size:0.95rem;color:#F5F7FA;'>{dia}</div>"
                f"<div style='color:#8494AC;font-size:0.78rem;margin-bottom:0.45rem;'>"
                f"{fecha_str} · {len(sub)} OC</div>"
                f"<div style='background:{color_kpi}1A;border:1px solid {color_kpi};"
                f"border-radius:8px;padding:0.4rem 0.6rem;margin-bottom:0.4rem;'>"
                f"<div style='font-size:0.68rem;color:#C7D2E0;font-weight:600;'>"
                f"🚚 {n_camiones_dia} camión(es)</div>"
                f"<div style='font-size:0.68rem;color:{color_kpi};font-weight:700;'>"
                f"{util_prom:.0f}% utilización promedio</div>"
                f"</div>"
                f"<div style='background:#141B2D;border:1px solid #232E45;border-radius:8px;"
                f"padding:0.4rem 0.6rem;margin-bottom:0.6rem;'>"
                f"<div style='font-size:0.68rem;color:#C7D2E0;font-weight:600;"
                f"margin-bottom:0.25rem;'>Facturados vs No</div>"
                f"<div style='display:flex;width:100%;height:8px;border-radius:4px;"
                f"overflow:hidden;background:#E4572E33;margin-bottom:0.25rem;'>"
                f"<div style='width:{pct_facturado:.0f}%;background:#1DB980;'></div>"
                f"</div>"
                f"<div style='display:flex;justify-content:space-between;font-size:0.65rem;'>"
                f"<span style='color:#1DB980;font-weight:700;'>✅ {n_facturados} facturadas</span>"
                f"<span style='color:#E4572E;font-weight:700;'>⏳ {n_no_facturados} pendientes</span>"
                f"</div>"
                f"</div>",
                unsafe_allow_html=True,
            )
            with st.container(height=600):
                for ventana, sub_v in sub.groupby("Ventana"):
                    total_pallets = sub_v["Pallets"].sum()
                    division_v = sub_v["División"].iloc[0] if "División" in sub_v.columns else ""
                    div_color = "#C084FC" if division_v == "FARMA" else "#FBBF24"
                    div_icono = "🧪" if division_v == "FARMA" else "🛒"
                    div_label = "FARMA" if division_v == "FARMA" else "CONSUMO"
                    div_badge = (
                        f"<span style='background:{div_color}26;color:{div_color};"
                        "font-size:0.6rem;font-weight:700;padding:0.1rem 0.5rem;"
                        "border-radius:999px;white-space:nowrap;display:inline-block;"
                        f"margin-top:0.25rem;'>{div_icono} {div_label}</span>"
                    ) if division_v else ""
                    ventana_html = (
                        "<div style='margin:0.5rem 0 0.35rem;'>"
                        "<div style='font-size:0.68rem;font-weight:700;color:#0B4F86;"
                        "letter-spacing:0.03em;'>"
                        f"VENTANA {ventana} · {total_pallets:.0f} pal"
                        "</div>"
                        f"{div_badge}"
                        "</div>"
                    )
                    st.markdown(ventana_html, unsafe_allow_html=True)
                    for _, row in sub_v.sort_values("Prioridad").iterrows():
                        color = _COLOR_PRIORIDAD.get(row["Prioridad"], "#888888")
                        chip = (
                            "<span style='background:#C6EFCE;color:#0b3d24;font-size:0.6rem;"
                            "font-weight:700;padding:0.05rem 0.4rem;border-radius:999px;"
                            "margin-left:0.4rem;'>FACTURADO</span>"
                        ) if row["Facturado"] == "Sí" else ""
                        tarjeta_html = (
                            f"<div style='background:#141B2D;border-left:4px solid {color};"
                            "border-radius:6px;padding:0.5rem 0.7rem;margin-bottom:0.5rem;"
                            "box-shadow:0 1px 2px rgba(0,0,0,0.2);'>"
                            "<div style='font-weight:700;font-size:0.82rem;color:#F5F7FA;'>"
                            f"Pedido {row['Pedido (OC)']}{chip}"
                            "</div>"
                            f"<div style='font-size:0.72rem;color:#8494AC;'>OC {row['OC']}</div>"
                            "<div style='display:flex;justify-content:space-between;"
                            "margin-top:0.3rem;font-size:0.76rem;color:#C7D2E0;'>"
                            f"<span>{formato_clp(row['Monto'])}</span>"
                            f"<span style='color:#3B9EFF;font-weight:600;'>{row['Pallets']:.0f} pal</span>"
                            "</div>"
                            "</div>"
                        )
                        st.markdown(tarjeta_html, unsafe_allow_html=True)


def render_tabla_camiones(resumen: pd.DataFrame, detalle: pd.DataFrame):
    """Tabla 'Plan de camiones' en HTML (para poder pintar en verde, dentro
    de la misma celda, los numeros de Pedido que ya estan Facturados) mas
    una columna extra de % Facturado por camion."""
    facturado_map = dict(zip(detalle["Pedido (OC)"].astype(str), detalle["Facturado"]))

    cols_base = ["Camión #", "Día", "Fecha", "Ventana", "División",
                 "Tipo camión (pallets)", "Pallets cargados", "Capacidad",
                 "Utilización %"]
    header_html = "".join(f"<th>{c}</th>" for c in cols_base) + \
        "<th>% Facturado</th><th>Pedidos incluidos</th>"

    filas_html = []
    for _, row in resumen.iterrows():
        pedidos = [p.strip() for p in str(row["Pedidos incluidos"]).split(",") if p.strip()]
        n_fact = sum(1 for p in pedidos if facturado_map.get(p) == "Sí")
        pct_fact = (n_fact / len(pedidos) * 100) if pedidos else 0
        pedidos_html = ", ".join(
            f"<span style='background:#C6EFCE;color:#0b3d24;font-weight:700;"
            f"border-radius:4px;padding:0 0.3rem;'>{p}</span>"
            if facturado_map.get(p) == "Sí" else f"<span>{p}</span>"
            for p in pedidos
        )
        celdas = "".join(f"<td>{row[c]}</td>" for c in cols_base[:5])
        celdas += (
            f"<td style='text-align:right;'>{row['Tipo camión (pallets)']:.0f}</td>"
            f"<td style='text-align:right;'>{row['Pallets cargados']:.0f}</td>"
            f"<td style='text-align:right;'>{row['Capacidad']:.0f}</td>"
            f"<td style='text-align:right;'>{row['Utilización %']:.1f}%</td>"
            f"<td style='text-align:right;color:#1DB980;font-weight:700;'>{pct_fact:.0f}%</td>"
            f"<td>{pedidos_html}</td>"
        )
        filas_html.append(f"<tr>{celdas}</tr>")

    tabla_html = (
        "<div style='overflow-x:auto;border:1px solid #232E45;border-radius:8px;'>"
        "<table style='border-collapse:collapse;width:100%;font-size:0.82rem;'>"
        "<thead>"
        f"<tr style='background:#0B4F86;color:#fff;text-align:left;'>{header_html}</tr>"
        "</thead>"
        f"<tbody>{''.join(filas_html)}</tbody>"
        "</table>"
        "</div>"
        "<style>"
        "table td, table th { padding:0.45rem 0.6rem; border-bottom:1px solid #232E45; "
        "white-space:nowrap; color:#F5F7FA; }"
        "table tbody tr:nth-child(even) { background:#0F1626; }"
        "table td:last-child { white-space:normal; }"
        "</style>"
    )
    st.markdown(tabla_html, unsafe_allow_html=True)


def _formatear_hoja_detalle(ws, df: pd.DataFrame):
    """Ajusta ancho de columnas al contenido, fechas cortas (dd-mm-aaaa) y
    Monto con separador de miles, para que no salga '####' ni números pegados."""
    from openpyxl.utils import get_column_letter

    col_fecha = {c for c in df.columns if "fecha" in c.lower()}
    col_monto = {c for c in df.columns if c.lower() == "monto"}

    for idx, col in enumerate(df.columns, start=1):
        letra = get_column_letter(idx)
        if col in col_fecha:
            for r in range(2, len(df) + 2):
                ws.cell(row=r, column=idx).number_format = "dd-mm-yyyy"
            ws.column_dimensions[letra].width = 13
        elif col in col_monto:
            for r in range(2, len(df) + 2):
                ws.cell(row=r, column=idx).number_format = "#,##0"
            ws.column_dimensions[letra].width = 14
        else:
            largo = max(
                [len(str(col))] + [len(str(v)) for v in df[col].astype(str)]
            ) if len(df) else len(str(col))
            ws.column_dimensions[letra].width = min(max(largo + 2, 10), 45)


def exportar_pendientes_excel(detalle: pd.DataFrame) -> bytes | None:
    """Excel SOLO con las OC que TODAVIA NO se facturan (Facturado = No, las
    filas blancas de la tabla): Hoja 'Resumen' = tabla pivote División (filas)
    x Día de despacho (columnas), igual estilo a la tabla de referencia
    (categorías al costado, arriba); luego una pestaña por día con el
    detalle completo."""
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

    pendientes = detalle[detalle["Facturado"] == "No"].copy()
    if pendientes.empty:
        return None

    orden_dias_cols = [d for d in _ORDEN_DIAS if d in pendientes["Día"].unique()]
    if "SIN VENTANA" in pendientes["Día"].unique():
        orden_dias_cols.append("SIN VENTANA")

    pivote = pd.pivot_table(
        pendientes, index="División", columns="Día",
        values="Pedido (OC)", aggfunc="count", fill_value=0,
    )
    pivote = pivote.reindex(columns=orden_dias_cols, fill_value=0)
    pivote["Total general"] = pivote.sum(axis=1)
    pivote.loc["Total general"] = pivote.sum(axis=0)

    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        pivote.to_excel(writer, sheet_name="Resumen", index=True)

        for dia in orden_dias_cols:
            sub = pendientes[pendientes["Día"] == dia].drop(columns=["Día"], errors="ignore")
            nombre_hoja = dia[:31]
            sub.to_excel(writer, sheet_name=nombre_hoja, index=False)
            _formatear_hoja_detalle(writer.sheets[nombre_hoja], sub)

        # --- Estilo hoja Resumen: encabezado azul Medcell + columna de
        # categorias resaltada, igual estructura que la tabla de referencia.
        ws = writer.sheets["Resumen"]
        header_fill = PatternFill("solid", fgColor="0B4F86")
        header_font = Font(bold=True, color="FFFFFF")
        cat_fill = PatternFill("solid", fgColor="E7F5EC")
        total_fill = PatternFill("solid", fgColor="EAF0F7")
        bold = Font(bold=True)
        thin = Side(style="thin", color="D9E2EC")
        border = Border(left=thin, right=thin, top=thin, bottom=thin)

        n_rows, n_cols = pivote.shape
        for c in range(1, n_cols + 2):
            cell = ws.cell(row=1, column=c)
            cell.fill, cell.font = header_fill, header_font
            cell.alignment = Alignment(horizontal="center")
            cell.border = border
        for r in range(2, n_rows + 2):
            cell = ws.cell(row=r, column=1)
            cell.fill, cell.font = cat_fill, bold
            cell.border = border
            for c in range(2, n_cols + 2):
                ws.cell(row=r, column=c).border = border
        for c in range(1, n_cols + 2):  # fila de totales
            cell = ws.cell(row=n_rows + 1, column=c)
            cell.font, cell.fill = bold, total_fill
        for r in range(1, n_rows + 2):  # columna de totales
            cell = ws.cell(row=r, column=n_cols + 1)
            cell.font, cell.fill = bold, total_fill
        for col in ws.columns:
            length = max(len(str(c.value)) if c.value is not None else 0 for c in col)
            ws.column_dimensions[col[0].column_letter].width = max(12, length + 2)

    return buf.getvalue()


def exportar_excel(resumen: pd.DataFrame, detalle: pd.DataFrame,
                    tabla_directos: pd.DataFrame) -> bytes:
    from openpyxl.styles import PatternFill

    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        resumen.to_excel(writer, sheet_name="Plan Despacho", index=False)
        detalle.to_excel(writer, sheet_name="Detalle Pedidos", index=False)
        if tabla_directos is not None and not tabla_directos.empty:
            tabla_directos.to_excel(writer, sheet_name="Directos (info, sin camion)", index=False)

        if "Facturado" in detalle.columns:
            ws = writer.sheets["Detalle Pedidos"]
            verde = PatternFill("solid", fgColor="C6EFCE")
            col_idx = detalle.columns.get_loc("Facturado") + 1  # numero de columna Facturado
            for r, val in enumerate(detalle["Facturado"], start=2):  # fila 1 = encabezado
                if val == "Sí":
                    for c in range(1, len(detalle.columns) + 1):
                        ws.cell(row=r, column=c).fill = verde

        # Formato numerico limpio (2 decimales) en vez del "general" de Excel
        if "Pallets" in detalle.columns:
            ws = writer.sheets["Detalle Pedidos"]
            col_pallets = detalle.columns.get_loc("Pallets") + 1
            for r in range(2, len(detalle) + 2):
                ws.cell(row=r, column=col_pallets).number_format = "0"
        if "Pallets cargados" in resumen.columns:
            ws = writer.sheets["Plan Despacho"]
            col_pallets = resumen.columns.get_loc("Pallets cargados") + 1
            col_util = resumen.columns.get_loc("Utilización %") + 1
            for r in range(2, len(resumen) + 2):
                ws.cell(row=r, column=col_pallets).number_format = "0"
                ws.cell(row=r, column=col_util).number_format = "0.0"
    return buf.getvalue()


# --------------------------------------------------------------------------
# 5. PAGINA DE STREAMLIT
# --------------------------------------------------------------------------

def _bytes_archivo_original(archivo) -> bytes | None:
    """Devuelve los bytes del Refresh que se esta usando (subido a mano o el
    data/Refresh.xlsx del repo), para poder ofrecerlo como descarga."""
    try:
        if hasattr(archivo, "getvalue"):
            return archivo.getvalue()
        if isinstance(archivo, str) and os.path.exists(archivo):
            with open(archivo, "rb") as f:
                return f.read()
    except Exception:
        pass
    return None



# --------------------------------------------------------------------------
# Panel de Stock y Caducidad (BBD STOCK) - reusa la misma logica y diseno
# que el dashboard de Stock y Caducidad de Medcell Almacenamiento, pero
# alimentado por la hoja "BBD STOCK" del MISMO Refresh que ya se usa aqui.
# --------------------------------------------------------------------------

def formato_unidades(valor):
    try:
        val_int = int(round(valor))
        return f"{val_int:,}".replace(",", ".")
    except (ValueError, TypeError):
        return "0"


def fmt_code(val):
    if pd.isna(val) or val == "" or val is None or str(val).lower() == "nan":
        return "S/N"
    val_str = str(val).strip()
    if val_str.endswith(".0"):
        val_str = val_str[:-2]
    return val_str


def limpiar_numero(val):
    if pd.isna(val) or val == "" or val is None or str(val).lower() == "nan":
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    val_str = str(val).strip().replace("$", "").replace(" ", "")
    if not val_str:
        return 0.0

    if "." in val_str and "," in val_str:
        if val_str.rfind(",") > val_str.rfind("."):
            val_str = val_str.replace(".", "").replace(",", ".")
        else:
            val_str = val_str.replace(",", "")
    elif "," in val_str:
        if val_str.count(",") == 1:
            val_str = val_str.replace(",", ".")
        else:
            val_str = val_str.replace(",", "")
    elif "." in val_str:
        partes = val_str.split(".")
        if len(partes) > 2:
            val_str = val_str.replace(".", "")
        elif len(partes) == 2:
            if len(partes[1]) == 3 and len(partes[0]) <= 3:
                val_str = val_str.replace(".", "")

    try:
        return float(val_str)
    except (ValueError, TypeError):
        return 0.0


def _dedent_html(html: str) -> str:
    return "\n".join(line.strip() for line in html.strip("\n").split("\n"))


_RENOMBRES_BBD_STOCK = {
    "Código": "codigo_articulo",
    "Codigo SB": "codigo_sb",
    "Código PU": "codigo_pu",
    "Descripción": "descripcion",
    "Lote Proveedor": "lote_proveedor",
    "Fecha vence": "fecha_expiracion",
    "Estado": "estado_lote",
    "Localizador": "localizador",
    "Cantidad": "cantidad",
}


@st.cache_data(show_spinner="Leyendo pestaña BBD STOCK...")
def cargar_bbd_stock(archivo) -> pd.DataFrame:
    """Lee la pestaña 'BBD STOCK' del Refresh y renombra sus columnas para
    que calcen con lo que espera render_stock()."""
    try:
        archivo.seek(0)
    except Exception:
        pass
    try:
        df = pd.read_excel(archivo, sheet_name="BBD STOCK")
    finally:
        try:
            archivo.seek(0)
        except Exception:
            pass
    df.columns = [str(c).strip() for c in df.columns]
    rename_map = {c: _RENOMBRES_BBD_STOCK[c] for c in df.columns if c in _RENOMBRES_BBD_STOCK}
    return df.rename(columns=rename_map)


def render_stock(df_stock_raw, key_ns: str = "stock", titulo: str = "📦 Dashboard de Fecha de Caducidad"):
    """Renderiza el dashboard de Stock / Fecha de Caducidad (pestaña 2).
    key_ns permite reutilizar esta misma funcion en mas de una pestaña
    (ej: STOCK y BBD STOCK) sin que sus widgets choquen entre si."""
    df = df_stock_raw.copy()

    st.markdown(f"### {titulo}")

    col_cod = next(
        (
            c
            for c in df.columns
            if c.strip().lower()
            in ["codigo_articulo", "id_producto", "sku", "codigo"]
        ),
        None,
    )
    col_estado_sub = next(
        (c for c in df.columns if c.strip().lower() == "estado_subin"), None
    ) or next(
        (
            c
            for c in df.columns
            if c.strip().lower() in ["sub_inventario", "estado sub inventario"]
        ),
        None,
    )
    col_estado_lote = next(
        (
            c
            for c in df.columns
            if c.strip().lower()
            in ["estado_lote", "estado lote", "estado_lote_prov"]
        ),
        None,
    )
    col_lote = next(
        (
            c
            for c in df.columns
            if c.strip().lower() == "lote_proveedor"
        ),
        None,
    ) or next(
        (
            c
            for c in df.columns
            if c.strip().lower() in ["lote", "lote_prov"]
        ),
        None,
    )
    col_loc = next(
        (
            c
            for c in df.columns
            if c.strip().lower() in ["localizador", "ubicacion"]
        ),
        None,
    )
    col_desc_stock = next(
        (c for c in df.columns if "descripcion" in c.lower()), None
    )
    if not col_desc_stock and len(df.columns) > 3:
      col_desc_stock = df.columns[3]
    col_fecha = next(
        (
            c
            for c in df.columns
            if c.strip().lower()
            in [
                "fecha_expiracion_lote",
                "vencimiento",
                "fecha expiracion",
                "fecha_expiracion",
            ]
        ),
        None,
    )
    col_cant = next(
        (
            c
            for c in df.columns
            if c.strip().lower() in ["cantidad", "stock", "unidades"]
        ),
        None,
    )

    if col_cod and col_cod in df.columns:
      df[col_cod] = df[col_cod].apply(fmt_code)

    col_sku_sb = next(
        (c for c in df.columns if c.strip().lower() == "codigo_sb"), None
    )
    col_sku_pu = next(
        (c for c in df.columns if c.strip().lower() == "codigo_pu"), None
    )
    if not col_sku_sb and len(df.columns) > 1:
      col_sku_sb = df.columns[1]
    if not col_sku_pu and len(df.columns) > 2:
      col_sku_pu = df.columns[2]

    if col_sku_sb and col_sku_sb in df.columns:
      df[col_sku_sb] = df[col_sku_sb].apply(fmt_code)
    if col_sku_pu and col_sku_pu in df.columns:
      df[col_sku_pu] = df[col_sku_pu].apply(fmt_code)

    if col_cant:
      df[col_cant] = df[col_cant].apply(limpiar_numero)

    hoy = pd.Timestamp.today()
    limite_6m = hoy + pd.DateOffset(months=6)
    limite_13m = hoy + pd.DateOffset(months=13)

    if col_fecha:
      df[col_fecha] = pd.to_datetime(df[col_fecha], errors="coerce")

      def calcular_alerta(fecha):
        if pd.isna(fecha):
          return "Sin Fecha"
        if fecha < hoy:
          return "Vencido"
        if fecha < limite_6m:
          return "Menos de 6 meses"
        elif fecha <= limite_13m:
          return "Pronto vence (6-13m)"
        else:
          return "Vigente (> 13m)"

      df["Alerta_Caducidad"] = df[col_fecha].apply(calcular_alerta)
    else:
      df["Alerta_Caducidad"] = "Sin Fecha"
      df[col_fecha] = "N/A"

    col_dash1, col_dash2 = st.columns([1, 2.3])

    key_codigo = f"sel_codigo_{key_ns}"
    key_sku_sb = f"sel_sku_sb_{key_ns}"
    key_sku_pu = f"sel_sku_pu_{key_ns}"

    def _limpiar_otros_filtros(keys_a_limpiar):
      for k in keys_a_limpiar:
        if k in st.session_state:
          st.session_state[k] = "Todos"

    with col_dash2:
      _pad_izq, filtro_codigo_col, filtro_sku_sb_col, filtro_sku_pu_col, _pad_der = (
          st.columns([0.3, 1, 1, 1, 0.3])
      )

      with filtro_codigo_col:
        if col_cod:
          lista_codigos = sorted(
              [str(x) for x in df[col_cod].dropna().unique() if str(x).strip() != ""]
          )
          codigo_sel = st.selectbox(
              "Código:",
              ["Todos"] + lista_codigos,
              key=key_codigo,
              on_change=_limpiar_otros_filtros,
              args=([key_sku_sb, key_sku_pu],),
          )
        else:
          codigo_sel = "Todos"

      with filtro_sku_sb_col:
        if col_sku_sb and col_sku_sb in df.columns:
          lista_sku_sb = sorted(
              [str(x) for x in df[col_sku_sb].dropna().unique() if str(x).strip() != "" and str(x) != "S/N"]
          )
          sku_sb_sel = st.selectbox(
              "SKU SB:",
              ["Todos"] + lista_sku_sb,
              key=key_sku_sb,
              on_change=_limpiar_otros_filtros,
              args=([key_codigo, key_sku_pu],),
          )
        else:
          sku_sb_sel = "Todos"

      with filtro_sku_pu_col:
        if col_sku_pu and col_sku_pu in df.columns:
          lista_sku_pu = sorted(
              [str(x) for x in df[col_sku_pu].dropna().unique() if str(x).strip() != "" and str(x) != "S/N"]
          )
          sku_pu_sel = st.selectbox(
              "SKU PU:",
              ["Todos"] + lista_sku_pu,
              key=key_sku_pu,
              on_change=_limpiar_otros_filtros,
              args=([key_codigo, key_sku_sb],),
          )
        else:
          sku_pu_sel = "Todos"

    df_dash = df.copy()
    if codigo_sel != "Todos" and col_cod:
      df_dash = df_dash[df_dash[col_cod].astype(str) == codigo_sel].copy()

    if sku_sb_sel != "Todos" and col_sku_sb and col_sku_sb in df_dash.columns:
      df_dash = df_dash[df_dash[col_sku_sb].astype(str) == sku_sb_sel].copy()

    if sku_pu_sel != "Todos" and col_sku_pu and col_sku_pu in df_dash.columns:
      df_dash = df_dash[df_dash[col_sku_pu].astype(str) == sku_pu_sel].copy()

    if codigo_sel != "Todos":
      prod_sel = codigo_sel
    elif sku_sb_sel != "Todos":
      prod_sel = sku_sb_sel
    elif sku_pu_sel != "Todos":
      prod_sel = sku_pu_sel
    else:
      prod_sel = "Seleccione..."

    if col_cant:
      total_unidades = df_dash[col_cant].sum()
      total_vencido = df_dash[
          df_dash["Alerta_Caducidad"] == "Vencido"
      ][col_cant].sum()
      total_menos_6m = df_dash[
          df_dash["Alerta_Caducidad"] == "Menos de 6 meses"
      ][col_cant].sum()
      total_pronto = df_dash[
          df_dash["Alerta_Caducidad"] == "Pronto vence (6-13m)"
      ][col_cant].sum()
      total_vigentes = df_dash[
          df_dash["Alerta_Caducidad"] == "Vigente (> 13m)"
      ][col_cant].sum()
    else:
      total_unidades = len(df_dash)
      total_vencido = len(df_dash[df_dash["Alerta_Caducidad"] == "Vencido"])
      total_menos_6m = len(
          df_dash[df_dash["Alerta_Caducidad"] == "Menos de 6 meses"]
      )
      total_pronto = len(
          df_dash[df_dash["Alerta_Caducidad"] == "Pronto vence (6-13m)"]
      )
      total_vigentes = len(
          df_dash[df_dash["Alerta_Caducidad"] == "Vigente (> 13m)"]
      )

    total_critico = total_vencido + total_menos_6m
    pct_critico = (
        (total_critico / total_unidades * 100) if total_unidades > 0 else 0.0
    )

    st.markdown(
        f"""
            <style>
            .critico-card {{
                border-radius:14px; padding:14px 20px; margin-bottom:15px;
                background:linear-gradient(160deg, {COLOR_CARD_BG} 0%, rgba(255,255,255,0.02) 100%);
                border:1px solid {COLOR_CARD_BORDER};
                display:flex; justify-content:space-between; align-items:center;
                box-shadow:0 2px 10px rgba(0,0,0,0.22);
            }}
            </style>
            """,
        unsafe_allow_html=True,
    )
    color_pct_critico = (
        COLOR_ROJO if pct_critico >= 15
        else COLOR_AMARILLO if pct_critico >= 5
        else COLOR_VERDE
    )
    st.markdown(
        '<div class="critico-card">'
        f'<span style="color:{COLOR_TEXT_MUTED}; font-weight:600; text-transform:uppercase; font-size:13px;">'
        '⚠️ % de Stock Crítico (vencido + vence en &lt; 6 meses)</span>'
        f'<span class="stock-card2-value" style="background-color:{color_pct_critico}22;'
        f'color:{color_pct_critico};border:1px solid {color_pct_critico}55;">'
        f'{pct_critico:.2f}%</span>'
        "</div>",
        unsafe_allow_html=True,
    )

    label_map_alerta = {
        "Todos": "Todos",
        "Vencido": "Vencido",
        "Vence en < 6 meses": "Menos de 6 meses",
        "Pronto vence (6-13m)": "Pronto vence (6-13m)",
        "Vigente (> 13m)": "Vigente (> 13m)",
    }
    key_alerta = f"radio_alerta_{key_ns}"
    etiqueta_sel = st.radio(
        "🔍 Filtrar por categoría de caducidad:",
        list(label_map_alerta.keys()),
        horizontal=True,
        key=key_alerta,
    )
    filtro_actual = label_map_alerta[etiqueta_sel]

    if filtro_actual != "Todos":
      df_dash_alerta = df_dash[df_dash["Alerta_Caducidad"] == filtro_actual].copy()
    else:
      df_dash_alerta = df_dash.copy()

    with col_dash1:

      def _ring(color, valor):
        return (
            f"box-shadow:0 0 0 2px {color}99, 0 2px 10px rgba(0,0,0,0.22);"
            if filtro_actual == valor
            else "box-shadow:0 2px 10px rgba(0,0,0,0.22);"
        )

      def stock_card(label, value_text, color, filtro_valor):
        st.markdown(
            f'<div class="stock-card2" style="{_ring(color, filtro_valor)}">'
            f'<p class="stock-card2-label">{label}</p>'
            f'<span class="stock-card2-value" style="background-color:{color}22;'
            f'color:{color};border:1px solid {color}55;">{value_text}</span>'
            f"</div>",
            unsafe_allow_html=True,
        )

      stock_card(
          "Unidades registradas",
          formato_unidades(total_unidades),
          COLOR_ACENTO_1,
          "Todos",
      )
      stock_card(
          "Vencido",
          formato_unidades(total_vencido),
          "#DC2626",
          "Vencido",
      )
      stock_card(
          "Vence en < 6 meses",
          formato_unidades(total_menos_6m),
          COLOR_ROJO,
          "Menos de 6 meses",
      )
      stock_card(
          "Pronto vence (6 a 13 meses)",
          formato_unidades(total_pronto),
          COLOR_AMARILLO,
          "Pronto vence (6-13m)",
      )
      stock_card(
          "Vigentes (> 13 meses)",
          formato_unidades(total_vigentes),
          COLOR_VERDE,
          "Vigente (> 13m)",
      )

    with col_dash2:
      st.markdown("#### Estado de caducidad")
      labels = ["Vencido", "< 6 meses", "6 a 13 meses", "Vigente (> 13m)"]
      values = [total_vencido, total_menos_6m, total_pronto, total_vigentes]
      colors = ["#DC2626", COLOR_ROJO, COLOR_AMARILLO, COLOR_VERDE]

      total_donut = sum(values)
      if total_donut > 0:
        textos_pct = [
            f"{lbl}<br>{(v / total_donut * 100):.2f}%"
            for lbl, v in zip(labels, values)
        ]
        fig_pie = go.Figure(
            data=[
                go.Pie(
                    labels=labels,
                    values=values,
                    hole=0.62,
                    marker=dict(colors=colors, line=dict(color="#0B0E1A", width=3)),
                    text=textos_pct,
                    texttemplate="%{text}",
                    textposition="outside",
                    textfont=dict(size=12, color=COLOR_TEXT_MUTED),
                )
            ]
        )
        fig_pie.update_layout(
            height=380,
            margin=dict(t=20, b=60, l=60, r=60),
            paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="rgba(0,0,0,0)",
            font=dict(color="#F2F2F2"),
            showlegend=True,
            legend=dict(
                orientation="h",
                y=-0.15,
                x=0.5,
                xanchor="center",
                yanchor="top",
                font=dict(color=COLOR_TEXT_MUTED, size=12),
            ),
            annotations=[
                dict(
                    text=(
                        f"<b style='font-size:26px;color:#F2F2F2;'>"
                        f"{formato_unidades(total_donut)}</b><br>"
                        f"<span style='font-size:11px;color:{COLOR_TEXT_MUTED};"
                        "letter-spacing:.5px;'>TOTAL</span>"
                    ),
                    x=0.5,
                    y=0.5,
                    showarrow=False,
                )
            ],
        )
        _pad_chart_izq, col_chart, _pad_chart_der = st.columns([0.3, 2, 0.3])
        with col_chart:
          st.plotly_chart(
              fig_pie, use_container_width=True, key=f"pie_{key_ns}"
          )
      else:
        st.info("Sin registros para mostrar.")

      if prod_sel != "Seleccione...":
        stock_actual = df_dash[col_cant].sum() if col_cant else 0
        prox_vencer = (
            df_dash[df_dash[col_fecha].notna()][col_fecha].min()
            if col_fecha
            else None
        )
        dias_vencer = (
            (prox_vencer - hoy).days if pd.notna(prox_vencer) else "N/A"
        )

        st.markdown(
            f'<div class="stock-card2" style="box-shadow:0 2px 10px rgba(0,0,0,0.22);">'
            f'<p class="stock-card2-label">Stock actual</p>'
            f'<span class="stock-card2-value" style="background-color:{COLOR_ACENTO_1}22;'
            f'color:{COLOR_ACENTO_1};border:1px solid {COLOR_ACENTO_1}55;">'
            f"{formato_unidades(stock_actual)}</span></div>",
            unsafe_allow_html=True,
        )

        if isinstance(dias_vencer, int):
          if prox_vencer < limite_6m:
            texto_vence = (
                f"Vence en {dias_vencer} días"
                if dias_vencer >= 0
                else f"Venció hace {abs(dias_vencer)} días"
            )
            color_vence = COLOR_ROJO
          elif prox_vencer <= limite_13m:
            texto_vence = f"Vence en {dias_vencer} días"
            color_vence = COLOR_AMARILLO
          else:
            texto_vence = f"Vence en {dias_vencer} días"
            color_vence = COLOR_VERDE
        else:
          texto_vence = "Sin fecha registrada"
          color_vence = COLOR_TEXT_MUTED

        st.markdown(
            f'<div class="stock-card2" style="box-shadow:0 2px 10px rgba(0,0,0,0.22);">'
            f'<p class="stock-card2-label">Plazo de vencimiento</p>'
            f'<span class="stock-card2-value" style="font-size:16px;'
            f'background-color:{color_vence}22;color:{color_vence};'
            f'border:1px solid {color_vence}55;">{texto_vence}</span></div>',
            unsafe_allow_html=True,
        )

    st.divider()

    if col_estado_lote and col_estado_lote in df_dash_alerta.columns:
      st.markdown("##### 🏷️ Cantidad de Unidades por Estado de Lote")
      df_est_grp = (
          df_dash_alerta.groupby(col_estado_lote, dropna=False)[col_cant]
          .sum()
          .reset_index()
          if col_cant
          else df_dash_alerta[col_estado_lote].value_counts().reset_index()
      )

      if not df_est_grp.empty:
        c_e, c_q = df_est_grp.columns[0], df_est_grp.columns[1]
        num_items = len(df_est_grp)
        cols_est = st.columns(min(num_items, 6))
        for idx_e, row_e in df_est_grp.iterrows():
          nombre_est = (
              str(row_e[c_e]) if pd.notna(row_e[c_e]) else "Sin Estado"
          )
          cant_est = row_e[c_q]

          with cols_est[idx_e % min(num_items, 6)]:
            st.markdown(
                _dedent_html(f"""<div style="background-color: #141414; border: 1px solid #0070f3; border-radius: 8px; padding: 10px; text-align: center; margin-bottom: 15px;">
                                  <div style="font-size: 12px; color: #aaaaaa; font-weight: 600; text-transform: uppercase;">{nombre_est}</div>
                                  <div style="font-size: 20px; font-weight: bold; color: #ffffff; margin-top: 3px;">{formato_unidades(cant_est)}</div>
                              </div>"""),
                unsafe_allow_html=True,
            )

      st.divider()

    # Filtro adicional por "Lote Proveedor", propio de la tabla de detalle.
    key_lote = f"sel_lote_proveedor_{key_ns}"
    if col_lote and col_lote in df_dash_alerta.columns:
      lista_lotes = sorted(
          [
              str(x)
              for x in df_dash_alerta[col_lote].dropna().unique()
              if str(x).strip() != ""
          ]
      )
    else:
      lista_lotes = []
    lote_sel = "Todos"

    detalle_filtro = "(General)"
    partes_filtro = []
    if codigo_sel != "Todos":
      partes_filtro.append(f"Código: {codigo_sel}")
    if sku_sb_sel != "Todos":
      partes_filtro.append(f"SKU SB: {sku_sb_sel}")
    if sku_pu_sel != "Todos":
      partes_filtro.append(f"SKU PU: {sku_pu_sel}")
    if filtro_actual != "Todos":
      partes_filtro.append(f"Caducidad: {filtro_actual}")

    st.subheader("📋 Detalle de Stock y Lotes")

    # Fila con el filtro de Lote Proveedor (a la izquierda) y el botón
    # de descarga a Excel (a la derecha), alineados con la tabla de abajo.
    col_filtro_lote, col_espacio, col_descarga = st.columns([1.3, 2.2, 1])

    with col_filtro_lote:
      if lista_lotes:
        lote_sel = st.selectbox(
            "Lote Proveedor:",
            ["Todos"] + lista_lotes,
            key=key_lote,
        )
      else:
        st.selectbox(
            "Lote Proveedor:",
            ["Todos"],
            key=key_lote,
            disabled=True,
        )

    if lote_sel != "Todos" and col_lote and col_lote in df_dash_alerta.columns:
      partes_filtro.append(f"Lote Proveedor: {lote_sel}")
      df_dash_alerta = df_dash_alerta[
          df_dash_alerta[col_lote].astype(str) == lote_sel
      ].copy()

    if partes_filtro:
      detalle_filtro = f"({' | '.join(partes_filtro)})"

    st.caption(detalle_filtro)

    cols_mostrar = []
    nombres_amigables = {}
    if col_cod:
      cols_mostrar.append(col_cod)
      nombres_amigables[col_cod] = "Código Artículo"
    if col_desc_stock and col_desc_stock in df_dash_alerta.columns:
      cols_mostrar.append(col_desc_stock)
      nombres_amigables[col_desc_stock] = "Descripción"
    if col_sku_sb and col_sku_sb in df_dash_alerta.columns:
      cols_mostrar.append(col_sku_sb)
      nombres_amigables[col_sku_sb] = "SKU SB"
    if col_sku_pu and col_sku_pu in df_dash_alerta.columns:
      cols_mostrar.append(col_sku_pu)
      nombres_amigables[col_sku_pu] = "SKU PU"
    if col_estado_sub:
      cols_mostrar.append(col_estado_sub)
      nombres_amigables[col_estado_sub] = "Estado Sub-Inv"
    if col_estado_lote:
      cols_mostrar.append(col_estado_lote)
      nombres_amigables[col_estado_lote] = "Estado Lote"
    if col_lote:
      cols_mostrar.append(col_lote)
      nombres_amigables[col_lote] = "Lote Proveedor"
    if col_loc:
      cols_mostrar.append(col_loc)
      nombres_amigables[col_loc] = "Localizador"
    if col_cant:
      cols_mostrar.append(col_cant)
      nombres_amigables[col_cant] = "Cantidad"
    if col_fecha:
      cols_mostrar.append(col_fecha)
      nombres_amigables[col_fecha] = "Fecha Expiración"
    cols_mostrar.append("Alerta_Caducidad")
    nombres_amigables["Alerta_Caducidad"] = "Rango Caducidad"

    df_vista_stock = df_dash_alerta[cols_mostrar].copy()
    df_vista_stock = df_vista_stock.rename(columns=nombres_amigables)

    if "Fecha Expiración" in df_vista_stock.columns:
      df_vista_stock["Fecha Expiración"] = pd.to_datetime(
          df_vista_stock["Fecha Expiración"], errors="coerce"
      ).dt.strftime("%d-%m-%Y")

    with col_descarga:

      @st.cache_data(show_spinner=False)
      def _construir_excel_stock(df_vista_stock):
        """Arma el Excel con estilo de reporte (encabezado azul marino,
        filas alternadas y bordes finos). Se cachea por contenido del
        DataFrame para no repetir el formateo celda a celda en cada
        rerun de Streamlit cuando los datos no cambiaron."""
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
          df_vista_stock.to_excel(writer, index=False, sheet_name="Stock")
          ws_stock = writer.sheets["Stock"]

          n_filas, n_cols = df_vista_stock.shape
          rango_tabla = None
          if n_filas > 0 and n_cols > 0:
            ultima_col = get_column_letter(n_cols)
            rango_tabla = f"A1:{ultima_col}{n_filas + 1}"

            # Estilo manual tipo "reporte": encabezado azul marino con
            # texto blanco en negrita, filas de datos alternando
            # blanco y gris muy claro, con bordes finos. Los objetos
            # de estilo se crean UNA sola vez y se reutilizan (openpyxl
            # los deduplica internamente), y se recorre con iter_rows
            # en vez de ws.cell() para evitar el costo de traducir
            # fila/columna a notación A1 en cada celda.
            RELLENO_ENCABEZADO = PatternFill(
                start_color="1F3864", end_color="1F3864", fill_type="solid"
            )
            RELLENO_FILA_PAR = PatternFill(
                start_color="FFFFFF", end_color="FFFFFF", fill_type="solid"
            )
            RELLENO_FILA_IMPAR = PatternFill(
                start_color="F2F2F2", end_color="F2F2F2", fill_type="solid"
            )
            FUENTE_ENCABEZADO = Font(
                name="Calibri", size=12, bold=True, color="FFFFFF"
            )
            FUENTE_DATO = Font(name="Calibri", size=12, color="000000")
            BORDE_FINO = Border(
                left=Side(style="thin", color="D9D9D9"),
                right=Side(style="thin", color="D9D9D9"),
                top=Side(style="thin", color="D9D9D9"),
                bottom=Side(style="thin", color="D9D9D9"),
            )
            ALINEACION_CENTRO = Alignment(
                horizontal="center", vertical="center"
            )

            fila_encabezado = next(
                ws_stock.iter_rows(min_row=1, max_row=1, max_col=n_cols)
            )
            for celda_enc in fila_encabezado:
              celda_enc.fill = RELLENO_ENCABEZADO
              celda_enc.font = FUENTE_ENCABEZADO
              celda_enc.alignment = ALINEACION_CENTRO
              celda_enc.border = BORDE_FINO

            for idx_fila, fila in enumerate(
                ws_stock.iter_rows(
                    min_row=2, max_row=n_filas + 1, max_col=n_cols
                ),
                start=2,
            ):
              relleno_fila = (
                  RELLENO_FILA_PAR
                  if idx_fila % 2 == 0
                  else RELLENO_FILA_IMPAR
              )
              for celda in fila:
                celda.fill = relleno_fila
                celda.font = FUENTE_DATO
                celda.alignment = ALINEACION_CENTRO
                celda.border = BORDE_FINO

            ws_stock.row_dimensions[1].height = 20

            # Altura de todas las filas de datos también en 20, a
            # tono con la fuente de tamaño 12.
            for idx_fila in range(2, n_filas + 2):
              ws_stock.row_dimensions[idx_fila].height = 20

          # Ancho de columna ajustado al contenido para que no quede
          # todo apretado ni con texto cortado al abrir el archivo.
          anchos_columnas = []
          for idx_col, col_name in enumerate(
              df_vista_stock.columns, start=1
          ):
            letra_col = get_column_letter(idx_col)
            largo_max = max(
                [len(str(col_name))]
                + [len(str(v)) for v in df_vista_stock[col_name]]
            ) if n_filas > 0 else len(str(col_name))
            ancho_col = min(largo_max + 4, 45)
            ws_stock.column_dimensions[letra_col].width = ancho_col
            anchos_columnas.append(ancho_col)

          ws_stock.freeze_panes = "A2"

          # Configuración de impresión: hoja Carta, horizontal,
          # centrada, y ajustada automáticamente a 1 página de ancho
          # (fitToWidth) para que la tabla siempre entre en el ancho
          # de la hoja y quede bien centrada.
          ws_stock.page_setup.orientation = "landscape"
          ws_stock.page_setup.paperSize = ws_stock.PAPERSIZE_LETTER

          ws_stock.sheet_properties.pageSetUpPr.fitToPage = True
          ws_stock.page_setup.fitToWidth = 1
          ws_stock.page_setup.fitToHeight = 0
          ws_stock.print_options.horizontalCentered = True
          ws_stock.print_options.verticalCentered = False
          # Márgenes descuadrados a propósito (izquierdo más chico,
          # derecho más grande) para correr el área de impresión
          # ~1 cm (0,4") hacia la izquierda, así la última columna
          # no sale cortada por el borde derecho de la hoja.
          ws_stock.page_margins.left = 0.05
          ws_stock.page_margins.right = 0.85
          ws_stock.page_margins.top = 0.5
          ws_stock.page_margins.bottom = 0.5
          if rango_tabla:
            ws_stock.print_area = rango_tabla
            ws_stock.print_title_rows = "1:1"

        buffer.seek(0)
        return buffer.getvalue()

      bytes_excel_stock = _construir_excel_stock(df_vista_stock)

      st.download_button(
          label="⬇️ Descargar Excel",
          data=bytes_excel_stock,
          file_name=f"detalle_stock_lotes_{date.today().strftime('%Y%m%d')}.xlsx",
          mime=(
              "application/vnd.openxmlformats-officedocument"
              ".spreadsheetml.sheet"
          ),
          key=f"btn_descarga_{key_ns}",
          use_container_width=True,
      )

    st.dataframe(df_vista_stock, hide_index=True, use_container_width=True)

    st.divider()

    if col_loc and col_loc in df_dash.columns:
      if filtro_actual != "Todos":
        titulo_loc = f"##### 📍 Top Localizadores — {filtro_actual}"
        df_critico = df_dash_alerta.copy()
      else:
        titulo_loc = "##### 📍 Top Localizadores con más Stock por Vencer"
        df_critico = df_dash[
            df_dash["Alerta_Caducidad"].isin(["Vencido", "Menos de 6 meses"])
        ].copy()

      st.markdown(titulo_loc)

      df_critico = df_critico[
          df_critico[col_loc].notna()
          & (df_critico[col_loc].astype(str).str.strip() != "")
      ].copy()

      if not df_critico.empty:
        cols_group = [col_loc]
        if col_desc_stock and col_desc_stock in df_critico.columns:
          cols_group.append(col_desc_stock)

        if col_cant:
          grp_loc = (
              df_critico.groupby(cols_group, dropna=False)[col_cant]
              .sum()
              .reset_index()
              .rename(columns={col_cant: "Cantidad"})
          )
        else:
          grp_loc = (
              df_critico.groupby(cols_group, dropna=False)
              .size()
              .reset_index(name="Cantidad")
          )

        grp_loc = grp_loc.sort_values(by="Cantidad", ascending=False).head(10)

        etiqueta_barra = (
            grp_loc[col_loc].astype(str)
            + (
                " — " + grp_loc[col_desc_stock].astype(str)
                if col_desc_stock and col_desc_stock in grp_loc.columns
                else ""
            )
        )
        grp_loc_sorted = grp_loc.assign(_etiqueta=etiqueta_barra).sort_values(
            by="Cantidad", ascending=True
        )
        fig_loc = px.bar(
            grp_loc_sorted,
            x="Cantidad",
            y="_etiqueta",
            orientation="h",
            text_auto=",.0f",
            color_discrete_sequence=["#e74c3c"],
        )
        fig_loc.update_traces(
            textfont_size=11, textposition="outside", cliponaxis=False
        )
        fig_loc.update_layout(
            template="plotly_dark",
            paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="rgba(0,0,0,0)",
            margin=dict(t=10, b=10, l=10, r=10),
            height=320,
            xaxis_title="",
            yaxis_title="",
        )
        st.plotly_chart(
            fig_loc, use_container_width=True, key=f"top_loc_{key_ns}"
        )

        rename_cols = {col_loc: "Localizador"}
        if col_desc_stock and col_desc_stock in grp_loc.columns:
          rename_cols[col_desc_stock] = "Descripción Producto"
        grp_loc_disp = grp_loc.rename(columns=rename_cols)
        st.dataframe(
            grp_loc_disp,
            column_config={
                "Cantidad": st.column_config.NumberColumn(
                    "Cantidad", format="%,d"
                ),
            },
            hide_index=True,
            use_container_width=True,
        )
      else:
        st.info(
            "No hay stock (con localizador registrado) para la categoría seleccionada."
        )



def render():
    _header()

    RUTA_REFRESH_DEFAULT = "data/Refresh.xlsx"

    with st.sidebar:
        st.markdown("### 📂 Fuente de datos")
        archivo_subido = st.file_uploader(
            "Reemplazar con un Excel más nuevo (opcional)", type=["xlsx"]
        )
        st.caption(
            f"Si no subes nada, se usa el archivo incluido en el repositorio "
            f"(`{RUTA_REFRESH_DEFAULT}`)."
        )
        if os.path.exists(RUTA_REFRESH_DEFAULT):
            mtime = datetime.datetime.fromtimestamp(os.path.getmtime(RUTA_REFRESH_DEFAULT))
            st.caption(f"🕒 Última modificación del archivo: {mtime.strftime('%d-%m-%Y')}")

    if archivo_subido is not None:
        archivo = archivo_subido
    elif os.path.exists(RUTA_REFRESH_DEFAULT):
        archivo = RUTA_REFRESH_DEFAULT
    else:
        st.info(
            "Sube el archivo Refresh (o deja uno guardado como "
            f"`{RUTA_REFRESH_DEFAULT}` en el repo) para continuar."
        )
        return


    tab_despacho, tab_bbd = st.tabs(
        ["🚚 Plan de Despacho", "🧬 Stock y Caducidad (BBD STOCK)"]
    )

    with tab_despacho:
        df = leer_sb(archivo)
        semanas_disp = sorted(df["Semana"].dropna().unique().tolist())

        semana = st.radio(
            "Semana a planificar", semanas_disp, horizontal=True,
            index=len(semanas_disp) - 1 if semanas_disp else 0,
        )

        # El año no se pide al usuario: se infiere de la Fecha vence de esa
        # misma semana en el Refresh (necesario solo para ubicar el Lunes ISO).
        _fechas_semana = pd.to_datetime(
            df.loc[df["Semana"] == semana, "Fecha vence"], errors="coerce"
        ).dropna()
        anio = int(_fechas_semana.dt.year.mode().iloc[0]) if not _fechas_semana.empty \
            else datetime.date.today().year

        with st.sidebar:
            st.markdown("### ⚙️ Configuración")
            pallet_col = st.radio(
                "Columna a usar para calcular pallets por OC",
                ["Pallets Pos.", "Pallets posibles"],
                help="'Pallets Pos.' viene acotada (0.3/1). 'Pallets posibles' es la "
                     "fracción real sin redondear.",
            )
            capacidades = st.multiselect("Capacidades de camión disponibles (pallets)",
                                          [13, 16], default=[13, 16])
            dias = st.multiselect("Días hábiles de despacho",
                                   ["Lunes", "Martes", "Miercoles", "Jueves", "Viernes", "Sabado"],
                                   default=["Lunes", "Martes", "Miercoles", "Jueves", "Viernes"])
            ventanas_por_dia = st.number_input("Ventanas de despacho por día", value=4, min_value=1)
            st.caption("Los productos con nombre en 'Directos' se sacan ANTES de armar los "
                       "camiones (no ocupan pallets/ventanas) y quedan solo como información. "
                       "Farma y Consumo Masivo nunca comparten camión. "
                       "Orden de carga del resto: 1) Solicitado=1er Posible, "
                       "2) cubierto con Pronto-vence, 5) parcial sin cobertura, 3) sin 1er Posible.")
            orden_prioridad = [1, 2, 5, 3]

        facturados = cargar_facturados_desde_refresh(archivo)
        if facturados:
            st.caption(
                f"✅ {len(facturados)} pedidos detectados como Facturados desde la pestaña "
                "OC del Refresh (automático, sin archivo aparte)."
            )

        if not capacidades or not dias:
            st.warning("Elige al menos una capacidad de camión y un día hábil.")
            return

        resumen, detalle, info, tabla_directos = generar_plan(
            df, semana, int(anio), pallet_col, capacidades, dias,
            int(ventanas_por_dia), orden_prioridad, facturados,
        )

        if resumen is None:
            st.warning(f"No hay OC para camión en la semana {semana} "
                        f"(revisa si todas quedaron como Directos).")
            if not tabla_directos.empty:
                st.subheader("Directos (información, no van en camión)")
                st.dataframe(tabla_directos, use_container_width=True, hide_index=True)
            return

        holgura = info["ventanas_disponibles"] - info["camiones"]
        render_kpi_cards([
            {"value": info["camiones"], "label": "Camiones necesarios"},
            {"value": info["ventanas_disponibles"], "label": "Ventanas disponibles"},
            {
                "value": holgura, "label": "Holgura",
                "badge_text": "Atención" if holgura < 0 else "OK",
                "badge_color": "#E4572E" if holgura < 0 else "#1DB980",
            },
            {"value": info["oc_directos"], "label": "OC 100% Directos"},
            {"value": info["oc_facturadas"], "label": "OC ya Facturadas"},
        ])
        if info["overflow"]:
            st.error("⚠️ No alcanzan las ventanas de la semana para todos los camiones "
                      "necesarios. Suma días/ventanas o revisa las capacidades.")

        excel_pendientes = exportar_pendientes_excel(detalle)
        col_desc_a, col_desc_b = st.columns(2)
        with col_desc_a:
            refresh_bytes = _bytes_archivo_original(archivo)
            if refresh_bytes:
                st.download_button(
                    "⬇️ Descargar Refresh de origen (Excel)",
                    data=refresh_bytes,
                    file_name="Refresh.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
        with col_desc_b:
            if excel_pendientes:
                n_pend = int((detalle["Facturado"] == "No").sum())
                st.download_button(
                    f"⬇️ Descargar no facturadas (Excel) · {n_pend} OC",
                    data=excel_pendientes,
                    file_name=f"Pendientes_S{semana}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
            else:
                st.success("✅ Todas las OC de esta semana ya aparecen como Facturadas.")

        st.subheader("Detalle por OC (van en camión)")
        tab_calendario, tab_tabla = st.tabs(["🗓️ Calendario", "📋 Tabla"])
        with tab_calendario:
            st.caption("Una columna por día, tarjetas con Pedido, OC, Monto y Pallets — "
                       "ordenadas por ventana y prioridad.")
            render_calendario(detalle, resumen)
        with tab_tabla:
            st.caption("Las filas en verde ya aparecen como Facturadas en la tabla externa.")
            st.dataframe(
                detalle.style.apply(_resaltar_facturado, axis=1).format({
                    "Pallets": "{:.0f}", "Monto": lambda v: formato_clp(v),
                }),
                use_container_width=True, hide_index=True,
            )

        st.subheader("Plan de camiones")
        st.caption("Los números de Pedido en verde ya aparecen como Facturados. "
                   "La columna % Facturado indica qué proporción de ese camión ya se despachó.")
        render_tabla_camiones(resumen, detalle)

        if not tabla_directos.empty:
            st.subheader("Directos (información, NO ocupan camión/ventana)")
            st.caption(f"{info['lineas_directos']} líneas / {info['oc_directos']} OC con "
                       "proveedor directo asignado.")
            st.dataframe(tabla_directos, use_container_width=True, hide_index=True)

        st.download_button(
            "⬇️ Descargar plan en Excel",
            data=exportar_excel(resumen, detalle, tabla_directos),
            file_name=f"Plan_Despacho_S{semana}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    with tab_bbd:
        try:
            df_bbd = cargar_bbd_stock(archivo)
        except Exception as e:
            st.error(f"No pude leer la pestaña 'BBD STOCK' de ese archivo: {e}")
        else:
            render_stock(
                df_bbd, key_ns="bbd",
                titulo="🧬 Dashboard de Stock y Caducidad (BBD STOCK)",
            )


if __name__ == "__main__":
    render()
