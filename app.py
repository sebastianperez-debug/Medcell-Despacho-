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
import math
import os
from io import BytesIO

import pandas as pd
import streamlit as st

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
                    st.markdown(
                        f"<div style='font-size:0.68rem;font-weight:700;color:#0B4F86;"
                        f"letter-spacing:0.03em;margin:0.5rem 0 0.35rem;'>"
                        f"VENTANA {ventana} · {total_pallets:.0f} pal</div>",
                        unsafe_allow_html=True,
                    )
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


if __name__ == "__main__":
    render()
