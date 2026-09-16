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

# Configuracion por hoja: nombres de columna y reglas de clasificacion.
# SB usa la regla completa (incluye cobertura por Pronto-vence). PU NO puede
# cubrir faltantes con stock por vencer, asi que esa regla se omite ahi.
HOJAS_CONFIG = {
    "SB": {
        "hoja": "SB", "semana": "Semana", "oc": "OC", "pedido": "Pedido",
        "fecha_vence": "Fecha vence", "solicitado": "Solicitado",
        "posible1": "1 Posible", "pronto_vence": "Pronto-vence",
        "solicitado_dolar": "Solicitado $", "directos": "Directos",
        "division": "Division", "descripcion": "Descripción",
        "sku": "SKU SB", "pallets_pos": "Pallets Pos.",
        "pallets_alt": "Pallets posibles", "usa_pronto_vence": True,
        "orden_prioridad": [1, 2, 5, 3],
        "capacidades_opciones": [13], "capacidades_default": [13],
        "dias_opciones": ["Lunes", "Martes", "Miercoles", "Jueves", "Viernes", "Sabado"],
        "dias_default": ["Lunes", "Martes", "Miercoles", "Jueves", "Viernes"],
        "usa_transportes": False, "n_transportes": None,
        # Flota real: camiones de 13 pallets para Consumo Masivo, con
        # capacidad para hacer 2 vueltas por dia -> 6 ventanas/dia. Farma
        # tiene SU PROPIA flota, de camiones de 16 pallets (ver
        # "capacidades_por_division" abajo), y un minimo garantizado de 2
        # de esos camiones por dia (flexible: solo si hay OC de Farma
        # esperando ese dia). Si el total de camiones de Consumo Masivo no
        # alcanza en las ventanas disponibles de la semana, como ULTIMO
        # RECURSO se fusionan (de a pares, empezando por los menos
        # prioritarios y siempre dentro de la misma division) en ramplas de
        # 27 pallets -> ver "capacidad_rescate" y _consolidar_con_rampla().
        "ventanas_por_dia_default": 6,
        "n_camiones_default": 3,
        "vueltas_por_camion_default": 2,
        "minimo_por_division": {"FARMA": 2},
        "capacidades_por_division": {"FARMA": [16]},
        # La rampla de 27 esta disponible para ambas divisiones (Consumo
        # Masivo la usa mas seguido; Farma solo si algun dia le falta
        # disponibilidad en sus 2 camiones de 16). Cuando se necesita una
        # rampla, se intenta ubicar en Miercoles o Jueves primero (que es
        # cuando en la practica se consigue ese transporte externo) -> ver
        # "dias_preferidos_rampla" y asignar_ventanas().
        "rescate_divisiones": {"CONSUMO MASIVO", "FARMA"},
        "dias_preferidos_rampla": {"Miercoles", "Miércoles", "Jueves"},
        "capacidad_rescate": 27,
        # Costos referenciales de flota (CLP). El camion normal (13 o 16
        # pallets) cobra un valor FIJO que ya incluye hasta 2 vueltas ese
        # dia (se haga 1 o 2, se paga igual); la rampla se cobra por vuelta,
        # y en este modelo siempre corresponde a 1 vuelta por rampla usada.
        "costo_camion": 170000,
        "costo_rampla": 270000,
        # Para SB el calculo de pallets SIEMPRE usa "Pallets Pos." (columna
        # AH del Refresh): no se ofrece alternativa en la UI para evitar que
        # alguien elija sin querer una columna distinta y el numero de
        # pallets mostrado deje de calzar con el Excel de origen.
        "forzar_pallet_col": True,
        # Columna "Producción" (BJ del Refresh): cuando trae algo (ej. el
        # texto "Producción"), esa OC son productos que hay que ESPERAR que
        # produccion fabrique, sin importar que la clasificacion de
        # prioridad diga "1 - Solicitado = 1er Posible (completo)". Por eso
        # esas OC se fuerzan a despacharse Miercoles o Jueves (mismos dias
        # que se usan para la rampla de rescate), nunca antes -> ver
        # "dias_preferidos_produccion" y asignar_ventanas().
        "produccion": "Producción",
        "dias_preferidos_produccion": {"Miercoles", "Miércoles", "Jueves"},
        # Columna "Campaña" (BQ del Refresh): se muestra tal cual en el
        # checklist de carga (columna "Campaña"), solo informativa.
        "campana": "Campaña",
    },
    "PU": {
        "hoja": "PU", "semana": "Sem", "oc": "OC", "pedido": "Pedido",
        "fecha_vence": "Fecha vence", "solicitado": "Solicitado",
        "posible1": "1er posible", "pronto_vence": None,
        "solicitado_dolar": "Solicitado $", "directos": "Directos",
        "division": "División", "descripcion": "Descripción",
        "sku": "Codigo PU", "pallets_pos": "Pallets Pos.",
        "pallets_alt": None, "usa_pronto_vence": False,
        "orden_prioridad": [1, 2, 3],
        # PU no se distribuye en la semana: solo se despacha el Viernes,
        # y ademas de camiones de 13/16 puede usar rampla de 27 pallets.
        "capacidades_opciones": [13, 16, 27], "capacidades_default": [13, 16, 27],
        "dias_opciones": ["Lunes", "Martes", "Miercoles", "Jueves", "Viernes", "Sabado"],
        "dias_default": ["Viernes"],
        # No hay "ventanas" fijas: son 3 transportes que hacen las vueltas
        # que hagan falta hasta completar todo el despacho de ese dia.
        "usa_transportes": True, "n_transportes": 3,
    },
}


def _version_archivo(archivo) -> str:
    """Identificador que cambia cuando el archivo cambia, para poder usarlo
    como parte de la clave de cache. Es necesario porque cuando 'archivo'
    es una ruta fija en disco (ej. data/Refresh.xlsx), el string de la ruta
    no cambia aunque el CONTENIDO del archivo se reemplace, y st.cache_data
    seguiria devolviendo el resultado viejo para siempre."""
    if isinstance(archivo, str):
        try:
            return str(os.path.getmtime(archivo))
        except OSError:
            return "0"
    # Archivo subido por el usuario (UploadedFile): Streamlit ya lo hashea
    # por contenido, pero devolvemos algo igual para mantener la firma
    # de cache consistente.
    return f"{getattr(archivo, 'name', '')}-{getattr(archivo, 'size', '')}"


@st.cache_data(show_spinner="Leyendo pestaña del Refresh...")
def leer_hoja(archivo, nombre_hoja: str, version: str = "") -> pd.DataFrame:
    """Lee una pestaña del Refresh (SB o PU) y devuelve un DataFrame limpio."""
    try:
        archivo.seek(0)
    except Exception:
        pass
    df = pd.read_excel(archivo, sheet_name=nombre_hoja, header=HEADER_ROW)
    try:
        archivo.seek(0)
    except Exception:
        pass
    # Normaliza nombres de columnas (a veces vienen con espacios extra)
    df.columns = [str(c).strip() for c in df.columns]
    return df


# --------------------------------------------------------------------------
# 1b. FACTURADOS
# --------------------------------------------------------------------------
# Se detecta automaticamente desde la pestana "OC" del MISMO Refresh:
#   "Pedido de Venta" = mismo numero que "Pedido" en SB/PU.
#   "Pendiente" = 0 (sumado por Pedido de Venta) -> ya se despacho/facturo.
# No requiere mantener ningun archivo aparte. Sirve para ambas pestañas.


@st.cache_data(show_spinner="Revisando pestaña OC del Refresh...")
def cargar_facturados_desde_refresh(archivo, version: str = "") -> dict:
    """Detecta el estado de facturacion de cada Pedido usando la pestana 'OC'
    del mismo Refresh ('Pedido de Venta', 'Despacho', 'Pendiente'):
    - 'Sí': todas sus lineas ya se despacharon (Pendiente = 0).
    - 'Parcial': al menos una linea se despacho, pero no todas (Despacho > 0
      y Pendiente > 0).
    - Si no aparece en el diccionario, se interpreta como 'No' (nada
      despachado)."""
    try:
        archivo.seek(0)
    except Exception:
        pass
    try:
        df_oc = pd.read_excel(archivo, sheet_name="OC")
    except Exception:
        return {}
    finally:
        try:
            archivo.seek(0)
        except Exception:
            pass

    df_oc.columns = [str(c).strip() for c in df_oc.columns]
    cols_necesarias = {"Pedido de Venta", "Despacho", "Pendiente"}
    if not cols_necesarias.issubset(df_oc.columns):
        return {}

    resumen = df_oc.groupby("Pedido de Venta")[["Despacho", "Pendiente"]].sum()
    estados = {}
    for pedido, fila in resumen.iterrows():
        if pd.isna(pedido):
            continue
        if fila["Despacho"] > 0 and fila["Pendiente"] <= 0:
            estados[str(int(pedido))] = "Sí"
        elif fila["Despacho"] > 0:
            estados[str(int(pedido))] = "Parcial"
    return estados


# --------------------------------------------------------------------------
# 2. CLASIFICACION DE PRIORIDAD POR OC
# --------------------------------------------------------------------------

def _clasificar(row, usa_pronto_vence: bool) -> tuple[int, str]:
    # Los Directos ya fueron separados antes de llegar aca: aqui solo quedan
    # las lineas que SI van por camion.
    sol, pos1 = row["sol"], row["pos1"]
    if pos1 == sol:
        return 1, "1 - Solicitado = 1er Posible (completo)"
    if usa_pronto_vence and (sol - pos1) <= row["pv"]:
        return 2, "2 - Diferencia cubierta por Pronto-vence"
    if pos1 == 0:
        return 3, "3 - Sin 1er Posible (sin stock disponible)"
    if usa_pronto_vence:
        return 5, "5 - Otros / parcial sin cobertura"
    # PU: el stock por vencer no se puede usar para completar faltantes,
    # asi que todo lo que no es completo ni cero queda como "parcial".
    return 2, "2 - No alcanza a completar el solicitado (parcial)"


def separar_directos(df: pd.DataFrame, semana: int, cfg: dict):
    """Separa las lineas con nombre en 'Directos' ANTES de armar camiones:
    esas lineas no ocupan espacio en ningun camion, solo se listan como info.
    Devuelve (df_camion, df_directos_info)."""
    d = df[df[cfg["semana"]] == semana].copy()
    d["directos_flag"] = d[cfg["directos"]].apply(
        lambda x: str(x).strip() not in ("", "-", "nan", "None")
    )
    df_directos = d[d["directos_flag"]].copy()
    df_camion = d[~d["directos_flag"]].copy()
    return df_camion, df_directos


def resumen_directos(df_directos: pd.DataFrame, cfg: dict) -> pd.DataFrame:
    """Tabla informativa de lineas Directas (no van en camion)."""
    if df_directos.empty:
        return df_directos
    cols = [cfg["pedido"], cfg["oc"], cfg["fecha_vence"], cfg["sku"],
            cfg["descripcion"], cfg["directos"], cfg["solicitado"], cfg["solicitado_dolar"]]
    cols = [c for c in cols if c in df_directos.columns]
    return df_directos[cols].rename(columns={
        cfg["directos"]: "Proveedor directo",
        cfg["solicitado_dolar"]: "Monto",
    })


def agrupar_por_oc(df_camion: pd.DataFrame, pallet_col: str, cfg: dict) -> pd.DataFrame:
    """Agrupa por Pedido (=OC) SOLO las lineas que van por camion (sin Directos)
    y clasifica prioridad."""
    if df_camion.empty:
        return df_camion

    agg_kwargs = dict(
        oc=(cfg["oc"], "first"),
        fecha_vence=(cfg["fecha_vence"], "first"),
        division=(cfg["division"], "first"),
        sol=(cfg["solicitado"], "sum"),
        pos1=(cfg["posible1"], "sum"),
        pallets=(pallet_col, "sum"),
        # "Posible actual $" se pone en $0 apenas la OC queda 100% despachada
        # (ya no hay nada "posible" pendiente), asi que para reflejar el
        # valor real de la OC (y cuanto se despacho de verdad) usamos
        # "Solicitado $", que no se resetea a 0.
        monto=(cfg["solicitado_dolar"], "sum"),
        n_sku=(cfg["pedido"], "count"),
    )
    if cfg["usa_pronto_vence"] and cfg["pronto_vence"] in df_camion.columns:
        agg_kwargs["pv"] = (cfg["pronto_vence"], "sum")

    # Columna "Producción" (BJ del Refresh): si CUALQUIER linea de la OC
    # trae algo ahi, es un producto que hay que esperar que produccion
    # fabrique -> esa OC entera se marca para forzarse a Miercoles/Jueves
    # en asignar_ventanas, sin importar que prioridad de despacho le toque.
    col_produccion = cfg.get("produccion")
    if col_produccion and col_produccion in df_camion.columns:
        df_camion = df_camion.copy()
        df_camion["_produccion_flag"] = df_camion[col_produccion].apply(
            lambda x: str(x).strip() not in ("", "-", "nan", "None")
        )
        agg_kwargs["requiere_produccion"] = ("_produccion_flag", "any")

    # Columna "Campaña" (BQ del Refresh): informativa, se lleva tal cual
    # (primer valor no vacio de la OC) hasta el checklist de carga.
    col_campana = cfg.get("campana")
    if col_campana and col_campana in df_camion.columns:
        agg_kwargs["campana"] = (col_campana, "first")

    agg = df_camion.groupby(cfg["pedido"]).agg(**agg_kwargs).reset_index()
    if "pv" not in agg.columns:
        agg["pv"] = 0
    if "campana" not in agg.columns:
        agg["campana"] = ""
    if "requiere_produccion" not in agg.columns:
        agg["requiere_produccion"] = False

    pr = agg.apply(lambda r: _clasificar(r, cfg["usa_pronto_vence"]), axis=1, result_type="expand")
    agg["prioridad"], agg["prioridad_label"] = pr[0], pr[1]
    # Regla de pallets:
    # - OC con mas de 1 pallet: se redondea HACIA ARRIBA individualmente
    #   (ej: 4.5 -> 5), igual que antes.
    # - OC con 1 pallet o menos (fragmentos chicos: medio pallet, un tercio,
    #   etc.): se deja como fraccion real (no se redondea sola a 1), para que
    #   el armado de camiones pueda COMBINAR varios fragmentos chicos en un
    #   mismo pallet fisico (ej: 2 OC de 0.5 = 1 pallet, no 2; 3 OC de 0.3 =
    #   1 pallet, no 3). El redondeo final hacia arriba se aplica recien al
    #   TOTAL del camion, no a cada fragmento por separado.
    agg["pallets_empaque"] = agg["pallets"].apply(
        lambda v: v if v <= 1 else math.ceil(round(v, 6))
    )
    # "pallets" se mantiene como el valor a MOSTRAR (siempre entero, para las
    # tablas/tarjetas), independiente de como se empaquen en el camion.
    agg["pallets"] = agg["pallets"].apply(lambda v: math.ceil(round(v, 6)))
    agg["monto"] = agg["monto"].round(0)
    agg = agg.rename(columns={cfg["pedido"]: "Pedido"})
    # OC con 0 pallets (linea unica que quedo en 0 tras excluir directos) no
    # necesita camion; se deja fuera del bin-packing.
    agg = agg[agg["pallets"] > 0].reset_index(drop=True)
    return agg


# --------------------------------------------------------------------------
# 3. ARMADO DE CAMIONES (bin packing, OC nunca se parte)
# --------------------------------------------------------------------------

# Tamano maximo (en pallets) de una OC para que pueda "rellenar" el hueco de un
# camion anterior ya casi lleno (ver empacar()). None = sin limite: cualquier OC
# puede adelantarse a un hueco. Si prefieres que solo se adelanten fragmentos
# chicos (ej: que una OC de 2 pal si pueda subirse al camion del lunes, pero una
# de 5 pal no), pon aca ese umbral: BACKFILL_MAX_PALLETS = 2.
BACKFILL_MAX_PALLETS = None

def _ventanas_del_dia(ventanas_por_dia, dia: str) -> int:
    """ventanas_por_dia puede ser un int fijo (mismo cupo todos los dias) o un
    dict {dia: cupo} cuando algun dia tiene menos disponibilidad (ej: se
    pierde un camion completo o una vuelta ese dia especifico)."""
    if isinstance(ventanas_por_dia, dict):
        return ventanas_por_dia.get(dia, 0)
    return ventanas_por_dia

def _intercalar_con_minimo_diario(bins_por_division: dict, dias: list[str],
                                   ventanas_por_dia: int, minimo_por_division: dict,
                                   orden_prioridad: list[int]) -> list[dict]:
    """Ordena los camiones dia por dia: primero reserva el minimo garantizado
    de cada division indicada en minimo_por_division (si hay camiones de esa
    division esperando), y despues completa el resto de las ventanas del dia
    con lo que siga en prioridad (de cualquier division). Es flexible: si una
    division no tiene camiones esperando, su cupo minimo simplemente no se usa
    ese dia (no se inventan camiones vacios)."""
    colas = {d: list(b) for d, b in bins_por_division.items()}

    def prioridad_bin(b):
        return min(orden_prioridad.index(it["prioridad"]) for it in b["items"])

    orden_final = []
    for dia in dias:
        cupo_dia = _ventanas_del_dia(ventanas_por_dia, dia)
        for div, minimo in minimo_por_division.items():
            n = 0
            while n < minimo and cupo_dia > 0 and colas.get(div):
                orden_final.append(colas[div].pop(0))
                cupo_dia -= 1
                n += 1
        while cupo_dia > 0 and any(colas.values()):
            mejor_div = min(
                (d for d in colas if colas[d]),
                key=lambda d: prioridad_bin(colas[d][0]),
            )
            orden_final.append(colas[mejor_div].pop(0))
            cupo_dia -= 1
        if not any(colas.values()):
            break

    for cola in colas.values():
        orden_final.extend(cola)
    return orden_final


def armar_camiones(agg: pd.DataFrame, capacidades: list[int],
                    orden_prioridad: list[int], dias: list[str] | None = None,
                    ventanas_por_dia: int | None = None,
                    minimo_por_division: dict | None = None,
                    capacidad_rescate: int | None = None,
                    capacidades_por_division: dict | None = None,
                    rescate_divisiones: set | None = None) -> list[dict]:
    """Arma camiones respetando 2 reglas duras:
    1) una OC nunca se parte entre camiones.
    2) Farma y Consumo Masivo NUNCA van en el mismo camion -> se empacan por
       separado (una division no le "presta" espacio a la otra).

    Si se entrega minimo_por_division (ej: {"FARMA": 2}) junto con dias y
    ventanas_por_dia, se reserva ese minimo de camiones de esa division
    dentro de CADA dia (si hay camiones de esa division esperando), en vez
    de solo intercalar por prioridad global.

    capacidad_rescate (ej: 27) se usa SOLO para etiquetar correctamente una
    OC que por si sola ya supera la capacidad normal (cap_max): no se
    fusiona nada aca (esa OC ya va sola, no se puede partir), solo se busca
    el tipo de vehiculo mas chico que igual le alcance. La fusion real de
    VARIOS camiones chicos en una rampla (el "ultimo recurso" cuando no
    alcanzan las ventanas de la semana) la hace _consolidar_con_rampla.

    capacidades_por_division (opcional, ej: {"FARMA": [16]}) le da a una
    division su PROPIA flota de capacidades, distinta de `capacidades`
    (que se sigue usando tal cual para cualquier otra division que no
    aparezca en el dict). Asi Farma puede despachar en camiones de 16
    pallets mientras Consumo Masivo sigue con los de 13 (+ rampla de
    rescate de 27), sin que una flota interfiera con la otra.

    rescate_divisiones (opcional, ej: {"CONSUMO MASIVO"}): si se entrega,
    la rampla de rescate SOLO se ofrece como opcion para esas divisiones
    (una division con flota propia, como Farma con sus camiones de 16, no
    tiene por que compartir la rampla externa de Consumo Masivo). Si se
    omite, el rescate queda disponible para todas las divisiones (mismo
    comportamiento que antes)."""
    capacidades_por_division = capacidades_por_division or {}

    def cerrar(b, capacidades_local, rescate_local):
        # El total del camion (suma de fragmentos, algunos fraccionarios) se
        # redondea HACIA ARRIBA reci�n aca, al cerrar el camion -- asi varios
        # fragmentos chicos (ej: 0.5 + 0.5) pueden compartir un mismo pallet
        # fisico en vez de que cada uno redondee a 1 por separado.
        b["total"] = math.ceil(round(b["total"], 6))
        opciones = sorted(set(capacidades_local) | ({rescate_local} if rescate_local else set()))
        for cap in opciones:
            if b["total"] <= cap:
                b["camion"] = cap
                return b
        # Ni la rampla de rescate alcanza: OC excepcionalmente grande. No se
        # parte igual (regla dura), pero se etiqueta con el tamano real
        # necesario para que quede visible que requiere transporte especial.
        b["camion"] = b["total"]
        return b

    def empacar(items, capacidades_local, rescate_local):
        """Empaque FIRST-FIT: mantiene todos los camiones abiertos y mete cada
        OC en el PRIMER camion donde quepa (no solo en el ultimo abierto).

        Antes esto era next-fit (un solo camion abierto): apenas una OC no
        cabia, el camion se cerraba y su espacio libre se perdia para siempre.
        Eso dejaba casos como un camion de 13 cerrado con 11 pal mientras OC
        chicas de 2 pal de la misma division armaban camiones aparte mas
        adelante en la semana. Con first-fit esas OC chicas rellenan el hueco
        y se adelantan al dia de ese camion.

        El orden de creacion de los camiones sigue siendo por prioridad, asi
        que las ventanas mas tempranas las toman igual las OC mas urgentes; lo
        unico que cambia es que una OC posterior puede subirse a un hueco
        anterior. El numero de camiones solo puede bajar, nunca subir."""
        cap_max_local = max(capacidades_local)
        items = sorted(items, key=lambda x: (orden_prioridad.index(x["prioridad"]), -x["pallets_empaque"]))
        abiertos = []
        for it in items:
            if it["pallets_empaque"] > cap_max_local:
                # OC que por si sola supera la capacidad normal: va sola y se
                # marca para que el first-fit no le meta nada (el relleno de
                # ramplas de rescate lo hace despues
                # _rellenar_ramplas_con_sobrantes, con camiones completos).
                abiertos.append({"items": [it], "total": it["pallets_empaque"], "_solo": True})
                continue
            puede_rellenar = (
                BACKFILL_MAX_PALLETS is None
                or it["pallets_empaque"] <= BACKFILL_MAX_PALLETS
            )
            destino = None
            if puede_rellenar:
                # Primer camion abierto (el mas prioritario/temprano) con espacio.
                destino = next(
                    (b for b in abiertos
                     if not b.get("_solo")
                     and b["total"] + it["pallets_empaque"] <= cap_max_local),
                    None,
                )
            elif abiertos and not abiertos[-1].get("_solo") \
                    and abiertos[-1]["total"] + it["pallets_empaque"] <= cap_max_local:
                # OC grande con backfill limitado: se comporta como antes
                # (solo intenta el ultimo camion abierto).
                destino = abiertos[-1]
            if destino is None:
                abiertos.append({"items": [it], "total": it["pallets_empaque"]})
            else:
                destino["items"].append(it)
                destino["total"] += it["pallets_empaque"]

        bins_local = []
        for b in abiertos:
            b.pop("_solo", None)
            # Si CUALQUIER OC del camion requiere esperar produccion, todo
            # el camion queda marcado para forzarse a Miercoles/Jueves en
            # asignar_ventanas (una OC nunca se parte, asi que el resto de
            # OC que comparten ese camion tambien esperan ese dia).
            b["requiere_produccion"] = any(
                it.get("requiere_produccion") for it in b["items"]
            )
            bins_local.append(cerrar(b, capacidades_local, rescate_local))
        return bins_local

    agg = agg.copy()
    agg["division"] = agg["division"].fillna("Sin división")
    bins_por_division = {}
    for division, sub in agg.groupby("division"):
        items = sub.to_dict("records")
        for it in items:
            it["division"] = division
        capacidades_local = capacidades_por_division.get(division, capacidades)
        rescate_local = capacidad_rescate if (rescate_divisiones is None or division in rescate_divisiones) else None
        bins_dv = empacar(items, capacidades_local, rescate_local)
        for b in bins_dv:
            b["division"] = division
        bins_por_division[division] = bins_dv

    # Antes de intercalar por prioridad/ventana: si alguna OC gigante ya
    # obligo a usar una rampla de rescate, aprovechamos el espacio libre
    # que le quedo con otros camiones normales completos de la misma
    # division (ver _rellenar_ramplas_con_sobrantes). Asi no se paga una
    # rampla cara a medio uso Y una ventana normal aparte para lo mismo.
    bins_por_division = _rellenar_ramplas_con_sobrantes(bins_por_division, capacidad_rescate)

    if minimo_por_division and dias and ventanas_por_dia:
        return _intercalar_con_minimo_diario(
            bins_por_division, dias, ventanas_por_dia, minimo_por_division, orden_prioridad,
        )

    # Sin minimo diario: se intercalan los camiones de todas las divisiones
    # segun prioridad global, para que las ventanas mas tempranas de la
    # semana las tomen las OC mas urgentes sin importar de que division sean.
    bins_all = [b for bins in bins_por_division.values() for b in bins]
    bins_all.sort(key=lambda b: min(orden_prioridad.index(it["prioridad"]) for it in b["items"]))
    return bins_all


def _rellenar_ramplas_con_sobrantes(bins_por_division: dict, capacidad_rescate: int | None) -> dict:
    """Aprovecha el espacio que quede libre en una rampla de rescate que ya
    se armo por FUERZA (porque una OC por si sola supera la capacidad normal
    y no se puede partir), sumandole ahi camiones normales COMPLETOS de la
    misma division que quepan enteros en ese espacio libre.

    Idea: si esa rampla (flota externa, mas cara) ya se va a pagar si o si
    por la OC gigante, es mejor llenarla al maximo con otras OC en vez de
    dejarla a medio uso Y ademas gastar una ventana/camion normal aparte
    para esas otras OC. No se crean ramplas nuevas aca (eso lo sigue
    haciendo _consolidar_con_rampla solo cuando faltan ventanas): esto solo
    reaprovecha ramplas que YA existian.

    Reglas duras que se respetan igual que en armar_camiones:
    - nunca se parte una OC (se mueven camiones COMPLETOS, no OC sueltas).
    - nunca se mezcla una division con otra (el barrido es division por
      division).
    """
    if not capacidad_rescate:
        return bins_por_division

    for division, bins in bins_por_division.items():
        ramplas = [b for b in bins if b["camion"] == capacidad_rescate]
        normales = [b for b in bins if b["camion"] != capacidad_rescate]
        if not ramplas or not normales:
            continue

        # First-fit-decreasing: probamos primero los camiones normales mas
        # grandes, asi el espacio libre de la rampla se llena mejor (menos
        # huecos) que si probamos en cualquier orden.
        normales.sort(key=lambda b: -b["total"])

        for rampla in ramplas:
            libre = capacidad_rescate - rampla["total"]
            if libre <= 0:
                continue
            i = 0
            while i < len(normales):
                candidato = normales[i]
                if candidato["total"] <= libre + 1e-6:
                    rampla["items"] = rampla["items"] + candidato["items"]
                    rampla["total"] = rampla["total"] + candidato["total"]
                    rampla["requiere_produccion"] = (
                        rampla.get("requiere_produccion", False)
                        or candidato.get("requiere_produccion", False)
                    )
                    libre -= candidato["total"]
                    normales.pop(i)
                    # no avanzamos i: puede que el siguiente (mas chico)
                    # tambien quepa en lo que sobro del espacio libre.
                else:
                    i += 1

        bins_por_division[division] = ramplas + normales

    return bins_por_division


def _consolidar_con_rampla(bins: list[dict], ventanas_disponibles: int,
                            capacidad_rescate: int,
                            divisiones_elegibles: set | None = None) -> list[dict]:
    """Ultimo recurso cuando los camiones normales no alcanzan a caber en
    las ventanas disponibles de la semana (dias x ventanas/dia): fusiona
    camiones de a pares -SIEMPRE dentro de la misma division, nunca
    partiendo una OC- en ramplas de 'capacidad_rescate' pallets (27),
    empezando por los camiones MENOS prioritarios (el final de la lista,
    que ya viene ordenada de mas a menos urgente), hasta que el plan quepa
    en las ventanas disponibles o ya no queden pares fusionables.

    divisiones_elegibles (opcional, ej: {"CONSUMO MASIVO"}): si se entrega,
    solo se fusionan camiones de esas divisiones. Una division con flota
    propia y sin rampla de respaldo (como Farma, que solo tiene camiones de
    16) queda afuera: si le faltan ventanas, sus camiones quedan igual como
    overflow / "SIN VENTANA" en vez de subirse a una rampla que en la
    realidad no existe para esa division.

    Si aun asi sobran camiones (division muy desbalanceada, por ejemplo),
    esos quedan igual que antes: como overflow / "SIN VENTANA" en
    asignar_ventanas."""
    if ventanas_disponibles <= 0 or len(bins) <= ventanas_disponibles:
        return bins

    bins = list(bins)
    cambiado = True
    while len(bins) > ventanas_disponibles and cambiado:
        cambiado = False
        for i in range(len(bins) - 1, 0, -1):
            b2 = bins[i]
            if divisiones_elegibles is not None and b2["division"] not in divisiones_elegibles:
                continue
            for j in range(i - 1, -1, -1):
                b1 = bins[j]
                if (b1["division"] == b2["division"]
                        and b1["total"] + b2["total"] <= capacidad_rescate):
                    b1["items"] = b1["items"] + b2["items"]
                    b1["total"] = b1["total"] + b2["total"]
                    b1["camion"] = capacidad_rescate
                    b1["requiere_produccion"] = (
                        b1.get("requiere_produccion", False)
                        or b2.get("requiere_produccion", False)
                    )
                    del bins[i]
                    cambiado = True
                    break
            if cambiado:
                break
    return bins


def asignar_ventanas(bins: list[dict], semana: int, anio: int,
                      dias: list[str], ventanas_por_dia: int,
                      usa_transportes: bool = False, n_transportes: int = 3,
                      capacidad_rescate: int | None = None,
                      dias_preferidos_rampla: set | None = None,
                      dias_preferidos_produccion: set | None = None):
    """Genera los "slots" (dia + ventana) donde se ubica cada camion.

    Modo normal (SB): dias x ventanas_por_dia es un tope fijo; si sobran
    camiones, quedan "SIN VENTANA" (overflow).

    Modo transportes (PU): no hay tope de ventanas por dia. Los camiones se
    reparten ciclicamente entre n_transportes hasta completar TODO el
    despacho ese dia (nunca hay overflow, cada transporte hace las vueltas
    que se necesiten).

    Si se entrega capacidad_rescate, cualquier camion que haya quedado con
    ese tamano (una rampla) intenta ubicarse PRIMERO en dias_preferidos_rampla
    (por defecto Miercoles/Jueves, que es cuando en la practica se consigue
    ese transporte externo). Si esos dias no tienen cupo o no estan dentro
    de los dias habilitados esa semana, la rampla igual se despacha, solo
    que cae en el resto de los dias como antes.

    De la misma forma, cualquier camion que traiga al menos una OC marcada
    "requiere_produccion" (columna "Producción" del Refresh: productos que
    hay que esperar que produccion fabrique) se fuerza IGUAL a ubicarse
    primero en dias_preferidos_produccion (por defecto tambien
    Miercoles/Jueves), sin importar que prioridad de despacho le haya
    tocado -aunque sea "1 - Solicitado = 1er Posible (completo)"-. Ambos
    tipos de camion forzado (rampla y produccion) compiten por el mismo cupo
    de dias preferidos, en el orden en que aparecen en `bins` (que ya viene
    ordenado por prioridad). El resto de los camiones (no forzados) rellenan
    los cupos que van quedando, en orden cronologico de dia/ventana."""
    lunes = datetime.date.fromisocalendar(anio, semana, 1)
    dia_offset = {"Lunes": 0, "Martes": 1, "Miercoles": 2, "Miércoles": 2,
                  "Jueves": 3, "Viernes": 4, "Sabado": 5, "Sábado": 5, "Domingo": 6}

    if usa_transportes:
        pares_dia_transporte = [
            (d, t) for d in dias for t in range(1, n_transportes + 1)
        ]
        slots = []
        for i in range(len(bins)):
            d, t = pares_dia_transporte[i % len(pares_dia_transporte)]
            fecha = lunes + datetime.timedelta(days=dia_offset.get(d, 0))
            slots.append({"dia": d, "fecha": fecha, "ventana": f"Transporte {t}"})
        overflow = False
        asignacion = {i: i for i in range(len(bins))}
    else:
        slots = []
        for d in dias:
            fecha = lunes + datetime.timedelta(days=dia_offset.get(d, 0))
            for v in range(1, _ventanas_del_dia(ventanas_por_dia, d) + 1):
                slots.append({"dia": d, "fecha": fecha, "ventana": v})
        overflow = len(bins) > len(slots)

        hay_produccion = any(b.get("requiere_produccion") for b in bins)
        if capacidad_rescate or hay_produccion:
            dias_pref = (
                set(dias_preferidos_rampla or set())
                | set(dias_preferidos_produccion or set())
            ) or {"Miercoles", "Miércoles", "Jueves"}
            idx_pref = [i for i, s in enumerate(slots) if s["dia"] in dias_pref]
            idx_resto = [i for i, s in enumerate(slots) if s["dia"] not in dias_pref]

            asignacion = {}
            # 1) las ramplas y/o los camiones con OC que requieren esperar
            # produccion (en el orden de prioridad que ya traian) se ubican
            # primero en los dias preferidos; si se acaban, siguen con el
            # resto de los dias en orden.
            for i, b in enumerate(bins):
                es_rampla = capacidad_rescate and b.get("camion") == capacidad_rescate
                es_produccion = b.get("requiere_produccion")
                if es_rampla or es_produccion:
                    if idx_pref:
                        asignacion[i] = idx_pref.pop(0)
                    elif idx_resto:
                        asignacion[i] = idx_resto.pop(0)
            # 2) el resto de los camiones (sin rampla ni produccion) rellenan
            # los cupos que vayan quedando, en orden cronologico de dia/ventana.
            libres = sorted(idx_pref + idx_resto)
            for i, b in enumerate(bins):
                if i not in asignacion:
                    if libres:
                        asignacion[i] = libres.pop(0)
        else:
            asignacion = {i: i for i in range(len(bins)) if i < len(slots)}

    for i, b in enumerate(bins):
        slot_idx = asignacion.get(i)
        slot = slots[slot_idx] if slot_idx is not None else {"dia": "SIN VENTANA", "fecha": None, "ventana": "-"}
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
                  orden_prioridad: list[int], facturados: dict | None = None,
                  cfg: dict | None = None):
    cfg = cfg or HOJAS_CONFIG["SB"]
    facturados = facturados or {}
    df_camion, df_directos = separar_directos(df, semana, cfg)
    tabla_directos = resumen_directos(df_directos, cfg)
    agg = agrupar_por_oc(df_camion, pallet_col, cfg)
    if agg.empty:
        return None, None, None, tabla_directos

    capacidad_rescate = cfg.get("capacidad_rescate")
    bins = armar_camiones(
        agg, capacidades, orden_prioridad, dias=dias, ventanas_por_dia=ventanas_por_dia,
        minimo_por_division=cfg.get("minimo_por_division"),
        capacidad_rescate=capacidad_rescate,
        capacidades_por_division=cfg.get("capacidades_por_division"),
        rescate_divisiones=cfg.get("rescate_divisiones"),
    )

    if capacidad_rescate and not cfg.get("usa_transportes", False):
        ventanas_disponibles = (
            sum(_ventanas_del_dia(ventanas_por_dia, d) for d in dias)
            if isinstance(ventanas_por_dia, dict) else len(dias) * ventanas_por_dia
        )
        bins = _consolidar_con_rampla(
            bins, ventanas_disponibles, capacidad_rescate,
            divisiones_elegibles=cfg.get("rescate_divisiones"),
        )

    bins, slots, overflow = asignar_ventanas(
        bins, semana, anio, dias, ventanas_por_dia,
        usa_transportes=cfg.get("usa_transportes", False),
        n_transportes=cfg.get("n_transportes", 3),
        capacidad_rescate=capacidad_rescate,
        dias_preferidos_rampla=cfg.get("dias_preferidos_rampla"),
        dias_preferidos_produccion=cfg.get("dias_preferidos_produccion"),
    )

    resumen = pd.DataFrame([{
        "Camión #": b["camion_num"], "Día": b["dia"], "Fecha": b["fecha"],
        "Ventana": b["ventana"], "División": b["division"],
        "Tipo camión (pallets)": b["camion"],
        "Pallets cargados": round(b["total"], 2), "Capacidad": b["camion"],
        "Utilización %": round(b["total"] / b["camion"] * 100, 1),
        "# OCs": len(b["items"]),
        "# SKUs": sum(it.get("n_sku", 0) for it in b["items"]),
        "Pedidos incluidos": ", ".join(str(it["Pedido"]) for it in b["items"]),
    } for b in bins])

    detalle_rows = []
    for b in bins:
        for it in b["items"]:
            estado_fact = facturados.get(str(it["Pedido"]).strip(), "No")
            detalle_rows.append({
                "Pedido (OC)": it["Pedido"], "OC": it["oc"], "Fecha vence": it["fecha_vence"],
                "División": it["division"],
                "Prioridad": it["prioridad"], "Descripción prioridad": it["prioridad_label"],
                "Solicitado": it.get("sol", 0), "1 Posible": it.get("pos1", 0),
                "Pronto-vence": it.get("pv", 0),
                "Pallets": it["pallets"], "Monto": it.get("monto", 0), "# SKUs": it["n_sku"],
                "Camión #": it["camion_num"], "Día": it["dia"], "Fecha": it.get("fecha"),
                "Ventana": it["ventana"], "Facturado": estado_fact,
                "Requiere Producción": "Sí" if it.get("requiere_produccion") else "No",
                "Campaña": it.get("campana") or "",
            })
    detalle = pd.DataFrame(detalle_rows)

    # Reordena AMBAS tablas (resumen y detalle) de forma CRONOLOGICA
    # (Fecha real -> Ventana), en vez de dejarlas en el orden en que se
    # crearon los camiones (que sigue la prioridad de despacho: 1, 2, 5, 3).
    # Sin esto, un camion de prioridad alta que igual sale mas tarde en la
    # semana (ej: por la regla de Miercoles/Jueves para rampla o para OC que
    # esperan produccion) podia aparecer ANTES en la tabla que camiones de
    # dias anteriores, lo que confundia al leerla de arriba a abajo. De paso
    # se renumera "Camión #" para que el numero tambien seas consecutivo en
    # ese mismo orden cronologico (Camión #1 = el primero que sale en la
    # semana), y se propaga el mapeo de numeros al detalle para que ambas
    # tablas (y el Excel exportado) queden consistentes entre si.
    def _clave_ventana(v):
        try:
            return (0, int(v))
        except (TypeError, ValueError):
            return (1, str(v))

    _fecha_max = datetime.date.max
    resumen = resumen.assign(
        _orden_fecha=resumen["Fecha"].apply(lambda f: f if pd.notna(f) else _fecha_max),
        _orden_ventana=resumen["Ventana"].apply(_clave_ventana),
    ).sort_values(by=["_orden_fecha", "_orden_ventana"]).reset_index(drop=True)

    mapa_camion = {int(old): i + 1 for i, old in enumerate(resumen["Camión #"])}
    resumen["Camión #"] = resumen["Camión #"].map(mapa_camion)
    resumen = resumen.drop(columns=["_orden_fecha", "_orden_ventana"])

    if not detalle.empty:
        detalle = detalle.assign(
            _orden_fecha=detalle["Fecha"].apply(lambda f: f if pd.notna(f) else _fecha_max),
            _orden_ventana=detalle["Ventana"].apply(_clave_ventana),
        ).sort_values(by=["_orden_fecha", "_orden_ventana", "Prioridad"]).reset_index(drop=True)
        detalle["Camión #"] = detalle["Camión #"].map(mapa_camion)
        detalle = detalle.drop(columns=["_orden_fecha", "_orden_ventana"])

    info = {"camiones": len(bins), "ventanas_disponibles": len(slots), "overflow": overflow,
            "oc_directos": tabla_directos["Pedido"].nunique() if not tabla_directos.empty else 0,
            "lineas_directos": len(tabla_directos),
            "oc_facturadas": int((detalle["Facturado"] == "Sí").sum()),
            "oc_parciales": int((detalle["Facturado"] == "Parcial").sum()),
            "oc_produccion": int(detalle.loc[detalle["Requiere Producción"] == "Sí", "Pedido (OC)"].nunique())
            if "Requiere Producción" in detalle.columns else 0}
    return resumen, detalle, info, tabla_directos


def _resaltar_facturado_factory(cap_max: float | None = None, cap_por_division: dict | None = None):
    """Devuelve una función de estilo por fila para el .style.apply() de la
    tabla 'Detalle por OC'. Si cap_max (o cap_por_division, para la
    división de esa fila puntual) viene informado, cualquier OC cuyos
    Pallets superen esa capacidad máxima de transporte disponible se pinta
    de rojo oscuro (tiene prioridad visual por sobre Facturado/Parcial,
    porque esa OC no puede despacharse tal como está — hay que corregirla
    o dividirla)."""
    cap_por_division = cap_por_division or {}

    def _fn(row):
        estado = row.get("Facturado")
        cap_row = cap_por_division.get(row.get("División"), cap_max)
        excede = cap_row is not None and row.get("Pallets", 0) > cap_row
        if excede:
            style = "background-color: #5C0A0A; color: #FFD9D9; font-weight: 700"
        elif estado == "Sí":
            style = "background-color: #C6EFCE; color: #0b3d24"
        elif estado == "Parcial":
            style = "background-color: #FDE3B8; color: #5C3A0B"
        else:
            style = ""
        return [style] * len(row)
    return _fn


# Se mantiene el nombre viejo por compatibilidad (sin el chequeo de capacidad).
_resaltar_facturado = _resaltar_facturado_factory()


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
        "<div style='display:flex;align-items:center;gap:0.4rem;margin-right:1.1rem;'>"
        "<span style='background:#C6EFCE;color:#0b3d24;font-size:0.6rem;"
        "font-weight:700;padding:0.05rem 0.4rem;border-radius:999px;'>FACTURADO</span>"
        "<span style='font-size:0.74rem;color:#C7D2E0;'>= 100% despachado</span>"
        "</div>"
        "<div style='display:flex;align-items:center;gap:0.4rem;margin-right:1.1rem;'>"
        "<span style='background:#FDE3B8;color:#5C3A0B;font-size:0.6rem;"
        "font-weight:700;padding:0.05rem 0.4rem;border-radius:999px;'>PARCIAL</span>"
        "<span style='font-size:0.74rem;color:#C7D2E0;'>= algunas líneas despachadas, no todas</span>"
        "</div>"
        "<div style='display:flex;align-items:center;gap:0.4rem;'>"
        "<span style='background:#5C0A0A;color:#FFD9D9;font-size:0.6rem;"
        "font-weight:700;padding:0.05rem 0.4rem;border-radius:999px;'>⚠️ EXCEDE CAPACIDAD</span>"
        "<span style='font-size:0.74rem;color:#C7D2E0;'>= una sola OC ya supera el transporte "
        "más grande disponible → debe cancelarse tal como está</span>"
        "</div>"
        "</div>"
    )
    st.markdown(leyenda_html, unsafe_allow_html=True)


def render_calendario(detalle: pd.DataFrame, resumen: pd.DataFrame | None = None,
                       cap_max: float | None = None, cap_por_division: dict | None = None):
    """Vista tipo calendario/kanban: una columna por dia, con un KPI de
    despacho arriba (camiones, utilizacion y Facturados vs No) y tarjetas por
    OC con Pedido, OC, Monto y Pallets."""
    cap_por_division = cap_por_division or {}
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

        # KPI Facturados vs Parcial vs No del dia
        n_facturados = int((sub["Facturado"] == "Sí").sum())
        n_parciales = int((sub["Facturado"] == "Parcial").sum())
        n_no_facturados = len(sub) - n_facturados - n_parciales
        total_sub = len(sub) if len(sub) else 1
        pct_facturado = n_facturados / total_sub * 100
        pct_parcial = n_parciales / total_sub * 100

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
                f"margin-bottom:0.25rem;'>Facturados vs Parcial vs No</div>"
                f"<div style='display:flex;width:100%;height:8px;border-radius:4px;"
                f"overflow:hidden;background:#E4572E33;margin-bottom:0.25rem;'>"
                f"<div style='width:{pct_facturado:.0f}%;background:#1DB980;'></div>"
                f"<div style='width:{pct_parcial:.0f}%;background:#F2994A;'></div>"
                f"</div>"
                f"<div style='display:flex;justify-content:space-between;font-size:0.62rem;'>"
                f"<span style='color:#1DB980;font-weight:700;'>✅ {n_facturados}</span>"
                f"<span style='color:#F2994A;font-weight:700;'>◐ {n_parciales}</span>"
                f"<span style='color:#E4572E;font-weight:700;'>⏳ {n_no_facturados}</span>"
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
                        cap_row = cap_por_division.get(row.get("División"), cap_max)
                        excede = cap_row is not None and row["Pallets"] > cap_row
                        color = "#5C0A0A" if excede else _COLOR_PRIORIDAD.get(row["Prioridad"], "#888888")
                        if row["Facturado"] == "Sí":
                            chip = (
                                "<span style='background:#C6EFCE;color:#0b3d24;font-size:0.6rem;"
                                "font-weight:700;padding:0.05rem 0.4rem;border-radius:999px;"
                                "white-space:nowrap;margin-left:0.4rem;'>FACTURADO</span>"
                            )
                        elif row["Facturado"] == "Parcial":
                            chip = (
                                "<span style='background:#FDE3B8;color:#5C3A0B;font-size:0.6rem;"
                                "font-weight:700;padding:0.05rem 0.4rem;border-radius:999px;"
                                "white-space:nowrap;margin-left:0.4rem;'>PARCIAL</span>"
                            )
                        else:
                            chip = ""
                        if row.get("Requiere Producción") == "Sí":
                            chip += (
                                "<span style='background:#7C3AED26;color:#C4B5FD;font-size:0.6rem;"
                                "font-weight:700;padding:0.05rem 0.4rem;border-radius:999px;"
                                "white-space:nowrap;margin-left:0.4rem;'>🏭 PRODUCCIÓN</span>"
                            )
                        # El aviso de "excede capacidad" va en su propia franja debajo
                        # del titulo (no como chip en linea) para que no se corte / envuelva.
                        banner_excede = (
                            "<div style='background:#5C0A0A40;border:1px solid #5C0A0A;"
                            "border-radius:4px;padding:0.2rem 0.5rem;margin-top:0.35rem;"
                            "font-size:0.68rem;font-weight:700;color:#FF8A80;"
                            "white-space:nowrap;overflow:hidden;text-overflow:ellipsis;'>"
                            f"⚠️ EXCEDE CAPACIDAD ({cap_row:.0f} pal máx.) — Cancelar OC"
                            "</div>"
                        ) if excede else ""
                        tarjeta_html = (
                            f"<div style='background:#141B2D;border-left:4px solid {color};"
                            "border-radius:6px;padding:0.5rem 0.7rem;margin-bottom:0.5rem;"
                            "box-shadow:0 1px 2px rgba(0,0,0,0.2);'>"
                            "<div style='font-weight:700;font-size:0.82rem;color:#F5F7FA;"
                            "display:flex;align-items:center;flex-wrap:wrap;gap:0.2rem;'>"
                            f"<span>Pedido {row['Pedido (OC)']}</span>{chip}"
                            "</div>"
                            f"<div style='font-size:0.72rem;color:#8494AC;'>OC {row['OC']}</div>"
                            "<div style='display:flex;justify-content:space-between;"
                            "margin-top:0.3rem;font-size:0.76rem;color:#C7D2E0;'>"
                            f"<span>{formato_clp(row['Monto'])}</span>"
                            f"<span style='color:{'#FF6B6B' if excede else '#3B9EFF'};font-weight:600;'>"
                            f"{row['Pallets']:.0f} pal</span>"
                            "</div>"
                            f"{banner_excede}"
                            "</div>"
                        )
                        st.markdown(tarjeta_html, unsafe_allow_html=True)


def render_tabla_camiones(resumen: pd.DataFrame, detalle: pd.DataFrame,
                           cap_max: float | None = None, cap_por_division: dict | None = None):
    """Tabla 'Plan de camiones' en HTML (para poder pintar en verde, dentro
    de la misma celda, los numeros de Pedido que ya estan Facturados) mas
    una columna extra de % Facturado por camion. Si cap_max (o, para cada
    fila, cap_por_division segun su División) viene informado, las
    filas/pedidos que superan esa capacidad maxima real se pintan en rojo
    para detectarlas al toque (ese "camion" es ficticio: el sistema le puso
    el tamaño de la OC porque no entraba en ningun transporte real)."""
    cap_por_division = cap_por_division or {}
    facturado_map = dict(zip(detalle["Pedido (OC)"].astype(str), detalle["Facturado"]))
    pallets_map = dict(zip(detalle["Pedido (OC)"].astype(str), detalle["Pallets"]))
    division_map = dict(zip(detalle["Pedido (OC)"].astype(str), detalle["División"]))

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
        cap_fila = cap_por_division.get(row.get("División"), cap_max)
        camion_excede = cap_fila is not None and row["Tipo camión (pallets)"] > cap_fila

        def _chip(p):
            cap_p = cap_por_division.get(division_map.get(p), cap_max)
            excede_p = cap_p is not None and pallets_map.get(p, 0) > cap_p
            estado = facturado_map.get(p)
            if excede_p:
                return (f"<span style='background:#5C0A0A;color:#FFD9D9;font-weight:700;"
                        f"border-radius:4px;padding:0 0.3rem;'>⚠️ {p}</span>")
            if estado == "Sí":
                return (f"<span style='background:#C6EFCE;color:#0b3d24;font-weight:700;"
                        f"border-radius:4px;padding:0 0.3rem;'>{p}</span>")
            if estado == "Parcial":
                return (f"<span style='background:#FDE3B8;color:#5C3A0B;font-weight:700;"
                        f"border-radius:4px;padding:0 0.3rem;'>{p}</span>")
            return f"<span>{p}</span>"

        pedidos_html = ", ".join(_chip(p) for p in pedidos)
        fila_style = " style='background:#5C0A0A26;'" if camion_excede else ""
        celdas = "".join(f"<td>{row[c]}</td>" for c in cols_base[:5])
        color_tipo = "color:#FF8A80;font-weight:700;" if camion_excede else ""
        celdas += (
            f"<td style='text-align:right;{color_tipo}'>{row['Tipo camión (pallets)']:.0f}"
            f"{' ⚠️' if camion_excede else ''}</td>"
            f"<td style='text-align:right;'>{row['Pallets cargados']:.0f}</td>"
            f"<td style='text-align:right;'>{row['Capacidad']:.0f}</td>"
            f"<td style='text-align:right;'>{row['Utilización %']:.1f}%</td>"
            f"<td style='text-align:right;color:#1DB980;font-weight:700;'>{pct_fact:.0f}%</td>"
            f"<td>{pedidos_html}</td>"
        )
        filas_html.append(f"<tr{fila_style}>{celdas}</tr>")

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
    """Excel con las OC que NO estan 100% facturadas (Facturado = No o
    Parcial): Hoja 'Resumen' = tabla pivote División (filas) x Día de
    despacho (columnas), igual estilo a la tabla de referencia (categorías
    al costado, arriba); luego una pestaña por día con el detalle completo
    (incluye la columna Facturado para distinguir Parcial de No)."""
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

    pendientes = detalle[detalle["Facturado"] != "Sí"].copy()
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
            ws_dia = writer.sheets[nombre_hoja]
            _formatear_hoja_detalle(ws_dia, sub)
            if "Facturado" in sub.columns:
                naranjo = PatternFill("solid", fgColor="FDE3B8")
                col_fact = sub.columns.get_loc("Facturado") + 1
                for r, val in enumerate(sub["Facturado"], start=2):
                    if val == "Parcial":
                        for c in range(1, len(sub.columns) + 1):
                            ws_dia.cell(row=r, column=c).fill = naranjo

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
            naranjo = PatternFill("solid", fgColor="FDE3B8")
            for r, val in enumerate(detalle["Facturado"], start=2):  # fila 1 = encabezado
                relleno = verde if val == "Sí" else naranjo if val == "Parcial" else None
                if relleno:
                    for c in range(1, len(detalle.columns) + 1):
                        ws.cell(row=r, column=c).fill = relleno

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


def _pestana_checklist_de(division: str) -> str:
    """A que pestaña del checklist de carga va cada division real del
    Refresh: todo lo que contenga 'FARMA' -> 'Farma'; cualquier otra
    division (Consumo Masivo, etc.) -> 'Consumo'."""
    return "Farma" if "FARMA" in str(division).upper() else "Consumo"


def exportar_checklist_carga(detalle: pd.DataFrame, semana, cfg: dict) -> bytes | None:
    """Genera el Excel de checklist de carga, con el mismo formato que usa
    Operaciones a mano en Google Sheets (título de cliente/división arriba,
    columna de verificación en blanco, numeración de carga): UNA pestaña
    por división ('Farma' para Farma, 'Consumo' para el resto), con
    TODAS las OC de la semana (esten o no facturadas). Cada pestaña trae:
    - 'Carga OC': numeración correlativa 1..N DENTRO de esa pestaña, en el
      mismo orden cronológico (Fecha -> Ventana -> Prioridad) que ya trae
      'detalle' desde generar_plan (no se reordena de nuevo aca).
    - 'Día': el día de despacho de esa OC (columna nueva, aparte del
      número de carga).
    - 'Verificador': casillero en blanco para marcar a mano al cargar
      físicamente el camión (Excel/openpyxl no soporta checkboxes nativos,
      así que se deja un recuadro vacío con borde marcado).
    - 'Camión': N° de camión de cada OC, y ademas cada vez que cambia de
      camión respecto a la fila anterior se marca con un borde superior
      grueso (reemplaza el rayado a mano que se hacía en Operaciones para
      separar los camiones).
    - 'Facturada': 'Sí' / 'Parcial' / vacío si aún no se ha facturado.
    Devuelve None si no hay OC para mostrar."""
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    if detalle is None or detalle.empty:
        return None

    d = detalle.copy()
    d["_pestana"] = d["División"].apply(_pestana_checklist_de)

    header_fill = PatternFill("solid", fgColor="0B4F86")
    header_font = Font(bold=True, color="FFFFFF")
    titulo_fill = PatternFill("solid", fgColor="FDE9D9")
    division_fill = PatternFill("solid", fgColor="FFF200")
    fact_fill = {
        "Sí": PatternFill("solid", fgColor="C6EFCE"),
        "Parcial": PatternFill("solid", fgColor="FDE3B8"),
    }
    thin = Side(style="thin", color="D9D9D9")
    thick_top = Side(style="thick", color="1F2937")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    border_check = Border(
        left=Side(style="medium"), right=Side(style="medium"),
        top=Side(style="medium"), bottom=Side(style="medium"),
    )
    # Mismo borde de siempre, pero con el lado superior grueso: marca donde
    # empieza un camion nuevo (reemplaza el rayado a mano que se hacia en
    # Operaciones para separar los camiones dentro de la pestaña).
    border_camion_nuevo = Border(left=thin, right=thin, top=thick_top, bottom=thin)
    border_check_camion_nuevo = Border(
        left=Side(style="medium"), right=Side(style="medium"),
        top=thick_top, bottom=Side(style="medium"),
    )
    centrado = Alignment(horizontal="center", vertical="center")

    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        hay_alguna_hoja = False
        for pestana in ["Farma", "Consumo"]:
            sub = d[d["_pestana"] == pestana].copy()
            if sub.empty:
                continue
            hay_alguna_hoja = True
            sub = sub.reset_index(drop=True)
            sub.insert(0, "Carga OC", range(1, len(sub) + 1))
            sub["Facturada"] = sub["Facturado"].map({"Sí": "Sí", "Parcial": "Parcial"}).fillna("")
            sub["Verificador"] = ""

            cols_orden = [
                "Carga OC", "Verificador", "Pedido (OC)", "OC", "Día",
                "Camión #", "Solicitado", "1 Posible", "Pronto-vence",
                "Pallets", "# SKUs", "Campaña", "Facturada",
            ]
            cols_orden = [c for c in cols_orden if c in sub.columns]
            vista = sub[cols_orden].rename(columns={
                "Pedido (OC)": "Pedido",
                "1 Posible": "Posible actual",
                "Pallets": "Pallet estimado",
                "# SKUs": "líneas",
                "Camión #": "Camión",
            })
            # Lista paralela de N° de camion (misma fila a fila que 'vista')
            # para poder marcar donde empieza cada camion nuevo, aunque la
            # columna "Camión" no se termine mostrando.
            camiones_fila = sub["Camión #"].tolist() if "Camión #" in sub.columns else [None] * len(sub)

            nombre_hoja = pestana[:31]
            vista.to_excel(writer, sheet_name=nombre_hoja, index=False, startrow=2)
            ws = writer.sheets[nombre_hoja]

            n_filas, n_cols = vista.shape
            ultima_letra = get_column_letter(n_cols) if n_cols else "A"

            # Fila 1: nombre de cliente/pestaña. Fila 2: división + semana.
            ws.merge_cells(f"A1:{ultima_letra}1")
            c1 = ws.cell(row=1, column=1, value=pestana)
            c1.font = Font(bold=True, size=16, color="1F2937")
            c1.alignment = centrado
            c1.fill = titulo_fill
            ws.row_dimensions[1].height = 26

            division_txt = "FARMA" if pestana == "Farma" else "CONSUMO MASIVO"
            ws.merge_cells(f"A2:{ultima_letra}2")
            c2 = ws.cell(row=2, column=1, value=f"{division_txt} · Semana {semana}")
            c2.font = Font(bold=True, size=12, color="1F2937")
            c2.alignment = centrado
            c2.fill = division_fill
            ws.row_dimensions[2].height = 20

            fila_header = 3  # startrow=2 (0-based) -> fila 3 real en Excel
            for c in range(1, n_cols + 1):
                cell = ws.cell(row=fila_header, column=c)
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = centrado
                cell.border = border

            col_facturada = vista.columns.get_loc("Facturada") + 1 if "Facturada" in vista.columns else None
            col_verif = vista.columns.get_loc("Verificador") + 1 if "Verificador" in vista.columns else None
            for i, r in enumerate(range(fila_header + 1, fila_header + n_filas + 1)):
                relleno_fila = None
                if col_facturada:
                    val = ws.cell(row=r, column=col_facturada).value
                    relleno_fila = fact_fill.get(val)
                # Nuevo camion = cambia el N° de camion respecto a la fila
                # anterior (la primera fila de la pestaña no cuenta, ya que
                # el encabezado ya la separa).
                camion_nuevo = i > 0 and camiones_fila[i] != camiones_fila[i - 1]
                for c in range(1, n_cols + 1):
                    cell = ws.cell(row=r, column=c)
                    cell.alignment = centrado
                    if col_verif and c == col_verif:
                        cell.border = border_check_camion_nuevo if camion_nuevo else border_check
                    else:
                        cell.border = border_camion_nuevo if camion_nuevo else border
                    if relleno_fila:
                        cell.fill = relleno_fila

            for idx_col, nombre_col in enumerate(vista.columns, start=1):
                letra = get_column_letter(idx_col)
                largo = max(
                    [len(str(nombre_col))] + [len(str(v)) for v in vista[nombre_col].astype(str)]
                ) if n_filas else len(str(nombre_col))
                ws.column_dimensions[letra].width = min(max(largo + 3, 10), 40)
            ws.column_dimensions["A"].width = max(ws.column_dimensions["A"].width or 0, 10)

            ws.freeze_panes = f"A{fila_header + 1}"

        if not hay_alguna_hoja:
            return None

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



def calcular_costos_flota(resumen: pd.DataFrame, cfg: dict, vueltas_por_camion: int) -> dict | None:
    """Estima el costo de la flota usada esta semana:
    - camion normal: valor FIJO (cfg['costo_camion']) que ya cubre hasta
      'vueltas_por_camion' vueltas ese dia (se pague 1 o 2 vueltas, el costo
      es el mismo) -> el numero de camiones FISICOS por dia es
      ceil(ventanas_normales_usadas_ese_dia / vueltas_por_camion).
    - rampla: valor por vuelta (cfg['costo_rampla']), y en este modelo cada
      rampla que aparece en el plan corresponde a 1 vuelta.
    Devuelve None si a la hoja le faltan los valores de costo en cfg."""
    costo_camion = cfg.get("costo_camion")
    costo_rampla = cfg.get("costo_rampla")
    capacidad_rescate = cfg.get("capacidad_rescate")
    if not (costo_camion and costo_rampla and capacidad_rescate):
        return None

    vpc = max(1, int(vueltas_por_camion or 1))
    es_rampla = resumen["Tipo camión (pallets)"] == capacidad_rescate
    resumen_rampla = resumen[es_rampla]
    resumen_normal = resumen[~es_rampla]

    n_ramplas = int(len(resumen_rampla))
    camiones_por_dia = (
        resumen_normal.groupby("Día").size().apply(lambda n: math.ceil(n / vpc))
        if not resumen_normal.empty else pd.Series(dtype=int)
    )
    n_camiones_fisicos = int(camiones_por_dia.sum())

    costo_camiones_total = n_camiones_fisicos * costo_camion
    costo_ramplas_total = n_ramplas * costo_rampla

    return {
        "n_camiones_fisicos": n_camiones_fisicos,
        "costo_camiones_total": costo_camiones_total,
        "n_ramplas": n_ramplas,
        "costo_ramplas_total": costo_ramplas_total,
        "costo_total": costo_camiones_total + costo_ramplas_total,
        "ramplas_por_dia": resumen_rampla["Día"].value_counts().to_dict(),
    }


def render_kpis_avanzados(resumen: pd.DataFrame, detalle: pd.DataFrame,
                           tabla_directos: pd.DataFrame, cfg: dict,
                           vueltas_por_camion: int, dias_orden: list[str], key_ns: str):
    """Bloque de KPIs/gráficos adicionales: costo de flota (camión vs
    rampla), monto y cantidad de OC por categoría, OC en riesgo por
    vencimiento, carga por día y complejidad de picking (SKUs/camión)."""

    # --- 1) Costo de flota (camiones vs ramplas) ---------------------------
    costos = calcular_costos_flota(resumen, cfg, vueltas_por_camion)
    if costos:
        st.markdown("##### 💰 Costo estimado de flota")
        dias_pref = cfg.get("dias_preferidos_rampla") or set()
        ramplas_pref = sum(v for d, v in costos["ramplas_por_dia"].items() if d in dias_pref)
        ramplas_otros = costos["n_ramplas"] - ramplas_pref
        render_kpi_cards([
            {"value": costos["n_camiones_fisicos"], "label": "Camiones físicos usados",
             "badge_text": formato_clp(costos["costo_camiones_total"]), "badge_color": "#3B9EFF"},
            {"value": costos["n_ramplas"], "label": "Ramplas usadas",
             "badge_text": formato_clp(costos["costo_ramplas_total"]),
             "badge_color": "#E4572E" if costos["n_ramplas"] else "#1DB980"},
            {"value": f"{ramplas_pref}/{costos['n_ramplas']}", "label": "Ramplas en día preferido",
             "badge_text": "Miér./Jue." if dias_pref else None, "badge_color": "#1DB980"},
            {"value": formato_clp(costos["costo_total"]), "label": "Costo total flota semana"},
        ])
        if costos["ramplas_por_dia"]:
            detalle_dias = ", ".join(f"{d}: {n}" for d, n in sorted(
                costos["ramplas_por_dia"].items(),
                key=lambda kv: dias_orden.index(kv[0]) if kv[0] in dias_orden else 99))
            st.caption(f"🆘 Ramplas por día → {detalle_dias}."
                       + (f" ⚠️ {ramplas_otros} fuera de Miércoles/Jueves." if ramplas_otros else ""))

    # --- 2) Monto y cantidad de OC por categoría ---------------------------
    st.markdown("##### 📦 OC y monto por categoría")
    if "Monto" in detalle.columns:
        filas = []
        total_oc, total_fact, total_pend = 0, 0.0, 0.0
        for div, sub in detalle.groupby("División"):
            n_oc = sub["Pedido (OC)"].nunique()
            m_fact = sub.loc[sub["Facturado"] == "Sí", "Monto"].sum()
            m_pend = sub.loc[sub["Facturado"] != "Sí", "Monto"].sum()
            filas.append({"Categoría": div, "# OC": n_oc,
                          "Monto facturado": m_fact, "Monto pendiente": m_pend})
            total_oc += n_oc
            total_fact += m_fact
            total_pend += m_pend
        n_oc_directo = tabla_directos["Pedido"].nunique() if not tabla_directos.empty else 0
        m_directo = tabla_directos["Monto"].sum() if ("Monto" in tabla_directos.columns and not tabla_directos.empty) else 0.0
        filas.append({"Categoría": "DIRECTOS (no van en camión)", "# OC": n_oc_directo,
                      "Monto facturado": None, "Monto pendiente": m_directo})
        filas.insert(0, {"Categoría": "TOTAL (camión + directos)", "# OC": total_oc + n_oc_directo,
                         "Monto facturado": total_fact, "Monto pendiente": total_pend + m_directo})
        df_cat = pd.DataFrame(filas)
        st.dataframe(
            df_cat.style.format({"Monto facturado": lambda v: formato_clp(v) if pd.notna(v) else "—",
                                 "Monto pendiente": lambda v: formato_clp(v) if pd.notna(v) else "—"}),
            use_container_width=True, hide_index=True,
        )

    # --- 3) OC en riesgo por vencimiento, aún sin facturar ------------------
    st.markdown("##### ⏰ OC en riesgo por vencimiento (sin facturar)")
    det_pend = detalle[detalle["Facturado"] != "Sí"].copy()
    det_pend["_venc"] = pd.to_datetime(det_pend["Fecha vence"], errors="coerce")
    det_pend = det_pend.dropna(subset=["_venc"])
    hoy = pd.Timestamp(datetime.date.today())
    det_pend["Días para vencer"] = (det_pend["_venc"] - hoy).dt.days
    riesgo = det_pend[det_pend["Días para vencer"] <= 7].sort_values("Días para vencer")
    if riesgo.empty:
        st.success("✅ Ninguna OC pendiente vence en los próximos 7 días.")
    else:
        monto_riesgo = riesgo["Monto"].sum() if "Monto" in riesgo.columns else 0
        n_vencidas = int((riesgo["Días para vencer"] < 0).sum())
        render_kpi_cards([
            {"value": len(riesgo), "label": "OC en riesgo (≤7 días, sin facturar)"},
            {"value": n_vencidas, "label": "Ya vencidas",
             "badge_text": "Revisar ya" if n_vencidas else None, "badge_color": "#E4572E"},
            {"value": formato_clp(monto_riesgo), "label": "Monto en riesgo"},
        ])
        cols_riesgo = ["Pedido (OC)", "División", "Prioridad", "Fecha vence",
                       "Días para vencer", "Monto", "Facturado"]
        cols_riesgo = [c for c in cols_riesgo if c in riesgo.columns]
        st.dataframe(
            riesgo[cols_riesgo].style.format({"Monto": lambda v: formato_clp(v)}),
            use_container_width=True, hide_index=True,
        )

    # --- 4) Carga por día (pallets cargados vs capacidad, por división) ----
    st.markdown("##### 📊 Carga por día")
    resumen_dia = resumen[resumen["Día"] != "SIN VENTANA"].copy()
    if not resumen_dia.empty:
        resumen_dia["Día"] = pd.Categorical(resumen_dia["Día"], categories=dias_orden, ordered=True)
        grp = resumen_dia.groupby(["Día", "División"], observed=True).agg(
            **{"Pallets cargados": ("Pallets cargados", "sum"), "Capacidad total": ("Capacidad", "sum")}
        ).reset_index()
        fig_carga = px.bar(
            grp, x="Día", y="Pallets cargados", color="División",
            barmode="stack", text_auto=",.0f",
        )
        cap_dia = resumen_dia.groupby("Día", observed=True)["Capacidad"].sum().reset_index()
        fig_carga.add_trace(go.Scatter(
            x=cap_dia["Día"], y=cap_dia["Capacidad"], mode="lines+markers",
            name="Capacidad disponible", line=dict(dash="dot", color="#F2C94C"),
        ))
        fig_carga.update_layout(
            template="plotly_dark", paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
            margin=dict(t=10, b=10, l=10, r=10), height=340, yaxis_title="Pallets",
        )
        st.plotly_chart(fig_carga, use_container_width=True, key=f"carga_dia_{key_ns}")

    # --- 5) Complejidad de picking (SKUs por camión) ------------------------
    if "# SKUs" in resumen.columns:
        st.markdown("##### 🧩 Complejidad de picking (SKUs por camión)")
        prom_sku = resumen["# SKUs"].mean()
        umbral = prom_sku + 1.5 * resumen["# SKUs"].std(ddof=0) if len(resumen) > 1 else prom_sku
        n_complejos = int((resumen["# SKUs"] > umbral).sum())
        render_kpi_cards([
            {"value": f"{prom_sku:.1f}", "label": "SKUs promedio por camión"},
            {"value": n_complejos, "label": "Camiones de picking complejo",
             "badge_text": f"> {umbral:.0f} SKUs" if n_complejos else None, "badge_color": "#E4572E"},
        ])
        top_sku = resumen.sort_values("# SKUs", ascending=False).head(10).sort_values("# SKUs")
        etiqueta = "Camión " + top_sku["Camión #"].astype(str) + " (" + top_sku["Día"].astype(str) + ", " + top_sku["División"].astype(str) + ")"
        fig_sku = px.bar(
            top_sku.assign(_etiqueta=etiqueta), x="# SKUs", y="_etiqueta", orientation="h",
            text_auto=",.0f",
            color=(top_sku["# SKUs"] > umbral).map({True: "Complejo", False: "Normal"}),
            color_discrete_map={"Complejo": "#E4572E", "Normal": "#3B9EFF"},
        )
        fig_sku.update_traces(textfont_size=11, textposition="outside", cliponaxis=False)
        fig_sku.update_layout(
            template="plotly_dark", paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
            margin=dict(t=10, b=10, l=10, r=10), height=320, xaxis_title="", yaxis_title="",
            legend_title="",
        )
        st.plotly_chart(fig_sku, use_container_width=True, key=f"sku_camion_{key_ns}")


def render_plan_hoja(archivo, cfg: dict):
    """Cuerpo completo del plan de despacho (KPIs, camiones, calendario,
    tabla, exportables) para UNA hoja del Refresh (SB o PU). key_ns evita que
    los widgets de ambas pestañas choquen entre si."""
    key_ns = cfg["hoja"].lower()

    df = leer_hoja(archivo, cfg["hoja"], version=_version_archivo(archivo))
    semanas_disp = sorted(df[cfg["semana"]].dropna().unique().tolist())

    semana = st.radio(
        "Semana a planificar", semanas_disp, horizontal=True,
        index=len(semanas_disp) - 1 if semanas_disp else 0,
        key=f"semana_{key_ns}",
    )

    # El año no se pide al usuario: se infiere de la Fecha vence de esa
    # misma semana en el Refresh (necesario solo para ubicar el Lunes ISO).
    _fechas_semana = pd.to_datetime(
        df.loc[df[cfg["semana"]] == semana, cfg["fecha_vence"]], errors="coerce"
    ).dropna()
    anio = int(_fechas_semana.dt.year.mode().iloc[0]) if not _fechas_semana.empty \
        else datetime.date.today().year

    opciones_pallets = [cfg["pallets_pos"]] + ([cfg["pallets_alt"]] if cfg["pallets_alt"] else [])

    vueltas_por_camion = 1  # default para hojas sin concepto de "vueltas" (ej: PU)
    with st.sidebar:
        st.markdown(f"### ⚙️ Configuración — {cfg['hoja']}")
        if cfg.get("forzar_pallet_col"):
            # SB: la columna de pallets queda fija en "Pallets Pos." (AH del
            # Refresh) sin opcion de cambiarla, para que el numero mostrado
            # SIEMPRE calce con esa columna del Excel de origen.
            pallet_col = cfg["pallets_pos"]
            st.caption(f"📦 Pallets por OC = columna **'{pallet_col}'** del Refresh (fija).")
        else:
            pallet_col = st.radio(
                "Columna a usar para calcular pallets por OC",
                opciones_pallets,
                help="'Pallets Pos.' viene acotada (0.3/1)."
                     + (" 'Pallets posibles' es la fracción real sin redondear."
                        if cfg["pallets_alt"] else ""),
                key=f"pallet_col_{key_ns}",
            )
        capacidades = st.multiselect(
            "Capacidades de camión disponibles (pallets)",
            cfg["capacidades_opciones"], default=cfg["capacidades_default"],
            key=f"capacidades_{key_ns}",
        )
        dias = st.multiselect(
            "Días hábiles de despacho",
            cfg["dias_opciones"], default=cfg["dias_default"],
            key=f"dias_{key_ns}",
        )
        if cfg.get("usa_transportes"):
            n_transportes = st.number_input(
                "Transportes disponibles", value=cfg.get("n_transportes", 3), min_value=1,
                help="No hay ventanas fijas: cada transporte hace las vueltas que "
                     "hagan falta hasta completar todo el despacho de ese día.",
                key=f"transportes_{key_ns}",
            )
            ventanas_por_dia = 1  # no se usa en modo transportes, pero debe existir
            cfg = {**cfg, "n_transportes": int(n_transportes)}
        elif cfg["hoja"] == "SB":
            # Flota real de SB: N camiones de 13 pallets, cada uno hace
            # V vueltas por dia -> ventanas/dia = N x V. Se pide asi (camiones
            # x vueltas) en vez de un numero de "ventanas" suelto, para que
            # quede claro y a prueba de error cual es la flota real.
            col_cam, col_vta = st.columns(2)
            with col_cam:
                n_camiones = st.number_input(
                    "Camiones disponibles", value=cfg.get("n_camiones_default", 3),
                    min_value=1, key=f"n_camiones_{key_ns}",
                    help="Camiones de 13 pallets con los que cuenta SB.",
                )
            with col_vta:
                vueltas_por_camion = st.number_input(
                    "Vueltas por camión al día", value=cfg.get("vueltas_por_camion_default", 2),
                    min_value=1, key=f"vueltas_{key_ns}",
                )
            ventanas_por_dia = int(n_camiones) * int(vueltas_por_camion)
            st.caption(f"= {ventanas_por_dia} ventanas/día ({n_camiones} camiones × {vueltas_por_camion} vueltas).")

            # Ajustes puntuales: dias donde se pierde un camion completo y/o
            # una vuelta suelta (ej: se presta un camion a otra area, se va a
            # mantencion, etc). Se descuentan ventanas SOLO ese dia; el resto
            # de la semana sigue con el cupo normal (n_camiones x vueltas).
            # Por defecto (se repite igual todas las semanas, editable aca):
            # Miercoles pierde 1 camion de 13 pallets completo, y Martes
            # pierde 1 vuelta suelta.
            _bajas_default = {"Miercoles": {"camiones": 1, "vueltas": 0},
                               "Martes": {"camiones": 0, "vueltas": 1}}
            with st.expander("🚧 Camiones/vueltas no disponibles algún día"):
                dias_con_baja = st.multiselect(
                    "Días con menor disponibilidad esta semana",
                    dias, default=[d for d in dias if d in _bajas_default],
                    key=f"dias_baja_{key_ns}",
                )
                bajas_por_dia = {}
                for d in dias_con_baja:
                    c1, c2 = st.columns(2)
                    camiones_perdidos = c1.number_input(
                        f"Camiones perdidos completos — {d}", min_value=0,
                        max_value=int(n_camiones),
                        value=_bajas_default.get(d, {}).get("camiones", 0),
                        key=f"cam_perdidos_{key_ns}_{d}",
                    )
                    vueltas_perdidas = c2.number_input(
                        f"Vueltas sueltas perdidas — {d}", min_value=0,
                        max_value=int(vueltas_por_camion),
                        value=_bajas_default.get(d, {}).get("vueltas", 0),
                        key=f"vta_perdidas_{key_ns}_{d}",
                    )
                    bajas_por_dia[d] = int(camiones_perdidos) * int(vueltas_por_camion) + int(vueltas_perdidas)

                if any(bajas_por_dia.values()):
                    ventanas_por_dia = {
                        d: max(0, ventanas_por_dia - bajas_por_dia.get(d, 0)) for d in dias
                    }
                    resumen_baja = "; ".join(
                        f"{d}: {ventanas_por_dia[d]} ventanas" for d in dias_con_baja
                    )
                    st.caption(f"⚠️ Ajustado: {resumen_baja} (resto de días sin cambio).")
        else:
            ventanas_por_dia = st.number_input(
                "Ventanas de despacho por día", value=cfg.get("ventanas_por_dia_default", 4),
                min_value=1, key=f"ventanas_{key_ns}",
            )
        if cfg.get("capacidad_rescate"):
            rescate_div_txt = (
                ", ".join(sorted(cfg["rescate_divisiones"])) if cfg.get("rescate_divisiones") else "todas las divisiones"
            )
            _total_ventanas_txt = (
                str(sum(ventanas_por_dia.values())) + " ventanas (con ajustes por día)"
                if isinstance(ventanas_por_dia, dict)
                else f"{len(dias)} días × {ventanas_por_dia} ventanas"
            )
            st.caption(
                f"🚛 Camiones normales de {cfg['capacidades_default'][0]} pallets. Si el total "
                f"de camiones no alcanza en las ventanas disponibles de la semana "
                f"({_total_ventanas_txt}) -es decir, no alcanza la "
                f"cubicación para despachar todo-, como ÚLTIMO RECURSO se fusionan los "
                f"camiones menos prioritarios (misma división) en ramplas de "
                f"{cfg['capacidad_rescate']} pallets. Aplica solo a: {rescate_div_txt}."
            )
        if cfg.get("produccion"):
            dias_prod_txt = ", ".join(sorted(cfg.get("dias_preferidos_produccion") or [])) or "Miercoles, Jueves"
            st.caption(
                f"🏭 Si la columna '{cfg['produccion']}' del Refresh trae algo para una OC "
                "(son productos que hay que esperar que produzca producción), esa OC se "
                "fuerza a despacharse el Miércoles o Jueves, aunque su prioridad sea "
                "'1 - Solicitado = 1er Posible (completo)'. Se marca con el chip 🏭 "
                f"PRODUCCIÓN. Días preferidos: {dias_prod_txt}."
            )
        if cfg.get("capacidades_por_division"):
            flota_txt = "; ".join(
                f"{div}: camiones de {'/'.join(str(c) for c in caps)} pallets"
                for div, caps in cfg["capacidades_por_division"].items()
            )
            st.caption(f"🚚 Flota propia por división — {flota_txt}.")
        if cfg["hoja"] == "PU":
            st.caption("📌 PU no se distribuye durante la semana: por defecto solo se "
                       "despacha el Viernes (ajustable arriba). Incluye rampla de 27 pallets. "
                       "No hay tope de ventanas: se reparte entre los transportes hasta "
                       "completar el despacho.")
        if cfg.get("minimo_por_division"):
            min_txt = ", ".join(f"{v} de {k}" for k, v in cfg["minimo_por_division"].items())
            st.caption(f"📌 Mínimo garantizado por día (si hay OC esperando): {min_txt}.")
        if cfg["usa_pronto_vence"]:
            st.caption(
                "Los productos con nombre en 'Directos' se sacan ANTES de armar los "
                "camiones (no ocupan pallets/ventanas) y quedan solo como información. "
                "Farma y Consumo Masivo nunca comparten camión. "
                "Orden de carga del resto: 1) Solicitado=1er Posible, "
                "2) cubierto con Pronto-vence, 5) parcial sin cobertura, 3) sin 1er Posible."
            )
        else:
            st.caption(
                "Los productos con nombre en 'Directos' se sacan ANTES de armar los "
                "camiones (no ocupan pallets/ventanas) y quedan solo como información. "
                f"En {cfg['hoja']} el stock por vencer NO se usa para completar faltantes. "
                "Orden de carga: 1) Solicitado=1er Posible (completo), "
                "2) no alcanza a completar el solicitado (parcial), 3) sin 1er Posible."
            )
        orden_prioridad = cfg["orden_prioridad"]

    facturados = cargar_facturados_desde_refresh(archivo, version=_version_archivo(archivo))
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
        ventanas_por_dia if isinstance(ventanas_por_dia, dict) else int(ventanas_por_dia),
        orden_prioridad, facturados, cfg,
    )

    if resumen is None:
        st.warning(f"No hay OC para camión en la semana {semana} "
                    f"(revisa si todas quedaron como Directos).")
        if not tabla_directos.empty:
            st.subheader("Directos (información, no van en camión)")
            st.dataframe(tabla_directos, use_container_width=True, hide_index=True)
        return

    if cfg.get("usa_transportes"):
        n_transportes_usados = min(info["camiones"], cfg.get("n_transportes", 3))
        render_kpi_cards([
            {"value": info["camiones"], "label": "Camiones / viajes necesarios"},
            {"value": cfg.get("n_transportes", 3), "label": "Transportes disponibles"},
            {
                "value": f"{n_transportes_usados}", "label": "Transportes en uso",
                "badge_text": "Sin tope de ventanas", "badge_color": "#1DB980",
            },
            {"value": info["oc_directos"], "label": "OC 100% Directos"},
            {"value": info["oc_facturadas"], "label": "OC ya Facturadas"},
        ])
    else:
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
            {
                "value": info.get("oc_produccion", 0), "label": "OC esperando Producción",
                "badge_text": "Miér./Jue." if info.get("oc_produccion", 0) else None,
                "badge_color": "#7C3AED",
            },
        ])
    if info["overflow"]:
        st.error("⚠️ No alcanzan las ventanas de la semana para todos los camiones "
                  "necesarios. Suma días/ventanas o revisa las capacidades.")

    excel_pendientes = exportar_pendientes_excel(detalle)
    excel_checklist = exportar_checklist_carga(detalle, semana, cfg)
    col_desc_a, col_desc_b, col_desc_c = st.columns(3)
    with col_desc_a:
        if excel_checklist:
            st.download_button(
                "⬇️ Descargar Plan de Despacho (Excel)",
                data=excel_checklist,
                file_name=f"Plan_de_Despacho_{cfg['hoja']}_S{semana}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                key=f"btn_plan_{key_ns}",
                help="Todas las OC de la semana (facturadas o no), separadas en pestañas "
                     "'Farma' y 'Consumo', con numeración de carga, día de "
                     "despacho, N° de camión con línea de separación entre camiones y "
                     "casillero de verificación, igual formato al que se usa a mano en "
                     "Operaciones.",
            )
        else:
            st.info("No hay OC para armar el Plan de Despacho de esta semana.")
    with col_desc_b:
        refresh_bytes = _bytes_archivo_original(archivo)
        if refresh_bytes:
            st.download_button(
                "⬇️ Descargar Refresh de origen (Excel)",
                data=refresh_bytes,
                file_name="Refresh.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                key=f"btn_refresh_{key_ns}",
            )
    with col_desc_c:
        if excel_pendientes:
            n_pend = int((detalle["Facturado"] != "Sí").sum())
            st.download_button(
                f"⬇️ Descargar no 100% facturadas (Excel) · {n_pend} OC",
                data=excel_pendientes,
                file_name=f"Pendientes_{cfg['hoja']}_S{semana}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                key=f"btn_pendientes_{key_ns}",
            )
        else:
            st.success("✅ Todas las OC de esta semana ya aparecen como Facturadas.")

    with st.expander("📈 KPIs y análisis adicionales", expanded=True):
        dias_orden = [d for d in cfg["dias_opciones"] if d in dias]
        render_kpis_avanzados(resumen, detalle, tabla_directos, cfg,
                               vueltas_por_camion, dias_orden, key_ns)

    # Capacidad máxima real disponible = el camión/rampla más grande que se
    # pueda usar (incluye la rampla de rescate). Una OC individual con más
    # pallets que esto NO puede despacharse tal cual, sin importar cómo se
    # arme el plan — se marca en rojo para que se revise/divida.
    cap_max = max(list(capacidades) + ([cfg["capacidad_rescate"]] if cfg.get("capacidad_rescate") else [])) \
        if capacidades else cfg.get("capacidad_rescate")

    # Algunas divisiones tienen su PROPIA flota (ej: Farma, con camiones de
    # 16 pallets propios, sin acceso a la rampla de rescate de Consumo
    # Masivo) -> su tope real de "OC individual demasiado grande" es
    # distinto al cap_max general. Se arma un tope por división para que el
    # aviso rojo de "excede capacidad" sea correcto para cada una.
    cap_por_division = {}
    for div, caps_div in (cfg.get("capacidades_por_division") or {}).items():
        rescate_div = cfg.get("capacidad_rescate") if div in (cfg.get("rescate_divisiones") or set()) else None
        cap_por_division[div] = max(list(caps_div) + ([rescate_div] if rescate_div else []))

    st.subheader("Detalle por OC (van en camión)")
    busqueda_oc = st.text_input(
        "🔍 Buscar por N° de Pedido o de OC",
        placeholder="Ej: 101668 u 8757515",
        key=f"buscar_oc_{key_ns}",
    )
    detalle_vista = detalle
    if busqueda_oc.strip():
        q = busqueda_oc.strip()
        match = (
            detalle["Pedido (OC)"].astype(str).str.contains(q, case=False, na=False)
            | detalle["OC"].astype(str).str.contains(q, case=False, na=False)
        )
        detalle_vista = detalle[match]
        if detalle_vista.empty:
            st.warning(f"No se encontró ningún Pedido ni OC que coincida con '{q}' en esta semana.")
        else:
            for _, r in detalle_vista.iterrows():
                fecha_r = r["Fecha"].strftime("%d-%m-%Y") if hasattr(r["Fecha"], "strftime") else ""
                st.info(
                    f"📦 Pedido **{r['Pedido (OC)']}** (OC {r['OC']}) → se despacha el "
                    f"**{r['Día']} {fecha_r}**, Camión #{r['Camión #']}, Ventana {r['Ventana']} "
                    f"· {r['Pallets']:.0f} pal."
                )

    tab_calendario, tab_tabla = st.tabs(["🗓️ Calendario", "📋 Tabla"])
    with tab_calendario:
        st.caption("Una columna por día, tarjetas con Pedido, OC, Monto y Pallets — "
                   "ordenadas por ventana y prioridad.")
        render_calendario(detalle_vista, resumen, cap_max, cap_por_division)
    with tab_tabla:
        st.caption("Verde = 100% Facturado, naranjo = despacho Parcial (según la pestaña OC del Refresh). "
                    "Rojo oscuro = la OC por sí sola supera la capacidad máxima de transporte disponible "
                    f"({cap_max:.0f} pal) y debe revisarse." if cap_max else "")
        st.dataframe(
            detalle_vista.style.apply(_resaltar_facturado_factory(cap_max, cap_por_division), axis=1).format({
                "Pallets": "{:.0f}", "Monto": lambda v: formato_clp(v),
            }),
            use_container_width=True, hide_index=True,
        )

    st.subheader("Plan de camiones")
    st.caption("Verde = 100% Facturado, naranjo = despacho Parcial. "
               "La columna % Facturado indica qué proporción de ese camión ya se despachó al 100%.")
    render_tabla_camiones(resumen, detalle, cap_max, cap_por_division)

    if not tabla_directos.empty:
        st.subheader("Directos (información, NO ocupan camión/ventana)")
        st.caption(f"{info['lineas_directos']} líneas / {info['oc_directos']} OC con "
                   "proveedor directo asignado.")
        st.dataframe(tabla_directos, use_container_width=True, hide_index=True)


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

    tab_sb, tab_pu, tab_bbd = st.tabs(
        ["SB", "PU", "🧬 Stock y Caducidad (BBD STOCK)"]
    )

    with tab_sb:
        render_plan_hoja(archivo, HOJAS_CONFIG["SB"])

    with tab_pu:
        render_plan_hoja(archivo, HOJAS_CONFIG["PU"])

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
