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
.medcell-header {
    background: linear-gradient(135deg, #0B4F86 0%, #12294A 100%);
    padding: 1.4rem 1.8rem;
    border-radius: 10px;
    margin-bottom: 1.4rem;
}
.medcell-header h1 {
    color: #FFFFFF;
    font-size: 1.9rem;
    font-weight: 700;
    margin: 0;
    letter-spacing: -0.01em;
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
</style>
"""


def _header():
    st.markdown(_BRAND_CSS, unsafe_allow_html=True)
    st.markdown(
        """
        <div class="medcell-header">
            <h1>🚚 Medcell Despacho <span class="tag">SB</span></h1>
            <p>Plan de despacho semanal: distribuye las OC en camiones de 13/16 pallets,
            respetando Farma / Consumo Masivo, Directos y Facturados.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

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
# Fuente principal (automatica): pestana "OC" del MISMO Refresh que ya subes.
#   "Pedido de Venta" = mismo numero que "Pedido" en SB.
#   "Pendiente" = 0 (sumado por Pedido de Venta) -> ya se despacho/facturo.
# Fuente opcional (respaldo/override manual): archivo facturados.xlsx en el
#   repo, o uno subido a mano en la app, por si hay algo mas reciente que
#   todavia no refleja el Refresh.

ARCHIVO_FACTURADOS_DEFAULT = "facturados.xlsx"


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


def cargar_facturados(archivo_subido=None, ruta_local: str = ARCHIVO_FACTURADOS_DEFAULT) -> set:
    """Devuelve el set de numeros de Pedido marcados como Facturados en el
    archivo externo (opcional): prioriza el subido manualmente en la app;
    si no hay, busca el archivo local (el que mantienes en el repo)."""
    import os

    df_fact = None
    fuente = None
    try:
        if archivo_subido is not None:
            fuente = archivo_subido
        elif os.path.exists(ruta_local):
            fuente = ruta_local

        if fuente is not None:
            nombre = getattr(fuente, "name", str(fuente))
            if str(nombre).lower().endswith(".csv"):
                df_fact = pd.read_csv(fuente)
            else:
                df_fact = pd.read_excel(fuente)
    except Exception as e:
        st.warning(f"No pude leer el archivo de Facturados: {e}")
        return set()

    if df_fact is None or df_fact.empty:
        return set()

    df_fact.columns = [str(c).strip() for c in df_fact.columns]
    col_pedido = next(
        (c for c in df_fact.columns if c.strip().lower() in
         ("pedido", "oc", "n° pedido", "numero pedido", "número pedido")),
        df_fact.columns[0],
    )
    return set(df_fact[col_pedido].dropna().astype(str).str.strip())


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
        n_sku=("Pedido", "count"),
    ).reset_index()

    pr = agg.apply(_clasificar, axis=1, result_type="expand")
    agg["prioridad"], agg["prioridad_label"] = pr[0], pr[1]
    agg["pallets"] = agg["pallets"].round(2)
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
                "Pallets": it["pallets"], "# SKUs": it["n_sku"],
                "Camión #": it["camion_num"], "Día": it["dia"], "Ventana": it["ventana"],
                "Facturado": "Sí" if es_facturado else "No",
            })
    detalle = pd.DataFrame(detalle_rows)

    info = {"camiones": len(bins), "ventanas_disponibles": len(slots), "overflow": overflow,
            "oc_directos": tabla_directos["Pedido"].nunique() if not tabla_directos.empty else 0,
            "lineas_directos": len(tabla_directos),
            "oc_facturadas": int((detalle["Facturado"] == "Sí").sum())}
    return resumen, detalle, info, tabla_directos


def _resaltar_facturado(row):
    color = "background-color: #C6EFCE" if row.get("Facturado") == "Sí" else ""
    return [color] * len(row)


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
                ws.cell(row=r, column=col_pallets).number_format = "0.00"
        if "Pallets cargados" in resumen.columns:
            ws = writer.sheets["Plan Despacho"]
            col_pallets = resumen.columns.get_loc("Pallets cargados") + 1
            col_util = resumen.columns.get_loc("Utilización %") + 1
            for r in range(2, len(resumen) + 2):
                ws.cell(row=r, column=col_pallets).number_format = "0.00"
                ws.cell(row=r, column=col_util).number_format = "0.0"
    return buf.getvalue()


# --------------------------------------------------------------------------
# 5. PAGINA DE STREAMLIT
# --------------------------------------------------------------------------

def render():
    _header()

    archivo = st.file_uploader("Sube el Refresh (Excel)", type=["xlsx"])
    if not archivo:
        st.info("Sube el archivo Refresh para continuar.")
        return

    df = leer_sb(archivo)
    semanas_disp = sorted(df["Semana"].dropna().unique().tolist())

    col1, col2 = st.columns(2)
    with col1:
        semana = st.selectbox("Semana a planificar", semanas_disp,
                               index=len(semanas_disp) - 1 if semanas_disp else 0)
    with col2:
        anio = st.number_input("Año", value=datetime.date.today().year, step=1)

    with st.expander("⚙️ Configuración (ajustable)"):
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

    with st.expander("🟩 Facturados"):
        st.caption(
            "Se detecta automáticamente desde la pestaña **OC** del mismo Refresh que "
            "subiste (Pedido de Venta con Pendiente = 0 → ya despachado/facturado). "
            "No necesitas mantener ningún archivo aparte."
        )
        usar_archivo_extra = st.checkbox(
            "Agregar/forzar Facturados desde un archivo externo (opcional, por si hay "
            "algo más reciente que el Refresh todavía no refleja)"
        )
        archivo_facturados = None
        if usar_archivo_extra:
            archivo_facturados = st.file_uploader(
                f"Subir tabla de Facturados (Excel/CSV, columna 'Pedido') — o se usa "
                f"`{ARCHIVO_FACTURADOS_DEFAULT}` del repo si no subes nada",
                type=["xlsx", "csv"], key="facturados_upload",
            )

    facturados_refresh = cargar_facturados_desde_refresh(archivo)
    facturados_extra = cargar_facturados(archivo_facturados) if usar_archivo_extra else set()
    facturados = facturados_refresh | facturados_extra
    if facturados:
        st.caption(
            f"✅ {len(facturados_refresh)} pedidos detectados como Facturados desde el "
            f"Refresh" + (f" + {len(facturados_extra)} desde archivo externo" if facturados_extra else "") + "."
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

    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Camiones necesarios", info["camiones"])
    c2.metric("Ventanas disponibles", info["ventanas_disponibles"])
    c3.metric("Holgura", info["ventanas_disponibles"] - info["camiones"])
    c4.metric("OC 100% Directos", info["oc_directos"])
    c5.metric("OC ya Facturadas", info["oc_facturadas"])
    if info["overflow"]:
        st.error("⚠️ No alcanzan las ventanas de la semana para todos los camiones "
                  "necesarios. Suma días/ventanas o revisa las capacidades.")

    st.subheader("Plan de camiones")
    st.dataframe(
        resumen.style.format({
            "Pallets cargados": "{:.2f}",
            "Capacidad": "{:.0f}",
            "Utilización %": "{:.1f}",
        }),
        use_container_width=True, hide_index=True,
    )

    st.subheader("Detalle por OC (van en camión)")
    st.caption("Las filas en verde ya aparecen como Facturadas en la tabla externa.")
    st.dataframe(
        detalle.style.apply(_resaltar_facturado, axis=1).format({"Pallets": "{:.2f}"}),
        use_container_width=True, hide_index=True,
    )

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
