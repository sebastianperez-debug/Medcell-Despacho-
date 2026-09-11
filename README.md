# Medcell Despacho

App de Streamlit para planificar el despacho semanal de OC en camiones de
13/16 pallets, a partir de la pestaña **SB** del Refresh (Excel de compras /
paletizado).

## Qué hace

1. Lees el Refresh y eliges la semana a planificar.
2. Separa primero las líneas con nombre en **Directos** (proveedor externo):
   esas no ocupan pallets ni ventanas de camión, solo quedan listadas como
   información.
3. Clasifica cada OC (columna "Pedido") en 4 categorías de prioridad:
   1. Solicitado = 1er Posible (completo)
   2. Diferencia cubierta por Pronto-vence
   5. Parcial sin cobertura
   3. Sin 1er Posible (sin stock disponible)
4. Arma camiones (13 o 16 pallets, se elige libremente el tamaño) sin partir
   nunca una OC entre dos camiones, y **sin mezclar nunca Farma con Consumo
   Masivo** en el mismo camión.
5. Asigna cada camión a un día × ventana de despacho de la semana (por
   defecto Lunes a Viernes, 4 ventanas por día — todo ajustable desde la
   app).
6. Marca en verde las OC que ya aparecen como **Facturadas** en
   `facturados.xlsx` (no se excluyen del camión, solo se destacan).
7. Descarga el resultado como Excel (`Plan Despacho`, `Detalle Pedidos`,
   `Directos`).

## Archivo de datos (Refresh)

Puedes dejar guardada una copia del Refresh directamente en el repo, en
`data/Refresh.xlsx`. Si no subes nada en la app, se usa automáticamente esa
copia (igual que en Medcell Almacenamiento). Para actualizarla cada semana:
reemplaza ese archivo en GitHub y haz commit — o simplemente sube uno más
nuevo desde el panel "📂 Fuente de datos" en la barra lateral sin tocar el
repo, si prefieres no dejarlo guardado ahí.

## Cómo correrla localmente

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Cómo actualizar los Facturados cada semana

Se detecta automáticamente desde la pestaña **OC** del mismo Refresh
(columna "Pedido de Venta" con "Pendiente" = 0 → ya despachado). No hay que
mantener ningún archivo aparte para esto.

## Despliegue (Streamlit Community Cloud)

1. Sube este repo a GitHub.
2. En [share.streamlit.io](https://share.streamlit.io), crea una nueva app
   apuntando a este repo, rama `main`, archivo principal `app.py`.
3. Streamlit detecta `requirements.txt` y `.streamlit/config.toml`
   automáticamente (tema de colores incluido).

## Estructura del repo

```
medcell-despacho/
├── app.py                   # App principal (toda la lógica + UI)
├── data/
│   └── Refresh.xlsx          # (opcional) copia del Refresh que mantienes en el repo
├── requirements.txt
├── .gitignore
├── README.md
└── .streamlit/
    └── config.toml           # Tema de colores (azul Medcell)
```

## Supuestos ajustables desde la UI

- Columna base de pallets: `Pallets Pos.` o `Pallets posibles` (sin
  redondear).
- Capacidades de camión disponibles (por defecto 13 y 16).
- Días hábiles de despacho (por defecto Lunes a Viernes).
- Ventanas de despacho por día (por defecto 4).
