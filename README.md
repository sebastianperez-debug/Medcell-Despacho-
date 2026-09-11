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

## Cómo correrla localmente

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Cómo actualizar los Facturados cada semana

Reemplaza el archivo `facturados.xlsx` en la raíz del repo (columna
`Pedido` con los números ya facturados) y haz commit/push. La app lo lee
automáticamente. Si alguna semana no alcanzas a actualizarlo en GitHub,
puedes subir uno manualmente desde el panel "🟩 Facturados (opcional)"
dentro de la app — ese sube por sobre el del repo, solo para esa sesión.

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
├── facturados.xlsx           # Tabla que mantienes tú, semana a semana
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
