# SISTEMA PRO V2.6.7 — versión WEB para Render

Esta carpeta es una copia de publicación. El V2.6.7 local original NO se modifica.

## Render
- Tipo: Web Service
- Runtime: Python
- Build: `pip install -r requirements.txt`
- Start: `python TABLERO_V2_6_4.py`
- Variable `RENDER=1` ya está definida en `render.yaml`.

## Datos
`BASE_RECORRIDOS_PRO_JCA.xlsx` es la copia inicial de datos. Para una publicación real,
la actualización automática del Excel desde el computador local debe resolverse aparte.

## KML
Las carpetas opcionales `Recorrido_Diurno` y `Recorrido_Nocturno` pueden colocarse junto al
programa para habilitar los KML mensuales en el entorno web.
