SISTEMA PRO V2.6.7 — CONEXION GOOGLE DRIVE

Esta versión conserva el tablero V2.6.7 y añade sincronización automática del Excel maestro desde Google Drive cuando se ejecuta en Render.

ARCHIVO GOOGLE DRIVE CONFIGURADO
BASE_RECORRIDOS_PRO_JCA.xlsx
Drive File ID: 14CeVb1lJGmh3XzkrQTBTuLORIOk028ze

COMPORTAMIENTO
- Render descarga una copia del XLSX público al iniciar.
- Vuelve a comprobar Google Drive cada 30 segundos.
- Si la descarga es válida, reemplaza la copia local de forma atómica.
- Si Google Drive falla temporalmente, conserva la última copia válida.
- El tablero mantiene su actualización normal.
- No es necesario cambiar el enlace ni los permisos si se carga una nueva versión del MISMO archivo en Google Drive.

IMPORTANTE
No borrar el archivo de Drive y crear otro. Para conservar el mismo ID/enlace, usar "Gestionar versiones" > "Subir nueva versión".

MAPAS
La sincronización del Excel queda preparada en esta versión. Las carpetas REC_DIURNO y REC_NOCTURNO se conectarán en el siguiente paso usando sus archivos KML públicos, sin alterar el diseño del tablero.
