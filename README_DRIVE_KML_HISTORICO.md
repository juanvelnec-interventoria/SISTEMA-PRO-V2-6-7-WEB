# SISTEMA PRO V2.6.7 — KML históricos por mes

El tablero conserva el historial de recorridos por mes. En Render, las carpetas públicas de Google Drive son la fuente de los KML:

- REC_DIURNO: 1DcsyWGZ_VI_4pNxNi6nJ7f7o8LymXpRp
- REC_NOCTURNO: 19iK5Fy3vgQwjmBwj74upPUwMjQyg0rDe

El sistema localiza automáticamente archivos con nombres `REC_DIUR_MESAAAA.kml` y `REC_NOCT_MESAAAA.kml` según el AÑO-MES seleccionado.

Al seleccionar un mes histórico, Render descarga ese KML bajo demanda y lo conserva en caché. El mes actual se vuelve a comprobar periódicamente. Por tanto, agregar un nuevo KML mensual en Drive no requiere cambiar el código ni hacer un nuevo despliegue.

Importante: las carpetas deben conservar acceso `Cualquier persona con el enlace → Lector`.
