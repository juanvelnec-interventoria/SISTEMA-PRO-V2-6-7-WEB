# -*- coding: utf-8 -*-
import json, os, re, threading, time, webbrowser, zipfile, xml.etree.ElementTree as ET, html as htmlmod
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import urlparse
import pandas as pd

# ==========================================================
# SISTEMA PRO V2.3 — TABLERO EJECUTIVO
# Excel maestro del usuario
# ==========================================================
# Configuración dual: mantiene el funcionamiento local original y permite publicar
# la misma aplicación en Render sin cambiar el tablero.
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
RENDER_MODE = os.environ.get("RENDER", "0") == "1"
if RENDER_MODE:
    EXCEL_PATH = os.environ.get("EXCEL_PATH", os.path.join(BASE_DIR, "BASE_RECORRIDOS_PRO_JCA.xlsx"))
    DIURNO_ROOT = os.environ.get("DIURNO_ROOT", os.path.join(BASE_DIR, "Recorrido_Diurno"))
    NOCTURNO_ROOT = os.environ.get("NOCTURNO_ROOT", os.path.join(BASE_DIR, "Recorrido_Nocturno"))
else:
    EXCEL_PATH = os.environ.get("EXCEL_PATH", r"C:\Users\USER\Desktop\INTERVENTORIA VELNEC\Recorridos\BASE_RECORRIDOS_PRO_JCA.xlsx")
    DIURNO_ROOT = os.environ.get("DIURNO_ROOT", r"C:\Users\USER\Desktop\INTERVENTORIA VELNEC\Recorridos\Recorrido_Diurno")
    NOCTURNO_ROOT = os.environ.get("NOCTURNO_ROOT", r"C:\Users\USER\Desktop\INTERVENTORIA VELNEC\Recorridos\Recorrido_Nocturno")

PORT = int(os.environ.get("PORT", "8765"))
HOST = os.environ.get("HOST", "0.0.0.0" if RENDER_MODE else "127.0.0.1")
REFRESH_SECONDS = int(os.environ.get("REFRESH_SECONDS", "5"))

MONTH_ABBR = {
    1:"ENE", 2:"FEB", 3:"MAR", 4:"ABR", 5:"MAY", 6:"JUN",
    7:"JUL", 8:"AGO", 9:"SEP", 10:"OCT", 11:"NOV", 12:"DIC"
}

def month_folder(year_month):
    """Convierte 2026-09 -> SEP2026."""
    try:
        y,m = str(year_month).split("-")
        return f"{MONTH_ABBR[int(m)]}{y}"
    except Exception:
        return None

def kml_candidates(year_month, turno):
    folder = month_folder(year_month)
    if not folder:
        return []
    y=str(year_month).split("-")[0]
    if str(turno).upper()=="DIURNO":
        root = os.path.join(DIURNO_ROOT, y, folder)
        names=[f"REC_DIUR_{folder}.kml",f"REC_DIUR_{folder}.kmz"]
    elif str(turno).upper()=="NOCTURNO":
        root = os.path.join(NOCTURNO_ROOT, folder)
        names=[f"REC_NOCT_{folder}.kml",f"REC_NOCT_{folder}.kmz"]
    else:
        return []
    out=[os.path.join(root,n) for n in names]
    search_root=DIURNO_ROOT if str(turno).upper()=="DIURNO" else NOCTURNO_ROOT
    if os.path.isdir(search_root):
        targets={n.lower() for n in names}
        for dirpath,_,files in os.walk(search_root):
            for f in files:
                if f.lower() in targets:
                    p=os.path.join(dirpath,f)
                    if p not in out: out.append(p)
    return out

def _strip(tag):
    return tag.split("}",1)[-1] if "}" in tag else tag

def _coords_to_points(raw):
    pts=[]
    for token in re.split(r"\s+", (raw or "").strip()):
        if not token: continue
        p=token.split(",")
        if len(p)>=2:
            try:
                lon=float(p[0]); lat=float(p[1])
                if -180<=lon<=180 and -90<=lat<=90:
                    pts.append([lat,lon])
            except: pass
    return pts

def read_kml_geojson(path):
    """Lee KML/KMZ y devuelve GeoJSON. Incluye respaldo para KML grandes,
    LineString, MultiGeometry y gx:Track/gx:coord."""
    if not path or not os.path.exists(path):
        return {"type":"FeatureCollection","features":[],"file":path,"exists":False}
    try:
        if path.lower().endswith(".kmz"):
            with zipfile.ZipFile(path) as z:
                names=[n for n in z.namelist() if n.lower().endswith(".kml")]
                if not names:
                    return {"type":"FeatureCollection","features":[],"file":path,"exists":True,"error":"KMZ sin KML"}
                raw=z.read(names[0])
        else:
            with open(path,'rb') as f: raw=f.read()
        features=[]
        # Primer intento: XML completo.
        try:
            root=ET.fromstring(raw)
            for pm in root.iter():
                if _strip(pm.tag)!="Placemark": continue
                name=""; desc=""
                for child in pm:
                    tag=_strip(child.tag)
                    if tag=="name": name=(child.text or "").strip()
                    elif tag=="description": desc=htmlmod.unescape(child.text or "").strip()
                # LineString y cualquier coordinates descendiente.
                for geom in pm.iter():
                    tag=_strip(geom.tag)
                    if tag=="LineString":
                        ce=None
                        for c in geom.iter():
                            if _strip(c.tag)=="coordinates": ce=c.text; break
                        pts=_coords_to_points(ce)
                        if len(pts)>=2:
                            features.append({"type":"Feature","geometry":{"type":"LineString","coordinates":[[p[1],p[0]] for p in pts]},"properties":{"name":name,"description":desc}})
                    elif tag=="Point":
                        ce=None
                        for c in geom.iter():
                            if _strip(c.tag)=="coordinates": ce=c.text; break
                        pts=_coords_to_points(ce)
                        if pts:
                            p=pts[0]
                            features.append({"type":"Feature","geometry":{"type":"Point","coordinates":[p[1],p[0]]},"properties":{"name":name,"description":desc}})
                    elif tag=="Track":
                        pts=[]
                        for c in geom.iter():
                            if _strip(c.tag)=="coord":
                                q=(c.text or '').strip().split()
                                if len(q)>=2:
                                    try:
                                        lon=float(q[0]);lat=float(q[1]);
                                        if -180<=lon<=180 and -90<=lat<=90: pts.append([lat,lon])
                                    except: pass
                        if len(pts)>=2:
                            features.append({"type":"Feature","geometry":{"type":"LineString","coordinates":[[p[1],p[0]] for p in pts]},"properties":{"name":name,"description":desc}})
        except Exception:
            pass
        # Respaldo tolerante: extrae geometrías aunque el XML tenga algún detalle inválido.
        if not features:
            try:
                txt=raw.decode('utf-8','ignore')
                # LineString coordinates
                for m in re.finditer(r'<(?:[^:>]+:)?LineString[^>]*>(.*?)</(?:[^:>]+:)?LineString>',txt,re.I|re.S):
                    block=m.group(1)
                    cm=re.search(r'<(?:[^:>]+:)?coordinates[^>]*>(.*?)</(?:[^:>]+:)?coordinates>',block,re.I|re.S)
                    if cm:
                        pts=_coords_to_points(cm.group(1))
                        if len(pts)>=2:
                            features.append({"type":"Feature","geometry":{"type":"LineString","coordinates":[[p[1],p[0]] for p in pts]},"properties":{"name":"Recorrido KML","description":""}})
                # gx:Track / gx:coord: formato lon lat alt
                for m in re.finditer(r'<(?:[^:>]+:)?Track[^>]*>(.*?)</(?:[^:>]+:)?Track>',txt,re.I|re.S):
                    pts=[]
                    for cm in re.finditer(r'<(?:[^:>]+:)?coord[^>]*>\s*([^<]+?)\s*</(?:[^:>]+:)?coord>',m.group(1),re.I|re.S):
                        q=cm.group(1).strip().split()
                        if len(q)>=2:
                            try:
                                lon=float(q[0]);lat=float(q[1])
                                if -180<=lon<=180 and -90<=lat<=90: pts.append([lat,lon])
                            except: pass
                    if len(pts)>=2:
                        features.append({"type":"Feature","geometry":{"type":"LineString","coordinates":[[p[1],p[0]] for p in pts]},"properties":{"name":"Recorrido KML","description":""}})
                # Puntos sueltos
                if not features:
                    for cm in re.finditer(r'<(?:[^:>]+:)?coordinates[^>]*>(.*?)</(?:[^:>]+:)?coordinates>',txt,re.I|re.S):
                        pts=_coords_to_points(cm.group(1))
                        if len(pts)>=2:
                            features.append({"type":"Feature","geometry":{"type":"LineString","coordinates":[[p[1],p[0]] for p in pts]},"properties":{"name":"Recorrido KML","description":""}})
            except Exception:
                pass
        if not features:
            raise ValueError("El KML existe, pero no se encontraron geometrías reconocibles (LineString, coordinates o gx:Track).")
        return {"type":"FeatureCollection","features":features,"file":path,"exists":True,
                "size":os.path.getsize(path),"modified":time.strftime("%d/%m/%Y %H:%M:%S",time.localtime(os.path.getmtime(path)))}
    except Exception as e:
        return {"type":"FeatureCollection","features":[],"file":path,"exists":True,"error":str(e),"size":os.path.getsize(path),"modified":time.strftime("%d/%m/%Y %H:%M:%S",time.localtime(os.path.getmtime(path)))}

def monthly_kml(year_month):
    result={"month":year_month,"diurno":[],"nocturno":[]}
    for turno,key in [("DIURNO","diurno"),("NOCTURNO","nocturno")]:
        for p in kml_candidates(year_month,turno):
            if os.path.exists(p):
                result[key]=read_kml_geojson(p)
                break
        if not result[key]:
            result[key]={"type":"FeatureCollection","features":[],"file":kml_candidates(year_month,turno)[0] if kml_candidates(year_month,turno) else None,"exists":False}
    return result


HTML = r"""
<!doctype html>
<html lang="es">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>SISTEMA PRO V2.6 · Control de Recorridos</title>
<link rel="stylesheet" href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css">
<style>
:root{
 --nav:#071c2d;--nav2:#0c2940;--blue:#1688e8;--blue2:#0e6fbe;
 --gold:#f4bd35;--green:#1fb46b;--red:#df3f4f;--orange:#f08b25;
 --purple:#7135c8;--cyan:#13a9b9;--ink:#152534;--muted:#71808d;
 --bg:#eef2f6;--card:#fff;--line:#dfe5ea;
}
*{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--ink);font-family:Segoe UI,Arial,sans-serif}
.top{height:112px;background:linear-gradient(100deg,#061a2b,#0a2942 65%,#061b2c);color:#fff;display:flex;align-items:center;padding:0 26px;gap:20px;box-shadow:0 3px 15px #0005}
.logo{width:56px;height:76px;position:relative}.lamp{position:absolute;left:4px;top:5px;width:35px;height:12px;border-radius:18px 18px 4px 4px;background:#f5c23b;transform:rotate(-12deg);box-shadow:0 10px 25px #f5c23b88}.pole{position:absolute;left:9px;top:13px;width:5px;height:60px;background:#dce4eb;border-radius:5px}.brand{flex:1}.brand h1{font-size:37px;margin:0;font-weight:900;letter-spacing:-1px}.brand h1 span{color:var(--gold)}.brand b{font-size:19px;display:block;margin-top:-3px}.brand small{font-size:11px;letter-spacing:.5px;opacity:.78}
.clock{min-width:220px;border-left:1px solid #ffffff2b;padding-left:20px}.clock .date{font-size:11px;opacity:.75}.clock .time{font-size:26px;font-weight:800}.connected{color:#66e0a2;font-size:11px;font-weight:700}.city{border-left:1px solid #ffffff2b;padding-left:20px;text-align:center;min-width:160px}.city strong{font-size:20px}.city small{display:block;font-size:10px;opacity:.7}
.nav{background:#062238;display:flex;gap:6px;padding:7px 16px;overflow:auto}.nav button{border:1px solid #ffffff18;background:#0d314a;color:#d9e8f2;border-radius:8px;padding:10px 18px;font-weight:700;white-space:nowrap;cursor:pointer}.nav button.active{background:#148ce9;color:#fff}
.wrap{max-width:1600px;margin:auto;padding:14px}
.filters{background:#fff;border:1px solid var(--line);border-radius:12px;padding:11px;display:flex;gap:9px;align-items:end;flex-wrap:wrap;box-shadow:0 3px 14px #102b3c0d;margin-bottom:13px}
label{font-size:9px;text-transform:uppercase;font-weight:800;color:#667582;display:flex;flex-direction:column;gap:4px}select,input{border:1px solid #c9d3dc;border-radius:7px;padding:8px 9px;min-width:120px;color:#1c2d3b;background:#fff}.filters button{border:0;background:#102a3c;color:#fff;padding:9px 14px;border-radius:7px;font-weight:700}.fstatus{margin-left:auto;font-size:10px;color:var(--muted)}
.page{display:none}.page.active{display:block}.hero{display:flex;justify-content:space-between;align-items:end;margin:7px 2px 11px}.hero h2{margin:0;font-size:20px}.hero p{margin:3px 0 0;color:var(--muted);font-size:10px}
.kpis{display:grid;grid-template-columns:repeat(8,1fr);gap:9px;margin-bottom:13px}.kpi{min-height:98px;color:#fff;border-radius:9px;padding:12px 13px;position:relative;overflow:hidden;box-shadow:0 5px 15px #0d2c4018}.kpi:after{content:"";position:absolute;width:70px;height:70px;border-radius:50%;right:-24px;top:-24px;background:#fff1}.kpi .ico{font-size:23px}.kpi .v{font-size:25px;font-weight:900;margin-top:2px}.kpi .l{font-size:9px;text-transform:uppercase;font-weight:800;margin-top:3px}.kpi .s{font-size:8px;opacity:.8;margin-top:4px}.kblue{background:#0d4f83}.kgreen{background:#0a9b55}.kcyan{background:#0793a5}.kgold{background:#d99108}.kred{background:#d52e40}.kpurple{background:#6630ba}.kgray{background:#5d6267}
.grid{display:grid;grid-template-columns:repeat(12,minmax(0,1fr));gap:12px}.card{background:var(--card);border:1px solid var(--line);border-radius:10px;padding:12px;box-shadow:0 4px 15px #102b3c0b;min-width:0;min-height:320px;overflow:hidden}.card h3{font-size:13px;margin:0 0 8px;line-height:1.25}.span2{grid-column:span 8}.span3{grid-column:1/-1}.chart{height:340px;min-height:300px;width:100%}.daily-axis-card .chart{height:350px;min-height:340px}.map{height:450px;border-radius:8px;overflow:hidden}.legend{font-size:9px;color:var(--muted);margin-top:6px;line-height:1.4}.design-panel{display:none;position:fixed;inset:0;background:#071c2dbb;z-index:5000;align-items:center;justify-content:center;padding:20px}.design-panel.open{display:flex}.design-box{background:#fff;border-radius:14px;width:min(900px,96vw);max-height:88vh;overflow:auto;padding:18px;box-shadow:0 20px 60px #0008}.design-box h2{margin:0 0 5px;font-size:20px}.design-box p{margin:0 0 14px;color:var(--muted);font-size:11px}.design-grid{display:grid;grid-template-columns:1fr 110px 110px;gap:8px;align-items:center}.design-grid .head{font-size:9px;text-transform:uppercase;font-weight:800;color:#667582}.design-grid select,.design-grid input{min-width:0;width:100%;padding:7px}.design-actions{display:flex;gap:8px;justify-content:flex-end;margin-top:14px}.design-actions button{border:0;border-radius:7px;padding:9px 14px;font-weight:800;cursor:pointer}.helpbox{background:#f5f8fa;border:1px solid var(--line);border-radius:9px;padding:10px;margin:10px 0;font-size:10px;color:#556570}.definition{background:#f7fafc;border-left:4px solid var(--blue);padding:9px 11px;border-radius:6px;font-size:10px;color:#596771;margin-bottom:10px}.kpi-note{font-size:8px;opacity:.86;margin-top:5px}.mini-kpis{display:grid;grid-template-columns:repeat(3,1fr);gap:8px;margin-bottom:10px}.mini-kpi{background:#f6f8fa;border:1px solid var(--line);border-radius:8px;padding:10px}.mini-kpi b{font-size:20px;display:block}.mini-kpi span{font-size:9px;color:var(--muted)}.status-good{color:#138653;font-weight:800}.status-bad{color:#c42d3c;font-weight:800}.status-info{color:#1688e8;font-weight:800}.status-warn{color:#a97713;font-weight:800}.maptab{border:1px solid #cbd6de;background:#fff;color:#153246;border-radius:7px;padding:8px 12px;font-weight:800;cursor:pointer}.maptab.active{background:#1688e8;color:#fff;border-color:#1688e8}.route-palette{display:inline-flex;gap:3px;align-items:center;margin-left:2px}.route-palette button{width:17px;height:17px;border:1px solid #fff;outline:1px solid #cbd6de;border-radius:50%;padding:0;cursor:pointer}.route-palette button:hover{transform:scale(1.15)}
.alertgrid{display:grid;grid-template-columns:repeat(4,1fr);gap:9px;margin-bottom:12px}.abox{color:#fff;border-radius:9px;padding:13px}.abox .n{font-size:25px;font-weight:900}.abox .t{font-size:9px;text-transform:uppercase;font-weight:800}.ar{background:#c52d3d}.ao{background:#d87917}.ay{background:#9d7b17}.ag{background:#20835a}
.scroll{overflow:auto;max-height:500px;border:1px solid var(--line);border-radius:8px}table{border-collapse:collapse;width:100%;font-size:10px}th,td{padding:7px 8px;border-bottom:1px solid #edf0f2;text-align:left;white-space:nowrap}th{background:#f4f7f9;color:#5f6d77;text-transform:uppercase;font-size:9px;position:sticky;top:0;z-index:2}.good{color:#138653;font-weight:800}.bad{color:#c42d3c;font-weight:800}.warn{color:#a97713;font-weight:800}.info{color:#1688e8;font-weight:800}.pill{padding:3px 7px;border-radius:12px;background:#edf2f5;font-weight:800}
.detailgrid{display:grid;grid-template-columns:1.1fr .9fr;gap:12px}.routebox{background:#f6f8fa;border:1px solid var(--line);border-radius:9px;padding:11px}.routebox h4{margin:0 0 7px}.routebox p{font-size:10px;margin:4px 0;color:#596771}
.footer{padding:16px 2px;color:#77848d;font-size:9px;text-align:center}
@media(max-width:1250px){.kpis{grid-template-columns:repeat(4,1fr)}.grid{grid-template-columns:repeat(12,minmax(0,1fr))}}
@media(max-width:750px){.top{height:auto;padding:13px}.clock,.city{display:none}.brand h1{font-size:28px}.grid{grid-template-columns:1fr}.card{grid-column:1/-1!important}.kpis{grid-template-columns:1fr 1fr}.alertgrid{grid-template-columns:1fr 1fr}.detailgrid{grid-template-columns:1fr}.design-grid{grid-template-columns:1fr 90px 90px}}
</style>
</head>
<body>
<header class="top">
<div class="logo"><div class="lamp"></div><div class="pole"></div></div>
<div class="brand"><h1>SISTEMA <span>PRO</span></h1><b>CONTROL DE RECORRIDOS</b><small>INTERVENTORÍA ALUMBRADO PÚBLICO · MEDELLÍN</small></div>
<div class="clock"><div class="date" id="date"></div><div class="time" id="time"></div><div class="connected" id="live">● Conectando Excel…</div></div>
<div class="city"><strong>MEDELLÍN</strong><small>Ciudad que ilumina · Interventoría</small></div>
</header>
<nav class="nav">
<button class="active" onclick="page('resumen',this)">⌂ Resumen</button>
<button onclick="page('tecnicos',this)">● Técnicos</button>
<button onclick="page('recorridos',this)">⌖ Recorridos</button>
<button onclick="page('mapa',this)">▣ Mapa</button>
<button onclick="page('alertas',this)">⚠ Alertas</button>
<button onclick="page('temporal',this)">◷ Análisis temporal</button>
<button onclick="page('gps',this)">⌖ Calidad GPS</button>
</nav>

<div id="designPanel" class="design-panel"><div class="design-box"><h2>⚙ Configuración visual del tablero</h2><p>Modifica el ancho y la altura de cada panel sin tocar los datos. Los cambios son solo de presentación y quedan guardados en este navegador.</p><div class="helpbox"><b>Ancho:</b> 25%, 33%, 50%, 67%, 75% o 100% de la fila. <b>Altura:</b> puedes escribir el valor en píxeles. Si un gráfico necesita más espacio, aumenta su ancho o altura.</div><div class="design-grid" id="designGrid"><div class="head">Panel</div><div class="head">Ancho</div><div class="head">Altura (px)</div></div><div class="design-actions"><button onclick="resetDesign()" style="background:#eef2f6;color:#223746">Restablecer diseño</button><button onclick="closeDesign()" style="background:#102a3c;color:#fff">Cerrar</button><button onclick="applyDesign()" style="background:#1688e8;color:#fff">Guardar cambios</button></div></div></div><main class="wrap">
<div class="filters">
<label>Año-Mes<select id="month"><option>Todos</option></select></label>
<label>Turno<select id="shift"><option>Todos</option><option>DIURNO</option><option>NOCTURNO</option></select></label>
<label>Técnico<select id="tech"><option>Todos</option></select></label>
<label>Fecha inicio<input id="d1" type="date"></label>
<label>Fecha fin<input id="d2" type="date"></label>
<label>Estado inicio<select id="si"><option>Todos</option><option>INICIO CUMPLE</option><option>INICIO TARDÍO</option><option>INICIO ANTES DEL HORARIO</option></select></label>
<label>Estado fin<select id="sf"><option>Todos</option><option>FIN CUMPLE</option><option>TERMINÓ ANTES</option><option>TERMINÓ DESPUÉS</option></select></label>
<button onclick="resetF()">LIMPIAR FILTROS</button><button onclick="openDesign()" id="designBtn" style="background:#7135c8">⚙ DISEÑO DEL TABLERO</button><button onclick="forceRefresh()" style="background:#1688e8">↻ ACTUALIZAR</button><span class="resize-note">Puedes definir el ancho de cada panel en % y su altura. El diseño queda guardado en este computador.</span><span class="fstatus" id="fstatus">—</span>
</div>

<section id="resumen" class="page active">
<div class="hero"><div><h2>Resumen ejecutivo</h2><p>Indicadores generales de control de recorridos.</p></div></div>
<div class="kpis">
<div class="kpi kblue"><div class="ico">⌖</div><div class="v" id="k1">-</div><div class="l">Total de recorridos</div><div class="s">Registros analizados</div></div>
<div class="kpi kgreen"><div class="ico">▤</div><div class="v" id="k2">-</div><div class="l">Kilómetros recorridos</div><div class="s">Distancia acumulada</div></div>
<div class="kpi kcyan"><div class="ico">◷</div><div class="v" id="k3">-</div><div class="l">Cumplimiento inicio</div><div class="s">Dentro del horario</div></div>
<div class="kpi kgold"><div class="ico">⚑</div><div class="v" id="k4">-</div><div class="l">Cumplimiento fin</div><div class="s">Fin cumple / posterior</div></div>
<div class="kpi kred"><div class="ico">!</div><div class="v" id="k5">-</div><div class="l">Inicios tardíos</div><div class="s">Requieren seguimiento</div></div>
<div class="kpi kred"><div class="ico">◷</div><div class="v" id="k6">-</div><div class="l">Terminó antes</div><div class="s">Requieren seguimiento</div></div>
<div class="kpi kpurple"><div class="ico">⌁</div><div class="v" id="k7">-</div><div class="l">Huecos GPS ≥10 min</div><div class="s">Calidad de trazabilidad</div></div>
<div class="kpi kgray"><div class="ico">⌛</div><div class="v" id="k8">-</div><div class="l">Desviación total</div><div class="s">Inicio tardío + fin anticipado</div></div>
</div>
<div class="definition"><b>Semáforo de lectura:</b> <span class="good">● Verde = cumple</span> · <span class="info">● Azul = antes/después del horario sin penalización</span> · <span class="bad">● Rojo = requiere seguimiento</span> · <span class="warn">● Amarillo = revisar información GPS o registro</span></div>
<div class="grid">
<div class="card" data-default-span="8"><h3>Cumplimiento por técnico</h3><div class="legend"><span style="color:#1688e8">■ Azul = % cumplimiento de inicio</span> &nbsp; <span style="color:#f08b25">■ Naranja = % cumplimiento de finalización</span><br><small>Inicio cumple = inició dentro o antes del horario. Fin cumple = terminó dentro o después del horario.</small></div><div id="techChart" class="chart"></div></div>
<div class="card daily-axis-card" data-default-span="4"><h3>Recorridos registrados por día</h3><div class="legend">Cantidad de recorridos registrados en el Excel para cada fecha.</div><div id="dailyChart" class="chart"></div></div>
<div class="card" data-default-span="6"><h3>Inicios tardíos por técnico</h3><div class="legend">Número de recorridos de cada técnico que comenzaron después de la hora programada.</div><div id="lateChart" class="chart"></div></div>
<div class="card" data-default-span="6"><h3>Terminaciones anticipadas por técnico</h3><div class="legend">Número de recorridos de cada técnico que finalizaron antes de la hora programada.</div><div id="earlyChart" class="chart"></div></div>
<div class="card span3" data-default-span="12"><div style="display:flex;gap:10px;align-items:end;justify-content:space-between;flex-wrap:wrap;margin-bottom:8px"><div><h3 style="margin-bottom:2px">Mapa de recorridos · selección de recorrido</h3><div class="legend" style="margin-top:0">Selecciona un turno o un recorrido específico para visualizar sus puntos de inicio y finalización.</div></div><label style="min-width:360px;flex:1;max-width:700px;text-transform:none">Visualizar recorrido<select id="summaryRouteSelect" onchange="changeSummaryRoute()"><option value="__all__">Todos los recorridos</option></select></label></div><div id="summaryRouteInfo" class="legend" style="margin:0 0 7px;font-weight:700;color:#334450"></div><div id="map1" class="map"></div><div class="legend">● Inicio &nbsp; ● Finalización &nbsp; La selección respeta los filtros superiores. Las coordenadas se toman directamente del Excel maestro.</div></div>
<div class="card" data-default-span="4"><h3>Distribución por turno</h3><div id="shiftChart" class="chart"></div></div>
<div class="card" data-default-span="4"><h3>Estado de inicio</h3><div id="startChart" class="chart"></div></div>
<div class="card" data-default-span="4"><h3>Estado de finalización</h3><div id="endChart" class="chart"></div></div>
<div class="card" data-default-span="8"><h3>Ranking de técnicos</h3><div class="scroll"><table id="rankTable"></table></div></div>
<div class="card" data-default-span="4"><h3>Alertas recientes</h3><div class="scroll" style="max-height:300px"><table id="recentTable"></table></div></div>
</div>
</section>

<section id="tecnicos" class="page">
<div class="hero"><div><h2>Control por técnico</h2><p>Ranking completo con indicadores y desviaciones acumuladas.</p></div></div>
<div class="definition"><b>¿Qué muestra esta página?</b> El desempeño de cada técnico según los recorridos registrados en el Excel. <b>% Inicio</b> = recorridos que iniciaron dentro o antes del horario programado. <b>% Fin</b> = recorridos que terminaron dentro o después del horario programado. <b>Desviación</b> = únicamente inicio tardío + terminación anticipada.</div><div class="card span3"><div class="scroll"><table id="techTable"></table></div></div>
</section>

<section id="recorridos" class="page">
<div class="hero"><div><h2>Recorridos</h2><p>Detalle de los registros filtrados.</p></div></div>
<div class="definition"><b>Detalle de recorridos:</b> aquí se puede revisar registro por registro lo que el sistema tomó del Excel. Los filtros de la parte superior se aplican a esta tabla.</div><div class="card span3"><div class="scroll" style="max-height:650px"><table id="routeTable"></table></div></div>
</section>

<section id="mapa" class="page">
<div class="hero"><div><h2>Mapa de recorridos · control espacial</h2><p>Selecciona un mapa. Cada vista muestra una sola fuente para facilitar la lectura.</p><div style="display:flex;gap:7px;margin-top:9px;flex-wrap:wrap"><button class="maptab active" onclick="showMapTab('diurno',this)">☀️ Recorrido diurno</button><button class="maptab" onclick="showMapTab('nocturno',this)">🌙 Recorrido nocturno</button><button class="maptab" onclick="showMapTab('coord',this)">📍 Inicio y finalización</button></div></div></div>
<div class="grid">
<div class="card span3 mapview" id="mapViewDiurno">
<h3>☀️ Recorrido DIURNO · KML mensual</h3>
<div style="display:flex;gap:16px;align-items:center;font-size:10px;font-weight:700;margin-bottom:7px"><span style="color:#15945a">● Trayectoria diurna</span><span id="kmlDStatus" style="margin-left:auto;color:#71808d">KML diurno: esperando…</span><label style="text-transform:none;display:inline-flex;flex-direction:row;align-items:center;gap:5px;margin-left:8px;font-size:9px">Color de línea <input id="colorDiurno" type="color" value="#15945a" title="Elegir color de la línea del recorrido diurno" onchange="setRouteColor('diurno',this.value)" style="min-width:38px;width:38px;height:27px;padding:2px;cursor:pointer"></label><span class="route-palette"><button title="Verde" onclick="setRouteColor('diurno','#15945a')" style="background:#15945a"></button><button title="Rojo" onclick="setRouteColor('diurno','#df3f4f')" style="background:#df3f4f"></button><button title="Azul" onclick="setRouteColor('diurno','#1688e8')" style="background:#1688e8"></button><button title="Morado" onclick="setRouteColor('diurno','#7135c8')" style="background:#7135c8"></button><button title="Naranja" onclick="setRouteColor('diurno','#f08b25')" style="background:#f08b25"></button><button title="Negro" onclick="setRouteColor('diurno','#222222')" style="background:#222222"></button></span></div>
<div id="mapDiurno" class="map" style="height:520px;min-height:520px"></div>
<div class="legend">Muestra la trayectoria real almacenada en <b>REC_DIUR_MESANO.kml</b>. Usa el selector <b>Color</b> para elegir manualmente cómo quieres visualizarla.</div>
</div>
<div class="card span3 mapview" id="mapViewNocturno" style="display:none">
<h3>🌙 Recorrido NOCTURNO · KML mensual</h3>
<div style="display:flex;gap:16px;align-items:center;font-size:10px;font-weight:700;margin-bottom:7px"><span style="color:#1688e8">● Trayectoria nocturna</span><span id="kmlNStatus" style="margin-left:auto;color:#71808d">KML nocturno: esperando…</span><label style="text-transform:none;display:inline-flex;flex-direction:row;align-items:center;gap:5px;margin-left:8px;font-size:9px">Color de línea <input id="colorNocturno" type="color" value="#1688e8" title="Elegir color de la línea del recorrido nocturno" onchange="setRouteColor('nocturno',this.value)" style="min-width:38px;width:38px;height:27px;padding:2px;cursor:pointer"></label><span class="route-palette"><button title="Verde" onclick="setRouteColor('nocturno','#15945a')" style="background:#15945a"></button><button title="Rojo" onclick="setRouteColor('nocturno','#df3f4f')" style="background:#df3f4f"></button><button title="Azul" onclick="setRouteColor('nocturno','#1688e8')" style="background:#1688e8"></button><button title="Morado" onclick="setRouteColor('nocturno','#7135c8')" style="background:#7135c8"></button><button title="Naranja" onclick="setRouteColor('nocturno','#f08b25')" style="background:#f08b25"></button><button title="Negro" onclick="setRouteColor('nocturno','#222222')" style="background:#222222"></button></span></div>
<div id="mapNocturno" class="map" style="height:520px;min-height:520px"></div>
<div class="legend">Muestra la trayectoria real almacenada en <b>REC_NOCT_MESANO.kml</b>. Usa el selector <b>Color</b> para elegir manualmente cómo quieres visualizarla.</div>
</div>
<div class="card span3 mapview" id="mapViewCoord" style="display:none">
<h3>📍 Inicio y finalización de recorridos</h3>
<div style="display:flex;gap:16px;align-items:center;font-size:10px;font-weight:700;margin-bottom:7px"><span style="color:#1688e8">● Inicio</span><span style="color:#df3f4f">● Finalización</span><span id="coordStatus" style="margin-left:auto;color:#71808d">Coordenadas del Excel</span></div>
<div id="map2" class="map" style="height:520px;min-height:520px"></div>
<div class="legend">Usa exclusivamente las coordenadas de inicio y finalización registradas en <b>BASE_RECORRIDOS_PRO_JCA.xlsx</b> y respeta los filtros.</div>
</div>
<div class="card"><h3>📁 KML DIURNO</h3><div id="kmlDFile" class="routebox"><p>—</p></div></div>
<div class="card"><h3>📁 KML NOCTURNO</h3><div id="kmlNFile" class="routebox"><p>—</p></div></div>
<div class="card"><h3>📊 Resumen espacial</h3><div class="routebox"><p><b id="kmld1">-</b> elementos KML diurno</p><p><b id="kmld2">-</b> elementos KML nocturno</p><p><b id="kmld3">-</b> recorridos Excel filtrados</p></div></div>
</div>
</section>

<section id="alertas" class="page">
<div class="hero"><div><h2>Centro de alertas</h2><p>Registros que requieren seguimiento operativo o revisión de trazabilidad.</p></div></div>
<div class="alertgrid">
<div class="abox ar"><div class="n" id="a1">-</div><div class="t">Inicios tardíos</div></div>
<div class="abox ao"><div class="n" id="a2">-</div><div class="t">Terminó antes</div></div>
<div class="abox ay"><div class="n" id="a3">-</div><div class="t">Huecos GPS</div></div>
<div class="abox ar"><div class="n" id="a4">-</div><div class="t">Requieren revisión</div></div>
</div>
<div class="definition"><b>Lectura rápida:</b> una alerta no significa automáticamente incumplimiento grave. Indica un registro que requiere seguimiento: inicio tardío, terminación anticipada, interrupción GPS o revisión del registro. La columna <b>Acción sugerida</b> orienta qué verificar.</div><div class="card span3"><div class="scroll" style="max-height:620px"><table id="alertTable"></table></div></div>
</section>

<section id="temporal" class="page">
<div class="hero"><div><h2>Análisis temporal</h2><p>Dos indicadores fáciles de interpretar: cantidad de recorridos registrados por fecha y tiempo acumulado de desviaciones de horario.</p></div></div>
<div class="definition"><b>¿Cómo leer esta página?</b> <b>Recorridos registrados por día</b> = cantidad de recorridos que aparecen en el Excel en cada fecha. <b>Tiempo acumulado de desviaciones</b> = suma del tiempo de <b>inicio tardío + terminación anticipada</b>. Se muestra en formato <b>HH:MM:SS</b>; no son horas trabajadas.</div>
<div class="grid">
<div class="card daily-axis-card" style="grid-column:span 6"><h3>Recorridos registrados por día</h3><div class="legend">Cantidad de recorridos registrados en el Excel para cada fecha seleccionada. El valor sobre cada punto corresponde al número de recorridos de ese día.</div><div id="daily2" class="chart"></div></div>
<div class="card" style="grid-column:span 6"><h3>Tiempo acumulado de desviaciones de horario por técnico</h3><div class="legend">Suma exclusivamente <b>inicio tardío + terminación anticipada</b> por técnico. Los valores se presentan como <b>HH:MM:SS</b>, por ejemplo <b>02:35:20</b> = 2 horas, 35 minutos y 20 segundos de desviaciones acumuladas.</div><div id="devChart" class="chart"></div></div>
</div>
</section>

<section id="gps" class="page">
<div class="hero"><div><h2>Calidad GPS</h2><p>Un “hueco GPS” es un intervalo de 10 minutos o más sin puntos GPS registrados.</p></div></div>
<div class="definition"><b>¿Qué es un hueco GPS?</b> Es un intervalo de <b>10 minutos o más sin puntos GPS registrados</b>. No significa necesariamente que el técnico se detuvo; significa que la trazabilidad GPS tuvo una interrupción que debe verificarse.</div><div class="mini-kpis"><div class="mini-kpi"><b id="gm1">-</b><span>Interrupciones GPS ≥10 min</span></div><div class="mini-kpi"><b id="gm2">-</b><span>Recorridos con alguna interrupción</span></div><div class="mini-kpi"><b id="gm3">-</b><span>% de recorridos con interrupción</span></div></div><div class="grid">
<div class="card" style="grid-column:span 7"><h3>Interrupciones GPS registradas por técnico</h3><div class="legend">Cantidad total de interrupciones de registro GPS de 10 minutos o más detectadas en los recorridos de cada técnico. No equivale automáticamente a una detención del técnico.</div><div id="gpsChart" class="chart"></div></div>
<div class="card" style="grid-column:span 5"><h3>Recorridos con alguna interrupción GPS</h3><div class="legend">Cantidad de recorridos de cada técnico en los que se detectó al menos una interrupción GPS de 10 minutos o más.</div><div id="gpsRoutes" class="chart"></div></div>
</div>
</section>

<div class="footer">SISTEMA PRO V2.6 · Control de Recorridos · Interventoría Alumbrado Público · Medellín · Excel maestro: BASE_RECORRIDOS_PRO_JCA.xlsx · Actualización automática cada 5 segundos.<br><span style="font-size:8px;opacity:.75">Autor: Juan Ayala · Derechos reservados</span></div>
</main>

<script src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js"></script>
<script>
let DATA=[],lastStamp='',map1,map2,mapDiurno,mapNocturno,layers1=[],layers2=[],layersD=[],layersN=[],kmlCache={},kmlStamp='';
let kmlLayer1=null,kmlLayer2=null;
const $=id=>document.getElementById(id);
function esc(v){return String(v??'').replace(/[&<>"']/g,m=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[m]))}
function pct(a,b){return b?((100*a/b).toFixed(1)+'%'):'N/D'}
function secFmt(s){s=Math.max(0,Math.round(s||0));let h=Math.floor(s/3600);s%=3600;let m=Math.floor(s/60);s%=60;return String(h).padStart(2,'0')+':'+String(m).padStart(2,'0')+':'+String(s).padStart(2,'0')}
function fmt(n){return (+n||0).toLocaleString('es-CO',{maximumFractionDigits:1})}
function filt(){let m=$('month').value,s=$('shift').value,t=$('tech').value,a=$('d1').value,b=$('d2').value,si=$('si').value,sf=$('sf').value;return DATA.filter(r=>(m==='Todos'||String(r['Año-Mes'])===m)&&(s==='Todos'||String(r.Turno)===s)&&(t==='Todos'||String(r['Técnico'])===t)&&(!a||r.Fecha>=a)&&(!b||r.Fecha<=b)&&(si==='Todos'||r['Estado inicio']===si)&&(sf==='Todos'||r['Estado fin']===sf))}
function setupFilters(){let ms=[...new Set(DATA.map(r=>r['Año-Mes']).filter(Boolean))].sort(),ts=[...new Set(DATA.map(r=>r['Técnico']).filter(Boolean))].sort();$('month').innerHTML='<option>Todos</option>'+ms.map(x=>`<option>${esc(x)}</option>`).join('');$('tech').innerHTML='<option>Todos</option>'+ts.map(x=>`<option>${esc(x)}</option>`).join('');['month','shift','tech','d1','d2','si','sf'].forEach(x=>$(x).onchange=render)}
function svgBox(id){$(id).innerHTML=''}
function bars(id,labels,series,names,max){let el=$(id);if(!el)return;let w=Math.max(el.clientWidth||0,760),h=Math.max(el.clientHeight||0,340),l=55,r=18,t=22,b=72,ph=h-t-b,n=labels.length,g=(w-l-r)/Math.max(1,n),bw=Math.min(28,g*.62/Math.max(1,series.length));let M=max||Math.max(1,...series.flat())*1.15;let s=`<svg width="100%" height="${h}" viewBox="0 0 ${w} ${h}" xmlns="http://www.w3.org/2000/svg"><g font-family="Segoe UI" font-size="9" fill="#68747e">`;for(let q=0;q<=4;q++){let y=t+ph-q/4*ph;let val=M*q/4;s+=`<line x1="${l}" y1="${y}" x2="${w-r}" y2="${y}" stroke="#e8edf1"/><text x="${l-7}" y="${y+3}" text-anchor="end">${Number(val).toFixed(val%1?1:0)}${max===100?'%':''}</text>`}series.forEach((arr,j)=>arr.forEach((v,i)=>{let x=l+i*g+g*.18+j*bw,y=t+ph-(v/M)*ph,hh=t+ph-y;s+=`<rect x="${x}" y="${y}" width="${Math.max(2,bw-2)}" height="${Math.max(0,hh)}" rx="2" fill="${j?'#f08b25':'#1688e8'}"/><title>${esc(String(labels[i]))}: ${Number(v).toFixed(v%1?1:0)}${max===100?'%':''}</title><text x="${x+(bw-2)/2}" y="${Math.max(12,y-5)}" text-anchor="middle" fill="#334450" font-weight="700">${Number(v).toFixed(v%1?1:0)}${max===100?'%':''}</text>`}));labels.forEach((x,i)=>{let xx=l+i*g+g/2;s+=`<text x="${xx}" y="${h-b+16}" text-anchor="middle" transform="rotate(-28 ${xx} ${h-b+16})">${esc(String(x).slice(0,18))}</text>`});s+='</g></svg>';el.innerHTML=s}
function hbars(id,labels,vals,color='#1688e8'){let el=$(id);if(!el)return;let w=Math.max(el.clientWidth||0,700),h=Math.max(el.clientHeight||0,340),l=155,r=55,t=12,b=15,row=Math.max(24,(h-t-b)/Math.max(1,labels.length)),M=Math.max(1,...vals)*1.12;let s=`<svg width="100%" height="${h}" viewBox="0 0 ${w} ${h}" xmlns="http://www.w3.org/2000/svg"><g font-family="Segoe UI" font-size="9">`;labels.forEach((lab,i)=>{let y=t+i*row+4,bw=(w-l-r)*(vals[i]/M);s+=`<text x="${l-8}" y="${y+10}" text-anchor="end" fill="#68747e">${esc(String(lab).slice(0,24))}</text><rect x="${l}" y="${y}" width="${Math.max(2,bw)}" height="15" rx="3" fill="${color}"/><text x="${Math.min(w-r+2,l+bw+6)}" y="${y+11}" fill="#334450" font-weight="700">${Number(vals[i]).toFixed(vals[i]%1?1:0)}</text>`});s+='</g></svg>';el.innerHTML=s}
function timeHbars(id,labels,vals){let el=$(id);if(!el)return;let w=Math.max(el.clientWidth||0,760),h=Math.max(el.clientHeight||0,340),l=175,r=85,t=12,b=15,row=Math.max(24,(h-t-b)/Math.max(1,labels.length)),M=Math.max(1,...vals)*1.12;let s=`<svg width="100%" height="${h}" viewBox="0 0 ${w} ${h}" xmlns="http://www.w3.org/2000/svg"><g font-family="Segoe UI" font-size="9">`;labels.forEach((lab,i)=>{let y=t+i*row+4,bw=(w-l-r)*(vals[i]/M);s+=`<text x="${l-8}" y="${y+10}" text-anchor="end" fill="#68747e">${esc(String(lab).slice(0,24))}</text><rect x="${l}" y="${y}" width="${Math.max(2,bw)}" height="15" rx="3" fill="#df3f4f"/><text x="${Math.min(w-r+2,l+bw+6)}" y="${y+11}" fill="#334450" font-weight="700">${secFmt(vals[i])}</text>`});s+='</g></svg>';el.innerHTML=s}
function lineChart(id,labels,vals){let el=$(id);if(!el)return;let w=Math.max(el.clientWidth||0,760),h=Math.max(el.clientHeight||0,350),l=52,r=18,t=24,b=88,ph=h-t-b,M=Math.max(1,...vals),step=(w-l-r)/Math.max(1,vals.length-1);let fmtDay=v=>{let x=String(v??'').trim();let m=x.match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})/);if(m)return String(m[3]).padStart(2,'0')+'/'+String(m[2]).padStart(2,'0');m=x.match(/^(\d{1,2})[-\/](\d{1,2})[-\/](\d{4})/);if(m)return String(m[1]).padStart(2,'0')+'/'+String(m[2]).padStart(2,'0');let d=new Date(x);if(!isNaN(d))return String(d.getDate()).padStart(2,'0')+'/'+String(d.getMonth()+1).padStart(2,'0');return x};let pts=vals.map((v,i)=>`${l+i*step},${t+ph-v/M*ph}`).join(' ');let s=`<svg width="100%" height="${h}" viewBox="0 0 ${w} ${h}"><g font-family="Segoe UI" font-size="10" fill="#68747e"><line x1="${l}" y1="${t+ph}" x2="${w-r}" y2="${t+ph}" stroke="#ccd5dc"/><polyline points="${pts}" fill="none" stroke="#1688e8" stroke-width="3"/>`;vals.forEach((v,i)=>{let x=l+i*step,y=t+ph-v/M*ph,lab=fmtDay(labels[i]);let dailyLabel=(id==='dailyChart'||id==='daily2');let ly=dailyLabel?h-34:h-18;let transform=dailyLabel?'':' transform="rotate(-35 '+x+' '+ly+')"';s+=`<line x1="${x}" y1="${t+ph}" x2="${x}" y2="${t+ph+6}" stroke="#aeb9c2"/><circle cx="${x}" cy="${y}" r="4" fill="#1688e8"/><text x="${x}" y="${y-9}" text-anchor="middle" fill="#334450" font-weight="700">${v}</text><text x="${x}" y="${ly}" text-anchor="middle" fill="#334450" font-size="10" font-weight="700"${transform}>${esc(lab)}</text>`});s+=`<text x="${(l+w-r)/2}" y="${h-2}" text-anchor="middle" font-size="10" font-weight="700" fill="#596771">Día</text><text x="14" y="${t+ph/2}" text-anchor="middle" font-size="10" font-weight="700" fill="#596771" transform="rotate(-90 14 ${t+ph/2})">Recorridos</text></g></svg>`;el.innerHTML=s}
function donut(id,labels,vals){let w=$(id).clientWidth||400,h=300,cx=w/2,cy=125,r=76,total=vals.reduce((a,b)=>a+b,0)||1,start=-Math.PI/2,colors=['#1688e8','#f08b25','#1fb46b','#df3f4f','#7135c8'];let s=`<svg width="100%" height="${h}" viewBox="0 0 ${w} ${h}"><g font-family="Segoe UI">`;vals.forEach((v,i)=>{let end=start+v/total*Math.PI*2,x1=cx+r*Math.cos(start),y1=cy+r*Math.sin(start),x2=cx+r*Math.cos(end),y2=cy+r*Math.sin(end),large=end-start>Math.PI?1:0;s+=`<path d="M ${x1} ${y1} A ${r} ${r} 0 ${large} 1 ${x2} ${y2}" fill="none" stroke="${colors[i%colors.length]}" stroke-width="42"/>`;start=end});s+=`<text x="${cx}" y="${cy+3}" text-anchor="middle" font-size="22" font-weight="900" fill="#24313b">${total}</text><text x="${cx}" y="${cy+20}" text-anchor="middle" font-size="10" fill="#71808d">recorridos</text>`;labels.forEach((l,i)=>s+=`<rect x="15" y="${215+i*16}" width="9" height="9" rx="2" fill="${colors[i%colors.length]}"/><text x="30" y="${223+i*16}" font-size="10" fill="#596771">${esc(l)} (${vals[i]})</text>`);s+='</g></svg>';$(id).innerHTML=s}
function showMapTab(tab,b){
 document.querySelectorAll('.maptab').forEach(x=>x.classList.remove('active'));b.classList.add('active');
 ['Diurno','Nocturno','Coord'].forEach(x=>{let el=$('mapView'+x);if(el)el.style.display='none'});
 let target=tab==='diurno'?'mapViewDiurno':tab==='nocturno'?'mapViewNocturno':'mapViewCoord';
 $(target).style.display='block';
 setTimeout(()=>{
   let m=tab==='diurno'?mapDiurno:tab==='nocturno'?mapNocturno:map2;
   if(tab==='coord'&&m){updateMap(m,filt(),layers2);}
   if(m){m.invalidateSize(true);setTimeout(()=>m.invalidateSize(true),250);setTimeout(()=>m.invalidateSize(true),700)}
   if(tab==='diurno'||tab==='nocturno'){let month=$('month').value;if(month==='Todos')month=latestMonth();if(month)loadKML(month)}
 },80);
}
function page(id,b){document.querySelectorAll('.page').forEach(x=>x.classList.remove('active'));$(id).classList.add('active');document.querySelectorAll('.nav button').forEach(x=>x.classList.remove('active'));b.classList.add('active');setTimeout(()=>{[map1,map2,mapDiurno,mapNocturno].forEach(m=>{if(m)m.invalidateSize(true)});if(id==='mapa'){if(map2)updateMap(map2,filt(),layers2);refreshMonthlyMap()}if(id==='resumen'&&map1)changeSummaryRoute();render()},180)}

function panelCards(){return [...document.querySelectorAll('.grid .card')];}
function spanToPct(span){return ({3:'25%',4:'33%',6:'50%',8:'67%',9:'75%',12:'100%'})[span]||'50%'}
function pctToSpan(v){return ({'25%':3,'33%':4,'50%':6,'67%':8,'75%':9,'100%':12})[v]||6}
function openDesign(){let grid=$('designGrid');grid.innerHTML='<div class="head">Panel</div><div class="head">Ancho</div><div class="head">Altura (px)</div>';panelCards().forEach((c,i)=>{let id='p'+i;c.dataset.pid=id;let title=c.querySelector('h3')?.textContent||('Panel '+(i+1));let span=c.dataset.defaultSpan||((c.classList.contains('span3'))?'12':(c.classList.contains('span2')?'8':'6'));let saved=null;try{saved=JSON.parse(localStorage.getItem('spro26-'+id)||'null')}catch(e){saved=null}let pct=saved?.pct||spanToPct(span);let h=saved?.h||parseInt(getComputedStyle(c).minHeight)||320;grid.insertAdjacentHTML('beforeend',`<div><b>${esc(title)}</b></div><select id="dw-${id}"><option ${pct==='25%'?'selected':''}>25%</option><option ${pct==='33%'?'selected':''}>33%</option><option ${pct==='50%'?'selected':''}>50%</option><option ${pct==='67%'?'selected':''}>67%</option><option ${pct==='75%'?'selected':''}>75%</option><option ${pct==='100%'?'selected':''}>100%</option></select><input id="dh-${id}" type="number" min="280" max="900" step="10" value="${h}">`)});$('designPanel').classList.add('open')}
function applyDesign(){panelCards().forEach((c,i)=>{let id=c.dataset.pid||'p'+i;let pct=$(('dw-'+id)).value;let h=Math.max(280,Math.min(900,Number($(('dh-'+id)).value)||320));c.style.gridColumn=`span ${pctToSpan(pct)}`;c.style.height=h+'px';localStorage.setItem('spro26-'+id,JSON.stringify({pct,h}))});closeDesign();setTimeout(()=>{render();[map1,map2,mapDiurno,mapNocturno].forEach(m=>{if(m)m.invalidateSize(true)})},100)}
function resetDesign(){Object.keys(localStorage).filter(k=>k.indexOf('spro26-')===0).forEach(k=>localStorage.removeItem(k));panelCards().forEach((c,i)=>{let span=c.dataset.defaultSpan||((c.classList.contains('span3'))?'12':(c.classList.contains('span2')?'8':'6'));c.style.height='';c.style.gridColumn=`span ${span}`});openDesign()}
function closeDesign(){$('designPanel').classList.remove('open')}
function restoreDesign(){panelCards().forEach((c,i)=>{let id='p'+i;c.dataset.pid=id;let saved=null;try{saved=JSON.parse(localStorage.getItem('spro26-'+id)||'null')}catch(e){saved=null}let span=c.dataset.defaultSpan||((c.classList.contains('span3'))?'12':(c.classList.contains('span2')?'8':'6'));if(saved){c.style.gridColumn=`span ${pctToSpan(saved.pct)}`;c.style.height=saved.h+'px'}else c.style.gridColumn=`span ${span}`})}
function resetF(){['month','shift','tech','si','sf'].forEach(x=>$(x).value='Todos');$('d1').value='';$('d2').value='';render()}
function initMap(id){let el=$(id);el.style.minHeight='450px';let m=L.map(id,{preferCanvas:true,zoomControl:true}).setView([6.2442,-75.5812],11);let sat=L.tileLayer('https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}',{maxZoom:19,attribution:'Tiles © Esri'}).addTo(m);let calles=L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png',{maxZoom:19,attribution:'© OpenStreetMap contributors'});L.control.layers({'Satélite':sat,'Calles':calles},null,{collapsed:true,position:'topright'}).addTo(m);setTimeout(()=>m.invalidateSize(true),250);return m;}
function updateMap(m,d,layers){layers.forEach(x=>m.removeLayer(x));layers.length=0;let pts=[];d.forEach((r,i)=>{let la=+r.LatIni,lo=+r.LonIni,laf=+r.LatFin,lof=+r.LonFin;if(Number.isFinite(la)&&Number.isFinite(lo)){let mk=L.circleMarker([la,lo],{radius:6,color:'#1688e8',fillColor:'#1688e8',fillOpacity:.9}).bindPopup(`<b>INICIO</b><br>${esc(r.Fecha)} · ${esc(r['Técnico'])}<br>${esc(r.Turno)}<br>${esc(r['Archivo'])}<br>Lat: ${la.toFixed(6)}<br>Lon: ${lo.toFixed(6)}`);mk.addTo(m);layers.push(mk);pts.push([la,lo])}if(Number.isFinite(laf)&&Number.isFinite(lof)){let mk=L.circleMarker([laf,lof],{radius:6,color:'#df3f4f',fillColor:'#df3f4f',fillOpacity:.9}).bindPopup(`<b>FINALIZACIÓN</b><br>${esc(r.Fecha)} · ${esc(r['Técnico'])}<br>${esc(r.Turno)}<br>${esc(r['Archivo'])}<br>Lat: ${laf.toFixed(6)}<br>Lon: ${lof.toFixed(6)}`);mk.addTo(m);layers.push(mk);pts.push([laf,lof])}if(Number.isFinite(la)&&Number.isFinite(lo)&&Number.isFinite(laf)&&Number.isFinite(lof)){let ln=L.polyline([[la,lo],[laf,lof]],{color:'#1688e8',weight:1,opacity:.35,dashArray:'4 4'}).addTo(m);layers.push(ln)}});if(pts.length)m.fitBounds(pts,{padding:[20,20],maxZoom:13})}
function summaryRouteKey(r){return String(r['Archivo']||'')+'|'+String(r.Fecha||'')+'|'+String(r['Técnico']||'')+'|'+String(r.Turno||'')}
function setupSummaryRouteSelector(d){
 const sel=$('summaryRouteSelect'); if(!sel)return;
 const previousFile=sel.dataset.selectedFile||'';
 let html='<option value="__all__">Todos los recorridos</option><option value="__diurno__">☀️ Todos los recorridos diurnos</option><option value="__nocturno__">🌙 Todos los recorridos nocturnos</option>';
 d.slice().sort((a,b)=>String(a.Fecha).localeCompare(String(b.Fecha))||String(a['Técnico']).localeCompare(String(b['Técnico']))||String(a.Archivo).localeCompare(String(b.Archivo))).forEach((r,i)=>{
   html+=`<option value="route-${i}" data-key="${esc(summaryRouteKey(r))}">${esc(String(r.Fecha||'')+' · '+String(r['Técnico']||'Sin técnico')+' · '+String(r.Turno||'')+' · '+String(r['Archivo']||''))}</option>`;
 });
 sel.innerHTML=html;
 if(previousFile){const opts=[...sel.options];const found=opts.find(o=>o.dataset.key===previousFile);if(found)sel.value=found.value;}
 changeSummaryRoute();
}
function changeSummaryRoute(){
 const sel=$('summaryRouteSelect');if(!sel||!map1)return;
 const d=filt();let val=sel.value,view=d, label='Todos los recorridos según los filtros superiores.';
 if(val==='__diurno__'){view=d.filter(r=>String(r.Turno).toUpperCase()==='DIURNO');label=`☀️ ${view.length} recorridos diurnos seleccionados.`;}
 else if(val==='__nocturno__'){view=d.filter(r=>String(r.Turno).toUpperCase()==='NOCTURNO');label=`🌙 ${view.length} recorridos nocturnos seleccionados.`;}
 else if(val!=='__all__'){
   const opts=[...sel.options],opt=opts.find(o=>o.value===val),key=opt?.dataset.key||'';
   const r=d.find(x=>summaryRouteKey(x)===key);
   view=r?[r]:[];
   if(r){sel.dataset.selectedFile=summaryRouteKey(r);label=`📍 ${r.Fecha||''} · ${r['Técnico']||''} · ${String(r.Turno||'').toUpperCase()} · ${r.Archivo||''}`;}
 } else sel.dataset.selectedFile='';
 $('summaryRouteInfo').textContent=label;
 updateMap(map1,view,layers1);
}


function latestMonth(){
  let vals=DATA.map(r=>String(r['Año-Mes']||'')).filter(x=>/^\d{4}-\d{2}$/.test(x)).sort();
  return vals.length?vals[vals.length-1]:null;
}
function routeColor(key,def){let v=localStorage.getItem('spro-route-'+key);return /^#[0-9A-Fa-f]{6}$/.test(v||'')?v:def}
function setRouteColor(key,color){if(!/^#[0-9A-Fa-f]{6}$/.test(color||''))return;localStorage.setItem('spro-route-'+key,color);if(key==='diurno'&&$('colorDiurno'))$('colorDiurno').value=color;if(key==='nocturno'&&$('colorNocturno'))$('colorNocturno').value=color;let month=$('month')?.value;if(month==='Todos')month=latestMonth();if(month&&kmlCache[month]){if(key==='diurno'&&kmlLayer1&&mapDiurno){mapDiurno.removeLayer(kmlLayer1);kmlLayer1=addKmlToMap(mapDiurno,kmlCache[month].diurno,color);if(kmlLayer1)try{mapDiurno.fitBounds(kmlLayer1.getBounds(),{padding:[20,20],maxZoom:14})}catch(e){}}if(key==='nocturno'&&kmlLayer2&&mapNocturno){mapNocturno.removeLayer(kmlLayer2);kmlLayer2=addKmlToMap(mapNocturno,kmlCache[month].nocturno,color);if(kmlLayer2)try{mapNocturno.fitBounds(kmlLayer2.getBounds(),{padding:[20,20],maxZoom:14})}catch(e){}}}else refreshMonthlyMap()}
function restoreRouteColors(){if($('colorDiurno'))$('colorDiurno').value=routeColor('diurno','#15945a');if($('colorNocturno'))$('colorNocturno').value=routeColor('nocturno','#1688e8')}
function clearKmlLayers(){
 if(kmlLayer1&&mapDiurno)mapDiurno.removeLayer(kmlLayer1);
 if(kmlLayer2&&mapNocturno)mapNocturno.removeLayer(kmlLayer2);
 kmlLayer1=null;kmlLayer2=null;
}
function addKmlToMap(map,obj,color){
 if(!obj||!obj.exists||obj.error||!obj.features||!obj.features.length)return null;
 return L.geoJSON(obj,{style:()=>({color:color,weight:5,opacity:.88}),
 pointToLayer:(f,latlng)=>L.circleMarker(latlng,{radius:3,color:'#ffffff',fillColor:'#5b6b78',fillOpacity:.65,weight:1}),
 onEachFeature:(f,l)=>{let p=f.properties||{};if(p.name||p.description)l.bindPopup(`<b>${esc(p.name||'Recorrido')}</b><br>${esc(p.description||'')}`)}}).addTo(map);
}
function drawMonthlyKML(month){
 if(!mapDiurno||!mapNocturno||!kmlCache[month])return;
 let j=kmlCache[month];
 if(kmlLayer1)mapDiurno.removeLayer(kmlLayer1);
 if(kmlLayer2)mapNocturno.removeLayer(kmlLayer2);
 kmlLayer1=addKmlToMap(mapDiurno,j.diurno,routeColor('diurno','#15945a'));
 kmlLayer2=addKmlToMap(mapNocturno,j.nocturno,routeColor('nocturno','#1688e8'));
 let dOk=j.diurno&&j.diurno.exists&&!j.diurno.error&&j.diurno.features&&j.diurno.features.length>0,nOk=j.nocturno&&j.nocturno.exists&&!j.nocturno.error&&j.nocturno.features&&j.nocturno.features.length>0;
 $('kmlDStatus').textContent='KML diurno: '+(dOk?'🟢 CONECTADO':(j.diurno&&j.diurno.exists?'🟠 ERROR DE LECTURA':'🔴 NO ENCONTRADO'));
 $('kmlNStatus').textContent='KML nocturno: '+(nOk?'🟢 CONECTADO':(j.nocturno&&j.nocturno.exists?'🟠 ERROR DE LECTURA':'🔴 NO ENCONTRADO'));
 $('kmlDFile').innerHTML=dOk?`<p><b>🟢 Conectado</b></p><p><small>${esc(j.diurno.file)}</small></p><p>Elementos: ${j.diurno.features.length}</p>`:`<p class="bad">${j.diurno&&j.diurno.exists?'El KML diurno existe, pero no pudo leerse: '+esc(j.diurno.error||'error desconocido'):'No se encontró el KML diurno para '+esc(month)}.</p><p><small>Ruta esperada: ${esc(kml_candidates_text(month,'DIURNO'))}</small></p>`;
 $('kmlNFile').innerHTML=nOk?`<p><b>🟢 Conectado</b></p><p><small>${esc(j.nocturno.file)}</small></p><p>Elementos: ${j.nocturno.features.length}</p>`:`<p class="bad">${j.nocturno&&j.nocturno.exists?'El KML nocturno existe, pero no pudo leerse: '+esc(j.nocturno.error||'error desconocido'):'No se encontró el KML nocturno para '+esc(month)}.</p><p><small>Ruta esperada: ${esc(kml_candidates_text(month,'NOCTURNO'))}</small></p>`;
 $('kmld1').textContent=dOk?j.diurno.features.length:0;$('kmld2').textContent=nOk?j.nocturno.features.length:0;
 $('coordStatus').textContent='Coordenadas del Excel · '+month;
 if(kmlLayer1){try{mapDiurno.fitBounds(kmlLayer1.getBounds(),{padding:[20,20],maxZoom:14})}catch(e){}}
 if(kmlLayer2){try{mapNocturno.fitBounds(kmlLayer2.getBounds(),{padding:[20,20],maxZoom:14})}catch(e){}}
}
function kml_candidates_text(month,turno){
 let p=String(month).split('-');if(p.length!==2)return 'mes no disponible';
 let y=p[0],m=Number(p[1]),abbr=['','ENE','FEB','MAR','ABR','MAY','JUN','JUL','AGO','SEP','OCT','NOV','DIC'][m]||'';
 let folder=abbr+y;
 return turno==='DIURNO'?`...\\Recorrido_Diurno\\${y}\\${folder}\\REC_DIUR_${folder}.kml`:`...\\Recorrido_Nocturno\\${folder}\\REC_NOCT_${folder}.kml`;
}
async function loadKML(month){
 if(!month)return;
 try{let r=await fetch('/api/kml?month='+encodeURIComponent(month)+'&x='+Date.now()),j=await r.json();kmlCache[month]=j;drawMonthlyKML(month)}
 catch(e){$('kmlDStatus').textContent='KML diurno: error';$('kmlNStatus').textContent='KML nocturno: error'}
}
function refreshMonthlyMap(){
 if(!mapDiurno||!mapNocturno)return;
 let month=$('month').value;if(month==='Todos')month=latestMonth();if(!month)return;
 if(!kmlCache[month])loadKML(month);else drawMonthlyKML(month);
 mapDiurno.invalidateSize(true);mapNocturno.invalidateSize(true);
}
function render(){
let d=filt();setupSummaryRouteSelector(d);let sa=d.filter(r=>!['NO DISPONIBLE','REVISAR','REVISAR GPX'].includes(r['Estado inicio'])),ea=d.filter(r=>!['NO DISPONIBLE','REVISAR','REVISAR GPX'].includes(r['Estado fin']));
let sok=sa.filter(r=>['INICIO CUMPLE','INICIO ANTES DEL HORARIO'].includes(r['Estado inicio'])).length,eok=ea.filter(r=>['FIN CUMPLE','TERMINÓ DESPUÉS'].includes(r['Estado fin'])).length;
let late=d.filter(r=>r['Estado inicio']==='INICIO TARDÍO'),early=d.filter(r=>r['Estado fin']==='TERMINÓ ANTES'),gps=d.filter(r=>(+r['Huecos GPS >=10 min']||0)>0),rev=d.filter(r=>String(r['Estado recorrido']||'').includes('REVISAR'));
let km=d.reduce((a,r)=>a+(+r.Kilómetros||0),0),gpsN=d.reduce((a,r)=>a+(+r['Huecos GPS >=10 min']||0),0),dev=d.reduce((a,r)=>a+(r['Estado inicio']==='INICIO TARDÍO'?+r.DifIniSec||0:0)+(r['Estado fin']==='TERMINÓ ANTES'?+r.DifFinSec||0:0),0);
$('k1').textContent=d.length;$('k2').textContent=fmt(km);$('k3').textContent=pct(sok,sa.length);$('k4').textContent=pct(eok,ea.length);$('k5').textContent=late.length;$('k6').textContent=early.length;$('k7').textContent=Math.round(gpsN);$('k8').textContent=secFmt(dev);
$('a1').textContent=late.length;$('a2').textContent=early.length;$('a3').textContent=Math.round(gpsN);$('a4').textContent=rev.length;$('gm1').textContent=Math.round(gpsN);$('gm2').textContent=gps.length;$('gm3').textContent=pct(gps.length,d.length);$('fstatus').textContent=`${d.length} recorridos · ${new Set(d.map(r=>r['Técnico'])).size} técnicos`;
let by={};d.forEach(r=>{let t=r['Técnico']||'Sin técnico';if(!by[t])by[t]={n:0,km:0,late:0,early:0,gps:0,okS:0,avS:0,okE:0,avE:0,dev:0};let x=by[t];x.n++;x.km+=+r.Kilómetros||0;x.gps+=+r['Huecos GPS >=10 min']||0;if(!['NO DISPONIBLE','REVISAR','REVISAR GPX'].includes(r['Estado inicio'])){x.avS++;if(['INICIO CUMPLE','INICIO ANTES DEL HORARIO'].includes(r['Estado inicio']))x.okS++}if(!['NO DISPONIBLE','REVISAR','REVISAR GPX'].includes(r['Estado fin'])){x.avE++;if(['FIN CUMPLE','TERMINÓ DESPUÉS'].includes(r['Estado fin']))x.okE++}if(r['Estado inicio']==='INICIO TARDÍO'){x.late++;x.dev+=+r.DifIniSec||0}if(r['Estado fin']==='TERMINÓ ANTES'){x.early++;x.dev+=+r.DifFinSec||0}});
let names=Object.keys(by);bars('techChart',names,[names.map(t=>by[t].avS?100*by[t].okS/by[t].avS:0),names.map(t=>by[t].avE?100*by[t].okE/by[t].avE:0)],['Inicio','Fin'],100);hbars('lateChart',names.slice().sort((a,b)=>by[b].late-by[a].late),names.slice().sort((a,b)=>by[b].late-by[a].late).map(t=>by[t].late),'#df3f4f');hbars('earlyChart',names.slice().sort((a,b)=>by[b].early-by[a].early),names.slice().sort((a,b)=>by[b].early-by[a].early).map(t=>by[t].early),'#f08b25');
let daily={};d.forEach(r=>daily[r.Fecha]=(daily[r.Fecha]||0)+1);let dates=Object.keys(daily).sort();lineChart('dailyChart',dates,dates.map(x=>daily[x]));lineChart('daily2',dates,dates.map(x=>daily[x]));
donut('shiftChart',['DIURNO','NOCTURNO'],['DIURNO','NOCTURNO'].map(x=>d.filter(r=>r.Turno===x).length));
let ss={};d.forEach(r=>ss[r['Estado inicio']||'SIN DATO']=(ss[r['Estado inicio']||'SIN DATO']||0)+1);donut('startChart',Object.keys(ss),Object.values(ss));
let es={};d.forEach(r=>es[r['Estado fin']||'SIN DATO']=(es[r['Estado fin']||'SIN DATO']||0)+1);donut('endChart',Object.keys(es),Object.values(es));
let st=names.slice().sort((a,b)=>by[b].km-by[a].km);let th='<thead><tr><th>#</th><th>Técnico</th><th>Recorridos</th><th>Km</th><th>% Inicio</th><th>% Fin</th><th>Tardíos</th><th>Terminó antes</th><th>GPS</th><th>Desviación</th></tr></thead><tbody>';st.forEach((t,i)=>{let x=by[t];th+=`<tr><td>${i+1}</td><td><b>${esc(t)}</b></td><td>${x.n}</td><td>${x.km.toFixed(1)}</td><td class="${x.okS<x.avS*.8?'bad':'good'}">${pct(x.okS,x.avS)}</td><td class="${x.okE<x.avE*.8?'bad':'good'}">${pct(x.okE,x.avE)}</td><td class="${x.late?'bad':'good'}">${x.late}</td><td class="${x.early?'bad':'good'}">${x.early}</td><td class="${x.gps?'warn':'good'}">${x.gps}</td><td>${secFmt(x.dev)}</td></tr>`});$('rankTable').innerHTML=th+'</tbody>';$('techTable').innerHTML=th+'</tbody>';
let ah='<thead><tr><th>Fecha</th><th>Técnico</th><th>Turno</th><th>Archivo</th><th>Inicio</th><th>Fin</th><th>GPS</th><th>Estado</th><th>Prioridad</th><th>Qué verificar</th></tr></thead><tbody>';let alerts=d.filter(r=>r['Estado inicio']==='INICIO TARDÍO'||r['Estado fin']==='TERMINÓ ANTES'||(+r['Huecos GPS >=10 min']||0)>0||String(r['Estado recorrido']||'').includes('REVISAR')).sort((a,b)=>String(b.Fecha).localeCompare(String(a.Fecha)));alerts.forEach(r=>{let si=String(r['Estado inicio']||''),sf=String(r['Estado fin']||''),sr=String(r['Estado recorrido']||''),g=+r['Huecos GPS >=10 min']||0;let ci=si==='INICIO CUMPLE'?'good':si==='INICIO ANTES DEL HORARIO'?'info':si==='INICIO TARDÍO'?'bad':'warn';let cf=sf==='FIN CUMPLE'?'good':sf==='TERMINÓ DESPUÉS'?'info':sf==='TERMINÓ ANTES'?'bad':'warn';let cg=g>0?'warn':'good';let reasons=[];if(si==='INICIO TARDÍO')reasons.push('Verificar hora real de inicio');if(sf==='TERMINÓ ANTES')reasons.push('Verificar finalización anticipada');if(g>0)reasons.push('Revisar continuidad GPS');if(sr.includes('REVISAR'))reasons.push('Revisar GPX/registro');let pr=(si==='INICIO TARDÍO'||sf==='TERMINÓ ANTES')?'ALTA':(g>0||sr.includes('REVISAR')?'MEDIA':'BAJA');let pc=pr==='ALTA'?'bad':pr==='MEDIA'?'warn':'good';ah+=`<tr><td>${r.Fecha||''}</td><td>${esc(r['Técnico'])}</td><td>${r.Turno||''}</td><td>${esc(r.Archivo)}</td><td class="${ci}">${esc(si||'SIN DATO')}</td><td class="${cf}">${esc(sf||'SIN DATO')}</td><td class="${cg}">${g}</td><td>${esc(sr)}</td><td class="${pc}">${pr}</td><td>${esc(reasons.join(' · ')||'Revisión general')}</td></tr>`});$('alertTable').innerHTML=ah+'</tbody>';$('recentTable').innerHTML=ah+'</tbody>';
let rh='<thead><tr><th>Fecha</th><th>Técnico</th><th>Turno</th><th>Inicio</th><th>Fin</th><th>Duración</th><th>Km</th><th>Vel.</th><th>Estado inicio</th><th>Estado fin</th><th>Estado recorrido</th></tr></thead><tbody>';d.slice().sort((a,b)=>String(b.Fecha).localeCompare(String(a.Fecha))).forEach(r=>{let si=String(r['Estado inicio']||''),sf=String(r['Estado fin']||''),sr=String(r['Estado recorrido']||'');let ci=si==='INICIO CUMPLE'?'good':si==='INICIO ANTES DEL HORARIO'?'info':si==='INICIO TARDÍO'?'bad':'warn';let cf=sf==='FIN CUMPLE'?'good':sf==='TERMINÓ DESPUÉS'?'info':sf==='TERMINÓ ANTES'?'bad':'warn';let ce=sr==='DENTRO DE HORARIO'?'good':sr.includes('ANTICIPADO')?'bad':sr.includes('POSTERIOR')?'info':sr.includes('REVISAR')?'warn':'warn';rh+=`<tr><td>${r.Fecha||''}</td><td>${esc(r['Técnico'])}</td><td>${r.Turno||''}</td><td>${esc(r['Inicio Colombia'])}</td><td>${esc(r['Fin Colombia'])}</td><td>${esc(r.Duración)}</td><td>${(+r.Kilómetros||0).toFixed(1)}</td><td>${(+r['Velocidad promedio (km/h)']||0).toFixed(1)}</td><td class="${ci}">${esc(si)}</td><td class="${cf}">${esc(sf)}</td><td class="${ce}">${esc(sr)}</td></tr>`});$('routeTable').innerHTML=rh+'</tbody>';
let gt={};d.forEach(r=>{let t=r['Técnico']||'Sin técnico';gt[t]=(gt[t]||0)+(+r['Huecos GPS >=10 min']||0)});let gn=Object.keys(gt).sort((a,b)=>gt[b]-gt[a]);hbars('gpsChart',gn,gn.map(x=>gt[x]),'#7135c8');let aff={};d.forEach(r=>{if((+r['Huecos GPS >=10 min']||0)>0){let t=r['Técnico']||'Sin técnico';aff[t]=(aff[t]||0)+1}});let an=Object.keys(aff).sort((a,b)=>aff[b]-aff[a]);hbars('gpsRoutes',an,an.map(x=>aff[x]),'#1688e8');timeHbars('devChart',names,names.map(t=>by[t].dev));let sh={};d.forEach(r=>{let h=String(r['Inicio Colombia']||'').slice(0,2);if(h)sh[h]=(sh[h]||0)+1});let sn=Object.keys(sh).sort();bars('startHour',sn,[sn.map(x=>sh[x])]);let eh={};d.forEach(r=>{let h=String(r['Fin Colombia']||'').slice(0,2);if(h)eh[h]=(eh[h]||0)+1});let en=Object.keys(eh).sort();bars('endHour',en,[en.map(x=>eh[x])]);
if(map1)changeSummaryRoute();if(map2)updateMap(map2,d,layers2);$('kmld3').textContent=d.length;if(document.getElementById('mapa')?.classList.contains('active'))refreshMonthlyMap();
}
function clock(){let d=new Date();$('date').textContent=d.toLocaleDateString('es-CO',{weekday:'long',day:'numeric',month:'long',year:'numeric'});$('time').textContent=d.toLocaleTimeString('es-CO',{hour12:false})}
async function load(){try{let r=await fetch('/api/data?x='+Date.now()),j=await r.json();if(j.error)throw Error(j.error);if(j.stamp!==lastStamp){let keep={m:$('month').value,t:$('tech').value,s:$('shift').value,a:$('d1').value,b:$('d2').value,si:$('si').value,sf:$('sf').value};DATA=j.rows;lastStamp=j.stamp;setupFilters();$('month').value=keep.m;$('tech').value=keep.t;$('shift').value=keep.s;$('d1').value=keep.a;$('d2').value=keep.b;$('si').value=keep.si;$('sf').value=keep.sf;render()}$('live').textContent='● Excel conectado · '+j.updated}catch(e){$('live').textContent='● Error leyendo Excel: '+String(e.message||e).slice(0,80)}}clock();setInterval(clock,1000);setInterval(load,5000);setTimeout(()=>{map1=initMap('map1');map2=initMap('map2');mapDiurno=initMap('mapDiurno');mapNocturno=initMap('mapNocturno');setTimeout(()=>{restoreDesign();restoreRouteColors();[map1,map2,mapDiurno,mapNocturno].forEach(m=>{if(m)m.invalidateSize(true)});render()},350)},350);load();
</script>
</body></html>
"""

def read_excel():
    if not os.path.exists(EXCEL_PATH):
        raise FileNotFoundError(EXCEL_PATH)
    try:
        df=pd.read_excel(EXCEL_PATH,sheet_name="RECORRIDOS")
    except Exception as exc:
        # Fallback para casos en que Excel esté abierto/bloqueado por otro proceso.
        from openpyxl import load_workbook
        wb=load_workbook(EXCEL_PATH,data_only=True,read_only=True)
        ws=wb["RECORRIDOS"]
        rows=list(ws.iter_rows(values_only=True))
        if not rows: raise exc
        headers=[str(x).strip() if x is not None else "" for x in rows[0]]
        df=pd.DataFrame(rows[1:],columns=headers)
        wb.close()
    required=["Fecha","Año-Mes","Técnico","Turno","Inicio Colombia","Fin Colombia","Diferencia inicio","Diferencia fin","Estado inicio","Estado fin","Kilómetros","Velocidad promedio (km/h)","Huecos GPS >=10 min","Estado recorrido","Duración","Archivo"]
    missing=[c for c in required if c not in df.columns]
    if missing: raise ValueError("Faltan columnas en RECORRIDOS: "+", ".join(missing))
    # Coordenadas: acepta nombres usados por las versiones anteriores del sistema.
    def coord_col(cands):
        for c in cands:
            if c in df.columns:return c
        return None
    ci=coord_col(["Latitud inicio","Latitud inicio recorrido","Lat Inicio","Latitud Inicio"])
    li=coord_col(["Longitud inicio","Longitud inicio recorrido","Lon Inicio","Longitud Inicio"])
    cf=coord_col(["Latitud fin","Latitud fin recorrido","Lat Fin","Latitud Fin"])
    lf=coord_col(["Longitud fin","Longitud fin recorrido","Lon Fin","Longitud Fin"])
    for name,col in [("LatIni",ci),("LonIni",li),("LatFin",cf),("LonFin",lf)]:
        df[name]=pd.to_numeric(df[col],errors="coerce") if col else float("nan")
    dt=pd.to_datetime(df["Fecha"],dayfirst=True,errors="coerce")
    df["Fecha"]=dt.dt.strftime("%Y-%m-%d")
    df["Turno"]=df["Turno"].astype(str).str.upper()
    for c in ["Diferencia inicio","Diferencia fin"]:
        df[c]=pd.to_timedelta(df[c],errors="coerce").dt.total_seconds().fillna(0)
    df["DifIniSec"]=df["Diferencia inicio"];df["DifFinSec"]=df["Diferencia fin"]
    return df.replace({pd.NA:None,float("nan"):None}).to_dict("records")

def stamp():
    try:return str(os.path.getmtime(EXCEL_PATH))
    except:return "NO_FILE"

class Handler(BaseHTTPRequestHandler):
    def log_message(self,*a):pass
    def do_GET(self):
        p=urlparse(self.path).path
        if p=="/":
            b=HTML.encode("utf-8");self.send_response(200);self.send_header("Content-Type","text/html; charset=utf-8");self.send_header("Content-Length",str(len(b)));self.end_headers();self.wfile.write(b);return
        if p=="/api/kml":
            try:
                from urllib.parse import parse_qs
                qs=parse_qs(urlparse(self.path).query)
                month=qs.get("month",[""])[0]
                b=json.dumps(monthly_kml(month),ensure_ascii=False,default=str).encode("utf-8")
                self.send_response(200);self.send_header("Content-Type","application/json; charset=utf-8")
                self.send_header("Cache-Control","no-store");self.end_headers();self.wfile.write(b)
            except Exception as e:
                b=json.dumps({"error":str(e)},ensure_ascii=False).encode("utf-8")
                self.send_response(500);self.send_header("Content-Type","application/json; charset=utf-8")
                self.end_headers();self.wfile.write(b)
            return
        if p=="/api/data":
            try:
                b=json.dumps({"stamp":stamp(),"updated":time.strftime("%d/%m/%Y %H:%M:%S"),"rows":read_excel()},ensure_ascii=False,default=str).encode("utf-8")
                self.send_response(200);self.send_header("Content-Type","application/json; charset=utf-8");self.send_header("Cache-Control","no-store");self.end_headers();self.wfile.write(b)
            except Exception as e:
                b=json.dumps({"error":str(e)},ensure_ascii=False).encode("utf-8");self.send_response(500);self.send_header("Content-Type","application/json; charset=utf-8");self.end_headers();self.wfile.write(b)

if __name__=="__main__":
    print("SISTEMA PRO V2.6 - TABLERO EJECUTIVO PROFESIONAL")
    print("Excel:",EXCEL_PATH)
    if not os.path.exists(EXCEL_PATH):
        print("ERROR: no se encontro el Excel.")
        input("ENTER para cerrar...")
    else:
        url=f"http://{HOST}:{PORT}/";print("Tablero:",url);print("Actualizacion:",REFRESH_SECONDS,"segundos")
        if not RENDER_MODE:
            threading.Timer(1.0,lambda:webbrowser.open(url)).start()
        ThreadingHTTPServer((HOST,PORT),Handler).serve_forever()
