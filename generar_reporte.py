#!/usr/bin/env python3
"""
Genera el reporte HTML autocontenido de estatus de distribuidores FOTON.

Uso:
    python3 generar_reporte.py                  # todos los grupos de config.json
    python3 generar_reporte.py --grupo velcen   # sólo un grupo
    python3 generar_reporte.py --config otra.json

Lee:
  - datos/KPIS.xlsx (Anexos, % Desarrollo de Proyecto, Software, Stock, Diagnóstico, Cummins)
  - Sistema de Personal FOTON (Personal: hojas Evaluación_Red, Regional_*, Cálculo_Red)
  - datos/fotos/fotos.json (fotos asignadas a conceptos y observaciones)
Escribe: salida/<archivo_salida>.html (un archivo por grupo, funciona sin internet).
"""
import argparse
import base64
import datetime as dt
import html
import io
import json
import os
import re
import sys
import unicodedata
import warnings

import openpyxl
from PIL import Image, ImageOps

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")
BASE = os.path.dirname(os.path.abspath(__file__))

ESTADO_TXT = {"verde": "Cumple", "amarillo": "En proceso", "rojo": "No cumple",
              "na": "No aplica", "sin": "Sin información"}
ESTADO_ICO = {"verde": "✓", "amarillo": "!", "rojo": "✕", "na": "–", "sin": "?"}


# --------------------------------------------------------------------------- utilidades
def norm(s):
    """Texto normalizado para comparar (sin acentos, minúsculas, espacios simples)."""
    s = unicodedata.normalize("NFKD", str(s or ""))
    s = "".join(c for c in s if not unicodedata.combining(c))
    return " ".join(s.lower().split())


def limpio(s):
    return " ".join(str(s or "").split())


def esc(s):
    return html.escape(str(s if s is not None else ""), quote=True)


def pct(v, dec=1):
    if v is None:
        return "—"
    txt = f"{v:.{dec}f}".rstrip("0").rstrip(".") if dec else f"{v:.0f}"
    return f"{txt}%"


MESES = ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic"]
MESES_LARGO = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto",
               "septiembre", "octubre", "noviembre", "diciembre"]


def fecha_corta(d):
    return f"{d.day}-{MESES[d.month - 1]}-{d.year}" if d else "—"


def fecha_larga(d):
    return f"{d.day} de {MESES_LARGO[d.month - 1]} de {d.year}"


def mes_txt(d):
    return f"{MESES[d.month - 1]}-{d.year}" if d else "—"


def semaforo(valor, regla):
    if valor is None:
        return "sin"
    if valor >= regla["verde_min"] - 1e-9:
        return "verde"
    if regla.get("amarillo_min") is not None and valor >= regla["amarillo_min"] - 1e-9:
        return "amarillo"
    return "rojo"


# --------------------------------------------------------------------------- KPI's
class LectorKPI:
    COLS = {"nombre": "nombre comercial", "fecha": "fecha", "tema": "tema",
            "subtema": "subtema", "concepto": "concepto", "req": "requerido", "adq": "adquirido"}

    def __init__(self, ruta):
        self.wb = openpyxl.load_workbook(ruta, data_only=True, read_only=True)
        self.cache = {}

    def filas(self, hoja):
        if hoja in self.cache:
            return self.cache[hoja]
        if hoja not in self.wb.sheetnames:
            raise SystemExit(f"ERROR: no existe la hoja «{hoja}» en el Excel de KPI's.")
        ws = self.wb[hoja]
        it = ws.iter_rows(values_only=True)
        enc = [norm(x) for x in next(it)]
        idx = {}
        for k, nombre in self.COLS.items():
            if nombre not in enc:
                raise SystemExit(f"ERROR: la hoja «{hoja}» no tiene la columna «{nombre}».")
            idx[k] = enc.index(nombre)
        out = []
        for n, fila in enumerate(it, start=2):
            if not fila or fila[idx["nombre"]] in (None, ""):
                continue
            r = {k: fila[i] if i < len(fila) else None for k, i in idx.items()}
            r["fila"] = n
            out.append(r)
        self.cache[hoja] = (out, idx)
        return self.cache[hoja]

    def indicador(self, hoja, nombre_kpi, corte, concepto=None):
        filas, idx = self.filas(hoja)
        propias = [r for r in filas if norm(r["nombre"]) == norm(nombre_kpi)
                   and isinstance(r["fecha"], dt.datetime) and r["fecha"].date() <= corte
                   and (concepto is None or norm(r["concepto"]) == norm(concepto))]
        if not propias:
            return None
        mes = max(r["fecha"] for r in propias)
        conceptos = []
        for r in propias:
            if r["fecha"] != mes:
                continue
            req = float(r["req"] or 0)
            adq = float(r["adq"] or 0)
            avance = min(adq / req, 1.0) if req > 0 else (1.0 if adq else 0.0)
            conceptos.append({
                "fila": r["fila"], "subtema": limpio(r["subtema"]), "concepto": limpio(r["concepto"]),
                "req": req, "adq": adq, "avance": avance,
            })
        valor = 100.0 * sum(c["avance"] for c in conceptos) / len(conceptos)
        letra = openpyxl.utils.get_column_letter
        return {
            "valor": valor, "mes": mes.date(), "conceptos": conceptos,
            "ubicacion": f"Hoja «{hoja}», filas {conceptos[0]['fila']}–{conceptos[-1]['fila']}"
                         f" (col. {letra(idx['concepto'] + 1)} concepto, {letra(idx['req'] + 1)} req., "
                         f"{letra(idx['adq'] + 1)} adq.)",
        }


# --------------------------------------------------------------------------- Personal
class LectorPersonal:
    def __init__(self, ruta):
        self.ruta = ruta
        self.wb = openpyxl.load_workbook(ruta, data_only=True)

    def _buscar_tabla(self, ws, encabezados):
        """Devuelve (fila_encabezado, col_inicio) de la tabla cuyo encabezado empieza con 'Distribuidor'."""
        for row in ws.iter_rows():
            for c in row:
                if norm(c.value) == "distribuidor":
                    vecinos = [norm(ws.cell(c.row, c.column + i).value) for i in range(1, 4)]
                    if all(e in vecinos for e in encabezados):
                        return c.row, c.column
        return None, None

    def resumen(self, nombre):
        ws = self.wb["Evaluación_Red"]
        fila, col = self._buscar_tabla(ws, ["categoria", "nivel de taller"])
        if fila is None:
            return None
        for r in range(fila + 1, ws.max_row + 1):
            if norm(ws.cell(r, col).value) == norm(nombre):
                v = [ws.cell(r, col + i).value for i in range(0, 8)]
                return {"categoria": v[1], "nivel": v[2], "req": v[3], "cub": v[4], "faltan": v[5],
                        "extra": v[6], "pct": (v[7] or 0) * 100 if isinstance(v[7], (int, float)) else None,
                        "ubicacion": f"Hoja «Evaluación_Red», fila {r}"}
        return None

    def por_area(self, nombre):
        for ws in self.wb.worksheets:
            if not ws.title.startswith("Regional_"):
                continue
            fila, col = self._buscar_tabla(ws, ["categoria", "reporte"])
            if fila is None:
                continue
            for r in range(fila + 1, ws.max_row + 1):
                if norm(ws.cell(r, col).value) == norm(nombre):
                    g = lambda i: ws.cell(r, col + i).value
                    areas = []
                    for etiqueta, i in (("Administrativo", 3), ("Comercial", 6), ("Postventa", 9)):
                        p = g(i + 2)
                        areas.append({"area": etiqueta, "req": g(i), "cub": g(i + 1),
                                      "pct": p * 100 if isinstance(p, (int, float)) else None})
                    return {"areas": areas, "ubicacion": f"Hoja «{ws.title}», fila {r}"}
        return None

    def puestos(self, nombre):
        if "Cálculo_Red" not in self.wb.sheetnames:
            return []
        ws = self.wb["Cálculo_Red"]
        enc = [norm(c.value) for c in ws[1]]
        i = {k: enc.index(k) for k in ("distribuidor", "area", "puesto", "requerido", "cubierto", "faltan", "extra")}
        out = []
        for n, fila in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
            if norm(fila[i["distribuidor"]]) != norm(nombre):
                continue
            out.append({"fila": n, "area": fila[i["area"]], "puesto": limpio(fila[i["puesto"]]),
                        "req": fila[i["requerido"]] or 0, "cub": fila[i["cubierto"]] or 0,
                        "faltan": fila[i["faltan"]] or 0, "extra": fila[i["extra"]] or 0})
        return out


# --------------------------------------------------------------------------- Fotos
def fecha_foto(ruta, entrada):
    if entrada.get("fecha"):
        return dt.date.fromisoformat(entrada["fecha"]), "indicada en fotos.json"
    try:
        exif = Image.open(ruta).getexif()
        for tag in (36867, 306):
            v = exif.get(tag) or exif.get_ifd(0x8769).get(tag)
            if v:
                return dt.datetime.strptime(str(v)[:10], "%Y:%m:%d").date(), "fecha de la foto (EXIF)"
    except Exception:
        pass
    m = re.search(r"(20\d\d)-(\d\d)-(\d\d)", os.path.basename(ruta))
    if m:
        return dt.date(int(m[1]), int(m[2]), int(m[3])), "fecha en el nombre del archivo"
    return dt.date.fromtimestamp(os.path.getmtime(ruta)), "fecha de modificación del archivo"


def foto_b64(ruta, ancho, calidad):
    im = ImageOps.exif_transpose(Image.open(ruta)).convert("RGB")
    im.thumbnail((ancho, ancho))
    buf = io.BytesIO()
    im.save(buf, "JPEG", quality=calidad, optimize=True, progressive=True)
    return "data:image/jpeg;base64," + base64.b64encode(buf.getvalue()).decode()


class Fotos:
    def __init__(self, ruta_json, cfg, corte):
        self.base = os.path.dirname(ruta_json)
        self.cfg, self.corte = cfg, corte
        self.datos = json.load(open(ruta_json, encoding="utf-8")) if os.path.exists(ruta_json) else {}
        self.usadas = []
        self._cache = {}

    def _preparar(self, e):
        ruta = os.path.join(self.base, e["archivo"])
        if not os.path.exists(ruta) or os.path.getsize(ruta) == 0:
            print(f"  aviso: foto no encontrada o vacía: {e['archivo']}", file=sys.stderr)
            return None
        if ruta not in self._cache:
            f, origen = fecha_foto(ruta, e)
            limite = self.corte - dt.timedelta(days=round(self.cfg["meses_foto_reciente"] * 30.44))
            self._cache[ruta] = {**e, "fecha": f, "origen_fecha": origen, "no_reciente": f < limite,
                                 "src": foto_b64(ruta, self.cfg["ancho_max_foto_px"], self.cfg["calidad_jpeg"]),
                                 "thumb": foto_b64(ruta, 160, 60)}
        return self._cache[ruta]

    def para_concepto(self, dist, indicador, concepto):
        cand = [e for e in self.datos.get("fotos", [])
                if e["distribuidor"] == dist and e["indicador"] == indicador and norm(e["concepto"]) == norm(concepto)]
        cand = [x for x in (self._preparar(e) for e in cand) if x]
        if not cand:
            return None
        return max(cand, key=lambda x: x["fecha"])  # la más reciente

    def observaciones(self, dist):
        out = []
        for o in self.datos.get("observaciones", []):
            if o["distribuidor"] != dist:
                continue
            o = dict(o)
            if o.get("archivo"):
                o["foto"] = self._preparar(o)
            out.append(o)
        return out


# --------------------------------------------------------------------------- cálculo
def cargar_ajustes(ruta):
    """Ajustes manuales (datos/ajustes.json): confirmaciones de campo que aún no están en el Excel."""
    if not ruta or not os.path.exists(ruta):
        return []
    return json.load(open(ruta, encoding="utf-8")).get("ajustes", [])


def aplicar_ajustes(x, ajustes):
    """Sustituye el 'Adquirido' de los conceptos indicados y recalcula el % del indicador."""
    for a in ajustes:
        c = next((c for c in x["conceptos"] if norm(c["concepto"]) == norm(a["concepto"])), None)
        if c is None:
            print(f"  aviso: ajuste sin concepto que coincida: {a['distribuidor']} / {a['indicador']} / {a['concepto']}",
                  file=sys.stderr)
            continue
        if a.get("adquirido") is not None:
            c["adq"] = float(a["adquirido"])
            c["avance"] = min(c["adq"] / c["req"], 1.0) if c["req"] > 0 else (1.0 if c["adq"] else 0.0)
        if a.get("falta"):
            c["falta"] = a["falta"]
        c["ajuste"] = a.get("nota", "Ajuste manual")
        x.setdefault("ajustes", []).append(
            f'{a.get("nota", "Ajuste manual")}' + (f' ({fecha_corta(dt.date.fromisoformat(a["fecha"]))})' if a.get("fecha") else ""))
    if ajustes:
        x["valor"] = 100.0 * sum(c["avance"] for c in x["conceptos"]) / len(x["conceptos"])


def evaluar(cfg, grupo, kpi, per, fotos, corte, ajustes=()):
    reglas = cfg["semaforo"]
    resultado = []
    for d in grupo["distribuidores"]:
        inds, pendientes = [], []
        for ind in cfg["indicadores"]:
            regla = reglas[ind["semaforo"]]
            r = {"id": ind["id"], "nombre": ind["nombre"], "regla": ind["semaforo"], "conceptos": [],
                 "valor": None, "estado": "sin", "fuente": "", "mes": None}
            if ind["id"] in d.get("no_aplica", []):
                r.update(estado="na", nota=d.get("motivo_no_aplica", "No aplica"))
                inds.append(r)
                continue
            if ind["fuente"] == "kpi":
                x = kpi.indicador(ind["hoja"], d["nombre_kpi"], corte, ind.get("concepto"))
                if x:
                    aplicar_ajustes(x, [a for a in ajustes if a["distribuidor"] == d["id"] and a["indicador"] == ind["id"]])
                    r.update(valor=x["valor"], mes=x["mes"], conceptos=x["conceptos"], fuente=x["ubicacion"],
                             ajustes=x.get("ajustes", []))
                    for c in x["conceptos"]:
                        c["foto"] = fotos.para_concepto(d["id"], ind["id"], c["concepto"])
                else:
                    r["nota"] = f"Sin información: «{d['nombre_kpi']}» no aparece en la hoja «{ind['hoja']}»."
            else:  # personal
                res = per.resumen(d["nombre_personal"])
                if res:
                    r.update(valor=res["pct"], personal=res, fuente=res["ubicacion"],
                             areas=per.por_area(d["nombre_personal"]), puestos=per.puestos(d["nombre_personal"]))
                else:
                    r["nota"] = f"Sin información: «{d['nombre_personal']}» no aparece en el archivo de personal."
            r["estado"] = semaforo(r["valor"], regla)
            inds.append(r)

            # ---- pendientes: todo lo que no esté al 100 %
            if r["valor"] is None:
                continue
            if ind["id"] == "personal":
                for p in r["puestos"]:
                    if p["faltan"] > 0:
                        pendientes.append({
                            "indicador": ind["nombre"], "concepto": p["puesto"], "area": p["area"],
                            "falta": f"Contratar {p['faltan']:g} plaza(s): requeridas {p['req']:g}, cubiertas {p['cub']:g}"})
            elif ind["con_conceptos"] or len(r["conceptos"]) > 1:
                for c in r["conceptos"]:
                    if c["avance"] < 1:
                        if c.get("falta"):
                            falta = c["falta"]
                        elif c["adq"] <= 0:
                            falta = "No cumple: implementar / obtener"
                        elif c["req"] > 1:
                            falta = f"Tiene {c['adq']:g} de {c['req']:g} requeridos ({pct(c['avance'] * 100, 0)})"
                        else:
                            falta = f"Avance {pct(c['avance'] * 100, 0)}: completar el {pct((1 - c['avance']) * 100, 0)} restante"
                        pendientes.append({"indicador": ind["nombre"], "concepto": c["concepto"],
                                           "area": c["subtema"], "falta": falta})
            elif r["valor"] < 100:
                c = r["conceptos"][0]
                pendientes.append({"indicador": ind["nombre"], "concepto": c["concepto"], "area": "",
                                   "falta": "No cumple: requerido y no acreditado"})

        validos = [i["valor"] for i in inds if i["estado"] not in ("na",) and i["valor"] is not None]
        sin_info = [i["nombre"] for i in inds if i["estado"] == "sin"]
        general = sum(validos) / len(validos) if validos else None
        for n, p in enumerate(pendientes, 1):
            p["id"] = f"{d['id']}-{n}"
        resultado.append({**d, "indicadores": inds, "pendientes": pendientes, "general": general,
                          "estado_general": semaforo(general, reglas["general"]),
                          "n_validos": len(validos), "sin_info": sin_info,
                          "observaciones": fotos.observaciones(d["id"])})
    return resultado


# --------------------------------------------------------------------------- HTML
ICONOS = {
    "imagen": '<path d="M3 21V8l9-5 9 5v13"/><path d="M9 21v-6h6v6"/>',
    "showroom": '<rect x="3" y="4" width="18" height="12" rx="1"/><path d="M8 20h8M12 16v4"/>',
    "almacen": '<path d="M3 8l9-5 9 5v12H3z"/><path d="M7 20v-7h10v7M7 16h10"/>',
    "taller": '<path d="M14.7 6.3a4 4 0 0 0-5.4 5.4L3 18v3h3l6.3-6.3a4 4 0 0 0 5.4-5.4l-2.6 2.6-2.4-.6-.6-2.4z"/>',
    "admin": '<rect x="3" y="7" width="18" height="13" rx="1"/><path d="M8 7V4h8v3"/>',
    "mobiliario": '<path d="M5 11V6a2 2 0 0 1 2-2h10a2 2 0 0 1 2 2v5"/><path d="M3 11h18v5H3zM5 16v4M19 16v4"/>',
    "generales": '<path d="M9 5h11M9 12h11M9 19h11"/><circle cx="4.5" cy="5" r="1"/><circle cx="4.5" cy="12" r="1"/><circle cx="4.5" cy="19" r="1"/>',
    "persona": '<circle cx="12" cy="8" r="4"/><path d="M4 21c0-4 4-6 8-6s8 2 8 6"/>',
    "doc": '<path d="M6 3h9l4 4v14H6z"/><path d="M14 3v5h5M9 13h7M9 17h7"/>',
    "motor": '<rect x="5" y="8" width="12" height="9" rx="1"/><path d="M17 11h3v3h-3M8 8V5h6v3M2 12h3"/>',
    "pc": '<rect x="3" y="4" width="18" height="12" rx="1"/><path d="M8 20h8M12 16v4"/>',
    "caja": '<path d="M3 7l9-4 9 4v10l-9 4-9-4z"/><path d="M3 7l9 4 9-4M12 11v10"/>',
    "diag": '<rect x="4" y="3" width="16" height="18" rx="2"/><path d="M8 13l2-3 2 4 2-2h2"/>',
}


def icono(nombre, cls="ico"):
    return (f'<svg class="{cls}" viewBox="0 0 24 24" aria-hidden="true" fill="none" stroke="currentColor" '
            f'stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round">{ICONOS.get(nombre, ICONOS["generales"])}</svg>')


def icono_subtema(sub):
    s = norm(sub)
    for clave, ico in (("imag", "imagen"), ("show", "showroom"), ("almac", "almacen"), ("taller", "taller"),
                       ("admin", "admin"), ("mobil", "mobiliario")):
        if clave in s:
            return ico
    return "generales"


ICONO_IND = {"anexos": "doc", "proyecto": "imagen", "software": "pc", "stock": "caja", "kit": "caja",
             "crm": "pc", "fotonet": "pc", "rifof": "doc", "planpiso": "doc", "pld": "doc",
             "diagnostico": "diag", "cummins": "motor", "personal": "persona"}


def pill(estado, texto=None):
    return (f'<span class="pill {estado}"><span class="pi" aria-hidden="true">{ESTADO_ICO[estado]}</span>'
            f'{esc(texto or ESTADO_TXT[estado])}</span>')


def meter(valor, estado):
    w = 0 if valor is None else max(0, min(100, valor))
    return f'<div class="meter {estado}" role="img" aria-label="{pct(valor)}"><span style="width:{w:.1f}%"></span></div>'


def figura(f, estado_concepto, concepto):
    if not f:
        return ""
    etiqueta = "Evidencia: cumple" if estado_concepto == "verde" else "Ejemplo de lo pendiente"
    badge = '<span class="badge-old">Foto no reciente</span>' if f["no_reciente"] else ""
    return (f'<figure class="foto {"ok" if estado_concepto == "verde" else "nok"}">'
            f'<img src="{f["src"]}" alt="{esc(concepto)}: {esc(f["descripcion"])}" loading="lazy">'
            f'<figcaption><b>{esc(etiqueta)}</b> · {esc(concepto)}<br>{esc(f["descripcion"])}<br>'
            f'<span class="fdate">Foto: {fecha_corta(f["fecha"])}</span> {badge}'
            f'<span class="fsrc">Drive: {esc(f["drive_carpeta"])} / {esc(os.path.basename(f["archivo"]))}</span>'
            f'</figcaption></figure>')


def estado_concepto(c):
    return "verde" if c["avance"] >= 1 else ("amarillo" if c["avance"] > 0 else "rojo")


def render_tarjeta(i, corte):
    if i["estado"] == "na":
        return (f'<article class="card na">{icono(ICONO_IND[i["id"]], "ico big")}<h4>{esc(i["nombre"])}</h4>'
                f'<div class="val">N/A</div>{pill("na")}<p class="note">{esc(i.get("nota", ""))}</p></article>')
    if i["valor"] is None:
        return (f'<article class="card sin">{icono(ICONO_IND[i["id"]], "ico big")}<h4>{esc(i["nombre"])}</h4>'
                f'<div class="val">—</div>{pill("sin")}<p class="note">{esc(i.get("nota", ""))}</p></article>')
    brecha = max(0.0, 100 - i["valor"])
    if i["id"] == "personal":
        p = i["personal"]
        detalle = f'{p["cub"]} de {p["req"]} plazas cubiertas · {p["faltan"]} vacantes'
    elif len(i["conceptos"]) > 1:
        ok = sum(1 for c in i["conceptos"] if c["avance"] >= 1)
        detalle = f'{ok} de {len(i["conceptos"])} conceptos completos'
    else:
        detalle = "Cumple" if i["valor"] >= 100 else "No acreditado"
    viejo = ""
    if i.get("mes") and (corte - i["mes"]).days > 62:
        viejo = f' <span class="stale" title="Último dato disponible en el Excel">dato a {mes_txt(i["mes"])}</span>'
    elif i.get("mes"):
        viejo = f' <span class="muted">dato a {mes_txt(i["mes"])}</span>'
    ajuste = "".join(f'<p class="ajuste">Ajuste: {esc(a)}</p>' for a in i.get("ajustes", []))
    return (f'<article class="card {i["estado"]}">{icono(ICONO_IND[i["id"]], "ico big")}<h4>{esc(i["nombre"])}</h4>'
            f'<div class="val">{pct(i["valor"])}</div>{meter(i["valor"], i["estado"])}'
            f'<div class="cardrow">{pill(i["estado"])}<span class="gap {"cero" if brecha == 0 else ""}">'
            f'Brecha {pct(brecha)}</span></div><p class="note">{esc(detalle)}{viejo}</p>{ajuste}</article>')


def render_conceptos(i, titulo_extra=""):
    """Desglose: primero lo que NO cumple (con foto o ícono), luego lo cumplido compacto."""
    if i["estado"] in ("na", "sin") or not i["conceptos"]:
        return ""
    nok = [c for c in i["conceptos"] if c["avance"] < 1]
    ok = [c for c in i["conceptos"] if c["avance"] >= 1]
    filas = []
    for c in nok:
        e = estado_concepto(c)
        visual = (f'<img class="thumb" src="{c["foto"]["thumb"]}" alt="">' if c["foto"]
                  else icono(icono_subtema(c["subtema"]), "ico mid"))
        avance = f'{c["adq"]:g} / {c["req"]:g}' if c["req"] != 1 else pct(c["avance"] * 100, 0)
        filas.append(f'<tr class="{e}"><td class="vis">{visual}</td><td><b>{esc(c["concepto"])}</b>'
                     f'<div class="muted">{esc(c["subtema"])}</div>'
                     + (f'<div class="ajuste">Ajuste: {esc(c["ajuste"])}</div>' if c.get("ajuste") else "")
                     + f'</td><td class="num">{avance}</td>'
                     f'<td>{pill(e, "Parcial" if e == "amarillo" else None)}</td></tr>')
    # cumplidos agrupados por subtema
    grupos = {}
    for c in ok:
        grupos.setdefault(c["subtema"], []).append(c)
    chips = "".join(
        f'<div class="okgrp"><span class="okhead">{icono(icono_subtema(s), "ico sm")} {esc(s)}</span>'
        + "".join(f'<span class="chip">✓ {esc(c["concepto"])}</span>' for c in cs) + "</div>"
        for s, cs in grupos.items())
    # resumen por subtema
    subt = {}
    for c in i["conceptos"]:
        s = subt.setdefault(c["subtema"], [0.0, 0])
        s[0] += c["avance"]; s[1] += 1
    resumen_sub = "".join(
        f'<div class="subbar"><span>{icono(icono_subtema(s), "ico sm")} {esc(s)}</span>'
        f'{meter(v / n * 100, semaforo(v / n * 100, {"verde_min": 100, "amarillo_min": 50}))}'
        f'<span class="num">{pct(v / n * 100, 0)}</span></div>' for s, (v, n) in subt.items()) if len(subt) > 1 else ""
    con_foto = sorted((c for c in i["conceptos"] if c["foto"]), key=lambda c: c["avance"] >= 1)  # pendientes primero
    fotos_html = "".join(figura(c["foto"], estado_concepto(c), c["concepto"]) for c in con_foto)
    return f"""
    <section class="block">
      <h3>{icono(ICONO_IND[i["id"]], "ico")} {esc(i["nombre"])} — desglose {titulo_extra}
        <span class="h3val">{pct(i["valor"])} {pill(i["estado"])}</span></h3>
      <p class="src">Fuente: {esc(i["fuente"])}. Cada concepto vale 1; los parciales cuentan su avance.</p>
      {f'<div class="subbars">{resumen_sub}</div>' if resumen_sub else ''}
      {f'<h4 class="nokh">No cumple / pendiente ({len(nok)})</h4><table class="tbl nokt"><thead><tr><th></th><th>Concepto</th><th class="num">Avance</th><th>Estatus</th></tr></thead><tbody>{"".join(filas)}</tbody></table>' if nok else '<p class="allok">Todos los conceptos cumplen.</p>'}
      {f'<details class="okdet"{" open" if len(ok) <= 12 else ""}><summary>Cumplidos ({len(ok)})</summary>{chips}</details>' if ok else ''}
      {f'<div class="fotos">{fotos_html}</div>' if fotos_html else ''}
    </section>"""


def render_personal(i):
    if i["estado"] in ("na", "sin"):
        return ""
    p = i["personal"]
    areas = i.get("areas") or {}
    filas_area = "".join(
        f'<tr><td>{esc(a["area"])}</td><td class="num">{a["req"]}</td><td class="num">{a["cub"]}</td>'
        f'<td class="num">{(a["req"] or 0) - (a["cub"] or 0)}</td><td class="num">{pct(a["pct"]) if a["pct"] is not None else "N/A"}</td></tr>'
        for a in areas.get("areas", []))
    vac = [x for x in i["puestos"] if x["faltan"] > 0]
    filas_vac = "".join(
        f'<tr class="rojo"><td class="vis">{icono("persona", "ico mid")}</td><td><b>{esc(x["puesto"])}</b>'
        f'<div class="muted">{esc(x["area"])}</div></td><td class="num">{x["req"]:g}</td><td class="num">{x["cub"]:g}</td>'
        f'<td class="num big">{x["faltan"]:g}</td></tr>' for x in vac)
    return f"""
    <section class="block">
      <h3>{icono("persona", "ico")} Personal <span class="h3val">{pct(i["valor"])} {pill(i["estado"])}</span></h3>
      <p class="src">Fuente: {esc(i["fuente"])}{"; " + esc(areas["ubicacion"]) if areas else ""}; detalle por puesto en hoja «Cálculo_Red». Valores tomados literal del archivo de personal.</p>
      <div class="kpirow">
        <div class="kpi"><span>Plazas requeridas</span><b>{p["req"]}</b></div>
        <div class="kpi"><span>Plazas cubiertas</span><b>{p["cub"]}</b></div>
        <div class="kpi bad"><span>Vacantes</span><b>{p["faltan"]}</b></div>
        <div class="kpi"><span>Personal extra</span><b>{p["extra"]}</b></div>
        <div class="kpi"><span>% Cumplimiento</span><b>{pct(p["pct"])}</b></div>
      </div>
      {f'<table class="tbl"><thead><tr><th>Área</th><th class="num">Requeridas</th><th class="num">Cubiertas</th><th class="num">Faltan</th><th class="num">%</th></tr></thead><tbody>{filas_area}</tbody></table>' if filas_area else ''}
      <h4 class="nokh">Vacantes por puesto ({len(vac)} puestos · {sum(x["faltan"] for x in vac):g} plazas)</h4>
      <table class="tbl nokt"><thead><tr><th></th><th>Puesto</th><th class="num">Req.</th><th class="num">Cub.</th><th class="num">Faltan</th></tr></thead>
      <tbody>{filas_vac}</tbody></table>
    </section>"""


def render_simple(i):
    """Indicadores Sí/No sin desglose (Anexos, Stock, Diagnóstico)."""
    if i["estado"] in ("na", "sin") or not i["conceptos"]:
        return ""
    c = i["conceptos"][0]
    e = estado_concepto(c)
    return (f'<tr class="{e}"><td class="vis">{icono(ICONO_IND[i["id"]], "ico mid")}</td><td><b>{esc(i["nombre"])}</b>'
            f'<div class="muted">{esc(c["concepto"])} · {esc(i["fuente"])} · dato a {mes_txt(i["mes"])}</div>'
            + (f'<div class="ajuste">Ajuste: {esc(c["ajuste"])}</div>' if c.get("ajuste") else "") + '</td>'
            f'<td>{pill(e)}</td></tr>')


def render_plan(d):
    filas = []
    for n, p in enumerate(d["pendientes"], 1):
        filas.append(
            f'<tr data-id="{p["id"]}"><td class="num">{n}</td><td>{esc(p["indicador"])}</td>'
            f'<td><b>{esc(p["concepto"])}</b>' + (f'<div class="muted">{esc(p["area"])}</div>' if p["area"] else "") + '</td>'
            f'<td>{esc(p["falta"])}</td><td><span class="prio">ALTA</span></td>'
            f'<td><input class="ed" data-f="resp" aria-label="Responsable"></td>'
            f'<td><input class="ed" type="date" data-f="fecha" aria-label="Fecha compromiso"></td>'
            f'<td><select class="ed" data-f="estatus" aria-label="Estatus"><option></option><option>Pendiente</option>'
            f'<option>En proceso</option><option>Cumplido</option></select></td>'
            f'<td><textarea class="ed" data-f="coment" rows="1" aria-label="Comentarios"></textarea></td></tr>')
    return f"""
    <section class="block plan" id="plan-{d["id"]}">
      <h3>Plan de trabajo — {esc(d["nombre"])} <span class="h3val">{len(d["pendientes"])} pendientes · prioridad ALTA</span></h3>
      <p class="src">Todo lo que no está al 100% es pendiente (aunque su semáforo esté en verde). Las columnas Responsable, Fecha, Estatus y Comentarios se capturan en pantalla y se guardan en este navegador.</p>
      <div class="tblwrap"><table class="tbl plantbl" data-dist="{d["id"]}" data-nombre="{esc(d["nombre"])}">
        <thead><tr><th class="num">#</th><th>Indicador</th><th>Concepto / puesto</th><th>Qué falta</th><th>Prioridad</th>
        <th>Responsable</th><th>Fecha compromiso</th><th>Estatus</th><th>Comentarios</th></tr></thead>
        <tbody>{"".join(filas)}</tbody></table></div>
    </section>"""


def render_observaciones(d):
    if not d["observaciones"]:
        return ""
    items = []
    for o in d["observaciones"]:
        f = o.get("foto")
        foto = ""
        if f:
            badge = '<span class="badge-old">Foto no reciente</span>' if f["no_reciente"] else ""
            foto = (f'<figure class="foto obs"><img src="{f["src"]}" alt="{esc(f["descripcion"])}" loading="lazy">'
                    f'<figcaption>{esc(f["descripcion"])}<br><span class="fdate">Foto: {fecha_corta(f["fecha"])}</span> {badge}'
                    f'<span class="fsrc">Drive: {esc(f["drive_carpeta"])} / {esc(os.path.basename(f["archivo"]))}</span></figcaption></figure>')
        items.append(f'<li><div>{esc(o["texto"])}</div>{foto}</li>')
    return (f'<section class="block avoid"><h3>{icono("generales", "ico")} Para verificar en la visita</h3>'
            f'<p class="src">Diferencias entre las evidencias de Drive y el Excel. No se modificó ningún dato: el reporte usa el valor del Excel.</p>'
            f'<ul class="obs">{"".join(items)}</ul></section>')


def render_distribuidor(d, grupo, corte):
    tarjetas = "".join(render_tarjeta(i, corte) for i in d["indicadores"])
    by = {i["id"]: i for i in d["indicadores"]}
    simples = "".join(render_simple(by[k]) for k in ("anexos", "crm", "fotonet", "rifof", "planpiso", "kit", "pld",
                                                     "diagnostico", "stock") if k in by)
    nota_gen = (f'Promedio simple de {d["n_validos"]} indicadores'
                + (f' (sin información: {", ".join(d["sin_info"])})' if d["sin_info"] else "")
                + (f'; no aplica: {", ".join(by[k]["nombre"] for k in d.get("no_aplica", []))} — {d.get("motivo_no_aplica", "")}'
                   if d.get("no_aplica") else ""))
    return f"""
  <section class="page dist" id="d-{d["id"]}">
    <header class="phead">
      <div><div class="eyebrow">{esc(grupo["nombre"])}</div><h2>{esc(d["nombre"])}</h2>
        <div class="corte">Fecha de corte: {fecha_larga(corte)}</div></div>
      <div class="hero {d["estado_general"]}"><span>% General</span><b>{pct(d["general"])}</b>{pill(d["estado_general"])}
        <small>{len(d["pendientes"])} pendientes</small></div>
      <div class="noprint actions"><button onclick="imprimir('{d["id"]}')">Imprimir / PDF</button>
        <button onclick="exportar(['{d["id"]}'])">Exportar plan a Excel</button></div>
    </header>
    <p class="src">{esc(nota_gen)}.</p>
    <div class="cards">{tarjetas}</div>
    {f'<section class="block avoid"><h3>Indicadores Sí / No</h3><table class="tbl nokt">{simples}</table></section>' if simples else ''}
    {render_conceptos(by["proyecto"]) if "proyecto" in by else ""}
    {render_conceptos(by["cummins"]) if "cummins" in by else ""}
    {render_conceptos(by["software"]) if "software" in by else ""}
    {render_personal(by["personal"]) if "personal" in by else ""}
    {render_observaciones(d)}
    {render_plan(d)}
    <footer class="pfoot">{esc(grupo["nombre"])} · {esc(d["nombre"])} · Corte {fecha_corta(corte)}</footer>
  </section>"""


def render_portada(grupo, dists, cfg, corte):
    cols = "".join(f'<th><a href="#d-{d["id"]}">{esc(d["nombre"])}</a></th>' for d in dists)
    filas = []
    for k, ind in enumerate(cfg["indicadores"]):
        celdas = []
        for d in dists:
            i = d["indicadores"][k]
            txt = "N/A" if i["estado"] == "na" else ("Sin info" if i["valor"] is None else pct(i["valor"]))
            celdas.append(f'<td class="cell {i["estado"]}"><span class="pi">{ESTADO_ICO[i["estado"]]}</span>{txt}</td>')
        filas.append(f'<tr><th scope="row">{icono(ICONO_IND[ind["id"]], "ico sm")} {esc(ind["nombre"])}</th>{"".join(celdas)}</tr>')
    gen = "".join(f'<td class="cell gen {d["estado_general"]}"><span class="pi">{ESTADO_ICO[d["estado_general"]]}</span>{pct(d["general"])}</td>' for d in dists)
    pend = "".join(f'<td class="num"><b>{len(d["pendientes"])}</b></td>' for d in dists)
    # Total del grupo: promedio de todos los indicadores que aplican en todas las agencias
    vals = [i["valor"] for d in dists for i in d["indicadores"] if i["estado"] != "na" and i["valor"] is not None]
    grupo_gen = sum(vals) / len(vals) if vals else None
    proy = [c["avance"] for d in dists for i in d["indicadores"] if i["id"] == "proyecto" for c in i["conceptos"]]
    grupo_proy = 100 * sum(proy) / len(proy) if proy else None
    eg, ep = semaforo(grupo_gen, cfg["semaforo"]["general"]), semaforo(grupo_proy, cfg["semaforo"]["proyecto"])
    total = f"""
      <div class="dcard {eg}">
        <h3>Total del grupo</h3>
        <div class="big">{pct(grupo_gen)}</div>{meter(grupo_gen, eg)}
        <div class="cardrow">{pill(eg)}<span><b>{sum(len(d["pendientes"]) for d in dists)}</b> pendientes</span></div>
        <p class="note">{len(vals)} indicadores de las {len(dists)} agencias (sin los que no aplican). Desarrollo de Proyecto del grupo: <b>{pct(grupo_proy)}</b> ({len(proy)} elementos) {pill(ep)}</p>
      </div>"""
    tarjetas = total + "".join(f"""
      <a class="dcard {d["estado_general"]}" href="#d-{d["id"]}">
        <h3>{esc(d["nombre"])}</h3>
        <div class="big">{pct(d["general"])}</div>{meter(d["general"], d["estado_general"])}
        <div class="cardrow">{pill(d["estado_general"])}<span><b>{len(d["pendientes"])}</b> pendientes</span></div>
        {f'<p class="note">No aplica: {esc(", ".join(i["nombre"] for i in d["indicadores"] if i["estado"] == "na"))}</p>' if any(i["estado"] == "na" for i in d["indicadores"]) else ''}
      </a>""" for d in dists)
    s = cfg["semaforo"]
    def regla_txt(r):
        return (f'Verde ≥ {r["verde_min"]}%' + (f' · Amarillo {r["amarillo_min"]}–{r["verde_min"]}%' if r["amarillo_min"] is not None else '')
                + f' · Rojo < {r["amarillo_min"] if r["amarillo_min"] is not None else r["verde_min"]}%')
    return f"""
  <section class="page portada" id="resumen">
    <header class="phead">
      <div><div class="eyebrow">Estatus de distribuidores FOTON</div><h1>{esc(grupo["nombre"])}</h1>
        <div class="corte">Fecha de corte: {fecha_larga(corte)} · {len(dists)} agencias</div></div>
      <div class="noprint actions"><button onclick="imprimir()">Imprimir / PDF (todo)</button>
        <button onclick="exportar()">Exportar planes a Excel</button></div>
    </header>
    <div class="dcards">{tarjetas}</div>
    <section class="block">
      <h3>Semáforo por indicador</h3>
      <div class="tblwrap"><table class="tbl matrix"><thead><tr><th>Indicador</th>{cols}</tr></thead>
        <tbody>{"".join(filas)}</tbody>
        <tfoot><tr><th scope="row">% General</th>{gen}</tr><tr><th scope="row">Pendientes</th>{pend}</tr></tfoot></table></div>
      <p class="src">Todos los indicadores pesan lo mismo. % General = promedio simple de los indicadores que aplican.
        Reglas: Sí/No {regla_txt(s["si_no"])}; Proyecto {regla_txt(s["proyecto"])}; Cummins {regla_txt(s["cummins"])};
        Personal {regla_txt(s["personal"])}; General {regla_txt(s["general"])}.</p>
    </section>
    <footer class="pfoot">{esc(grupo["nombre"])} · Resumen · Corte {fecha_corta(corte)}</footer>
  </section>"""


CSS = r"""
:root{--ink:#14171c;--ink2:#4a5160;--muted:#717887;--line:#e3e5ea;--bg:#f5f6f8;--card:#fff;--navy:#1d2a3a;
--good:#0ca30c;--warn:#fab219;--crit:#d03b3b;--na:#9aa0aa;--goodbg:#e7f6e7;--warnbg:#fff4dc;--critbg:#fbe9e9;}
*{box-sizing:border-box}html{scroll-behavior:smooth}
body{margin:0;background:var(--bg);color:var(--ink);font:16px/1.45 system-ui,-apple-system,"Segoe UI",Roboto,Arial,sans-serif}
.topbar{position:sticky;top:0;z-index:10;background:var(--navy);color:#fff;display:flex;gap:16px;align-items:center;padding:10px 24px;flex-wrap:wrap}
.topbar b{font-size:18px}.topbar nav{display:flex;gap:4px;flex-wrap:wrap}.topbar a{color:#d8e1ec;text-decoration:none;padding:6px 12px;border-radius:6px}
.topbar a:hover{background:rgba(255,255,255,.12);color:#fff}.topbar .tcorte{margin-left:auto;font-size:14px;color:#c9d3df}
main{max-width:1280px;margin:0 auto;padding:16px 24px 48px}
.page{background:var(--card);border-radius:12px;box-shadow:0 1px 3px rgba(0,0,0,.06);padding:28px 32px;margin:20px 0}
.phead{display:flex;gap:24px;align-items:flex-start;justify-content:space-between;flex-wrap:wrap;border-bottom:2px solid var(--line);padding-bottom:16px;margin-bottom:12px}
.eyebrow{text-transform:uppercase;letter-spacing:.08em;font-size:13px;color:var(--muted);font-weight:600}
h1{font-size:40px;margin:2px 0}h2{font-size:34px;margin:2px 0}h3{font-size:21px;margin:0 0 6px;display:flex;align-items:center;gap:8px;flex-wrap:wrap}
h4{margin:0}.corte{color:var(--ink2);font-weight:600}
.actions{display:flex;gap:8px;flex-wrap:wrap}
button{font:inherit;font-size:15px;border:1px solid var(--navy);background:#fff;color:var(--navy);border-radius:8px;padding:8px 14px;cursor:pointer}
button:hover{background:var(--navy);color:#fff}
.hero{display:grid;grid-template-columns:auto;justify-items:center;gap:4px;padding:10px 22px;border-radius:12px;border:2px solid var(--line);min-width:180px}
.hero>span:first-child{font-size:13px;color:var(--muted);text-transform:uppercase;letter-spacing:.06em}.hero b{font-size:44px;line-height:1}
.hero small{color:var(--ink2)}
.hero.rojo{border-color:var(--crit)}.hero.amarillo{border-color:var(--warn)}.hero.verde{border-color:var(--good)}
.pill{display:inline-flex;align-items:center;gap:6px;border-radius:999px;padding:3px 12px 3px 4px;font-size:14px;font-weight:600;white-space:nowrap;background:#eef0f3;color:var(--ink)}
.pill .pi{display:inline-grid;place-items:center;width:20px;height:20px;border-radius:50%;color:#fff;font-size:12px;background:var(--na)}
.pill.verde{background:var(--goodbg)}.pill.verde .pi{background:var(--good)}
.pill.amarillo{background:var(--warnbg)}.pill.amarillo .pi{background:var(--warn);color:#0b0b0b}
.pill.rojo{background:var(--critbg)}.pill.rojo .pi{background:var(--crit)}
.meter{height:8px;border-radius:99px;background:#e9ebef;overflow:hidden;margin:6px 0}.meter span{display:block;height:100%;border-radius:99px;background:var(--na)}
.meter.verde span{background:var(--good)}.meter.amarillo span{background:var(--warn)}.meter.rojo span{background:var(--crit)}
.cards{display:grid;grid-template-columns:repeat(auto-fill,minmax(220px,1fr));gap:14px;margin:12px 0 8px}
.card{border:1px solid var(--line);border-radius:12px;padding:14px 16px;position:relative;background:#fff;border-top:5px solid var(--na)}
.card.verde{border-top-color:var(--good)}.card.amarillo{border-top-color:var(--warn)}.card.rojo{border-top-color:var(--crit);background:#fffafa}
.card.na,.card.sin{background:#fafafa;color:var(--ink2)}
.card h4{font-size:15px;color:var(--ink2);padding-right:34px;min-height:40px}.card .val{font-size:36px;font-weight:700;line-height:1.1}
.card .ico.big{position:absolute;right:12px;top:12px;width:26px;height:26px;color:var(--muted)}
.cardrow{display:flex;justify-content:space-between;align-items:center;gap:8px;flex-wrap:wrap}
.gap{font-weight:700;color:var(--crit)}.gap.cero{color:var(--muted);font-weight:500}
.note{font-size:13px;color:var(--ink2);margin:8px 0 0}.muted{color:var(--muted);font-size:13px}
.ajuste{font-size:12.5px;color:#1f4e86;background:#e8f0fa;border-radius:4px;padding:2px 6px;margin:4px 0 0;display:inline-block}
.stale{background:#eef0f3;border-radius:4px;padding:1px 6px;font-size:12px;color:var(--ink2);font-weight:600}
.block{margin:26px 0 0;padding-top:14px;border-top:1px solid var(--line)}
.h3val{margin-left:auto;display:inline-flex;gap:8px;align-items:center;font-size:18px}
.src{font-size:13px;color:var(--muted);margin:2px 0 10px}
.tbl{width:100%;border-collapse:collapse;font-size:15px}.tbl th,.tbl td{padding:8px 10px;border-bottom:1px solid var(--line);text-align:left;vertical-align:middle}
.tbl thead th{font-size:13px;text-transform:uppercase;letter-spacing:.04em;color:var(--muted);border-bottom:2px solid var(--line)}
.num{text-align:right!important;font-variant-numeric:tabular-nums}.big{font-size:20px;font-weight:700}
.nokh{color:var(--crit);margin:14px 0 6px;font-size:17px}
.nokt tr.rojo td:first-child{box-shadow:inset 4px 0 var(--crit)}.nokt tr.amarillo td:first-child{box-shadow:inset 4px 0 var(--warn)}.nokt tr.verde td:first-child{box-shadow:inset 4px 0 var(--good)}
.vis{width:76px}.thumb{width:68px;height:50px;object-fit:cover;border-radius:6px;display:block}
.ico{width:20px;height:20px;flex:none;vertical-align:middle}.ico.mid{width:28px;height:28px;color:var(--ink2);margin-left:14px}.ico.sm{width:16px;height:16px;color:var(--muted)}
.subbars{display:grid;grid-template-columns:repeat(auto-fill,minmax(260px,1fr));gap:4px 20px;margin:6px 0 10px}
.subbar{display:grid;grid-template-columns:1fr 110px 44px;gap:10px;align-items:center;font-size:14px}.subbar span:first-child{display:flex;gap:6px;align-items:center}
.okdet{margin:10px 0;font-size:14px}.okdet summary{cursor:pointer;color:var(--ink2);font-weight:600;padding:4px 0}
.okgrp{margin:6px 0}.okhead{display:inline-flex;gap:4px;align-items:center;font-weight:600;color:var(--ink2);margin-right:6px}
.chip{display:inline-block;border:1px solid var(--line);border-radius:999px;padding:1px 10px;margin:2px;color:var(--ink2);font-size:13px}
.allok{color:var(--ink2)}
.fotos{display:grid;grid-template-columns:repeat(auto-fill,minmax(260px,1fr));gap:14px;margin-top:14px}
.foto{margin:0;border:1px solid var(--line);border-radius:10px;overflow:hidden;background:#fff}
.foto img{width:100%;height:190px;object-fit:cover;display:block}
.foto.nok{border:2px solid var(--crit)}.foto.ok{border:2px solid var(--good)}
.foto figcaption{padding:8px 10px;font-size:13px;color:var(--ink2)}
.fdate{font-weight:700;color:var(--ink)}.fsrc{display:block;color:var(--muted);font-size:11px;margin-top:3px;word-break:break-word}
.badge-old{display:inline-block;background:#3a3f48;color:#fff;border-radius:4px;padding:0 6px;font-size:12px;font-weight:700;margin-left:4px}
.kpirow{display:flex;gap:12px;flex-wrap:wrap;margin:8px 0 14px}.kpi{border:1px solid var(--line);border-radius:10px;padding:8px 14px;min-width:130px}
.kpi span{display:block;font-size:13px;color:var(--muted)}.kpi b{font-size:26px}.kpi.bad b{color:var(--crit)}
.obs{padding-left:18px}.obs li{margin:8px 0 14px}.foto.obs{max-width:340px;margin-top:6px}
.prio{background:var(--crit);color:#fff;font-weight:700;font-size:12px;border-radius:4px;padding:2px 8px}
.tblwrap{overflow-x:auto}.plantbl{font-size:14px;min-width:1100px}.plantbl td{vertical-align:top}
.ed{font:inherit;font-size:14px;width:100%;min-width:110px;border:1px solid #cfd3da;border-radius:6px;padding:5px 7px;background:#fffef6}
textarea.ed{min-width:180px;resize:vertical}
.dcards{display:grid;grid-template-columns:repeat(auto-fit,minmax(260px,1fr));gap:16px;margin:16px 0}
.dcard{display:block;text-decoration:none;color:inherit;border:1px solid var(--line);border-radius:14px;padding:18px 20px;border-top:6px solid var(--na)}
.dcard.verde{border-top-color:var(--good)}.dcard.amarillo{border-top-color:var(--warn)}.dcard.rojo{border-top-color:var(--crit)}
.dcard h3{font-size:22px}.dcard .big{font-size:52px;font-weight:700;line-height:1.1}
.matrix th[scope=row]{font-weight:600;white-space:nowrap}.matrix{min-width:560px}.matrix thead th{text-align:center;font-size:15px}
.matrix thead th a{color:var(--ink);text-decoration:none}
.cell{text-align:center!important;font-weight:700;font-size:17px}.cell .pi{display:inline-grid;place-items:center;width:22px;height:22px;border-radius:50%;color:#fff;font-size:12px;margin-right:8px;background:var(--na)}
.cell.verde{background:var(--goodbg)}.cell.verde .pi{background:var(--good)}
.cell.amarillo{background:var(--warnbg)}.cell.amarillo .pi{background:var(--warn);color:#0b0b0b}
.cell.rojo{background:var(--critbg)}.cell.rojo .pi{background:var(--crit)}
.cell.na,.cell.sin{color:var(--muted);font-weight:500}.cell.gen{font-size:20px}
.matrix tfoot th,.matrix tfoot td{border-top:2px solid var(--line)}
.pfoot{display:none}
@media (max-width:700px){main{padding:8px 12px}.tbl{display:block;overflow-x:auto}.vis{width:56px}.ico.mid{margin-left:4px}.page{padding:18px 16px}h1{font-size:30px}h2{font-size:26px}.topbar .tcorte{margin-left:0}}
@media print{
 @page{size:letter;margin:11mm}
 body{background:#fff;font-size:11.5px}main{max-width:none;padding:0}
 .topbar,.noprint{display:none!important}
 .page{box-shadow:none;border-radius:0;padding:0;margin:0;break-before:page}.page:first-child{break-before:auto}
 body.solo .page{display:none}body.solo .page.imprimir{display:block;break-before:auto}
 .avoid{break-inside:avoid}h3,h4{break-after:avoid}.card,.foto,tr{break-inside:avoid}
 h1{font-size:28px}h2{font-size:24px}.hero b{font-size:32px}.card .val{font-size:26px}.dcard .big{font-size:36px}
 .cards{grid-template-columns:repeat(4,1fr);gap:8px}.fotos{grid-template-columns:repeat(3,1fr)}.foto img{height:130px}
 .okdet{display:block}.okdet>*{display:block}.okdet summary{list-style:none}
 .plantbl{min-width:0;font-size:10.5px}.ed{border:none;background:none;padding:0;min-width:0;resize:none}
 select.ed{appearance:none}.tblwrap{overflow:visible}
 .pill,.cell,.meter span,.pi,.prio,.badge-old,.card,.kpi{-webkit-print-color-adjust:exact;print-color-adjust:exact}
 .pfoot{display:block;margin-top:14px;font-size:10px;color:var(--muted);border-top:1px solid var(--line);padding-top:4px}
}
"""

JS = r"""
const CLAVE='estatus-dist-'+document.body.dataset.grupo+'-'+document.body.dataset.corte;
function leer(){try{return JSON.parse(localStorage.getItem(CLAVE)||'{}')}catch(e){return {}}}
function guardar(d){try{localStorage.setItem(CLAVE,JSON.stringify(d))}catch(e){}}
(function(){const d=leer();document.querySelectorAll('.plantbl tbody tr').forEach(tr=>{const v=d[tr.dataset.id]||{};
 tr.querySelectorAll('.ed').forEach(el=>{if(v[el.dataset.f]!=null)el.value=v[el.dataset.f];
  el.addEventListener('input',()=>{const all=leer();all[tr.dataset.id]=all[tr.dataset.id]||{};all[tr.dataset.id][el.dataset.f]=el.value;guardar(all)});});});})();
function imprimir(id){document.querySelectorAll('.page').forEach(p=>p.classList.remove('imprimir'));
 document.querySelectorAll('details').forEach(x=>x.open=true);
 if(id){document.body.classList.add('solo');document.getElementById('d-'+id).classList.add('imprimir')}else document.body.classList.remove('solo');
 window.print();}
window.addEventListener('afterprint',()=>document.body.classList.remove('solo'));
// ---- Exportación a .xlsx (sin librerías: ZIP sin compresión + SpreadsheetML)
const CRC=(()=>{let t=new Uint32Array(256);for(let n=0;n<256;n++){let c=n;for(let k=0;k<8;k++)c=c&1?0xEDB88320^(c>>>1):c>>>1;t[n]=c>>>0}return t})();
function crc32(b){let c=0xFFFFFFFF;for(let i=0;i<b.length;i++)c=CRC[(c^b[i])&255]^(c>>>8);return (c^0xFFFFFFFF)>>>0}
function zip(files){const enc=new TextEncoder(),parts=[],cen=[];let off=0;
 for(const [name,txt] of files){const nb=enc.encode(name),data=enc.encode(txt),crc=crc32(data);
  const h=new DataView(new ArrayBuffer(30));h.setUint32(0,0x04034b50,true);h.setUint16(4,20,true);h.setUint16(6,0x0800,true);
  h.setUint32(14,crc,true);h.setUint32(18,data.length,true);h.setUint32(22,data.length,true);h.setUint16(26,nb.length,true);
  parts.push(new Uint8Array(h.buffer),nb,data);
  const c=new DataView(new ArrayBuffer(46));c.setUint32(0,0x02014b50,true);c.setUint16(4,20,true);c.setUint16(6,20,true);c.setUint16(8,0x0800,true);
  c.setUint32(16,crc,true);c.setUint32(20,data.length,true);c.setUint32(24,data.length,true);c.setUint16(28,nb.length,true);c.setUint32(42,off,true);
  cen.push(new Uint8Array(c.buffer),nb);off+=30+nb.length+data.length;}
 const csize=cen.reduce((a,b)=>a+b.length,0),e=new DataView(new ArrayBuffer(22));e.setUint32(0,0x06054b50,true);
 e.setUint16(8,files.length,true);e.setUint16(10,files.length,true);e.setUint32(12,csize,true);e.setUint32(16,off,true);
 return new Blob([...parts,...cen,new Uint8Array(e.buffer)],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});}
const X=s=>String(s==null?'':s).replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c])).replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g,'');
function col(i){let s='';i++;while(i){const m=(i-1)%26;s=String.fromCharCode(65+m)+s;i=Math.floor((i-1)/26)}return s}
function hoja(filas,anchos){const cols='<cols>'+anchos.map((w,i)=>`<col min="${i+1}" max="${i+1}" width="${w}" customWidth="1"/>`).join('')+'</cols>';
 const rows=filas.map((f,r)=>`<row r="${r+1}">`+f.map((v,c)=>`<c r="${col(c)}${r+1}" t="inlineStr"${r===2?' s="1"':(r===0?' s="2"':'')}><is><t xml:space="preserve">${X(v)}</t></is></c>`).join('')+'</row>').join('');
 return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetViews><sheetView workbookViewId="0"><pane ySplit="3" topLeftCell="A4" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>${cols}<sheetData>${rows}</sheetData><autoFilter ref="A3:${col(filas[2].length-1)}${filas.length}"/></worksheet>`;}
function exportar(ids){const tablas=[...document.querySelectorAll('.plantbl')].filter(t=>!ids||ids.includes(t.dataset.dist));
 const corte=document.body.dataset.cortetxt,grupo=document.body.dataset.gruponom;const hojas=[];
 tablas.forEach(t=>{const filas=[[grupo+' — Plan de trabajo '+t.dataset.nombre],['Fecha de corte: '+corte],
   ['#','Indicador','Concepto / puesto','Área / subtema','Qué falta','Prioridad','Responsable','Fecha compromiso','Estatus','Comentarios']];
  t.querySelectorAll('tbody tr').forEach(tr=>{const td=tr.children,v=f=>tr.querySelector(`[data-f="${f}"]`).value;
   const conc=td[2].querySelector('b').textContent,area=(td[2].querySelector('.muted')||{}).textContent||'';
   filas.push([td[0].textContent,td[1].textContent,conc,area,td[3].textContent,'ALTA',v('resp'),v('fecha'),v('estatus'),v('coment')]);});
  hojas.push([t.dataset.nombre.replace(/[\\\/\?\*\[\]:]/g,'').slice(0,31),hoja(filas,[5,24,48,22,48,10,22,16,14,48])]);});
 const ct=`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>${hojas.map((h,i)=>`<Override PartName="/xl/worksheets/sheet${i+1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('')}</Types>`;
 const rels=`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`;
 const wb=`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>${hojas.map((h,i)=>`<sheet name="${X(h[0])}" sheetId="${i+1}" r:id="rId${i+1}"/>`).join('')}</sheets>${hojas.map((h,i)=>`<definedNames><definedName name="_xlnm._FilterDatabase" localSheetId="${i}" hidden="1">'${X(h[0])}'!$A$3:$J$3</definedName></definedNames>`).slice(0,0).join('')}</workbook>`;
 const wbr=`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${hojas.map((h,i)=>`<Relationship Id="rId${i+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i+1}.xml"/>`).join('')}<Relationship Id="rId${hojas.length+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`;
 const st=`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="3"><font><sz val="11"/><name val="Calibri"/></font><font><b/><color rgb="FFFFFFFF"/><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="14"/><name val="Calibri"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF1D2A3A"/></patternFill></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf/></cellStyleXfs><cellXfs count="3"><xf/><xf fontId="1" fillId="2" applyFont="1" applyFill="1"/><xf fontId="2" applyFont="1"/></cellXfs></styleSheet>`;
 const files=[['[Content_Types].xml',ct],['_rels/.rels',rels],['xl/workbook.xml',wb],['xl/_rels/workbook.xml.rels',wbr],['xl/styles.xml',st],...hojas.map((h,i)=>[`xl/worksheets/sheet${i+1}.xml`,h[1]])];
 const a=document.createElement('a');a.href=URL.createObjectURL(zip(files));
 a.download=`Plan_de_trabajo_${grupo.replace(/\s+/g,'_')}${ids&&ids.length===1?'_'+tablas[0].dataset.nombre.replace(/\s+/g,'_'):''}_${document.body.dataset.corte}.xlsx`;
 document.body.appendChild(a);a.click();setTimeout(()=>{URL.revokeObjectURL(a.href);a.remove()},1000);}
"""


def render_html(grupo, dists, cfg, corte):
    nav = "".join(f'<a href="#d-{d["id"]}">{esc(d["nombre"])}</a>' for d in dists)
    cuerpo = render_portada(grupo, dists, cfg, corte) + "".join(render_distribuidor(d, grupo, corte) for d in dists)
    return f"""<!doctype html>
<html lang="es"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>{esc(grupo["nombre"])} — Estatus de distribuidores</title><style>{CSS}</style></head>
<body data-grupo="{esc(grupo["id"])}" data-gruponom="{esc(grupo["nombre"])}" data-corte="{corte.isoformat()}" data-cortetxt="{fecha_larga(corte)}">
<div class="topbar"><b>{esc(grupo["nombre"])}</b><nav><a href="#resumen">Resumen</a>{nav}</nav>
<span class="tcorte">Fecha de corte: {fecha_larga(corte)}</span></div>
<main>{cuerpo}</main>
<script>{JS}</script></body></html>"""


# --------------------------------------------------------------------------- consola
def imprimir_resumen(grupo, dists):
    print(f"\n== {grupo['nombre']} ==")
    for d in dists:
        print(f"  {d['nombre']}: general {pct(d['general'])} ({ESTADO_TXT[d['estado_general']]}), "
              f"{len(d['pendientes'])} pendientes")
        for i in d["indicadores"]:
            print(f"     - {i['nombre']:<28} {('N/A' if i['estado'] == 'na' else pct(i['valor'])):>8}  "
                  f"{ESTADO_TXT[i['estado']]:<15} {i.get('fuente', '')}")


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--config", default=os.path.join(BASE, "config.json"))
    ap.add_argument("--grupo", help="id del grupo a generar (por defecto, todos)")
    a = ap.parse_args()
    cfg = json.load(open(a.config, encoding="utf-8"))
    ruta = lambda p: p if os.path.isabs(p) else os.path.join(BASE, p)
    corte = dt.date.fromisoformat(cfg["fecha_corte"])
    kpi = LectorKPI(ruta(cfg["archivos"]["kpis"]))
    per = LectorPersonal(ruta(cfg["archivos"]["personal_sistema"]))
    fotos = Fotos(ruta(cfg["archivos"]["fotos"]), cfg, corte)
    ajustes = cargar_ajustes(ruta(cfg["archivos"]["ajustes"])) if cfg["archivos"].get("ajustes") else []
    salida = ruta(cfg["carpeta_salida"])
    os.makedirs(salida, exist_ok=True)
    grupos = [g for g in cfg["grupos"] if not a.grupo or g["id"] == a.grupo]
    if not grupos:
        raise SystemExit(f"No existe el grupo «{a.grupo}» en config.json")
    for g in grupos:
        dists = evaluar(cfg, g, kpi, per, fotos, corte, ajustes)
        out = os.path.join(salida, g["archivo_salida"])
        with open(out, "w", encoding="utf-8") as fh:
            fh.write(render_html(g, dists, cfg, corte))
        imprimir_resumen(g, dists)
        print(f"  -> {os.path.relpath(out, BASE)} ({os.path.getsize(out) / 1024:.0f} KB)")


if __name__ == "__main__":
    main()
