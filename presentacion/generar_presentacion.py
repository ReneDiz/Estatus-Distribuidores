#!/usr/bin/env python3
"""
Genera la presentación ejecutiva en PowerPoint con el template oficial FOTON.

Uso:
    python3 presentacion/generar_presentacion.py
    python3 presentacion/generar_presentacion.py --salida presentacion/otra.pptx

Lee los mismos datos que la herramienta (config.json, KPI's, personal, fotos, ajustes)
más datos/explicaciones.json, y usa template/Foton_Mexico_-_Template_Base-2026.pptx
(fondos, logo, tipografías y colores). Las diapositivas de muestra del template no se
incluyen en el resultado.
"""
import argparse
import io
import os
import sys
import zipfile

from lxml import etree
from PIL import Image, ImageOps
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION, XL_LEGEND_POSITION
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, RAIZ)
import analisis as A  # noqa: E402
import generar_reporte as g  # noqa: E402

TEMPLATE = os.path.join(RAIZ, "template", "Foton_Mexico_-_Template_Base-2026.pptx")

# ---- paleta del template FOTON
BLANCO = "F2F2F2"
CIAN = "17DFE9"
AZUL = "48C8E5"
AZUL_OSC = "1C4587"
GRIS = "A9B3B7"
GRIS_CL = "D1D6DB"
TARJETA = "16233F"
TARJETA2 = "1E2F52"
LINEA = "2B3E63"
SEM = {"verde": "2FBF71", "amarillo": "F2B544", "rojo": "E5484D", "na": "6B7A90", "sin": "6B7A90"}
SEM_TXT = {"verde": "Cumple", "amarillo": "En proceso", "rojo": "No cumple", "na": "No aplica", "sin": "Sin información"}

# medios del template (ppt/media)
FONDO_NAVY = "image10.png"
FONDO_MAPA = "image13.png"
FOTO_PORTADA = "image8.png"
FOTO_CIERRE = "image9.png"
LOGO = "image2.png"

X0 = 0.4          # margen izquierdo
ANCHO = 9.2       # ancho útil
Y_CONT = 1.55     # inicio del contenido
Y_FIN = 5.3       # fin del contenido


def rgb(h):
    return RGBColor.from_string(h)


# =========================================================================== deck
class Deck:
    def __init__(self, template):
        self.prs = Presentation(template)
        self.n_muestra = len(self.prs.slides)
        with zipfile.ZipFile(template) as z:
            self.media = {os.path.basename(n): z.read(n) for n in z.namelist() if n.startswith("ppt/media/")}
        self.layout = next(l for l in self.prs.slide_layouts if l.name == "BLANK")
        self._fotos = {}

    # ---------------------------------------------------------------- diapositivas
    def nueva(self, fondo=FONDO_NAVY, notas=None):
        s = self.prs.slides.add_slide(self.layout)
        if fondo:
            self.fondo(s, self.jpg(fondo) if fondo != FONDO_NAVY else self.media[fondo])
        if notas:
            s.notes_slide.notes_text_frame.text = notas
        return s

    def fondo(self, s, data):
        _, rid = s.part.get_or_add_image_part(io.BytesIO(data))
        csld = s._element.find(qn("p:cSld"))
        for viejo in csld.findall(qn("p:bg")):
            csld.remove(viejo)
        bg = etree.SubElement(csld, qn("p:bg"))
        csld.remove(bg)
        csld.insert(0, bg)
        bgpr = etree.SubElement(bg, qn("p:bgPr"))
        bf = etree.SubElement(bgpr, qn("a:blipFill"))
        etree.SubElement(bf, qn("a:blip")).set(qn("r:embed"), rid)
        etree.SubElement(etree.SubElement(bf, qn("a:stretch")), qn("a:fillRect"))
        etree.SubElement(bgpr, qn("a:effectLst"))

    def terminar(self, salida):
        lista = self.prs.slides._sldIdLst
        for sid in list(lista)[: self.n_muestra]:
            self.prs.part.drop_rel(sid.rId)
            lista.remove(sid)
        self.prs.save(salida)

    # ---------------------------------------------------------------- imágenes
    def imagen(self, s, data, x, y, w, h, nombre=None, alt=""):
        pic = s.shapes.add_picture(io.BytesIO(data), Inches(x), Inches(y), Inches(w), Inches(h))
        if nombre:
            pic.name = nombre
        if alt:
            pic._element.nvPicPr.cNvPr.set("descr", alt)
        return pic

    def foto(self, ruta, w, h, max_px=1600):
        """JPEG recortado al marco (sin deformar) y reducido para que el archivo no pese de más."""
        clave = (ruta, round(w / h, 3))
        if clave not in self._fotos:
            im = ImageOps.exif_transpose(Image.open(ruta))
            if "icc_profile" in im.info:
                try:
                    from PIL import ImageCms
                    src = ImageCms.ImageCmsProfile(io.BytesIO(im.info["icc_profile"]))
                    im = ImageCms.profileToProfile(im, src, ImageCms.createProfile("sRGB"), outputMode="RGB")
                except Exception:
                    pass
            im = im.convert("RGB")
            px_w = min(max_px, im.width)
            im = ImageOps.fit(im, (px_w, max(1, round(px_w * h / w))), Image.LANCZOS)
            buf = io.BytesIO()
            im.save(buf, "JPEG", quality=82, optimize=True)
            self._fotos[clave] = buf.getvalue()
        return self._fotos[clave]

    def jpg(self, nombre):
        """Versión JPEG de una imagen grande del template (pesa menos en el archivo final)."""
        clave = ("media", nombre)
        if clave not in self._fotos:
            im = Image.open(io.BytesIO(self.media[nombre]))
            if im.mode in ("RGBA", "LA", "P"):
                im = im.convert("RGBA")
                base = Image.new("RGB", im.size, (11, 22, 43))
                base.paste(im, mask=im.split()[-1])
                im = base
            buf = io.BytesIO()
            im.convert("RGB").save(buf, "JPEG", quality=88, optimize=True)
            self._fotos[clave] = buf.getvalue()
        return self._fotos[clave]

    def logo(self, s):
        return self.imagen(s, self.media[LOGO], 0.29, 0.23, 1.31, 0.39, "Logo FOTON", "FOTON")


# =========================================================================== primitivas
def caja(s, x, y, w, h, fill=TARJETA, radio=0.06, linea=None, nombre=None, alpha=None, forma=None):
    tipo = forma or (MSO_SHAPE.ROUNDED_RECTANGLE if radio else MSO_SHAPE.RECTANGLE)
    sh = s.shapes.add_shape(tipo, Inches(x), Inches(y), Inches(w), Inches(h))
    if radio and tipo == MSO_SHAPE.ROUNDED_RECTANGLE:
        sh.adjustments[0] = min(0.5, radio / min(w, h))
    if fill:
        sh.fill.solid()
        sh.fill.fore_color.rgb = rgb(fill)
        if alpha is not None:
            sf = sh.fill._xPr.find(qn("a:solidFill"))[0]
            etree.SubElement(sf, qn("a:alpha")).set("val", str(int(alpha * 100000)))
    else:
        sh.fill.background()
    if linea:
        sh.line.color.rgb = rgb(linea)
        sh.line.width = Pt(0.75)
    else:
        sh.line.fill.background()
    sh.shadow.inherit = False
    if nombre:
        sh.name = nombre
    sh.text_frame.text = ""
    return sh


def punto(s, x, y, d, estado):
    return caja(s, x, y, d, d, SEM[estado], radio=0, forma=MSO_SHAPE.OVAL, nombre=f"Semáforo {SEM_TXT[estado]}")


def texto(s, x, y, w, h, parrafos, size=10, color=BLANCO, bold=False, font="Arial", align="l",
          anchor="t", espacio=0, interlinea=None, nombre=None, auto=False):
    """parrafos: str | lista de (str | lista de runs (texto, {size,color,bold,font,italic}) | dict)."""
    tb = s.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    if nombre:
        tb.name = nombre
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = {"t": MSO_ANCHOR.TOP, "m": MSO_ANCHOR.MIDDLE, "b": MSO_ANCHOR.BOTTOM}[anchor]
    if auto:
        tf.auto_size = True
    if isinstance(parrafos, str):
        parrafos = [parrafos]
    for k, par in enumerate(parrafos):
        p = tf.paragraphs[0] if k == 0 else tf.add_paragraph()
        opts = {}
        if isinstance(par, dict):
            opts = par
            par = par["runs"]
        if isinstance(par, str):
            par = [(par, {})]
        p.alignment = {"l": PP_ALIGN.LEFT, "c": PP_ALIGN.CENTER, "r": PP_ALIGN.RIGHT}[opts.get("align", align)]
        esp = opts.get("espacio", espacio)
        if esp:
            p.space_after = Pt(esp)
        if opts.get("antes"):
            p.space_before = Pt(opts["antes"])
        if interlinea:
            p.line_spacing = interlinea
        for t, o in par:
            r = p.add_run()
            r.text = t
            f = r.font
            f.name = o.get("font", opts.get("font", font))
            f.size = Pt(o.get("size", opts.get("size", size)))
            f.bold = o.get("bold", opts.get("bold", bold))
            f.italic = o.get("italic", False)
            f.color.rgb = rgb(o.get("color", opts.get("color", color)))
    return tb


def titulo(s, t, sub=None):
    texto(s, X0, 0.82, 8.6, 0.4, t, size=20, bold=True, color=BLANCO, nombre="Título")
    if sub:
        texto(s, X0, 1.18, 8.8, 0.3, sub, size=12, bold=True, color=CIAN, nombre="Subtítulo")


def barra(s, x, y, w, h, valor, estado, fondo=LINEA):
    caja(s, x, y, w, h, fondo, radio=h / 2)
    if valor and valor > 0:
        caja(s, x, y, max(h, w * min(valor, 100) / 100), h, SEM[estado], radio=h / 2)


def pie(s, t):
    texto(s, X0, 5.33, 9.2, 0.2, t, size=7.5, color=GRIS, nombre="Nota al pie")


def pc(v, dec=1):
    return A.fmt(v, dec)


def tabla(s, x, y, w, filas, anchos, alto_fila=0.27, size=8.5, encabezado=True, alto_enc=None):
    """Tabla nativa con estilo FOTON. filas: lista de listas de celdas; celda = str o dict(t, color, bold, fill, align)."""
    nf, nc = len(filas), len(filas[0])
    alto_enc = alto_enc or alto_fila
    shp = s.shapes.add_table(nf, nc, Inches(x), Inches(y), Inches(w), Inches(alto_enc + alto_fila * (nf - 1)))
    tbl = shp.table
    tblpr = tbl._tbl.tblPr
    for k in ("bandRow", "firstRow"):
        tblpr.set(k, "0")
    est = tblpr.find(qn("a:tableStyleId"))
    if est is not None:
        tblpr.remove(est)
    for j, a in enumerate(anchos):
        tbl.columns[j].width = Inches(a)
    for i, fila in enumerate(filas):
        tbl.rows[i].height = Inches(alto_enc if i == 0 and encabezado else alto_fila)
        for j, c in enumerate(fila):
            c = c if isinstance(c, dict) else {"t": str(c)}
            cell = tbl.cell(i, j)
            es_enc = encabezado and i == 0
            fill = c.get("fill", AZUL_OSC if es_enc else (TARJETA if i % 2 else TARJETA2))
            cell.fill.solid()
            cell.fill.fore_color.rgb = rgb(fill)
            cell.margin_left = cell.margin_right = Inches(0.06)
            cell.margin_top = cell.margin_bottom = Inches(0.02)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            tf = cell.text_frame
            tf.word_wrap = True
            runs = c.get("runs") or [(c.get("t", ""), {})]
            p = tf.paragraphs[0]
            p.alignment = {"l": PP_ALIGN.LEFT, "c": PP_ALIGN.CENTER, "r": PP_ALIGN.RIGHT}[c.get("align", "l")]
            for t, o in runs:
                r = p.add_run()
                r.text = t
                r.font.name = "Arial"
                r.font.size = Pt(o.get("size", c.get("size", size)))
                r.font.bold = o.get("bold", c.get("bold", es_enc))
                r.font.color.rgb = rgb(o.get("color", c.get("color", BLANCO)))
            # bordes finos
            tcpr = cell._tc.get_or_add_tcPr()
            for lado in ("a:lnL", "a:lnR", "a:lnT", "a:lnB"):
                ln = etree.SubElement(tcpr, qn(lado), w="6350")
                sf = etree.SubElement(ln, qn("a:solidFill"))
                etree.SubElement(sf, qn("a:srgbClr")).set("val", "0B162B")
            # el orden del esquema exige lnX antes del relleno
            fillel = tcpr.find(qn("a:solidFill"))
            if fillel is not None:
                tcpr.remove(fillel)
                tcpr.append(fillel)
    return shp


def estilo_grafica(ch, size=9):
    ch.font.name = "Arial"
    ch.font.size = Pt(size)
    ch.font.color.rgb = rgb(GRIS_CL)


# =========================================================================== diapositivas
def portada(dk, ctx):
    s = dk.nueva(fondo=None, notas="Fuente: KPI's FOTON (corte 30-oct-2026), Sistema de Personal FOTON (noviembre) y "
                                   "fotos de visitas. Mismo peso para los 11 indicadores.")
    dk.imagen(s, dk.jpg(FOTO_PORTADA), 0, 0, 10, 5.625, "Foto portada")
    dk.logo(s)
    texto(s, 0.38, 4.18, 7.5, 0.5, ctx["grupo"]["nombre"].upper(), size=28, bold=True, font="Arial Black",
          color="212121", nombre="Título portada")
    texto(s, 0.4, 4.72, 7.5, 0.3, "Estatus de distribuidores y plan de trabajo", size=13, bold=True, color="212121")
    texto(s, 0.4, 5.0, 7.5, 0.3, f"Corte al {g.fecha_larga(ctx['corte'])} · "
          + " · ".join(d["corto"] for d in ctx["dists"]), size=10, color="3A495A")


def resumen_ejecutivo(dk, ctx):
    t = ctx["total"]
    s = dk.nueva(notas=(
        f"Total del grupo: promedio de los {t['n_indicadores']} indicadores aplicables de las tres agencias "
        f"(Allende incluido, sin sus 5 indicadores que no aplican). Proyecto del grupo: suma del avance de los "
        f"{t['elementos']} elementos entre {t['elementos']}. Semáforo general: verde desde 80%, amarillo 50-79%. "
        f"Proyecto: verde desde 80%, amarillo 50-79%."))
    titulo(s, "Resumen ejecutivo", f"{ctx['grupo']['nombre']} · corte {g.fecha_corta(ctx['corte'])} · "
           f"11 indicadores con el mismo peso")
    # bloque del grupo
    caja(s, X0, Y_CONT, 2.75, 3.7, TARJETA, nombre="Total del grupo")
    texto(s, X0 + 0.2, Y_CONT + 0.15, 2.4, 0.25, "TOTAL DEL GRUPO", size=10, bold=True, color=CIAN)
    for k, (lab, val, est, det) in enumerate([
            ("Calificación general", t["general"], t["estado_general"], f"{t['n_indicadores']} indicadores"),
            ("Desarrollo de proyecto", t["proyecto"], t["estado_proyecto"], f"{t['elementos']} elementos")]):
        y = Y_CONT + 0.5 + k * 1.55
        texto(s, X0 + 0.2, y, 2.4, 0.22, lab, size=10, color=GRIS_CL)
        texto(s, X0 + 0.2, y + 0.22, 1.9, 0.6, pc(t[["general", "proyecto"][k]]), size=32, bold=True,
              font="Arial Black", color=BLANCO)
        punto(s, X0 + 0.22, y + 0.92, 0.16, est)
        texto(s, X0 + 0.45, y + 0.89, 2.2, 0.22, f"{SEM_TXT[est]} · {det}", size=9, color=GRIS_CL)
    # tarjetas por agencia
    w = 2.0
    for k, d in enumerate(ctx["dists"]):
        x = X0 + 2.95 + k * (w + 0.12)
        caja(s, x, Y_CONT, w, 3.7, TARJETA, nombre=f"Agencia {d['corto']}")
        texto(s, x + 0.15, Y_CONT + 0.15, w - 0.3, 0.25, d["nombre"].upper(), size=10, bold=True, color=CIAN)
        for j, (lab, val, est) in enumerate([("General", d["general"], d["estado_general"]),
                                             ("Proyecto", d["proyecto"]["valor"], d["proyecto"]["estado"])]):
            y = Y_CONT + 0.5 + j * 1.0
            punto(s, x + 0.15, y + 0.1, 0.3, est)
            texto(s, x + 0.55, y, w - 0.65, 0.22, lab, size=9, color=GRIS_CL)
            texto(s, x + 0.55, y + 0.2, w - 0.65, 0.4, pc(val), size=20, bold=True, color=BLANCO)
            texto(s, x + 0.55, y + 0.6, w - 0.65, 0.2, SEM_TXT[est], size=8.5, color=SEM[est], bold=True)
        texto(s, x + 0.15, Y_CONT + 2.55, w - 0.3, 0.4, [[(f"{d['n_pendientes']}", {"size": 20, "bold": True}),
                                                         ("  pendientes", {"size": 9, "color": GRIS_CL})]])
        nota = d.get("mensaje_clave") or ("; ".join(f"{a['nombre']} {pc(a['valor'], 0)}" for a in
                                                     sorted([a for a in d["areas"] if a["faltan"]], key=lambda a: a["valor"])[:2])
                                          + " son lo más atrasado.")
        texto(s, x + 0.15, Y_CONT + 3.0, w - 0.3, 0.6, nota, size=8.5, color=GRIS_CL)
    pie(s, "Semáforo general y de proyecto: verde ≥ 80%, amarillo 50–79%, rojo < 50%.")


def matriz(dk, ctx):
    s = dk.nueva(notas="Cada celda es el % del indicador al corte. Binarios: 100% o 0%. Cummins y proyecto: promedio "
                       "de sus conceptos. Personal: plazas cubiertas / requeridas del Sistema de Personal. "
                       "En Allende no aplican CRM, Fotonet, Kit, Diagnóstico ni Cummins (showroom).")
    titulo(s, "Los 11 indicadores por agencia", "Lo que está en rojo o amarillo es pendiente de prioridad alta")
    ds = ctx["dists"]
    filas = [[{"t": "Indicador"}] + [{"t": d["corto"], "align": "c"} for d in ds]]
    for k, ind in enumerate(ds[0]["resumen"]):
        fila = [{"t": ind["nombre"].replace("% ", "")}]
        for d in ds:
            e = d["resumen"][k]
            if e["estado"] == "na":
                fila.append({"t": "No aplica", "color": GRIS, "align": "c"})
            else:
                fila.append({"runs": [("●  ", {"color": SEM[e["estado"]], "size": 13}),
                                      (pc(e["valor"]), {"bold": e["estado"] != "verde"})], "align": "c"})
        filas.append(fila)
    fila = [{"t": "Calificación general", "bold": True, "fill": AZUL_OSC}]
    for d in ds:
        fila.append({"runs": [("●  ", {"color": SEM[d["estado_general"]], "size": 13}), (pc(d["general"]), {"bold": True})],
                     "align": "c", "fill": AZUL_OSC})
    filas.append(fila)
    tabla(s, X0, Y_CONT, 6.1, filas, [2.5, 1.2, 1.2, 1.2], alto_fila=0.268, size=9)
    # lectura rápida
    x = 6.8
    caja(s, x, Y_CONT, 2.8, 3.48, TARJETA)
    texto(s, x + 0.18, Y_CONT + 0.15, 2.45, 0.25, "LECTURA RÁPIDA", size=10, bold=True, color=CIAN)
    pars = []
    for d in ds:
        pend = [e["nombre"].replace("% ", "") for e in d["resumen"] if e["falta"]]
        ok = sum(1 for e in d["resumen"] if e["estado"] == "verde")
        apl = sum(1 for e in d["resumen"] if e["estado"] != "na")
        pars.append({"runs": [(d["corto"], {"bold": True, "color": BLANCO})], "size": 10, "espacio": 1})
        pars.append({"runs": [(f"Cumple {ok} de {apl}. Falta: " + ", ".join(pend) + ".", {})],
                     "size": 9, "color": GRIS_CL, "espacio": 9})
    texto(s, x + 0.18, Y_CONT + 0.48, 2.45, 2.9, pars)
    pie(s, "Semáforo: binarios verde sólo al 100%; Cummins amarillo 50–99%; proyecto, personal y general verde ≥ 80%, amarillo 50–79%.")


def concentracion(dk, ctx):
    ds = ctx["dists"]
    s = dk.nueva(notas="Conteo de pendientes por agencia: cada elemento de proyecto, requisito de Cummins, puesto con "
                       "vacantes o indicador binario no cumplido cuenta como un pendiente.")
    titulo(s, "Dónde se concentra lo pendiente", f"{ctx['total']['pendientes']} pendientes en el grupo, todos con prioridad alta")
    grupos = [("Proyecto", lambda p: p["indicador"].endswith("Proyecto")),
              ("Personal", lambda p: p["indicador"] == "Personal"),
              ("Cummins", lambda p: p["indicador"] == "Cummins"),
              ("Otros", lambda p: not (p["indicador"].endswith("Proyecto") or p["indicador"] in ("Personal", "Cummins")))]
    cd = CategoryChartData()
    cd.categories = [f"{d['corto']} ({d['n_pendientes']})" for d in ds]
    valores = []
    for nom, f in grupos:
        valores.append([sum(1 for p in d["pendientes"] if f(p)) for d in ds])
        cd.add_series(nom, valores[-1])
    gf = s.shapes.add_chart(XL_CHART_TYPE.BAR_STACKED, Inches(X0), Inches(Y_CONT), Inches(5.2), Inches(3.6), cd)
    gf.name = "Gráfica pendientes por agencia"
    ch = gf.chart
    estilo_grafica(ch)
    ch.has_legend = True
    ch.legend.position = XL_LEGEND_POSITION.TOP
    ch.legend.include_in_layout = False
    ch.legend.font.color.rgb = rgb(GRIS_CL)
    colores = [AZUL, CIAN, "7E8FA6", GRIS_CL]
    for ser, col in zip(ch.plots[0].series, colores):
        ser.format.fill.solid()
        ser.format.fill.fore_color.rgb = rgb(col)
    pl = ch.plots[0]
    pl.gap_width = 60
    pl.overlap = 100
    pl.has_data_labels = True
    pl.data_labels.position = XL_LABEL_POSITION.CENTER
    pl.data_labels.font.size = Pt(9)
    pl.data_labels.font.bold = True
    pl.data_labels.font.color.rgb = rgb("0B162B")
    ch.value_axis.visible = False
    ch.value_axis.has_major_gridlines = False
    ch.category_axis.format.line.fill.background()
    ch.category_axis.tick_labels.font.size = Pt(10)
    ch.category_axis.tick_labels.font.color.rgb = rgb(BLANCO)
    # sin etiqueta en segmentos muy cortos (no caben); el total va en la categoría
    for ser, vals in zip(pl.series, valores):
        dls = ser.data_labels
        dls.show_value = True
        dls.position = XL_LABEL_POSITION.CENTER
        dls.font.size = Pt(9)
        dls.font.bold = True
        dls.font.color.rgb = rgb("0B162B")
        for k, v in enumerate(vals):
            if v < 4:
                dl = ser.points[k].data_label._get_or_add_dLbl()
                for hijo in list(dl)[1:]:
                    dl.remove(hijo)
                etree.SubElement(dl, qn("c:delete")).set("val", "1")
    ch.category_axis.reverse_order = True
    # mensajes
    x = 5.85
    mob = sum(a["faltan"] for d in ds for a in d["areas"] if g.norm(a["original"]) == "mobiliario")
    vac = sum(d["personal"]["faltan"] for d in ds if d["personal"])
    peor = min(((d, a) for d in ds for a in d["areas"] if a["faltan"] and g.norm(a["original"]) != "mobiliario"),
               key=lambda x: (x[1]["valor"], -x[1]["faltan"]))
    taller = [(d, a) for d in ds for a in d["areas"] if g.norm(a["original"]) == "taller" and a["faltan"] >= 10]
    cum = [d for d in ds for e in d["resumen"] if e["id"] == "cummins" and e["valor"] is not None and e["valor"] == 0]
    msgs = [("Mobiliario", f"{mob} elementos en 0% en las tres agencias. Validar si el mobiliario actual cumple el estándar.")]
    msgs.append(("Personal", f"{vac:g} vacantes en el grupo. La mayor parte son asesores de ventas."))
    for d, a in taller:
        msgs.append((f"Taller {d['corto']}", f"{pc(a['valor'], 0)}: faltan {a['faltan']} de {a['n']} elementos."))
    for d in cum:
        msgs.append((f"Cummins {d['corto']}", "0 de 7 requisitos. Arrancar contrato, crédito y QSOL."))
    for d in ds:
        if d.get("mensaje_clave"):
            msgs.append((d["corto"], d["mensaje_clave"]))
    y = Y_CONT
    alto = 3.6 / len(msgs)
    for nom, txt in msgs:
        caja(s, x, y, 3.75, alto - 0.08, TARJETA)
        texto(s, x + 0.15, y + 0.07, 3.45, alto - 0.2, [{"runs": [(nom, {"bold": True, "color": CIAN})], "size": 10, "espacio": 2},
                                                         {"runs": [(txt, {})], "size": 9, "color": GRIS_CL}], anchor="m")
        y += alto


def escenario_mobiliario(dk, ctx):
    t = ctx["total"]
    ds = ctx["dists"]
    s = dk.nueva(fondo=FONDO_MAPA, notas=(
        "Escenario: los elementos de mobiliario se toman como cumplidos (100%) porque el mobiliario actual se valida "
        "como OK. El resto de los conceptos e indicadores no cambia. Proyecto = suma del avance de los elementos / "
        "número de elementos. General = promedio de los indicadores aplicables con el nuevo proyecto. En Allende sólo "
        "aplica si se acepta el mobiliario de la oficina móvil."))
    titulo(s, "Mobiliario: cuánto sube si se valida el actual",
           f"Hoy los {t['mobiliario']} elementos de mobiliario están en 0% en el KPI")
    filas = [["Agencia", "Elementos", "Proyecto hoy", "Proyecto validado", "General hoy", "General validado"]]
    for d in ds:
        e = d["escenario"]
        filas.append([{"t": d["corto"], "bold": True}, {"t": str(e["elementos"]), "align": "c"},
                      {"t": pc(e["proyecto_hoy"]), "align": "c"},
                      {"runs": [(pc(e["proyecto_validado"]), {"bold": True, "color": CIAN}),
                                (f"  +{e['proyecto_validado'] - e['proyecto_hoy']:.1f}", {"size": 8, "color": GRIS_CL})], "align": "c"},
                      {"t": pc(e["general_hoy"]), "align": "c"},
                      {"runs": [(pc(e["general_validado"]), {"bold": True, "color": CIAN}),
                                (f"  +{e['general_validado'] - e['general_hoy']:.1f}", {"size": 8, "color": GRIS_CL})], "align": "c"}])
    filas.append([{"t": "Grupo", "bold": True, "fill": AZUL_OSC}, {"t": str(t["mobiliario"]), "align": "c", "fill": AZUL_OSC},
                  {"t": pc(t["proyecto"]), "align": "c", "fill": AZUL_OSC},
                  {"runs": [(pc(t["proyecto_validado"]), {"bold": True, "color": CIAN}),
                            (f"  +{t['proyecto_validado'] - t['proyecto']:.1f}", {"size": 8, "color": GRIS_CL})],
                   "align": "c", "fill": AZUL_OSC},
                  {"t": pc(t["general"]), "align": "c", "fill": AZUL_OSC},
                  {"runs": [(pc(t["general_validado"]), {"bold": True, "color": CIAN}),
                            (f"  +{t['general_validado'] - t['general']:.1f}", {"size": 8, "color": GRIS_CL})],
                   "align": "c", "fill": AZUL_OSC}])
    tabla(s, X0, Y_CONT + 0.05, 9.2, filas, [1.6, 1.1, 1.5, 1.75, 1.5, 1.75], alto_fila=0.5, size=11, alto_enc=0.38)
    y = Y_CONT + 0.05 + 0.38 + 0.5 * 4 + 0.25
    caja(s, X0, y, 9.2, Y_FIN - y, TARJETA)
    texto(s, X0 + 0.2, y + 0.1, 8.8, Y_FIN - y - 0.2, anchor="m", parrafos=[
        {"runs": [("Qué hay que decidir. ", {"bold": True, "color": CIAN}),
                  ("Revisar en la visita si escritorios, mueble de recepción, mesas y exhibidores actuales cumplen el "
                   "estándar FOTON. Si cumplen, se registran en el KPI; si no, se cotiza el faltante.", {})],
         "size": 10, "espacio": 4},
        {"runs": [("El general sube poco porque el mobiliario es sólo una parte del indicador de proyecto. "
                   "Allende sólo cuenta si se acepta el mobiliario de la oficina móvil.", {})], "size": 9, "color": GRIS_CL}])


def seccion(dk, d, ruta_foto, nota_foto=None):
    s = dk.nueva(fondo=None, notas=f"{d['nombre']}: general {pc(d['general'], 2)}, proyecto {pc(d['proyecto']['valor'], 2)}, "
                                   f"{d['n_pendientes']} pendientes.")
    dk.imagen(s, dk.foto(ruta_foto, 10, 5.625, 2000), 0, 0, 10, 5.625, "Fachada", f"Fachada {d['nombre']}")
    # panel oscuro con degradado, igual al del template (título de sección)
    sh = caja(s, 0, 0, 7.18, 5.625, None, radio=0, nombre="Panel")
    sppr = sh._element.spPr
    for el in sppr.findall(qn("a:noFill")):
        sppr.remove(el)
    grad = etree.fromstring(
        '<a:gradFill xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" rotWithShape="1"><a:gsLst>'
        '<a:gs pos="0"><a:srgbClr val="19193B"><a:alpha val="0"/></a:srgbClr></a:gs>'
        '<a:gs pos="100000"><a:srgbClr val="101026"/></a:gs></a:gsLst><a:lin ang="10800000" scaled="0"/></a:gradFill>')
    sppr.insert(list(sppr).index(sppr.find(qn("a:prstGeom"))) + 1, grad)
    # degradado inferior para que el texto se lea sobre cualquier foto
    sh2 = caja(s, 0, 3.3, 10, 2.325, None, radio=0, nombre="Sombra inferior")
    sppr = sh2._element.spPr
    for el in sppr.findall(qn("a:noFill")):
        sppr.remove(el)
    grad = etree.fromstring(
        '<a:gradFill xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" rotWithShape="1"><a:gsLst>'
        '<a:gs pos="0"><a:srgbClr val="101026"><a:alpha val="0"/></a:srgbClr></a:gs>'
        '<a:gs pos="100000"><a:srgbClr val="101026"><a:alpha val="85000"/></a:srgbClr></a:gs></a:gsLst>'
        '<a:lin ang="5400000" scaled="0"/></a:gradFill>')
    sppr.insert(list(sppr).index(sppr.find(qn("a:prstGeom"))) + 1, grad)
    dk.logo(s)
    texto(s, 0.38, 3.55, 7.5, 0.6, d["nombre"].upper(), size=36, bold=True, font="Arial Black", color="FFFFFF")
    texto(s, 0.4, 4.2, 9.0, 0.3, [[("General ", {}), (pc(d["general"]), {"color": SEM[d["estado_general"]]}),
                                   ("   ·   Proyecto ", {}), (pc(d["proyecto"]["valor"]), {"color": SEM[d["proyecto"]["estado"]]}),
                                   (f"   ·   {d['n_pendientes']} pendientes", {})]], size=14, bold=True, color="FFFFFF")
    lineas = []
    if d.get("mensaje_clave"):
        lineas.append({"runs": [(d["mensaje_clave"], {"bold": True, "color": CIAN})], "size": 13, "espacio": 3})
    na = [i["nombre"] for i in d["indicadores"] if i["estado"] == "na"]
    if na:
        lineas.append({"runs": [(f"{d.get('motivo_no_aplica', 'No aplica')}. No aplican: {', '.join(na)}.", {})],
                       "size": 10, "color": GRIS_CL})
    if lineas:
        texto(s, 0.4, 4.6, 8.8, 0.7, lineas)
    if nota_foto:
        caja(s, 8.45, 0.25, 1.3, 0.28, "000000", radio=0.06, alpha=0.55)
        texto(s, 8.45, 0.25, 1.3, 0.28, nota_foto, size=8, bold=True, color="FFFFFF", align="c", anchor="m")


def indicadores(dk, d):
    s = dk.nueva(notas="Indicadores al corte. Se muestra en detalle sólo lo que no está al 100%. "
                       + " ".join(f"{e['nombre']}: {pc(e['valor'])}." for e in d["resumen"] if e["estado"] != "na"))
    titulo(s, f"{d['corto']}: qué falta por indicador",
           f"Calificación general {pc(d['general'])} · {SEM_TXT[d['estado_general']].lower()}")
    ok = [e for e in d["resumen"] if e["estado"] == "verde"]
    na = [e for e in d["resumen"] if e["estado"] == "na"]
    pend = [e for e in d["resumen"] if e["falta"]]
    # franja compacta de lo que ya cumple
    caja(s, X0, Y_CONT, ANCHO, 0.42, TARJETA)
    runs = [("Cumplen al 100%:  ", {"bold": True, "color": SEM["verde"]})]
    runs.append(("  ·  ".join(e["nombre"].replace("% ", "") for e in ok), {"color": GRIS_CL}))
    if na:
        runs.append(("     No aplican: ", {"bold": True, "color": GRIS}))
        runs.append((", ".join(e["nombre"] for e in na), {"color": GRIS}))
    texto(s, X0 + 0.15, Y_CONT, ANCHO - 0.3, 0.42, [runs], size=9, anchor="m")
    # tarjetas de pendientes: alto según contenido, no estiradas
    n = len(pend)
    cols = n if n <= 3 else 2
    filas_n = (n + cols - 1) // cols
    gap = 0.15
    w = (ANCHO - gap * (cols - 1)) / cols
    y0 = Y_CONT + 0.57

    def alto_tarjeta(e):
        chars = len(e["falta"]) + len(e["necesita"]) + (110 if e["id"] == "personal" else 30)
        return 0.75 + 0.19 * (-(-chars // int((w - 0.4) * 15.5)) + (1 if e["necesita"] else 0))

    mensaje = d.get("mensaje_clave")
    disponible = Y_FIN - y0 - (0.85 if mensaje else 0)
    altos = [max(alto_tarjeta(e) for e in pend[r * cols:(r + 1) * cols]) for r in range(filas_n)]
    escala = min(1.0, (disponible - gap * (filas_n - 1)) / sum(altos)) if altos else 1
    altos = [max(1.25, h * escala) if escala < 1 else max(1.25, h) for h in altos]
    y = y0
    for r in range(filas_n):
        h = altos[r]
        for c, e in enumerate(pend[r * cols:(r + 1) * cols]):
            x = X0 + c * (w + gap)
            caja(s, x, y, w, h, TARJETA, nombre=f"Pendiente {e['nombre']}")
            punto(s, x + 0.18, y + 0.2, 0.2, e["estado"])
            nombre = "Proyecto" if (e["id"] == "proyecto" and cols > 2) else e["nombre"].replace("% ", "")
            texto(s, x + 0.5, y + 0.14, w - 1.9, 0.32, nombre, size=12, bold=True)
            texto(s, x + w - 1.4, y + 0.1, 1.25, 0.4, pc(e["valor"]), size=18, bold=True, align="r", color=SEM[e["estado"]])
            pars = [{"runs": [("Qué falta: ", {"bold": True, "color": CIAN}), (e["falta"], {})], "size": 10, "espacio": 4}]
            necesita = e["necesita"]
            if e["id"] == "personal" and d["personal"]:  # sólo en la presentación: puestos con más vacantes
                necesita += " Principales: " + ", ".join(f"{v['puesto']} {v['faltan']:g}" for v in d["personal"]["vacantes"][:3]) + "."
            if necesita:
                lab = "Detalle: " if e["id"] in ("proyecto", "cummins") else "Qué se necesita: "
                pars.append({"runs": [(lab, {"bold": True, "color": CIAN}), (necesita, {})], "size": 10})
            texto(s, x + 0.2, y + 0.55, w - 0.4, h - 0.65, pars, color=GRIS_CL)
        y += h + gap
    if mensaje:
        caja(s, X0, y, ANCHO, 0.7, AZUL_OSC, nombre="Mensaje clave")
        texto(s, X0 + 0.25, y, ANCHO - 0.5, 0.7, [[("Clave: ", {"bold": True, "color": CIAN}),
                                                 (mensaje, {"bold": True, "color": "FFFFFF"})]], size=13, anchor="m")


def proyecto_grafica(dk, d):
    p = d["proyecto"]
    s = dk.nueva(fondo=FONDO_MAPA, notas=(
        "% por área = suma del avance de sus elementos / número de elementos. Un elemento parcial cuenta su avance "
        "(Adquirido / Requerido). " + " ".join(f"{a['nombre']}: {a['cumple']} de {a['n']} completos, {pc(a['valor'])}." for a in d["areas"])))
    titulo(s, f"{d['corto']}: desarrollo de proyecto por área",
           f"{pc(p['valor'])} · {p['cumple']} de {p['n']} elementos completos · faltan {p['faltan']}")
    areas = d["areas"]
    cd = CategoryChartData()
    cd.categories = [a["nombre"] for a in areas]
    cd.add_series("% de avance", [round(a["valor"], 1) for a in areas])
    alto = min(3.65, 0.5 + 0.47 * len(areas))
    gf = s.shapes.add_chart(XL_CHART_TYPE.BAR_CLUSTERED, Inches(X0), Inches(Y_CONT), Inches(4.7), Inches(alto), cd)
    gf.name = "Gráfica avance por área"
    ch = gf.chart
    estilo_grafica(ch)
    ch.has_title = False
    ch.has_legend = False
    pl = ch.plots[0]
    pl.gap_width = 55
    pl.vary_by_categories = False
    for k, a in enumerate(areas):
        pt = pl.series[0].points[k]
        pt.format.fill.solid()
        pt.format.fill.fore_color.rgb = rgb(SEM[a["estado"]])
    pl.has_data_labels = True
    dl = pl.data_labels
    dl.number_format = '0.0"%"'
    dl.number_format_is_linked = False
    dl.position = XL_LABEL_POSITION.OUTSIDE_END
    dl.font.size = Pt(9)
    dl.font.bold = True
    dl.font.color.rgb = rgb(BLANCO)
    va = ch.value_axis
    va.minimum_scale, va.maximum_scale = 0, 115
    va.visible = False
    va.has_major_gridlines = False
    ca = ch.category_axis
    ca.reverse_order = True
    ca.format.line.fill.background()
    ca.tick_labels.font.size = Pt(9.5)
    ca.tick_labels.font.color.rgb = rgb(BLANCO)
    # tabla de áreas
    filas = [["Área", "Elementos", "Cumple", "Faltan", "%"]]
    for a in areas:
        filas.append([{"t": a["nombre"]}, {"t": str(a["n"]), "align": "c"}, {"t": str(a["cumple"]), "align": "c"},
                      {"t": str(a["faltan"]) if a["faltan"] else "—", "align": "c", "bold": bool(a["faltan"]),
                       "color": SEM["rojo"] if a["faltan"] else GRIS},
                      {"runs": [("● ", {"color": SEM[a["estado"]], "size": 11}), (pc(a["valor"]), {})], "align": "c"}])
    filas.append([{"t": "Total", "bold": True, "fill": AZUL_OSC}, {"t": str(p["n"]), "align": "c", "fill": AZUL_OSC, "bold": True},
                  {"t": str(p["cumple"]), "align": "c", "fill": AZUL_OSC, "bold": True},
                  {"t": str(p["faltan"]), "align": "c", "fill": AZUL_OSC, "bold": True},
                  {"t": pc(p["valor"]), "align": "c", "fill": AZUL_OSC, "bold": True}])
    tabla(s, 5.3, Y_CONT + 0.05, 4.3, filas, [1.45, 0.82, 0.62, 0.58, 0.83], alto_fila=0.29, size=9)
    yb = Y_CONT + 0.05 + 0.29 * (len(filas)) + 0.2
    completas = [a for a in areas if not a["faltan"]]
    txt = []
    if completas:
        txt.append({"runs": [("Completas: ", {"bold": True, "color": SEM["verde"]}),
                             (", ".join(f"{a['nombre']} ({a['n']} de {a['n']})" for a in completas) + ".", {})],
                    "size": 9, "espacio": 3})
    e = d["escenario"]
    if e:
        txt.append({"runs": [("Si el mobiliario actual se valida como OK: ", {"bold": True, "color": CIAN}),
                             (f"proyecto {pc(e['proyecto_hoy'])} → {pc(e['proyecto_validado'])}; general "
                              f"{pc(e['general_hoy'])} → {pc(e['general_validado'])}.", {})], "size": 9})
    if txt:
        texto(s, 5.3, yb, 4.3, Y_FIN - yb, txt, color=GRIS_CL)
    pie(s, "Verde ≥ 80%, amarillo 50–79%, rojo < 50%. Un elemento parcial cuenta su avance (por ejemplo, 1 de 2 = 50%).")


def _alto_item(p, w):
    chars = len(p["titulo"]) + len(p["texto"] or "") + len(p["avance_txt"]) + 6
    por_linea = int(w * 19)
    return 0.075 + 0.14 * max(1, -(-chars // por_linea))


def _alto_area(a, w, cols=1):
    if cols == 1:
        return 0.62 + sum(_alto_item(p, w - 0.3) for p in a["pendientes"])
    mitad = (len(a["pendientes"]) + 1) // 2
    ww = (w - 0.45) / 2
    return 0.62 + max(sum(_alto_item(p, ww) for p in a["pendientes"][:mitad]),
                      sum(_alto_item(p, ww) for p in a["pendientes"][mitad:]))


def bloque_area(s, a, x, y, w, h, cols=1):
    caja(s, x, y, w, h, TARJETA, nombre=f"Área {a['nombre']}")
    punto(s, x + 0.15, y + 0.15, 0.18, a["estado"])
    texto(s, x + 0.42, y + 0.1, w - 1.4, 0.28, a["nombre"], size=12, bold=True)
    texto(s, x + w - 1.0, y + 0.08, 0.85, 0.3, pc(a["valor"]), size=14, bold=True, align="r", color=SEM[a["estado"]])
    texto(s, x + 0.42, y + 0.36, w - 0.6, 0.2, f"{a['n']} elementos · cumple {a['cumple']} · faltan {a['faltan']}",
          size=8.5, color=GRIS)
    items = a["pendientes"]
    grupos = [items] if cols == 1 else [items[: (len(items) + 1) // 2], items[(len(items) + 1) // 2:]]
    ww = (w - 0.3) if cols == 1 else (w - 0.45) / 2
    for c, grupo in enumerate(grupos):
        pars = []
        for p in grupo:
            runs = [(p["titulo"], {"bold": True, "color": BLANCO})]
            if p["parcial"]:
                runs.append((f" ({p['avance_txt']})", {"color": SEM["amarillo"], "bold": True}))
            if p["texto"]:
                runs.append(("  " + p["texto"], {"color": GRIS_CL}))
            pars.append({"runs": runs, "size": 9.5 if cols == 2 else 8.5, "espacio": 6 if cols == 2 else 3.5})
        texto(s, x + 0.15 + c * (ww + 0.15), y + 0.62, ww, h - 0.7, pars)


def areas_detalle(dk, d):
    """Cada área con pendientes es su propio bloque; se acomodan en dos columnas por diapositiva."""
    areas = [a for a in d["areas"] if a["faltan"] and g.norm(a["original"]) != "mobiliario"]
    disponible = Y_FIN - Y_CONT
    w = (ANCHO - 0.15) / 2
    pags, col, cols, usado = [], [], [], 0
    for a in areas:
        h1 = _alto_area(a, w)
        if h1 > disponible:  # área grande: diapositiva propia, ancho completo y dos columnas de pendientes
            if col:
                cols.append(col)
            if cols:
                pags.append(cols)
            pags.append([[("full", a)]])
            col, cols, usado = [], [], 0
            continue
        if usado + h1 > disponible and col:
            cols.append(col)
            col, usado = [], 0
            if len(cols) == 2:
                pags.append(cols)
                cols = []
        col.append(("half", a, h1))
        usado += h1 + 0.12
    if col:
        cols.append(col)
    if cols:
        pags.append(cols)
    # una sola columna con varias áreas: repartir en dos columnas
    for cols in pags:
        if len(cols) == 1 and cols[0][0][0] == "half" and len(cols[0]) >= 2:
            c = cols[0]
            cols[:] = [c[: (len(c) + 1) // 2], c[(len(c) + 1) // 2:]]
    for k, cols in enumerate(pags):
        nombres = [it[1]["nombre"] for c in cols for it in c]
        s = dk.nueva(notas="Pendientes por área: qué falta y qué implica. Entre paréntesis, el avance de los elementos parciales.")
        titulo(s, f"{d['corto']}: pendientes de proyecto" + (f" ({k + 1}/{len(pags)})" if len(pags) > 1 else ""),
               " · ".join(nombres))
        for j, c in enumerate(cols):
            y = Y_CONT
            for it in c:
                if it[0] == "full":
                    bloque_area(s, it[1], X0, y, ANCHO, disponible, cols=2)
                else:
                    x = X0 + j * (w + 0.15)
                    h = it[2]
                    if len(c) == 1:
                        h = max(h, min(disponible, h))
                    bloque_area(s, it[1], x, y, w, h)
                    y += h + 0.12


def mobiliario(dk, d):
    a = next((a for a in d["areas"] if g.norm(a["original"]) == "mobiliario" and a["faltan"]), None)
    if not a:
        return
    e = d["escenario"]
    s = dk.nueva(notas="Mobiliario: elementos del estándar FOTON registrados en 0 en el KPI. El escenario supone que el "
                       "mobiliario actual se valida como OK (cada elemento al 100%).")
    titulo(s, f"{d['corto']}: mobiliario", f"{a['faltan']} de {a['n']} elementos en 0% · área en {pc(a['valor'], 0)}")
    # lista compacta en tres columnas
    cols = 3 if a["faltan"] > 8 else 2
    porcol = -(-a["faltan"] // cols)
    caja(s, X0, Y_CONT, 5.9, Y_FIN - Y_CONT, TARJETA)
    texto(s, X0 + 0.2, Y_CONT + 0.12, 5.5, 0.25, "Elementos registrados en 0%", size=10, bold=True, color=CIAN)
    cw = (5.9 - 0.4) / cols
    for c in range(cols):
        grupo = a["pendientes"][c * porcol:(c + 1) * porcol]
        texto(s, X0 + 0.2 + c * cw, Y_CONT + 0.48, cw - 0.1, 3.2,
              [{"runs": [("✕  ", {"color": SEM["rojo"], "bold": True}), (p["titulo"], {})], "size": 9.5, "espacio": 7}
               for p in grupo], color=GRIS_CL)
    # escenario
    x = X0 + 6.1
    caja(s, x, Y_CONT, 3.1, Y_FIN - Y_CONT, AZUL_OSC)
    texto(s, x + 0.2, Y_CONT + 0.12, 2.7, 0.45, "Si se valida el mobiliario actual como OK", size=10, bold=True, color=CIAN)
    for k, (lab, hoy, val) in enumerate([("Proyecto", e["proyecto_hoy"], e["proyecto_validado"]),
                                         ("General", e["general_hoy"], e["general_validado"])]):
        y = Y_CONT + 0.7 + k * 1.05
        texto(s, x + 0.2, y, 2.7, 0.22, lab, size=9.5, color=GRIS_CL)
        texto(s, x + 0.2, y + 0.22, 2.7, 0.45, [[(pc(hoy), {"color": GRIS_CL, "size": 15}), ("  →  ", {"color": GRIS, "size": 13}),
                                                 (pc(val), {"bold": True, "size": 20})]], color=BLANCO)
        texto(s, x + 0.2, y + 0.68, 2.7, 0.2, f"+{val - hoy:.1f} puntos", size=9, color=CIAN, bold=True)
    texto(s, x + 0.2, Y_CONT + 2.85, 2.7, 0.85,
          "Revisar en la visita qué muebles cumplen el estándar. Lo que no cumpla se cotiza y entra al plan." +
          (" Aplica sólo si se acepta el mobiliario de la oficina móvil." if d.get("mensaje_clave") else ""),
          size=8.5, color=GRIS_CL)


def personal(dk, d):
    pe = d["personal"]
    if not pe or not pe["faltan"]:
        return
    s = dk.nueva(notas=f"Fuente: Sistema de Personal FOTON (noviembre). Requeridas {pe['req']:g}, cubiertas {pe['cub']:g}, "
                       f"faltan {pe['faltan']:g}, extra {pe['extra']:g}. El % es cubiertas / requeridas.")
    titulo(s, f"{d['corto']}: personal", f"{pe['cub']:g} de {pe['req']:g} plazas cubiertas ({pc(pe['valor'])}) · "
                                         f"{pe['faltan']:g} vacantes")
    areas = pe["areas"]
    cd = CategoryChartData()
    cd.categories = [a["area"] for a in areas]
    cd.add_series("Requeridas", [a["req"] for a in areas])
    cd.add_series("Cubiertas", [a["cub"] for a in areas])
    gf = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(X0), Inches(Y_CONT), Inches(4.3), Inches(3.7), cd)
    gf.name = "Gráfica plazas por área"
    ch = gf.chart
    estilo_grafica(ch)
    ch.has_legend = True
    ch.legend.position = XL_LEGEND_POSITION.TOP
    ch.legend.include_in_layout = False
    ch.legend.font.color.rgb = rgb(GRIS_CL)
    for ser, col in zip(ch.plots[0].series, ["5B6B85", AZUL]):
        ser.format.fill.solid()
        ser.format.fill.fore_color.rgb = rgb(col)
    pl = ch.plots[0]
    pl.gap_width = 70
    pl.overlap = -10
    pl.has_data_labels = True
    pl.data_labels.position = XL_LABEL_POSITION.OUTSIDE_END
    pl.data_labels.font.size = Pt(9)
    pl.data_labels.font.bold = True
    pl.data_labels.font.color.rgb = rgb(BLANCO)
    ch.value_axis.visible = False
    ch.value_axis.has_major_gridlines = False
    ch.category_axis.format.line.color.rgb = rgb(LINEA)
    ch.category_axis.tick_labels.font.color.rgb = rgb(BLANCO)
    # vacantes por puesto
    x = X0 + 4.55
    filas = [["Puesto", "Área", "Vacantes"]]
    for v in pe["vacantes"]:
        filas.append([{"t": v["puesto"]}, {"t": v["area"], "color": GRIS_CL}, {"t": f"{v['faltan']:g}", "align": "c", "bold": True}])
    alto = min(0.25, (Y_FIN - Y_CONT - 0.3) / len(filas))
    tabla(s, x, Y_CONT, 4.65, filas, [2.65, 1.15, 0.85], alto_fila=alto, size=8.5 if len(filas) > 10 else 9)


def cummins(dk, d):
    e = next((e for e in d["resumen"] if e["id"] == "cummins"), None)
    if not e or len(e["items"]) < 3:
        return
    s = dk.nueva(notas="Requisitos Cummins del KPI. El % es el promedio de los 7 requisitos; la certificación de técnicos "
                       "pide 2 personas (1 de 2 = 50%).")
    titulo(s, f"{d['corto']}: Cummins", f"{pc(e['valor'])} · {e['falta'].rstrip('.')}")
    n = len(e["items"])
    alto = (Y_FIN - Y_CONT - 0.95) / n
    for k, it in enumerate(e["items"]):
        y = Y_CONT + k * alto
        caja(s, X0, y, 5.9, alto - 0.07, TARJETA)
        caja(s, X0 + 0.12, y + (alto - 0.07 - 0.3) / 2, 0.3, 0.3, AZUL_OSC, radio=0, forma=MSO_SHAPE.OVAL)
        texto(s, X0 + 0.12, y + (alto - 0.07 - 0.3) / 2, 0.3, 0.3, str(k + 1), size=9, bold=True, align="c", anchor="m")
        texto(s, X0 + 0.55, y, 5.2, alto - 0.07, [[(it["titulo"], {})] + ([(f"  ({it['avance_txt']})", {"color": SEM["amarillo"]})]
                                                                      if it["avance_txt"] else [])],
              size=10, anchor="m")
    obs = [o.get("texto_presentacion") or o["texto"] for o in d["observaciones"] if "Cummins" in o["texto"]]
    y = Y_FIN - 0.85
    if obs:
        caja(s, X0, y, 5.9, 0.85, AZUL_OSC)
        texto(s, X0 + 0.15, y + 0.08, 5.6, 0.7, [[("Para validar: ", {"bold": True, "color": CIAN}), (obs[0], {})]],
              size=9, color=GRIS_CL, anchor="m")
    # orden sugerido
    x = X0 + 6.1
    caja(s, x, Y_CONT, 3.1, Y_FIN - Y_CONT, TARJETA)
    texto(s, x + 0.2, Y_CONT + 0.12, 2.7, 3.5, [
        {"runs": [("Orden sugerido", {"bold": True, "color": CIAN})], "size": 10, "espacio": 8},
        {"runs": [("1. ", {"bold": True}), ("Contrato, línea de crédito y cuenta QSOL: son trámites y abren lo demás.", {})],
         "size": 9, "espacio": 6},
        {"runs": [("2. ", {"bold": True}), ("Inversión en herramienta y partes.", {})], "size": 9, "espacio": 6},
        {"runs": [("3. ", {"bold": True}), ("Certificar al analista de garantías y a 2 técnicos o jefes.", {})], "size": 9, "espacio": 10},
        {"runs": [("Pedir en la visita fecha compromiso para cada paso.", {"color": GRIS_CL})], "size": 9}], color=BLANCO)
    cumplidos = len([c for c in d["indicadores"] if c["id"] == "cummins"][0]["conceptos"]) - n
    texto(s, x + 0.2, Y_FIN - 1.05, 2.7, 0.5, [[(f"{cumplidos} de {cumplidos + n}", {"size": 26, "bold": True, "font": "Arial Black"})]],
          color=SEM[e["estado"]])
    texto(s, x + 0.2, Y_FIN - 0.45, 2.7, 0.3, "requisitos cumplidos", size=9, color=GRIS_CL)


def evidencia(dk, d, fotos, ctx):
    if not fotos:
        return
    s = dk.nueva(fondo=FONDO_MAPA, notas="Fotos de la visita más reciente disponible. Fotos con más de 6 meses se marcan "
                                         "como no recientes.")
    titulo(s, f"{d['corto']}: evidencia", "Lo que se ve en la agencia")
    n = len(fotos)
    gap = 0.15
    w = (ANCHO - gap * (n - 1)) / n
    h = min(2.75, w * 0.75)
    for k, (ruta, pie_txt, estado, no_rec) in enumerate(fotos):
        x = X0 + k * (w + gap)
        dk.imagen(s, dk.foto(ruta, w, h), x, Y_CONT, w, h, f"Foto {k + 1}", pie_txt)
        if no_rec:
            caja(s, x + 0.08, Y_CONT + 0.08, 1.25, 0.26, "000000", radio=0.05, alpha=0.6)
            texto(s, x + 0.08, Y_CONT + 0.08, 1.25, 0.26, "Foto no reciente", size=8, bold=True, align="c", anchor="m", color="FFFFFF")
        punto(s, x, Y_CONT + h + 0.14, 0.14, estado)
        texto(s, x + 0.22, Y_CONT + h + 0.1, w - 0.22, 0.75, pie_txt, size=9.5, color=BLANCO)


def validar(dk, ctx):
    s = dk.nueva(notas="Puntos donde lo que se ve en campo y el KPI no coinciden, o falta información. Confirmar en la visita.")
    titulo(s, "Puntos a validar en la visita", "Lo que se ve en campo no coincide con el KPI o falta información")
    items = [(d["corto"], o.get("texto_presentacion") or o["texto"]) for d in ctx["dists"] for o in d["observaciones"]]
    n = len(items)
    alto = (Y_FIN - Y_CONT) / n
    for k, (ag, t) in enumerate(items):
        y = Y_CONT + k * alto
        caja(s, X0, y, ANCHO, alto - 0.1, TARJETA)
        texto(s, X0 + 0.2, y, 1.4, alto - 0.1, ag, size=11, bold=True, color=CIAN, anchor="m")
        texto(s, X0 + 1.6, y, ANCHO - 1.8, alto - 0.1, t, size=10, color=GRIS_CL, anchor="m")


def _alto_fila_plan(p):
    lineas = max(-(-len(p["falta"]) // 85), -(-len(p["pendiente"]) // 20), 1)
    return max(0.34, 0.115 * lineas + 0.13)


def plan(dk, d):
    disponible = Y_FIN - Y_CONT - 0.32
    pags, act, usado = [], [], 0
    for p in d["plan"]:
        h = _alto_fila_plan(p)
        if act and usado + h > disponible:
            pags.append(act)
            act, usado = [], 0
        act.append(p)
        usado += h
    if act:
        pags.append(act)
    for k, grupo in enumerate(pags):
        s = dk.nueva(notas="Plan de trabajo agrupado por área o indicador. Todo lo que no está al 100% es prioridad alta. "
                           "Responsable, fecha compromiso, estatus y comentarios se llenan con el gerente en la visita. "
                           "El detalle elemento por elemento está en la herramienta HTML.")
        titulo(s, f"{d['corto']}: plan de trabajo" + (f" ({k + 1}/{len(pags)})" if len(pags) > 1 else ""),
               "Llenar responsable, fecha y estatus con el gerente")
        filas = [["Pendiente", "Qué falta", "Prioridad", "Responsable", "Fecha compromiso", "Estatus", "Comentarios"]]
        for p in grupo:
            filas.append([{"t": p["pendiente"], "bold": True}, {"t": p["falta"], "size": 7.5, "color": GRIS_CL},
                          {"t": p["prioridad"], "bold": True, "color": SEM["rojo"], "align": "c"}, "", "", "", ""])
        shp = tabla(s, X0, Y_CONT, ANCHO, filas, [1.45, 3.35, 0.68, 1.0, 0.85, 0.75, 1.12], alto_fila=0.4, size=8.5, alto_enc=0.32)
        for i, p in enumerate(grupo, 1):
            shp.table.rows[i].height = Inches(_alto_fila_plan(p))


def cierre(dk, ctx):
    s = dk.nueva(fondo=None, notas="Siguientes pasos después de las visitas.")
    dk.imagen(s, dk.jpg(FOTO_CIERRE), 0, 0, 10, 5.625, "Foto cierre")
    caja(s, 0, 0, 5.6, 5.625, "0B162B", radio=0, alpha=0.72)
    dk.logo(s)
    texto(s, 0.6, 1.35, 4.8, 0.5, "Siguientes pasos", size=24, bold=True, color="FFFFFF")
    pasos = ["Revisar el plan con el gerente de cada agencia.",
             "Registrar responsable, fecha compromiso y estatus en la herramienta.",
             "Validar mobiliario, tótem de Escobedo y Cummins de Apodaca en campo.",
             "Allende: definir el proyecto de la agencia.",
             "Nuevo corte mensual y revisión a 30 días."]
    texto(s, 0.6, 2.0, 4.7, 3.0, [{"runs": [(f"{k + 1}.  ", {"bold": True, "color": CIAN}), (p, {})], "size": 11, "espacio": 7}
                                  for k, p in enumerate(pasos)], color=BLANCO)


def evidencia_validar(dk, ctx, fotos):
    """Una sola lámina: fotos de lo pendiente o por validar y los puntos que hay que confirmar en la visita."""
    s = dk.nueva(fondo=FONDO_MAPA, notas="Fotos de la visita más reciente disponible (más de 6 meses: «Foto no reciente»). "
                                         "Abajo, lo que se ve en campo y no coincide con el KPI o falta información.")
    titulo(s, "Evidencia y puntos a validar", "Lo que se ve en campo y lo que hay que confirmar en la visita")
    n = len(fotos)
    gap = 0.15
    w = (ANCHO - gap * (n - 1)) / n
    h = 1.55
    for k, (ruta, pie_txt, estado, no_rec) in enumerate(fotos):
        x = X0 + k * (w + gap)
        dk.imagen(s, dk.foto(ruta, w, h), x, Y_CONT, w, h, f"Foto {k + 1}", pie_txt)
        if no_rec:
            caja(s, x + 0.06, Y_CONT + 0.06, 1.15, 0.24, "000000", radio=0.05, alpha=0.6)
            texto(s, x + 0.06, Y_CONT + 0.06, 1.15, 0.24, "Foto no reciente", size=7.5, bold=True, align="c", anchor="m", color="FFFFFF")
        punto(s, x, Y_CONT + h + 0.1, 0.13, estado)
        texto(s, x + 0.2, Y_CONT + h + 0.06, w - 0.2, 0.45, pie_txt, size=8.5, color=BLANCO)
    items = [(d["corto"], o.get("texto_presentacion") or o["texto"]) for d in ctx["dists"] for o in d["observaciones"]]
    y0 = Y_CONT + h + 0.6
    cols = 2
    filas_n = -(-len(items) // cols)
    wc = (ANCHO - 0.15) / cols
    hc = (Y_FIN - y0 - 0.1 * (filas_n - 1)) / max(filas_n, 1)
    for k, (ag, t) in enumerate(items):
        x = X0 + (k % cols) * (wc + 0.15)
        y = y0 + (k // cols) * (hc + 0.1)
        caja(s, x, y, wc, hc, TARJETA)
        texto(s, x + 0.15, y, wc - 0.3, hc, [[(ag + ". ", {"bold": True, "color": CIAN}), (t, {})]], size=9, color=GRIS_CL, anchor="m")


EVIDENCIA_GRUPO = [
    ("apodaca/WhatsApp Image 2026-07-02 at 6.14.58 PM (1).jpeg", "Apodaca. Taller: piso de concreto, falta pintura epóxica.", "rojo"),
    ("apodaca/WhatsApp Image 2026-07-02 at 6.14.53 PM.jpeg", "Apodaca. Acceso: lona temporal en lugar del tótem.", "rojo"),
    ("escobedo/WhatsApp Image 2026-08-21 at 9.51.06 AM (1).jpeg", "Escobedo. Tótem en el acceso: validar si cumple.", "amarillo"),
    ("escobedo/WhatsApp Image 2026-07-02 at 10.25.28 AM (5).jpeg", "Escobedo. Mobiliario actual: validar estándar.", "amarillo"),
    ("allende/WhatsApp Image 2025-06-27 at 12.56.02 PM (2).jpeg", "Allende. Opera en oficina móvil, sin imagen FOTON.", "rojo"),
]


# =========================================================================== armado
EVIDENCIAS = {
    "apodaca": [
        ("apodaca/WhatsApp Image 2026-07-02 at 6.14.58 PM (1).jpeg", "Taller: piso de concreto, falta pintura epóxica.", "rojo"),
        ("apodaca/WhatsApp Image 2026-07-02 at 6.14.53 PM.jpeg", "Acceso: lona temporal en lugar del tótem.", "rojo"),
        ("apodaca/WhatsApp Image 2026-07-02 at 6.14.55 PM (1).jpeg", "Letrero Total Care: ya instalado.", "verde"),
    ],
    "escobedo": [
        ("escobedo/WhatsApp Image 2026-08-21 at 9.51.07 AM.jpeg", "Estacionamiento: faltan cajones reservados para clientes.", "rojo"),
        ("escobedo/WhatsApp Image 2026-08-21 at 9.51.06 AM (1).jpeg", "Tótem en el acceso: el KPI lo marca en 0. Validar si cumple.", "amarillo"),
        ("escobedo/WhatsApp Image 2026-07-02 at 10.25.28 AM (5).jpeg", "Mobiliario actual: validar si cumple el estándar.", "amarillo"),
    ],
    "allende": [
        ("allende/WhatsApp Image 2025-06-27 at 12.56.02 PM (2).jpeg", "Opera en oficina móvil: sin fachada ni imagen FOTON.", "rojo"),
        ("allende/WhatsApp Image 2025-06-27 at 12.56.01 PM (1).jpeg", "Unidades exhibidas a la intemperie, sin área techada.", "rojo"),
    ],
}
FACHADAS = {
    "apodaca": ("fotos/editadas/apodaca_fachada_editada.jpg", None),
    "escobedo": ("fotos/editadas/escobedo_fachada_editada.jpg", None),
    "allende": ("datos/fotos/allende/WhatsApp Image 2025-06-27 at 12.56.02 PM (2).jpeg", "Foto no reciente"),
}


def _fotos(lista, ctx):
    base = os.path.join(RAIZ, "datos", "fotos")
    limite = ctx["corte"].toordinal() - round(ctx["cfg"]["meses_foto_reciente"] * 30.44)
    out = []
    for arch, pie_txt, est in lista:
        r = os.path.join(base, arch)
        if os.path.exists(r):
            f, _ = g.fecha_foto(r, {"archivo": arch})
            out.append((r, pie_txt, est, f.toordinal() < limite))
    return out


def armar(salida, completa=False):
    """Versión ejecutiva (25 láminas) o, con completa=True, la versión extendida con mobiliario, Cummins,
    personal y evidencia por agencia."""
    ctx = A.cargar()
    dk = Deck(TEMPLATE)
    portada(dk, ctx)
    resumen_ejecutivo(dk, ctx)
    matriz(dk, ctx)
    if completa:
        concentracion(dk, ctx)
    escenario_mobiliario(dk, ctx)
    for d in ctx["dists"]:
        ruta_fachada, nota = FACHADAS.get(d["id"], (None, None))
        if ruta_fachada and os.path.exists(os.path.join(RAIZ, ruta_fachada)):
            seccion(dk, d, os.path.join(RAIZ, ruta_fachada), nota)
        indicadores(dk, d)
        proyecto_grafica(dk, d)
        areas_detalle(dk, d)
        if completa:
            mobiliario(dk, d)
            cummins(dk, d)
            personal(dk, d)
            evidencia(dk, d, _fotos(EVIDENCIAS.get(d["id"], []), ctx), ctx)
    if completa:
        validar(dk, ctx)
    else:
        evidencia_validar(dk, ctx, _fotos(EVIDENCIA_GRUPO, ctx))
    for d in ctx["dists"]:
        plan(dk, d)
    cierre(dk, ctx)
    dk.terminar(salida)
    print(f"  -> {os.path.relpath(salida, RAIZ)}: {len(dk.prs.slides)} diapositivas, "
          f"{os.path.getsize(salida) / 1e6:.1f} MB")


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--salida", default=None)
    ap.add_argument("--completa", action="store_true",
                    help="versión extendida (mobiliario, Cummins, personal y evidencia por agencia)")
    a = ap.parse_args()
    nombre = "Grupo_Velcen_Estatus_nov2026_completa.pptx" if a.completa else "Grupo_Velcen_Estatus_nov2026.pptx"
    armar(a.salida or os.path.join(RAIZ, "presentacion", nombre), a.completa)


if __name__ == "__main__":
    main()
