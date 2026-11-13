"""Actualiza las cifras de personal (Apodaca y Escobedo) en la presentación V3 del usuario,
sin tocar el resto de sus textos."""
import sys
sys.path.insert(0, "/home/user/Estatus-Distribuidores/presentacion")
sys.path.insert(0, "/home/user/Estatus-Distribuidores")
from pptx import Presentation
from pptx.util import Emu
from pptx.dml.color import RGBColor
import generar_presentacion as P

SRC, DST = sys.argv[1], sys.argv[2]
prs = Presentation(SRC)
S = lambda n: prs.slides[n - 1]
VERDE, AMARILLO = RGBColor.from_string("2FBF71"), RGBColor.from_string("F2B544")
hechos = []


def runs(slide):
    for sh in slide.shapes:
        if sh.has_text_frame:
            for p in sh.text_frame.paragraphs:
                for r in p.runs:
                    yield sh, p, r
        if sh.has_table:
            for row in sh.table.rows:
                for c in row.cells:
                    for p in c.text_frame.paragraphs:
                        for r in p.runs:
                            yield c, p, r


def exacto(n, viejo, nuevo, veces=1, color=None):
    k = 0
    for _, p, r in runs(S(n)):
        if r.text == viejo and k < veces:
            r.text = nuevo
            if color is not None:
                r.font.color.rgb = color
            k += 1
    assert k == veces, (n, viejo, k)
    hechos.append(f"{n}: {viejo} -> {nuevo}")


def sub(n, viejo, nuevo):
    k = 0
    for _, p, r in runs(S(n)):
        if viejo in r.text:
            r.text = r.text.replace(viejo, nuevo)
            k += 1
    if k == 0:  # texto repartido en varios fragmentos: se reemplaza en el párrafo, conservando el formato de cada fragmento
        vistos = set()
        for _, p, r in runs(S(n)):
            if p._p in vistos:
                continue
            vistos.add(p._p)
            completo = "".join(x.text for x in p.runs)
            if viejo in completo:
                ini = completo.index(viejo)
                pos = 0
                for x in p.runs:
                    a0, a1 = pos, pos + len(x.text)
                    pos = a1
                    t = x.text
                    # parte del fragmento que cae dentro del texto a reemplazar
                    i0, i1 = max(a0, ini), min(a1, ini + len(viejo))
                    if i0 < i1:
                        rel0, rel1 = i0 - a0, i1 - a0
                        x.text = t[:rel0] + (nuevo if i0 == ini else "") + t[rel1:]
                k += 1
    assert k == 1, (n, viejo, k)
    hechos.append(f"{n}: «{viejo}» -> «{nuevo}»")


def celda(n, viejo_inicio, nuevo):
    for sh in S(n).shapes:
        if sh.has_table:
            for row in sh.table.rows:
                for c in row.cells:
                    if c.text_frame.text.startswith(viejo_inicio):
                        p = c.text_frame.paragraphs[0]
                        p.runs[0].text = nuevo
                        for r in p.runs[1:]:
                            r.text = ""
                        return
    raise AssertionError((n, viejo_inicio))


# ---------------------------------------------------------------- 2 resumen ejecutivo
exacto(2, "73.5%", "76.2%")
exacto(2, "92.2%", "95.4%")
exacto(2, "83.5%", "85.9%")
exacto(2, "70", "63")
exacto(2, "47", "40")

# ---------------------------------------------------------------- 3 matriz
tabla = next(sh for sh in S(3).shapes if sh.has_table).table
for row in tabla.rows:
    nombre = row.cells[0].text_frame.text.strip()
    if nombre == "Personal":
        for j, val in ((1, "85.0%"), (2, "90.0%")):
            rs = row.cells[j].text_frame.paragraphs[0].runs
            rs[0].font.color.rgb = VERDE
            rs[1].text = val
            rs[1].font.bold = False
    if nombre == "Calificación general":
        row.cells[1].text_frame.paragraphs[0].runs[1].text = "76.2%"
        row.cells[2].text_frame.paragraphs[0].runs[1].text = "95.4%"
exacto(3, " 8 de 11. ", " 9 de 11. ")   # Escobedo (primero, para no tocar el de Apodaca ya cambiado)
exacto(3, " 7 de 11. ", " 8 de 11. ")   # Apodaca

# ---------------------------------------------------------------- 4 mobiliario
exacto(4, "73.5%", "76.2%")
exacto(4, "75.3%", "78.0%")
exacto(4, "92.2%", "95.4%")
exacto(4, "94.0%", "97.2%")
exacto(4, "83.5%", "85.9%")
exacto(4, "85.5%", "87.9%")


# ---------------------------------------------------------------- 5 y 12 agencia
def tile_personal(n, valor):
    """La tarjeta de Personal es la última de la fila de indicadores (x > 8.9 in)."""
    s = S(n)
    for sh in s.shapes:
        x, y = Emu(sh.left).inches, Emu(sh.top).inches
        if x > 8.85 and y > 3.9:
            if sh.has_text_frame and sh.text_frame.text in ("55.0%",):
                sh.text_frame.paragraphs[0].runs[0].text = valor
            elif sh.has_text_frame and sh.text_frame.text == "En proceso":
                r = sh.text_frame.paragraphs[0].runs[0]
                r.text = "Cumple"
                r.font.color.rgb = VERDE
            elif not sh.has_text_frame or sh.text_frame.text == "":
                if Emu(sh.height).inches < 0.06:  # franja de color de la tarjeta
                    sh.fill.solid()
                    sh.fill.fore_color.rgb = VERDE
    hechos.append(f"{n}: tarjeta Personal -> {valor} Cumple")


exacto(5, "73.5%", "76.2%")
sub(5, "70 pendientes", "63 pendientes")
tile_personal(5, "85.0%")
exacto(12, "92.2%", "95.4%")
sub(12, "47 pendientes", "40 pendientes")
tile_personal(12, "90.0%")

# ---------------------------------------------------------------- 6 y 13 proyecto por área
sub(6, "general 73.5% → 75.3%", "general 76.2% → 78.0%")
sub(13, "general 92.2% → 94.0%", "general 95.4% → 97.2%")


# ---------------------------------------------------------------- 11 y 17 personal (se redibuja)
def personal(n, d):
    s = S(n)
    for sh in list(s.shapes):
        sh._element.getparent().remove(sh._element)
    X0 = P.X0
    reglas = {"verde_min": 80, "amarillo_min": 50}
    P.texto(s, X0, 0.88, 6, 0.22, d["tag"], size=9.5, bold=True, color=P.CIAN)
    P.texto(s, X0, 1.1, 9.2, 0.45, d["titulo"], size=22, bold=True, color="FFFFFF")
    xi, wi = X0, 5.55
    P.texto(s, xi, 1.68, wi, 0.28, [[("Personal  ·  ", {"bold": True}), (d["resumen"], {"bold": True})]], size=12, color="FFFFFF")
    y = 2.02
    for area, cub, req, peor in d["areas"]:
        pct = 100 * cub / req
        est = P.g.semaforo(pct, reglas)
        P.texto(s, xi, y, 1.3, 0.22, area, size=9.5, bold=peor, color="FFFFFF" if peor else P.GRIS_CL, anchor="m")
        P.barra(s, xi + 1.35, y + 0.06, 3.45, 0.1, pct, est)
        P.texto(s, xi + 4.85, y, 0.7, 0.22, f"{cub} / {req}", size=9.5, bold=True, align="r", anchor="m")
        y += 0.29
    y += 0.08
    total = sum(v for _, v in d["vacantes"])
    P.texto(s, xi, y, wi, 0.22, f"Vacantes por puesto ({total} plazas)", size=9.5, bold=True, color=P.SEM["rojo"])
    y += 0.3
    vac = d["vacantes"]
    mitad = -(-len(vac) // 2)
    cw = (wi - 0.12) / 2
    alto = 0.3
    for i, (puesto, n_) in enumerate(vac):
        x = xi + (i // mitad) * (cw + 0.12)
        yy = y + (i % mitad) * (alto + 0.05)
        P.caja(s, x, yy, cw, alto, P.TARJETA, radio=0.03)
        P.texto(s, x + 0.1, yy, cw - 0.5, alto, puesto, size=8.5, color=P.GRIS_CL, anchor="m")
        P.texto(s, x + cw - 0.4, yy, 0.3, alto, str(n_), size=9, bold=True, align="r", anchor="m")
    xd, wd = X0 + 5.85, 3.35
    P.caja(s, xd, 1.7, wd, 1.65, "C0392B", radio=0.06, nombre="Vacantes del área más baja")
    P.texto(s, xd + 0.25, 1.78, wd - 0.4, 0.95, str(d["vac_area"]), size=48, bold=True, font="Arial Black", color="FFFFFF", anchor="m")
    P.texto(s, xd + 0.25, 2.8, wd - 0.4, 0.35, d["vac_label"], size=13, color="FFFFFF")
    P.caja(s, xd, 3.5, wd, 1.2, P.TARJETA, radio=0.06, nombre="Personas extra")
    P.texto(s, xd + 0.22, 3.6, wd - 0.4, 1.0, [
        {"runs": [(d["extra"], {"bold": True, "color": "FFFFFF"})], "size": 13, "espacio": 4},
        {"runs": [(d["extra_txt"], {})], "size": 9.5, "color": P.GRIS_CL}])
    hechos.append(f"{n}: lámina de personal redibujada")


personal(11, {
    "tag": "FOTON APODACA  ·  PERSONAL",
    "titulo": "Administrativo: 7 de 9 plazas cubiertas",
    "resumen": "34 de 40 plazas (85%)",
    "areas": [("Administrativo", 7, 9, True), ("Comercial", 12, 14, False), ("Postventa", 15, 17, False)],
    "vacantes": [("Asesor de ventas HDT", 2), ("Recepcionista", 1), ("Marketing", 1),
                 ("Asesor de servicio", 1), ("Técnico VE (eléctricos)", 1)],
    "vac_area": 2, "vac_label": "vacantes en administrativo",
    "extra": "4 personas extra",
    "extra_txt": "El reporte actual no trae el detalle por puesto. Revisar si alguna puede cubrir vacantes.",
})
personal(17, {
    "tag": "FOTON ESCOBEDO  ·  PERSONAL",
    "titulo": "Ventas: 11 de 14 plazas cubiertas",
    "resumen": "36 de 40 plazas (90%)",
    "areas": [("Administrativo", 8, 9, False), ("Comercial", 11, 14, True), ("Postventa", 17, 17, False)],
    "vacantes": [("Asesor de ventas HDT", 2), ("Marketing", 1), ("Asesor digital (Biky)", 1)],
    "vac_area": 3, "vac_label": "vacantes en ventas",
    "extra": "17 personas extra",
    "extra_txt": "Caja (4), asesores Pick Ups, Vanes y MT (2), asesor de servicio (1) y técnicos A, B y C (10).",
})

# ---------------------------------------------------------------- 23 y 24 plan de trabajo
celda(23, "Personal (18 vacantes)", "Personal (6 vacantes)")
celda(23, "Asesor de ventas L&MD 4", "Asesor de ventas HDT 2, Recepcionista 1, Marketing 1, Asesor de servicio 1, Técnico VE (eléctricos) 1.")
celda(24, "Personal (18 vacantes)", "Personal (4 vacantes)")
celda(24, "Asesor de ventas Pick Ups", "Asesor de ventas HDT 2, Marketing 1, Asesor digital (Biky) 1.")

prs.save(DST)
print("\n".join(hechos))
