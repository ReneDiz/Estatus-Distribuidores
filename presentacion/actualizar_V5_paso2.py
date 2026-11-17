"""Paso 1: Escobedo personal (igual que Apodaca) y renglón Generales en el plan de Allende."""
import copy, sys
sys.path.insert(0, "/home/user/Estatus-Distribuidores/presentacion")
from pptx import Presentation
from pptx.util import Emu, Inches
import generar_presentacion as P

SRC, DST = sys.argv[1], sys.argv[2]
prs = Presentation(SRC)
S = lambda n: prs.slides[n - 1]


def caja_txt(s, pred):
    hits = [sh for sh in s.shapes if sh.has_text_frame and pred(sh)]
    assert len(hits) == 1, len(hits)
    return hits[0]


def poner(sh, texto):
    p = sh.text_frame.paragraphs[0]
    p.runs[0].text = texto
    for r in p.runs[1:]:
        r._r.getparent().remove(r._r)


y = lambda sh: Emu(sh.top).inches
x = lambda sh: Emu(sh.left).inches
s = S(17)
# posiciones con el formato de la imagen (área comercial con prefijo «Comercial:», cantidad «X de Y»)
poner(caja_txt(s, lambda sh: sh.text_frame.text == "Asesor de ventas HDT"), "Comercial: Asesor de ventas HDT")
poner(caja_txt(s, lambda sh: sh.text_frame.text == "Asesor digital (Biky)"), "Comercial: Asesor digital (Biky)")
for pred, txt in ((lambda sh: sh.text_frame.text == "2 de 2", "2 de 2"),
                  (lambda sh: sh.text_frame.text == "1" and x(sh) < 3 and abs(y(sh) - 3.62) < 0.05, "1 de 1"),
                  (lambda sh: sh.text_frame.text == "1" and x(sh) > 5, "1 de 1")):
    sh = caja_txt(s, pred)
    poner(sh, txt)
    sh.left = sh.left - Inches(0.11)
    sh.width = Inches(0.41)
# cuadro rojo y nota
poner(caja_txt(s, lambda sh: sh.text_frame.text == "3" and x(sh) > 6), "10%")
poner(caja_txt(s, lambda sh: sh.text_frame.text == "vacantes en ventas"), "personal faltante  ·  4 vacantes")
P.texto(s, 6.25, 3.5, 3.35, 0.5, "Personal reportado por Liliana Ozoria el jueves 12 de noviembre.", size=10,
        color=P.GRIS_CL, nombre="Nota reporte de personal")

# plan de Allende: renglón Generales (copia del renglón Mobiliario)
s = S(25)
tbl = next(sh for sh in s.shapes if sh.has_table).table
tr = tbl._tbl.tr_lst[-1]
nuevo = copy.deepcopy(tr)
tr.addnext(nuevo)
fila = tbl.rows[len(tbl.rows) - 1]
txts = ["Generales (11 de 11)",
        "BP firmado, Sistema CRM, PLD, Línea de plan piso, Seguro de plan piso, Pedido inicial de unidades "
        "(3 meses de venta), Capacitación comercial y de producto, Capacitación técnica, Capacitación de "
        "procesos de posventa, Estacionamiento de clientes, Cajones reservados."]
for j, t in enumerate(txts):
    p = fila.cells[j].text_frame.paragraphs[0]
    p.runs[0].text = t
    for r in p.runs[1:]:
        r._r.getparent().remove(r._r)
prs.save(DST)
print("slide 21 part:", prs.slides[20].part.partname)
