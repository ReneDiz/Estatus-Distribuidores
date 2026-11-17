"""Paso 2: la copia de la lámina 21 se convierte en «Allende: pendientes de proyecto (3/3) · Generales»."""
import sys
sys.path.insert(0, "/home/user/Estatus-Distribuidores/presentacion")
from pptx import Presentation
from pptx.util import Emu
import generar_presentacion as P

SRC, DST = sys.argv[1], sys.argv[2]
prs = Presentation(SRC)
s20, s21, s22 = prs.slides[19], prs.slides[20], prs.slides[21]
assert s21.shapes[0].text_frame.text.startswith("Allende: pendientes de proyecto (2/2)") or any(
    sh.has_text_frame and sh.text_frame.text.startswith("Allende: pendientes de proyecto (2/2)") for sh in s21.shapes)


def cambia(s, viejo, nuevo):
    for sh in s.shapes:
        if sh.has_text_frame:
            for p in sh.text_frame.paragraphs:
                for r in p.runs:
                    if viejo in r.text:
                        r.text = r.text.replace(viejo, nuevo)
                        return
    raise AssertionError(viejo)


cambia(s20, "(1/2)", "(1/3)")
cambia(s21, "(2/2)", "(2/3)")
cambia(s22, "(2/2)", "(3/3)")
# subtítulo y bloque de la lámina nueva
for sh in list(s22.shapes):
    if Emu(sh.top).inches > 1.45:
        sh._element.getparent().remove(sh._element)
for sh in s22.shapes:
    if sh.has_text_frame and sh.text_frame.text.strip() == "Showroom":
        sh.text_frame.paragraphs[0].runs[0].text = "Generales"
        for r in sh.text_frame.paragraphs[0].runs[1:]:
            r.text = ""
titulos = ["BP firmado", "Sistema CRM", "PLD", "Línea de plan piso", "Seguro de plan piso",
           "Pedido inicial de unidades (3 meses de venta)", "Capacitación comercial y de producto",
           "Capacitación técnica", "Capacitación de procesos de posventa", "Estacionamiento de clientes",
           "Cajones reservados"]
area = {"nombre": "Generales", "n": 11, "cumple": 0, "faltan": 11, "valor": 0.0, "estado": "rojo",
        "pendientes": [{"titulo": t, "texto": None, "parcial": False, "avance_txt": ""} for t in titulos]}
h = P._alto_area(area, P.ANCHO, 3)
P.bloque_area(s22, area, P.X0, P.Y_CONT, P.ANCHO, h, cols=3)
prs.save(DST)
print("lámina nueva: 22; alto del bloque %.2f in" % h)
