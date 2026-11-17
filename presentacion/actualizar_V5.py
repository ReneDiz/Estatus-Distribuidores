"""V5: cambios pedidos sobre la V4 final del usuario. Edita el mismo archivo; no rehace láminas."""
import copy
import sys
sys.path.insert(0, "/home/user/Estatus-Distribuidores/presentacion")
from pptx import Presentation
from pptx.util import Emu, Inches, Pt
from pptx.dml.color import RGBColor
from pptx.chart.data import CategoryChartData
import generar_presentacion as P

SRC, DST = sys.argv[1], sys.argv[2]
prs = Presentation(SRC)
S = lambda n: prs.slides[n - 1]
ROJO, AMARILLO, GRIS = RGBColor.from_string("E5484D"), RGBColor.from_string("F2B544"), RGBColor.from_string("A9B3B7")
log = []


def todos_runs(s):
    for sh in s.shapes:
        if sh.has_text_frame:
            for p in sh.text_frame.paragraphs:
                for r in p.runs:
                    yield sh, p, r


def run_exacto(n, viejo, nuevo, color=None, filtro=lambda sh: True):
    hits = [(sh, r) for sh, p, r in todos_runs(S(n)) if r.text == viejo and filtro(sh)]
    assert len(hits) == 1, (n, viejo, len(hits))
    hits[0][1].text = nuevo
    if color is not None:
        hits[0][1].font.color.rgb = color
    log.append(f"{n}: «{viejo}» → «{nuevo}»")
    return hits[0]


def x_de(sh):
    return Emu(sh.left).inches


def y_de(sh):
    return Emu(sh.top).inches


def agregar_parrafo(sh, modelo_p, texto, antes=False, color=None, bold=None):
    """Copia un párrafo existente (mismo formato) con otro texto."""
    nuevo = copy.deepcopy(modelo_p._p)
    if antes:
        modelo_p._p.addprevious(nuevo)
    else:
        modelo_p._p.addnext(nuevo)
    from pptx.text.text import _Paragraph
    p = _Paragraph(nuevo, modelo_p._parent)
    rs = p.runs
    rs[0].text = texto
    for r in rs[1:]:
        r._r.getparent().remove(r._r)
    if color is not None:
        rs[0].font.color.rgb = color
    if bold is not None:
        rs[0].font.bold = bold
    return p


def celda_texto(cell):
    return cell.text_frame.text.strip()


# ============================================================ 2 · Resumen ejecutivo
s = S(2)
allende = lambda sh: x_de(sh) > 7.6
run_exacto(2, "86.1%", "0.0%", filtro=allende)
run_exacto(2, "Cumple", "No cumple", color=ROJO, filtro=lambda sh: allende(sh) and abs(y_de(sh) - 2.65) < 0.05)
run_exacto(2, "30.6%", "0.0%", filtro=allende)
run_exacto(2, "25", "36", filtro=allende)
for sh in s.shapes:  # semáforo de la calificación general de Allende → rojo
    if sh.shape_type == 1 and x_de(sh) > 7.6 and x_de(sh) < 8.1 and 2.0 < y_de(sh) < 2.4:
        sh.fill.solid(); sh.fill.fore_color.rgb = ROJO
        log.append("2: semáforo general Allende → rojo")
nota = next(sh for sh in s.shapes if sh.has_text_frame and sh.text_frame.text.startswith("Se requiere"))
p0 = nota.text_frame.paragraphs[0]
agregar_parrafo(nota, p0, "Pendiente la definición del proyecto.", antes=True, color=AMARILLO, bold=True)
nota.height = Inches(0.6)
log.append("2: Allende: agregado «Pendiente la definición del proyecto.»")
run_exacto(2, "85.9%", "82.1%")
run_exacto(2, "Cumple · 27 indicadores", "Cumple · 23 indicadores")
run_exacto(2, "55.6%", "50.9%")

# ============================================================ 3 · 11 indicadores + lectura rápida
s = S(3)
tbl_shape = next(sh for sh in s.shapes if sh.has_table)
tbl = tbl_shape.table
modelo_na = None
for row in tbl.rows:
    if celda_texto(row.cells[3]) == "No aplica":
        modelo_na = row.cells[3].text_frame.paragraphs[0]._p
        break
for row in tbl.rows:
    nom = celda_texto(row.cells[0])
    c = row.cells[3]
    if nom == "Desarrollo de Proyecto":
        c.text_frame.paragraphs[0].runs[1].text = "0.0%"
        log.append("3: Allende Desarrollo de Proyecto → 0.0%")
    elif nom == "Calificación general":
        rs = c.text_frame.paragraphs[0].runs
        rs[0].font.color.rgb = ROJO
        rs[1].text = "0.0%"
        log.append("3: Allende Calificación general → 0.0% (rojo)")
    elif nom != "Indicador" and celda_texto(c) != "No aplica":
        p = c.text_frame.paragraphs[0]._p
        p.getparent().replace(p, copy.deepcopy(modelo_na))
        log.append(f"3: Allende {nom} → No aplica")
# leyenda bajo la columna de Allende
alto_tabla = sum(r.height for r in tbl.rows)
y_ley = Emu(tbl_shape.top + alto_tabla).inches + 0.06
x_col = Emu(tbl_shape.left + sum(tbl.columns[j].width for j in range(3))).inches
w_col = Emu(tbl.columns[3].width).inches
P.texto(s, x_col - 0.25, y_ley, w_col + 0.5, 0.2, "No se cuenta con el proyecto", size=8, bold=True,
        color="E5484D", align="c", nombre="Leyenda Allende")
log.append("3: leyenda «No se cuenta con el proyecto» bajo la columna de Allende")
lect = next(sh for sh in s.shapes if sh.has_text_frame and "Escobedo" in sh.text_frame.text and "Cumple 9" in sh.text_frame.text)
ps = lect.text_frame.paragraphs
i_esc = next(i for i, p in enumerate(ps) if p.text.strip() == "Escobedo")
p_tit = agregar_parrafo(lect, ps[i_esc + 1], "Allende")
# el título copia el formato del título «Escobedo»
p_tit._p.getparent().replace(p_tit._p, copy.deepcopy(ps[i_esc]._p))
ps = lect.text_frame.paragraphs
p_tit = ps[i_esc + 2]
p_tit.runs[0].text = "Allende"
for r in p_tit.runs[1:]:
    r._r.getparent().remove(r._r)
agregar_parrafo(lect, ps[i_esc + 1], "Es necesario contar con el proyecto de inmediato.")
# el detalle debe ir después del título de Allende
ps = lect.text_frame.paragraphs
det = ps[i_esc + 2]; tit = ps[i_esc + 3]
tit._p.addnext(det._p)
log.append("3: lectura rápida: Allende · «Es necesario contar con el proyecto de inmediato.»")

# ============================================================ 4 · mobiliario
s = S(4)
tbl = next(sh for sh in s.shapes if sh.has_table).table
for row in tbl.rows:
    nom = celda_texto(row.cells[0])
    if nom == "Allende":
        row.cells[2].text_frame.paragraphs[0].runs[0].text = "0.0%"
        row.cells[4].text_frame.paragraphs[0].runs[0].text = "0.0%"
        for j in (3, 5):
            rs = row.cells[j].text_frame.paragraphs[0].runs
            rs[0].text = "No aplica"; rs[0].font.color.rgb = GRIS; rs[0].font.bold = False
            for r in rs[1:]:
                r.text = ""
        log.append("4: Allende: proyecto y general 0.0%; escenario «No aplica»")
    if nom == "Grupo":
        row.cells[2].text_frame.paragraphs[0].runs[0].text = "50.9%"
        rs = row.cells[3].text_frame.paragraphs[0].runs; rs[0].text = "67.2%"; rs[1].text = "  +16.2"
        row.cells[4].text_frame.paragraphs[0].runs[0].text = "82.1%"
        rs = row.cells[5].text_frame.paragraphs[0].runs; rs[0].text = "83.8%"; rs[1].text = "  +1.7"
        log.append("4: Grupo: proyecto 50.9% → 67.2%; general 82.1% → 83.8%")


# ============================================================ 8 · Apodaca: Almacén justo antes de Taller
def bloque(s, nombre):
    card = next(sh for sh in s.shapes if sh.name == f"Área {nombre}")
    y0, y1 = card.top, card.top + card.height
    return [sh for sh in s.shapes if y0 - Emu(Inches(0.01)) <= sh.top < y1], card


s = S(8)
alm, c_alm = bloque(s, "Almacén de refacciones")
gen, c_gen = bloque(s, "Generales")
y_top = c_alm.top
gap = c_gen.top - (c_alm.top + c_alm.height)
d_gen = y_top - c_gen.top
d_alm = (y_top + c_gen.height + gap) - c_alm.top
for sh in gen:
    sh.top = sh.top + d_gen
for sh in alm:
    sh.top = sh.top + d_alm
run_exacto(8, "Almacén de refacciones  ·  Generales", "Generales  ·  Almacén de refacciones")
log.append("8: Generales arriba y Almacén abajo, para que Almacén quede inmediatamente antes de Taller (lámina 9)")

# ============================================================ 11 · Apodaca personal
s = S(11)
run_exacto(11, "2", "15%", filtro=lambda sh: x_de(sh) > 6)
run_exacto(11, "vacantes en administrativo", "personal faltante  ·  6 vacantes")
P.texto(s, 6.25, 3.5, 3.35, 0.5, "Personal reportado por Liliana Ozoria el jueves 12 de noviembre.", size=10,
        color=P.GRIS_CL, nombre="Nota reporte de personal")
log.append("11: cuadro rojo 15% · 6 vacantes; nota de Liliana Ozoria")

# ============================================================ 18 · Allende (agencia)
s = S(18)
run_exacto(18, "86.1%", "0.0%", color=ROJO)
run_exacto(18, "Cumple", "No cumple", color=ROJO, filtro=lambda sh: y_de(sh) < 2.5)
run_exacto(18, "25 ", "36 ")
run_exacto(18, "Calificación general  ·  5 indicadores aplicables", "Calificación general  ·  solo aplica desarrollo de proyecto")
run_exacto(18, "30.6%", "0.0%", filtro=lambda sh: y_de(sh) < 3.5)
run_exacto(18, "faltan", "pendiente la definición del proyecto")
run_exacto(18, " 25 de 36 ", "")
run_exacto(18, "elementos", "")
caja = next(sh for sh in s.shapes if sh.has_text_frame and sh.text_frame.text.startswith("Proyecto "))
agregar_parrafo(caja, caja.text_frame.paragraphs[-1], "Es necesario contar con el proyecto de inmediato.")
log.append("18: agregado «Es necesario contar con el proyecto de inmediato.»")
# tarjeta de Proyecto a 0%; se quitan las 10 tarjetas que no aplican
run_exacto(18, "30.6%", "0.0%", filtro=lambda sh: y_de(sh) > 4.0)
quitar = [sh for sh in s.shapes if y_de(sh) > 4.0 and x_de(sh) > 1.2]
for sh in quitar:
    sh._element.getparent().remove(sh._element)
log.append(f"18: quitadas las 10 tarjetas de indicadores que no aplican ({len(quitar)} formas)")

# ============================================================ 19 · Allende proyecto por área
s = S(19)
run_exacto(19, "30.6% · 11 de 36 elementos completos · faltan 25", "0.0% · 0 de 36 elementos completos · faltan 36")
graf = next(sh for sh in s.shapes if sh.has_chart).chart
cd = CategoryChartData()
cd.categories = ["Imagen corporativa", "Showroom", "Mobiliario", "Generales"]
cd.add_series("% de avance", [0.0, 0.0, 0.0, 0.0])
graf.replace_data(cd)
tbl = next(sh for sh in s.shapes if sh.has_table).table
for row in tbl.rows:
    nom = celda_texto(row.cells[0])
    if nom == "Generales":
        row.cells[2].text_frame.paragraphs[0].runs[0].text = "0"
        r = row.cells[3].text_frame.paragraphs[0].runs[0]
        r.text = "11"; r.font.color.rgb = ROJO; r.font.bold = True
        rs = row.cells[4].text_frame.paragraphs[0].runs
        rs[0].font.color.rgb = ROJO; rs[1].text = "0.0%"
    if nom == "Total":
        row.cells[2].text_frame.paragraphs[0].runs[0].text = "0"
        row.cells[3].text_frame.paragraphs[0].runs[0].text = "36"
        row.cells[4].text_frame.paragraphs[0].runs[0].text = "0.0%"
log.append("19: gráfica y tabla en 0% (Generales 0 de 11; total 0 de 36)")
txt = next(sh for sh in s.shapes if sh.has_text_frame and sh.text_frame.text.startswith("Completas"))
p1, p2 = txt.text_frame.paragraphs[0], txt.text_frame.paragraphs[1]
p1.runs[0].text = "No se cuenta con el proyecto."; p1.runs[0].font.color.rgb = ROJO
for r in p1.runs[1:]:
    r.text = ""
p2.runs[0].text = "Es necesario contar con el proyecto de inmediato."
for r in p2.runs[1:]:
    r.text = ""
log.append("19: texto inferior: «No se cuenta con el proyecto.» / «Es necesario contar con el proyecto de inmediato.» (quitado el escenario de mobiliario)")

# ============================================================ 25 · Allende plan
s = S(25)
tbl = next(sh for sh in s.shapes if sh.has_table).table
c = tbl.cell(1, 1)
assert c.text_frame.text.startswith("Es necesario definir"), c.text_frame.text
r = c.text_frame.paragraphs[0].runs[-1]
r.text = r.text + " Es necesario contar con el proyecto de inmediato."
log.append("25: renglón «Proyecto de la agencia»: agregado «Es necesario contar con el proyecto de inmediato.»")

prs.save(DST)
print("\n".join(log))
