"""Versión final: Almacén + Taller en una lámina (Apodaca y Escobedo) y vacantes clasificadas por área.
Sólo toca las láminas 8, 9, 11, 14, 15 y 17."""
import sys
sys.path.insert(0, "/home/user/Estatus-Distribuidores/presentacion")
from pptx import Presentation
from pptx.util import Emu, Inches
import generar_presentacion as P
import generar_reporte as g

SRC, DST = sys.argv[1], sys.argv[2]
prs = Presentation(SRC)
S = lambda n: prs.slides[n - 1]
IN = lambda e: Emu(e).inches
log = []


# ------------------------------------------------------------------ leer un bloque tal como está
def leer_bloque(s, nombre):
    card = next(sh for sh in s.shapes if sh.name == f"Área {nombre}")
    y0, y1 = card.top, card.top + card.height
    formas = [sh for sh in s.shapes if y0 - 9144 <= sh.top < y1]
    txt = [sh for sh in formas if sh.has_text_frame and sh.text_frame.text.strip()]
    pct = next(sh for sh in txt if sh.text_frame.text.strip().endswith("%") and IN(sh.left) > 8).text_frame.text.strip()
    conteo = next(sh for sh in txt if "elementos" in sh.text_frame.text).text_frame.text.strip()
    cols = sorted([sh for sh in txt if IN(sh.top) > IN(card.top) + 0.6], key=lambda sh: IN(sh.left))
    items = []
    for sh in cols:
        for p in sh.text_frame.paragraphs:
            t = "".join(r.text for r in p.runs)
            if not t.strip():
                continue
            if any(r.font.bold for r in p.runs):
                tit, _, av = t.partition("  ·  ")
                items.append({"titulo": tit.strip(), "avance": av.strip(), "texto": ""})
            else:
                items[-1]["texto"] = (items[-1]["texto"] + " " + t).strip()
    valor = float(pct.rstrip("%"))
    return {"nombre": nombre, "pct": pct, "valor": valor, "conteo": conteo, "items": items,
            "estado": g.semaforo(valor, {"verde_min": 80, "amarillo_min": 50})}, formas


def quitar(formas):
    for sh in formas:
        sh._element.getparent().remove(sh._element)


# ------------------------------------------------------------------ bloque compacto (mismo diseño)
CAB = 0.58


def alto_item(it, ww, ti, te):
    tit = it["titulo"] + ("\u00a0\u00a0·\u00a0\u00a0" + it["avance"].replace(" ", "\u00a0") if it["avance"] else "")
    h = P.lineas(tit, ti, ww, True) * ti * 1.17 / 72
    h += P.lineas(it["texto"], te, ww) * te * 1.17 / 72
    return h + 0.1 * ti / 10


def columnas(items, n):
    por = -(-len(items) // n)
    return [items[i * por:(i + 1) * por] for i in range(n)]


def alto_bloque(a, w, cols, ti, te):
    ww = (w - 0.4 - 0.25 * (cols - 1)) / cols
    return CAB + 0.08 + max(sum(alto_item(it, ww, ti, te) for it in c) for c in columnas(a["items"], cols)) + 0.04


def dibujar(s, a, x, y, w, h, cols, ti, te):
    P.caja(s, x, y, w, h, P.TARJETA, nombre=f"Área {a['nombre']}")
    P.punto(s, x + 0.2, y + 0.16, 0.16, a["estado"])
    P.texto(s, x + 0.45, y + 0.09, w - 1.6, 0.28, a["nombre"], size=12, bold=True)
    P.texto(s, x + w - 1.25, y + 0.07, 1.05, 0.3, a["pct"], size=14, bold=True, align="r", color=P.SEM[a["estado"]])
    P.texto(s, x + 0.45, y + 0.32, w - 0.7, 0.16, a["conteo"], size=8, color=P.GRIS)
    P.barra(s, x + 0.2, y + 0.51, w - 0.4, 0.04, a["valor"], a["estado"])
    ww = (w - 0.4 - 0.25 * (cols - 1)) / cols
    sep = 0.1 * ti / 10 * 72
    for c, grupo in enumerate(columnas(a["items"], cols)):
        pars = []
        for it in grupo:
            runs = [(it["titulo"], {"bold": True, "color": P.BLANCO})]
            if it["avance"]:
                runs.append(("\u00a0\u00a0·\u00a0\u00a0" + it["avance"].replace(" ", "\u00a0"), {"bold": True, "color": P.SEM["amarillo"]}))
            pars.append({"runs": runs, "size": ti, "espacio": 1 if it["texto"] else sep})
            if it["texto"]:
                pars.append({"runs": [(it["texto"], {})], "size": te, "color": P.GRIS_CL, "espacio": sep})
        P.texto(s, x + 0.2 + c * (ww + 0.25), y + CAB + 0.06, ww, h - CAB - 0.08, pars)


def juntar(n_origen, n_taller, sub_origen, sub_taller):
    so, st = S(n_origen), S(n_taller)
    alm, f_alm = leer_bloque(so, "Almacén de refacciones")
    tal, f_tal = leer_bloque(st, "Taller")
    disp, gap = P.Y_FIN - P.Y_CONT, 0.12
    elegido = None
    grandes = ((10.5, 9.5), (10, 9), (9.5, 8.5), (9, 8))
    chicas = ((8.5, 7.5), (8, 7))
    opciones = [(ti, te, 3) for ti, te in grandes] + [(ti, te, 4) for ti, te in grandes] + \
               [(ti, te, c) for ti, te in chicas for c in (3, 4)]
    for ti, te, cols in opciones:  # se prefieren 3 columnas y letra más grande
        ha = alto_bloque(alm, P.ANCHO, 1, ti, te)
        ht = alto_bloque(tal, P.ANCHO, cols, ti, te)
        if ha + gap + ht <= disp - 0.02:
            elegido = (ti, te, cols, ha, ht)
            break
    assert elegido, "no cabe"
    ti, te, cols, ha, ht = elegido
    quitar(f_alm)
    quitar(f_tal)
    for s, viejo, nuevo in ((so, sub_origen[0], sub_origen[1]), (st, sub_taller[0], sub_taller[1])):
        ok = False
        for sh in s.shapes:
            if sh.has_text_frame and sh.text_frame.text.strip() == viejo:
                p = sh.text_frame.paragraphs[0]
                p.runs[0].text = nuevo
                for r in p.runs[1:]:
                    r._r.getparent().remove(r._r)
                ok = True
        assert ok, viejo
    dibujar(st, alm, P.X0, P.Y_CONT, P.ANCHO, ha, 1, ti, te)
    dibujar(st, tal, P.X0, P.Y_CONT + ha + gap, P.ANCHO, ht, cols, ti, te)
    log.append(f"{n_taller}: Almacén ({len(alm['items'])}) + Taller ({len(tal['items'])}) juntos; "
               f"letra {ti}/{te} pt, Taller en {cols} columnas; alto {ha + gap + ht:.2f} de {disp:.2f} in")
    log.append(f"{n_origen}: se quitó Almacén (pasó a la lámina {n_taller})")


juntar(8, 9, ("Generales  ·  Almacén de refacciones", "Generales"), ("Taller", "Almacén de refacciones  ·  Taller"))
juntar(14, 15, ("Imagen corporativa  ·  Almacén de refacciones", "Imagen corporativa"),
       ("Taller", "Almacén de refacciones  ·  Taller"))


# ------------------------------------------------------------------ personal: vacantes por área, «X de Y»
def vacantes(n, filas):
    s = S(n)
    hdr = next(sh for sh in s.shapes if sh.has_text_frame and sh.text_frame.text.startswith("Vacantes por puesto"))
    y0 = IN(hdr.top) + 0.25
    quitar([sh for sh in s.shapes if IN(sh.top) >= y0 and IN(sh.left) < 6.0])
    xi, wi = P.X0, 5.55
    mitad = -(-len(filas) // 2)
    cw = (wi - 0.12) / 2
    alto = 0.3
    y = IN(hdr.top) + 0.3
    for i, (area, puesto, falta, req) in enumerate(filas):
        x = xi + (i // mitad) * (cw + 0.12)
        yy = y + (i % mitad) * (alto + 0.05)
        P.caja(s, x, yy, cw, alto, P.TARJETA, radio=0.03)
        P.texto(s, x + 0.1, yy, cw - 0.65, alto, [[(f"{area}: ", {"bold": True, "color": P.BLANCO}), (puesto, {})]],
                size=8.5, color=P.GRIS_CL, anchor="m")
        P.texto(s, x + cw - 0.6, yy, 0.5, alto, f"{falta} de {req}", size=9, bold=True, align="r", anchor="m")
    log.append(f"{n}: vacantes por área: " + "; ".join(f"{a}: {p} {f} de {r}" for a, p, f, r in filas))


vacantes(11, [("Administrativo", "Recepcionista", 1, 1), ("Administrativo", "Marketing", 1, 1),
              ("Comercial", "Asesor de ventas HDT", 2, 2),
              ("Postventa", "Asesor de servicio", 1, 3), ("Postventa", "Técnico VE (eléctricos)", 1, 1)])
vacantes(17, [("Administrativo", "Marketing", 1, 1), ("Comercial", "Asesor de ventas HDT", 2, 2),
              ("Comercial", "Asesor digital (Biky)", 1, 1)])

prs.save(DST)
print("\n".join(log))
