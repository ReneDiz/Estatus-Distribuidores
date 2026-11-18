#!/usr/bin/env python3
"""
Lector del «Check list de seguimiento de proyecto» (Excel) con fotos por elemento.

Lee cada renglón de elemento (Lineamiento, Elemento, Fotografías, % de cumplimiento,
Fecha de revisión, Comentario) y las fotos que tiene ligadas. Entiende las dos formas en que
Excel guarda fotos:
  - imagen flotante anclada a la celda de «Fotografías» (xl/drawings), y
  - imagen dentro de la celda («Colocar en celda»; xl/richData), que en el texto aparece como #VALUE!.

Uso rápido:
    python3 checklist.py datos/checklists/<archivo>.xlsx
"""
import io
import json
import os
import re
import sys
import zipfile

import openpyxl

import generar_reporte as g

BASE = os.path.dirname(os.path.abspath(__file__))
NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def _num(v):
    if isinstance(v, (int, float)):
        return float(v)
    if isinstance(v, str):
        t = v.strip().replace("%", "")
        if t.upper() in ("N/A", "NA"):
            return None
        try:
            return float(t) / (100 if "%" in v else 1)
        except ValueError:
            return None
    return None


def _fotos_en_celda(ruta, hoja_idx):
    """Fotos «dentro de la celda» (richData). Regresa {(fila, col): [bytes, ...]}."""
    out = {}
    with zipfile.ZipFile(ruta) as z:
        nombres = set(z.namelist())
        if "xl/richData/rdrichvalue.xml" not in nombres or "xl/metadata.xml" not in nombres:
            return out
        txt = lambda n: z.read(n).decode("utf-8", "replace")
        # relaciones richValueRel -> media
        rels = {}
        rp = "xl/richData/_rels/richValueRel.xml.rels"
        if rp in nombres:
            for m in re.finditer(r'<Relationship[^>]*Id="([^"]+)"[^>]*Target="([^"]+)"', txt(rp)):
                rels[m.group(1)] = os.path.normpath(os.path.join("xl/richData", m.group(2))).replace("\\", "/")
        rel_ids = re.findall(r'<rel[^>]*r:id="([^"]+)"', txt("xl/richData/richValueRel.xml")) \
            if "xl/richData/richValueRel.xml" in nombres else []
        # estructura de cada rich value: qué posición trae el índice de imagen
        claves = []
        if "xl/richData/rdrichvaluestructure.xml" in nombres:
            for s in re.findall(r"<s [^>]*>(.*?)</s>", txt("xl/richData/rdrichvaluestructure.xml"), re.S):
                claves.append(re.findall(r'<k n="([^"]+)"', s))
        valores = []
        for m in re.finditer(r'<rv s="(\d+)"[^>]*>(.*?)</rv>', txt("xl/richData/rdrichvalue.xml"), re.S):
            s, vs = int(m.group(1)), re.findall(r"<v[^>]*>(.*?)</v>", m.group(2), re.S)
            k = claves[s] if s < len(claves) else []
            pos = next((i for i, n in enumerate(k) if "LocalImageIdentifier" in n), 0)
            try:
                valores.append(int(vs[pos]))
            except (IndexError, ValueError):
                valores.append(None)
        # metadata: vm (1-based) -> índice de rich value
        meta = txt("xl/metadata.xml")
        fut = re.search(r'<futureMetadata name="XLRICHVALUE"[^>]*>(.*?)</futureMetadata>', meta, re.S)
        rv_de_bk = [int(x) for x in re.findall(r'<xlrd:rvb i="(\d+)"', fut.group(1))] if fut else []
        vm_bk = []
        vm = re.search(r"<valueMetadata[^>]*>(.*?)</valueMetadata>", meta, re.S)
        if vm:
            for bk in re.findall(r"<bk>(.*?)</bk>", vm.group(1), re.S):
                v = re.search(r'<rc [^>]*v="(\d+)"', bk)
                vm_bk.append(int(v.group(1)) if v else None)
        # hoja
        wbx = txt("xl/workbook.xml")
        hojas = re.findall(r'<sheet [^>]*r:id="([^"]+)"', wbx)
        wrels = {m.group(1): m.group(2) for m in re.finditer(
            r'<Relationship[^>]*Id="([^"]+)"[^>]*Target="([^"]+)"', txt("xl/_rels/workbook.xml.rels"))}
        destino = wrels.get(hojas[hoja_idx], "")
        destino = destino.lstrip("/")
        hoja = destino if destino.startswith("xl/") else "xl/" + destino
        if hoja not in nombres:
            return out
        for m in re.finditer(r'<c r="([A-Z]+)(\d+)"[^>]*\bvm="(\d+)"', txt(hoja)):
            col, fila, vmi = m.group(1), int(m.group(2)), int(m.group(3))
            try:
                rv = rv_de_bk[vm_bk[vmi - 1]]
                rel = rel_ids[valores[rv]]
                data = z.read(rels[rel])
            except (IndexError, KeyError, TypeError):
                continue
            ncol = 0
            for ch in col:
                ncol = ncol * 26 + (ord(ch) - 64)
            out.setdefault((fila, ncol), []).append(data)
    return out


def leer(ruta):
    """Renglones de elementos del checklist con sus fotos (bytes)."""
    wb = openpyxl.load_workbook(ruta)
    ws, idx = wb.worksheets[0], 0
    # columnas por encabezado
    cab = None
    for r in range(1, min(ws.max_row, 40) + 1):
        vals = [g.norm(ws.cell(r, c).value) for c in range(1, ws.max_column + 1)]
        if "elemento" in vals and "lineamiento" in vals:
            cab = r
            col = {n: vals.index(n) + 1 for n in ("lineamiento", "elemento") if n in vals}
            col["fotos"] = next((i + 1 for i, v in enumerate(vals) if v.startswith("fotograf")), None)
            col["pct"] = next((i + 1 for i, v in enumerate(vals) if v.startswith("% de cumplimiento")), None)
            col["fecha"] = next((i + 1 for i, v in enumerate(vals) if v.startswith("fecha de revision")), None)
            arriba = [g.norm(ws.cell(r - 1, c).value) for c in range(1, ws.max_column + 1)]
            col["coment"] = next((i + 1 for i, v in enumerate(arriba) if v.startswith("comentario")), None)
            break
    if not cab:
        raise ValueError("No encontré el encabezado «Lineamiento / Elemento» en el checklist")
    flot = {}
    for im in getattr(ws, "_images", []):
        f = im.anchor._from
        flot.setdefault(f.row + 1, []).append(im._data())
    celda = _fotos_en_celda(ruta, idx)
    filas = []
    for r in range(cab + 1, ws.max_row + 1):
        el = ws.cell(r, col["elemento"]).value
        no = ws.cell(r, col["elemento"] - 2).value
        if not el or not isinstance(no, (int, float)):
            continue
        fotos = list(flot.get(r, []))
        if col.get("fotos"):
            fotos += celda.get((r, col["fotos"]), [])
        fecha = ws.cell(r, col["fecha"]).value if col.get("fecha") else None
        filas.append({
            "fila": r, "lineamiento": g.limpio(ws.cell(r, col["lineamiento"]).value),
            "elemento": g.limpio(el), "pct": _num(ws.cell(r, col["pct"]).value) if col.get("pct") else None,
            "fecha": fecha.date().isoformat() if hasattr(fecha, "date") else (str(fecha) if fecha else None),
            "comentario": g.limpio(ws.cell(r, col["coment"]).value) if col.get("coment") else "",
            "fotos": fotos})
    return filas


def empatar(conceptos, filas, equivalencias=None):
    """Liga cada concepto del KPI con su renglón del checklist (nombre igual o equivalencia declarada)."""
    eq = {g.norm(k): g.norm(v) for k, v in (equivalencias or {}).items()}
    por_nombre = {}
    for f in filas:
        por_nombre.setdefault(g.norm(f["elemento"]), f)
    out = {}
    for c in conceptos:
        n = g.norm(c)
        f = por_nombre.get(n) or por_nombre.get(eq.get(n, ""))
        if f:
            out[c] = f
    return out


def config():
    p = os.path.join(BASE, "datos", "checklists", "checklists.json")
    return json.load(open(p, encoding="utf-8")) if os.path.exists(p) else {}


if __name__ == "__main__":
    filas = leer(sys.argv[1])
    print(f"{len(filas)} elementos, {sum(len(f['fotos']) for f in filas)} fotos")
    for f in filas:
        print(f"  {f['fila']:>4} {f['lineamiento'][:20]:20} {f['elemento'][:45]:45} {f['pct']!s:6} fotos={len(f['fotos'])}")
