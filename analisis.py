#!/usr/bin/env python3
"""
Análisis compartido por la presentación (presentacion/generar_presentacion.py) y la
propuesta de herramienta (propuesta/generar_propuesta.py).

Usa los mismos cálculos de generar_reporte.py (evaluar) y agrega:
  - % Desarrollo de Proyecto por área (elementos, cumple, faltan, %).
  - Explicación corta de cada pendiente (datos/explicaciones.json).
  - Escenario «mobiliario validado como OK».
  - Total del grupo (promedio de todos los indicadores aplicables, Allende incluido).
  - Renglones del plan de trabajo agrupados por área / indicador.
"""
import datetime as dt
import json
import os
import sys

import generar_reporte as g

BASE = os.path.dirname(os.path.abspath(__file__))


def ruta(p):
    return p if os.path.isabs(p) else os.path.join(BASE, p)


def fmt(v, dec=1):
    """Porcentaje con coma decimal fija (73.5%)."""
    return "—" if v is None else f"{v:.{dec}f}%"


class Explicaciones:
    def __init__(self, archivo):
        self.d = json.load(open(archivo, encoding="utf-8")) if os.path.exists(archivo) else {}
        self.c = {g.norm(k): v for k, v in self.d.get("conceptos", {}).items()}
        self.cu = {g.norm(k): v for k, v in self.d.get("cummins", {}).items()}
        self.cuc = {g.norm(k): v for k, v in self.d.get("cummins_corto", {}).items()}
        self.pu = {g.norm(k): v for k, v in self.d.get("puestos", {}).items()}
        self.ar = {g.norm(k): v for k, v in self.d.get("areas", {}).items()}
        self.sin_texto = set()

    def concepto(self, nombre):
        e = self.c.get(g.norm(nombre))
        if not e:
            self.sin_texto.add(nombre)
            return g.limpio(nombre), None
        return e["titulo"], e.get("texto")

    def cummins(self, nombre):
        return self.cu.get(g.norm(nombre))

    def cummins_corto(self, nombre):
        return self.cuc.get(g.norm(nombre), g.limpio(nombre))

    def indicador(self, ind_id):
        return self.d.get("indicadores", {}).get(ind_id, {})

    def puesto(self, nombre):
        return self.pu.get(g.norm(nombre), g.limpio(nombre))

    def area(self, nombre):
        return self.ar.get(g.norm(nombre), g.limpio(nombre))


def cargar(config=None, grupo_id=None):
    cfg = json.load(open(config or ruta("config.json"), encoding="utf-8"))
    corte = dt.date.fromisoformat(cfg["fecha_corte"])
    kpi = g.LectorKPI(ruta(cfg["archivos"]["kpis"]))
    per = g.LectorPersonal(ruta(cfg["archivos"]["personal_sistema"]))
    fotos = g.Fotos(ruta(cfg["archivos"]["fotos"]), cfg, corte)
    ajustes = g.cargar_ajustes(ruta(cfg["archivos"]["ajustes"])) if cfg["archivos"].get("ajustes") else []
    grupo = next(x for x in cfg["grupos"] if not grupo_id or x["id"] == grupo_id)
    dists = g.evaluar(cfg, grupo, kpi, per, fotos, corte, ajustes)
    exp = Explicaciones(ruta(cfg["archivos"].get("explicaciones", "datos/explicaciones.json")))
    ana = [analizar(d, cfg, exp) for d in dists]
    if exp.sin_texto:
        print("  aviso: pendientes sin explicación en datos/explicaciones.json: "
              + "; ".join(sorted(exp.sin_texto)), file=sys.stderr)
    return {"cfg": cfg, "corte": corte, "grupo": grupo, "dists": ana, "fotos": fotos,
            "total": totales(ana, cfg)}


def _avance_txt(c):
    if c["req"] > 1:
        return f"tiene {c['adq']:g} de {c['req']:g}"
    return f"avance {c['avance'] * 100:.0f}%"


def analizar(d, cfg, exp):
    reglas = cfg["semaforo"]
    ind = {i["id"]: i for i in d["indicadores"]}

    # ---- proyecto por área (cada área es su propio bloque, en el orden del Excel)
    pr = ind.get("proyecto")
    areas = []
    if pr and pr["conceptos"]:
        orden = {}
        for c in pr["conceptos"]:
            orden.setdefault(c["subtema"], []).append(c)
        for nombre, cs in orden.items():
            pend = []
            for c in cs:
                if c["avance"] < 1:
                    titulo, texto = exp.concepto(c["concepto"])
                    pend.append({"concepto": c["concepto"], "titulo": titulo, "texto": texto,
                                 "avance": c["avance"], "parcial": c["avance"] > 0,
                                 "avance_txt": _avance_txt(c) if c["avance"] > 0 else "",
                                 "foto": c.get("foto")})
            valor = 100 * sum(c["avance"] for c in cs) / len(cs)
            areas.append({"nombre": exp.area(nombre), "original": nombre, "n": len(cs),
                          "cumple": sum(1 for c in cs if c["avance"] >= 1), "faltan": len(pend),
                          "valor": valor, "estado": g.semaforo(valor, reglas["proyecto"]), "pendientes": pend})
    proyecto = None
    if pr and pr["conceptos"]:
        n = len(pr["conceptos"])
        proyecto = {"valor": pr["valor"], "estado": pr["estado"], "n": n,
                    "cumple": sum(1 for c in pr["conceptos"] if c["avance"] >= 1),
                    "suma": sum(c["avance"] for c in pr["conceptos"])}
        proyecto["faltan"] = n - proyecto["cumple"]

    # ---- escenario: el mobiliario actual se valida como OK (cada elemento al 100 %)
    escenario = None
    mob = next((a for a in areas if g.norm(a["original"]) == "mobiliario"), None)
    if mob and proyecto:
        mob_suma = mob["valor"] * mob["n"] / 100
        proy_val = 100 * (proyecto["suma"] - mob_suma + mob["n"]) / proyecto["n"]
        vals = [proy_val if i["id"] == "proyecto" else i["valor"] for i in d["indicadores"]
                if i["estado"] != "na" and i["valor"] is not None]
        gen_val = sum(vals) / len(vals)
        escenario = {"elementos": mob["n"], "pendientes": mob["faltan"],
                     "proyecto_hoy": proyecto["valor"], "proyecto_validado": proy_val,
                     "general_hoy": d["general"], "general_validado": gen_val,
                     "suma_proyecto_validado": proyecto["suma"] - mob_suma + mob["n"]}

    # ---- personal
    personal = None
    pe = ind.get("personal")
    if pe and pe.get("personal"):
        r = pe["personal"]
        vac = sorted([{"puesto": exp.puesto(p["puesto"]), "area": p["area"], "faltan": p["faltan"],
                       "req": p["req"], "cub": p["cub"]} for p in pe["puestos"] if p["faltan"] > 0],
                     key=lambda p: -p["faltan"])
        personal = {"valor": pe["valor"], "estado": pe["estado"], "req": r["req"], "cub": r["cub"],
                    "faltan": r["faltan"], "extra": r["extra"], "vacantes": vac,
                    "areas": [a for a in (pe.get("areas") or {}).get("areas", []) if a["req"]]}

    # ---- resumen de los 11 indicadores
    resumen = []
    for i in d["indicadores"]:
        e = {"id": i["id"], "nombre": i["nombre"], "valor": i["valor"], "estado": i["estado"],
             "nota": i.get("nota", ""), "falta": "", "necesita": "", "items": []}
        if i["estado"] not in ("na", "sin") and i["valor"] is not None and i["valor"] < 100 - 1e-9:
            if i["id"] == "proyecto":
                peor = sorted([a for a in areas if a["faltan"]], key=lambda a: a["valor"])[:3]
                e["falta"] = f"Faltan {proyecto['faltan']} de {proyecto['n']} elementos."
                e["necesita"] = "Mayor rezago: " + ", ".join(f"{a['nombre']} {fmt(a['valor'], 0)}" for a in peor) + "."
            elif i["id"] == "personal":
                e["falta"] = f"{personal['cub']:g} de {personal['req']:g} plazas cubiertas."
                e["necesita"] = (f"Contratar {personal['faltan']:g} vacantes en {len(personal['vacantes'])} puestos. Principales: "
                                 + ", ".join(f"{v['puesto']} {v['faltan']:g}" for v in personal["vacantes"][:3]) + ".")
            elif i["id"] == "cummins":
                for c in i["conceptos"]:
                    if c["avance"] < 1:
                        accion = exp.cummins(c["concepto"]) or g.limpio(c["concepto"])
                        if c.get("falta"):
                            accion = c["falta"]
                        e["items"].append({"titulo": accion, "corto": exp.cummins_corto(c["concepto"]),
                                           "avance_txt": _avance_txt(c) if c["avance"] > 0 else ""})
                pend = len(e["items"])
                e["falta"] = (f"{pend} de {len(i['conceptos'])} requisitos pendientes." if pend > 1
                              else e["items"][0]["titulo"] + "." if pend else "")
                corto = [x["corto"] for x in e["items"]]
                txt = ", ".join(corto[:-1]) + " y " + corto[-1] + "."
                e["necesita"] = txt[0].upper() + txt[1:] if pend > 1 else ""
            else:
                x = exp.indicador(i["id"])
                e["falta"] = x.get("falta", "No cumple: requerido y no acreditado.")
                e["necesita"] = x.get("necesita", "")
        resumen.append(e)

    # ---- plan de trabajo agrupado (lo detallado vive en la herramienta)
    plan = []
    if d.get("mensaje_clave"):
        plan.append({"pendiente": "Proyecto de la agencia", "falta": d["mensaje_clave"], "prioridad": "ALTA"})
    for a in areas:
        if not a["pendientes"]:
            continue
        items = [p["titulo"] + (f" ({p['avance_txt']})" if p["parcial"] else "") for p in a["pendientes"]]
        falta = ", ".join(items) + "."
        if g.norm(a["original"]) == "mobiliario" and a["faltan"] > 8:
            falta = (f"Validar el mobiliario actual contra el estándar FOTON y cotizar lo que no cumpla: "
                     f"recepción, escritorios ({sum(1 for p in a['pendientes'] if 'scritorio' in p['titulo'])}), "
                     f"mesas, periqueras, exhibidor, mueble de café y tótem de ficha técnica.")
        plan.append({"pendiente": f"{a['nombre']} ({a['faltan']} de {a['n']})",
                     "falta": falta, "prioridad": "ALTA", "indicador": "proyecto"})
    for e in resumen:
        if e["id"] == "proyecto" or not e["falta"]:
            continue
        if e["id"] == "personal":
            falta = ", ".join(f"{v['puesto']} {v['faltan']:g}" for v in personal["vacantes"]) + "."
            plan.append({"pendiente": f"Personal ({personal['faltan']:g} vacantes)", "falta": falta,
                         "prioridad": "ALTA", "indicador": "personal"})
        elif e["id"] == "cummins":
            plan.append({"pendiente": "Cummins", "falta": e["necesita"] or e["falta"], "prioridad": "ALTA",
                         "indicador": "cummins"})
        else:
            plan.append({"pendiente": e["nombre"], "falta": f"{e['falta']} {e['necesita']}".strip(),
                         "prioridad": "ALTA", "indicador": e["id"]})

    return {**d, "areas": areas, "proyecto": proyecto, "escenario": escenario, "personal": personal,
            "resumen": resumen, "plan": plan, "n_pendientes": len(d["pendientes"]),
            "corto": d["nombre"].replace("FOTON ", "")}


def totales(dists, cfg):
    """Total del grupo: promedio de TODOS los indicadores aplicables de las agencias (Allende incluido)."""
    reglas = cfg["semaforo"]
    vals = [i["valor"] for d in dists for i in d["indicadores"] if i["estado"] != "na" and i["valor"] is not None]
    general = sum(vals) / len(vals) if vals else None
    n = sum(d["proyecto"]["n"] for d in dists if d["proyecto"])
    suma = sum(d["proyecto"]["suma"] for d in dists if d["proyecto"])
    proyecto = 100 * suma / n if n else None
    # escenario mobiliario validado, a nivel grupo
    vals_v = []
    for d in dists:
        for i in d["indicadores"]:
            if i["estado"] == "na" or i["valor"] is None:
                continue
            vals_v.append(d["escenario"]["proyecto_validado"] if i["id"] == "proyecto" and d["escenario"] else i["valor"])
    suma_v = sum((d["escenario"]["suma_proyecto_validado"] if d["escenario"] else d["proyecto"]["suma"])
                 for d in dists if d["proyecto"])
    return {"general": general, "estado_general": g.semaforo(general, reglas["general"]),
            "n_indicadores": len(vals), "proyecto": proyecto,
            "estado_proyecto": g.semaforo(proyecto, reglas["proyecto"]), "elementos": n,
            "pendientes": sum(d["n_pendientes"] for d in dists),
            "general_validado": sum(vals_v) / len(vals_v) if vals_v else None,
            "proyecto_validado": 100 * suma_v / n if n else None,
            "mobiliario": sum(d["escenario"]["elementos"] for d in dists if d["escenario"])}


if __name__ == "__main__":
    ctx = cargar()
    t = ctx["total"]
    print(f"Grupo: general {fmt(t['general'], 2)} ({t['n_indicadores']} indicadores), proyecto {fmt(t['proyecto'], 2)} "
          f"({t['elementos']} elementos), {t['pendientes']} pendientes. Mobiliario validado: "
          f"general {fmt(t['general_validado'], 2)}, proyecto {fmt(t['proyecto_validado'], 2)}")
    for d in ctx["dists"]:
        e = d["escenario"]
        print(f"\n{d['nombre']}: general {fmt(d['general'], 2)}, proyecto {fmt(d['proyecto']['valor'], 2)}, "
              f"{d['n_pendientes']} pendientes; mobiliario OK -> proyecto {fmt(e['proyecto_validado'])}, "
              f"general {fmt(e['general_validado'], 2)}")
        for a in d["areas"]:
            print(f"   {a['nombre']:<24} {fmt(a['valor']):>7} {a['n']:>3} elem  cumple {a['cumple']:>2}  faltan {a['faltan']}")
        for p in d["plan"]:
            print(f"   PLAN {p['pendiente']}: {p['falta'][:110]}")
