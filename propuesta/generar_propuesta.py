#!/usr/bin/env python3
"""
Genera la PROPUESTA de herramienta nueva: propuesta/herramienta_estatus_PROPUESTA.html

Uso:
    python3 propuesta/generar_propuesta.py

Un solo archivo HTML que funciona sin internet, con el estilo del template FOTON:
  - Navegación por sucursal (Resumen, Apodaca, Escobedo, Allende, Plan de trabajo).
  - Filtros por indicador y por estado; clic en un pendiente abre su detalle.
  - Plan de trabajo editable (Responsable, Fecha compromiso, Estatus, Comentarios),
    exportable a Excel e imprimible / PDF.
  - Modo presentación a pantalla completa (flechas para avanzar, Esc para salir).

No modifica la herramienta actual (generar_reporte.py / salida/Grupo_Velcen.html):
sólo reutiliza sus cálculos a través de analisis.py.
"""
import base64
import html
import io
import json
import os
import re
import sys
import zipfile

from PIL import Image, ImageOps

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, RAIZ)
import analisis as A  # noqa: E402
import generar_reporte as g  # noqa: E402

TEMPLATE = os.path.join(RAIZ, "template", "Foton_Mexico_-_Template_Base-2026.pptx")
SALIDA = os.path.join(RAIZ, "propuesta", "herramienta_estatus_PROPUESTA.html")
FACHADAS = {
    "apodaca": ("fotos/editadas/apodaca_fachada_editada.jpg", False),
    "escobedo": ("fotos/editadas/escobedo_fachada_editada.jpg", False),
    "allende": ("datos/fotos/allende/WhatsApp Image 2025-06-27 at 12.56.02 PM (2).jpeg", True),
}


def slug(s):
    return re.sub(r"[^a-z0-9]+", "-", g.norm(s)).strip("-")[:90]


def b64_jpg(ruta, ancho, calidad=72):
    im = ImageOps.exif_transpose(Image.open(ruta))
    if "icc_profile" in im.info:
        try:
            from PIL import ImageCms
            src = ImageCms.ImageCmsProfile(io.BytesIO(im.info["icc_profile"]))
            im = ImageCms.profileToProfile(im, src, ImageCms.createProfile("sRGB"), outputMode="RGB")
        except Exception:
            pass
    im = im.convert("RGB")
    if im.width > ancho:
        im = im.resize((ancho, round(im.height * ancho / im.width)), Image.LANCZOS)
    buf = io.BytesIO()
    im.save(buf, "JPEG", quality=calidad, optimize=True)
    return "data:image/jpeg;base64," + base64.b64encode(buf.getvalue()).decode()


def media_template(nombre):
    with zipfile.ZipFile(TEMPLATE) as z:
        return "data:image/png;base64," + base64.b64encode(z.read("ppt/media/" + nombre)).decode()


def datos(ctx):
    fotos = {}

    def foto(f):
        """Registra una foto (de Fotos de generar_reporte) y regresa su id; sin fechas ni rutas."""
        if not f:
            return None
        k = "f" + str(len(fotos) + 1)
        fotos[k] = {"src": f["src"], "pie": f.get("descripcion", ""), "no_reciente": bool(f.get("no_reciente"))}
        return k

    out = []
    for d in ctx["dists"]:
        pend = []

        def agrega(ind, indicador, area, titulo, falta, texto, avance, estado, f=None, clave=None):
            pend.append({"id": f"{d['id']}-{ind}-{slug(clave or titulo)}", "dist": d["id"], "ind": ind,
                         "indicador": indicador, "area": area, "titulo": titulo, "falta": falta,
                         "texto": texto or "", "avance": round(avance * 100), "estado": estado, "foto": f})

        if d.get("mensaje_clave"):
            agrega("proyecto", "Desarrollo de Proyecto", "Proyecto de la agencia", "Definir el proyecto de la agencia",
                   d["mensaje_clave"], "Sin proyecto definido no se puede programar imagen, showroom ni mobiliario.",
                   0, "rojo", clave="definir-proyecto")
        areas = []
        for a in d["areas"]:
            items = []
            for p in a["pendientes"]:
                fid = foto(p["foto"]) if p.get("foto") else None
                falta = f"Parcial: {p['avance_txt']}" if p["parcial"] else "No cumple"
                agrega("proyecto", "Desarrollo de Proyecto", a["nombre"], p["titulo"], falta, p["texto"], p["avance"],
                       "amarillo" if p["parcial"] else "rojo", fid, clave=a["nombre"] + "-" + p["concepto"])
                items.append(pend[-1]["id"])
            areas.append({k: a[k] for k in ("nombre", "n", "cumple", "faltan", "valor", "estado")} | {"items": items,
                          "mobiliario": g.norm(a["original"]) == "mobiliario"})
        resumen = []
        for e in d["resumen"]:
            r = {k: e[k] for k in ("id", "nombre", "valor", "estado", "falta", "necesita")}
            r["nombre"] = r["nombre"].replace("% ", "")
            resumen.append(r)
            if e["id"] == "cummins":
                for it in e["items"]:
                    agrega("cummins", "Cummins", "Requisitos Cummins", it["titulo"],
                           f"Parcial: {it['avance_txt']}" if it["avance_txt"] else "No cumple", "",
                           0.5 if it["avance_txt"] else 0, "amarillo" if it["avance_txt"] else "rojo", clave=it["corto"])
            elif e["id"] not in ("proyecto", "personal") and e["falta"]:
                agrega(e["id"], r["nombre"], "", r["nombre"], e["falta"], e["necesita"], 0, "rojo")
        per = d["personal"]
        if per:
            for v in per["vacantes"]:
                agrega("personal", "Personal", v["area"], v["puesto"], f"Contratar {v['faltan']:g} (requeridas {v['req']:g}, cubiertas {v['cub']:g})",
                       "", v["cub"] / v["req"] if v["req"] else 0, "amarillo" if v["cub"] else "rojo")
        fach, no_rec = FACHADAS.get(d["id"], (None, False))
        fachada = None
        if fach and os.path.exists(os.path.join(RAIZ, fach)):
            k = "fach-" + d["id"]
            fotos[k] = {"src": b64_jpg(os.path.join(RAIZ, fach), 1400), "pie": f"Fachada {d['nombre']}", "no_reciente": no_rec}
            fachada = k
        obs = []
        for o in d["observaciones"]:
            obs.append({"texto": o.get("texto_presentacion") or o["texto"],
                        "foto": foto({**o["foto"], "descripcion": o.get("descripcion", "")}) if o.get("foto") else None})
        out.append({
            "id": d["id"], "nombre": d["nombre"], "corto": d["corto"], "general": d["general"],
            "estado_general": d["estado_general"], "proyecto": d["proyecto"], "escenario": d["escenario"],
            "mensaje": d.get("mensaje_clave", ""), "motivo_na": d.get("motivo_no_aplica", ""),
            "resumen": resumen, "areas": areas, "personal": per, "pendientes": pend, "fachada": fachada,
            "observaciones": obs, "n_pendientes": d["n_pendientes"]})
    t = ctx["total"]
    return {"grupo": ctx["grupo"]["nombre"], "grupo_id": ctx["grupo"]["id"], "corte": ctx["corte"].isoformat(),
            "corte_txt": g.fecha_larga(ctx["corte"]), "total": t, "dists": out, "fotos": fotos,
            "semaforo": ctx["cfg"]["semaforo"]}


CSS = r"""
:root{--bg:#0B162B;--bg2:#0F192E;--card:#16233F;--card2:#1E2F52;--line:#2B3E63;--tx:#F2F2F2;--mut:#A9B3B7;--mut2:#D1D6DB;
 --cian:#17DFE9;--azul:#48C8E5;--azul2:#1C4587;--v:#2FBF71;--a:#F2B544;--r:#E5484D;--na:#6B7A90;--rad:10px}
*{box-sizing:border-box}html,body{margin:0}html{background:#0B162B}
body{min-height:100vh;font-family:Arial,Helvetica,sans-serif;background:linear-gradient(160deg,#0B162B 0%,#0F192E 55%,#0E1A31 100%);color:var(--tx);font-size:14px;line-height:1.45}
button,select,input,textarea{font:inherit;color:inherit}
a{color:var(--cian)}
.top{position:sticky;top:0;z-index:20;background:rgba(11,22,43,.96);border-bottom:1px solid var(--line);backdrop-filter:blur(6px)}
.top .in{max-width:1240px;margin:0 auto;padding:10px 16px;display:flex;align-items:center;gap:16px;flex-wrap:wrap}
.logo{height:26px;width:auto;display:block}
.grp{display:flex;flex-direction:column;line-height:1.2}.grp b{font-size:15px}.grp span{font-size:12px;color:var(--mut)}
nav{display:flex;gap:4px;flex-wrap:wrap;margin-left:auto}
nav button{background:none;border:1px solid transparent;border-radius:999px;padding:7px 14px;cursor:pointer;color:var(--mut2);font-weight:700;font-size:13px}
nav button:hover{border-color:var(--line)}nav button[aria-current="page"]{background:var(--azul2);color:#fff}
nav .pres{border-color:var(--cian);color:var(--cian)}
main{max-width:1240px;margin:0 auto;padding:22px 16px 60px}
h1,h2,h3{margin:0;font-weight:700}
h1{font-family:'Arial Black',Arial,sans-serif;font-size:30px;letter-spacing:.2px}
h2{font-size:20px;margin:30px 0 4px}h3{font-size:16px}
.sub{color:var(--cian);font-weight:700;font-size:13px;margin:2px 0 14px}
.mut{color:var(--mut)}.mut2{color:var(--mut2)}
.card{background:var(--card);border-radius:var(--rad);padding:16px}
.grid{display:grid;gap:12px}
.g4{grid-template-columns:1.3fr 1fr 1fr 1fr}.g3{grid-template-columns:repeat(3,1fr)}.g2{grid-template-columns:repeat(2,1fr)}
.sem{display:inline-flex;align-items:center;gap:6px;font-weight:700;font-size:12px}
.sem i{width:11px;height:11px;border-radius:50%;display:inline-block;background:var(--na);flex:none}
.s-verde i{background:var(--v)}.s-amarillo i{background:var(--a)}.s-rojo i{background:var(--r)}
.c-verde{color:var(--v)}.c-amarillo{color:var(--a)}.c-rojo{color:var(--r)}.c-na,.c-sin{color:var(--na)}
.big{font-family:'Arial Black',Arial,sans-serif;font-size:34px;line-height:1.1}
.mid{font-size:24px;font-weight:700}
.lab{font-size:12px;color:var(--mut2);margin-top:8px}
.tag{font-size:11px;font-weight:700;letter-spacing:.6px;text-transform:uppercase;color:var(--cian)}
.dcard{cursor:pointer;border:1px solid transparent;text-align:left;display:flex;flex-direction:column;gap:6px;padding:0;overflow:hidden}
.dcard:hover,.dcard:focus-visible{border-color:var(--cian);outline:none}
.dcard .img{height:120px;background:#000 center/cover no-repeat;position:relative}
.dcard .body{padding:12px 16px 16px;display:flex;flex-direction:column;gap:6px}
.row{display:flex;gap:16px;align-items:baseline;flex-wrap:wrap}
.msg{background:var(--azul2);border-radius:var(--rad);padding:12px 16px;font-weight:700;margin:12px 0}
.msg b{color:var(--cian)}
table{border-collapse:collapse;width:100%}
th,td{padding:8px 10px;text-align:left;border-bottom:1px solid var(--line);vertical-align:top}
th{background:var(--azul2);font-size:12px;position:sticky;top:0}
tbody tr:nth-child(odd) td{background:rgba(30,47,82,.45)}
.mat td.c{text-align:center}.mat td button{background:none;border:0;cursor:pointer;padding:2px 6px;border-radius:6px}
.mat td button:hover{background:var(--card2)}
.tw{overflow-x:auto;border-radius:var(--rad)}
.hero{position:relative;border-radius:14px;overflow:hidden;min-height:300px;background:#000 center/cover no-repeat;display:flex;align-items:flex-end}
.hero:before{content:"";position:absolute;inset:0;background:linear-gradient(90deg,rgba(16,16,38,.95) 0%,rgba(16,16,38,.75) 45%,rgba(25,25,59,0) 75%),linear-gradient(180deg,rgba(16,16,38,0) 40%,rgba(16,16,38,.85) 100%)}
.hero .in{position:relative;padding:24px;max-width:780px}
.hero h1{font-size:40px}
.badge{position:absolute;top:12px;right:12px;background:rgba(0,0,0,.6);color:#fff;font-size:11px;font-weight:700;padding:4px 8px;border-radius:6px}
.okstrip{background:var(--card);border-radius:var(--rad);padding:10px 14px;margin-top:14px;font-size:13px;color:var(--mut2)}
.okstrip b{color:var(--v)}
.filtros{display:flex;flex-wrap:wrap;gap:8px;align-items:center;margin:18px 0 6px;position:sticky;top:52px;z-index:5;background:rgba(11,22,43,.96);padding:8px 0}
.filtros select,.filtros input{background:var(--card);border:1px solid var(--line);border-radius:8px;padding:7px 10px}
.chip{background:var(--card);border:1px solid var(--line);border-radius:999px;padding:6px 12px;cursor:pointer;font-weight:700;font-size:12.5px;color:var(--mut2)}
.chip[aria-pressed="true"]{background:var(--azul2);border-color:var(--azul2);color:#fff}
.icards{display:grid;grid-template-columns:repeat(auto-fill,minmax(250px,1fr));gap:12px;margin-top:12px}
.icard{background:var(--card);border-radius:var(--rad);padding:14px;border:1px solid transparent;cursor:pointer;text-align:left}
.icard:hover,.icard[aria-pressed="true"]{border-color:var(--cian)}
.icard .val{font-size:22px;font-weight:700;float:right}
.icard p{margin:8px 0 0;font-size:13px;color:var(--mut2)}
.area{background:var(--card);border-radius:var(--rad);padding:14px 16px;margin-top:12px}
.area .hd{display:flex;align-items:center;gap:10px;flex-wrap:wrap}
.area .hd .pc{margin-left:auto;font-size:20px;font-weight:700}
.bar{height:10px;background:var(--line);border-radius:6px;overflow:hidden;margin:10px 0 4px}
.bar span{display:block;height:100%;border-radius:6px}
.items{display:grid;grid-template-columns:repeat(auto-fill,minmax(300px,1fr));gap:8px;margin-top:10px}
.item{background:var(--card2);border:1px solid transparent;border-radius:8px;padding:10px 12px;text-align:left;cursor:pointer;display:flex;flex-direction:column;gap:3px}
.item:hover,.item:focus-visible{border-color:var(--cian);outline:none}
.item .t{display:flex;gap:8px;align-items:center;font-weight:700}
.item .x{font-size:12.5px;color:var(--mut2)}
.item .pl{font-size:11.5px;color:var(--cian);font-weight:700}
.mini{font-size:12px;color:var(--mut)}
.esc{border:1px solid var(--cian);border-radius:8px;padding:10px 12px;margin-top:10px;font-size:13px}
.esc b{color:var(--cian)}
.fotos{display:grid;grid-template-columns:repeat(auto-fill,minmax(240px,1fr));gap:12px;margin-top:10px}
.foto{background:var(--card);border-radius:var(--rad);overflow:hidden;position:relative}
.foto img{width:100%;height:180px;object-fit:cover;display:block}
.foto p{margin:0;padding:8px 10px;font-size:13px}
.vacs{display:flex;flex-direction:column;gap:6px}
.pbar{display:grid;grid-template-columns:110px 1fr 70px;gap:10px;align-items:center;font-size:13px}
.plan td{font-size:13px}.plan .ed{width:100%;background:var(--bg);border:1px solid var(--line);border-radius:6px;padding:5px 7px;min-width:110px}
.pv{display:none}.plan textarea.ed{min-width:170px;resize:vertical;min-height:34px}
.prio{color:var(--r);font-weight:700;font-size:12px}
.btn{background:var(--azul2);border:0;border-radius:8px;padding:8px 14px;font-weight:700;cursor:pointer;color:#fff}
.btn.sec{background:var(--card);border:1px solid var(--line)}
.btn:hover{filter:brightness(1.15)}
.acciones{display:flex;gap:8px;flex-wrap:wrap;margin-left:auto}
.vacio{padding:24px;text-align:center;color:var(--mut)}
/* detalle */
.velo{position:fixed;inset:0;background:rgba(0,0,0,.55);z-index:40;display:none}
.det{position:fixed;top:0;right:0;bottom:0;width:min(480px,100%);background:var(--bg2);z-index:41;transform:translateX(100%);transition:transform .2s;overflow-y:auto;border-left:1px solid var(--line);padding:20px}
.det{visibility:hidden}.abierto .velo{display:block}.abierto .det{transform:none;visibility:visible}
.det h3{font-size:20px;margin:6px 0 8px}
.det dl{display:grid;grid-template-columns:110px 1fr;gap:6px 10px;margin:12px 0;font-size:13px}.det dt{color:var(--mut)}
.det img{width:100%;border-radius:8px;margin-top:8px}
.det label{display:block;font-size:12px;color:var(--mut);margin-top:10px}
.det .ed{width:100%;background:var(--card);border:1px solid var(--line);border-radius:6px;padding:7px 9px}
.det .nav{display:flex;justify-content:space-between;gap:8px;margin-top:16px}
.cerrar{float:right;background:none;border:1px solid var(--line);border-radius:8px;padding:4px 10px;cursor:pointer}
/* presentación */
#show{position:fixed;inset:0;background:#05070d;z-index:60;display:none;align-items:center;justify-content:center}
#show.on{display:flex}
#stage{width:1280px;height:720px;position:relative;overflow:hidden;transform-origin:center center;background:linear-gradient(160deg,#0B162B,#0F192E);border-radius:4px}
.sl{position:absolute;inset:0;padding:64px 72px 48px;display:flex;flex-direction:column}
.sl .logo{position:absolute;top:28px;left:36px;height:32px}
.sl h1{font-size:36px;margin-top:40px}.sl .sub{font-size:18px;margin-bottom:24px}
.sl .card{padding:22px}.sl .big{font-size:56px}.sl .mid{font-size:34px}.sl .lab{font-size:16px}
.sl .item{font-size:16px}.sl .x{font-size:14px}
.sl.foto-bg{background:#000 center/cover no-repeat;justify-content:flex-end}
.sl.foto-bg:before{content:"";position:absolute;inset:0;background:linear-gradient(90deg,rgba(16,16,38,.95),rgba(16,16,38,.6) 55%,rgba(25,25,59,0) 80%),linear-gradient(180deg,rgba(16,16,38,0) 45%,rgba(16,16,38,.9))}
.sl.foto-bg>*{position:relative}
.sl.foto-bg h1{font-size:60px}
.sbar{display:grid;grid-template-columns:260px 1fr 90px;gap:14px;align-items:center;font-size:18px;margin:8px 0}
.sbar .bar{height:22px;margin:0}
.ctl{position:fixed;bottom:14px;left:50%;transform:translateX(-50%);display:flex;gap:8px;z-index:61;align-items:center;color:#fff;font-size:13px}
.ctl button{background:rgba(255,255,255,.12);border:0;border-radius:8px;padding:8px 14px;color:#fff;cursor:pointer;font-weight:700}
@media (max-width:900px){.g4{grid-template-columns:1fr 1fr}.g3{grid-template-columns:1fr}}
@media (max-width:700px){.g4,.g2{grid-template-columns:1fr}h1{font-size:24px}.hero h1{font-size:28px}.hero{min-height:240px}
 nav{margin-left:0;width:100%;overflow-x:auto;flex-wrap:nowrap}nav button{white-space:nowrap}.filtros{top:96px}
 .items{grid-template-columns:1fr}.pbar{grid-template-columns:90px 1fr 56px}}
@media print{
 html{background:#fff}body{background:#fff;color:#14202E;font-size:11px}.top,.filtros,.acciones,.noprint,#show,.velo,.det{display:none!important}
 main{max-width:none;padding:0}.card,.area,.item,.icard,.okstrip,.foto{background:#fff;border:1px solid #ccd3dc;color:#14202E;break-inside:avoid}
 .mut2,.item .x,.icard p,.okstrip{color:#3A495A}.sub,.tag,.item .pl{color:#1C4587}th{background:#1C4587;color:#fff;position:static}
 tbody tr:nth-child(odd) td{background:#f3f5f8}.plan .ed{border:none;background:none;padding:0;min-width:0}
 .plan .ed{display:none!important}.pv{display:block!important;white-space:pre-wrap}
 .hero{min-height:200px}.msg{background:#e8eef6;color:#14202E}.msg b{color:#1C4587}
 .sem i,.bar span,.hero,th{-webkit-print-color-adjust:exact;print-color-adjust:exact}
}
"""

JS = r"""
const D=JSON.parse(document.getElementById('datos').textContent);
const F=D.fotos, DS=D.dists, POR={};DS.forEach(d=>POR[d.id]=d);
const PEND={};DS.forEach(d=>d.pendientes.forEach(p=>PEND[p.id]=p));
const TXT={verde:'Cumple',amarillo:'En proceso',rojo:'No cumple',na:'No aplica',sin:'Sin información'};
const CLAVE='estatus-propuesta-'+D.grupo_id+'-'+D.corte;
const ST={vista:'resumen',fInd:'',fEst:'',pDist:'',pInd:'',pEst:'',pPlan:'',q:''};
const $=s=>document.querySelector(s);
const E=s=>String(s==null?'':s).replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const pc=(v,d=1)=>v==null?'—':v.toFixed(d)+'%';
const sem=(e,t)=>`<span class="sem s-${e}"><i></i>${t===undefined?TXT[e]:t}</span>`;
const bar=(v,e)=>`<div class="bar"><span style="width:${Math.max(0,Math.min(100,v||0))}%;background:var(--${e==='verde'?'v':e==='amarillo'?'a':'r'})"></span></div>`;
function leer(){try{return JSON.parse(localStorage.getItem(CLAVE)||'{}')}catch(e){return {}}}
function guardar(d){try{localStorage.setItem(CLAVE,JSON.stringify(d))}catch(e){}}
let PLAN=leer();
function setPlan(id,f,v){PLAN[id]=PLAN[id]||{};PLAN[id][f]=v;guardar(PLAN)}
const plan=id=>PLAN[id]||{};
function planTxt(id){const p=plan(id),x=[];if(p.estatus)x.push(p.estatus);if(p.resp)x.push(p.resp);if(p.fecha)x.push(p.fecha.split('-').reverse().join('/'));return x.join(' · ')}
const fotoSrc=k=>k&&F[k]?F[k].src:'';

// ------------------------------------------------------------------ vistas
function vResumen(){const t=D.total;
 let h=`<h1>${E(D.grupo)}</h1><p class="sub">Corte al ${E(D.corte_txt)} · 11 indicadores con el mismo peso · ${t.pendientes} pendientes, todos de prioridad alta</p>`;
 h+=`<div class="grid g4"><div class="card"><div class="tag">Total del grupo</div>
  <div class="lab">Calificación general</div><div class="big">${pc(t.general)}</div>${sem(t.estado_general,TXT[t.estado_general]+' · '+t.n_indicadores+' indicadores')}
  <div class="lab">Desarrollo de proyecto</div><div class="big">${pc(t.proyecto)}</div>${sem(t.estado_proyecto,TXT[t.estado_proyecto]+' · '+t.elementos+' elementos')}
  <div class="lab" style="margin-top:14px"><b style="font-size:22px;color:#fff">${t.pendientes}</b> pendientes en el grupo</div><button class="btn" style="margin-top:12px" data-go="plan">Ir al plan de trabajo</button></div>`;
 DS.forEach(d=>{h+=`<button class="card dcard" data-go="${d.id}" aria-label="Ver ${E(d.nombre)}">
  <div class="img" style="background-image:url('${fotoSrc(d.fachada)}')"></div><div class="body">
  <div class="tag">${E(d.nombre)}</div>
  <div class="row"><div><div class="lab">General</div><div class="mid c-${d.estado_general}">${pc(d.general)}</div>${sem(d.estado_general)}</div>
  <div><div class="lab">Proyecto</div><div class="mid c-${d.proyecto.estado}">${pc(d.proyecto.valor)}</div>${sem(d.proyecto.estado)}</div></div>
  <div><b>${d.n_pendientes}</b> <span class="mut">pendientes</span></div>${d.mensaje?`<div class="c-amarillo" style="font-weight:700;font-size:13px">${E(d.mensaje)}</div>`:''}</div></button>`});
 h+=`</div><h2>Los 11 indicadores</h2><p class="sub">Clic en una celda para ver los pendientes de ese indicador</p><div class="tw"><table class="mat"><thead><tr><th>Indicador</th>${DS.map(d=>`<th style="text-align:center">${E(d.corto)}</th>`).join('')}</tr></thead><tbody>`;
 DS[0].resumen.forEach((r,i)=>{h+=`<tr><td>${E(r.nombre)}</td>`+DS.map(d=>{const e=d.resumen[i];
   return e.estado==='na'?`<td class="c mut">No aplica</td>`:`<td class="c"><button data-go="${d.id}" data-ind="${e.id}">${sem(e.estado,pc(e.valor))}</button></td>`}).join('')+'</tr>'});
 h+=`<tr><td><b>Calificación general</b></td>${DS.map(d=>`<td class="c"><b>${sem(d.estado_general,pc(d.general))}</b></td>`).join('')}</tr></tbody></table></div>`;
 h+=`<h2>Mobiliario: cuánto sube si se valida el actual</h2><p class="sub">Hoy los ${t.mobiliario} elementos de mobiliario están en 0% en el KPI</p><div class="tw"><table><thead><tr><th>Agencia</th><th>Elementos</th><th>Proyecto hoy → validado</th><th>General hoy → validado</th></tr></thead><tbody>`;
 DS.forEach(d=>{const e=d.escenario;if(!e)return;h+=`<tr><td><b>${E(d.corto)}</b></td><td>${e.elementos}</td><td>${pc(e.proyecto_hoy)} → <b class="c-verde">${pc(e.proyecto_validado)}</b> <span class="mut">(+${(e.proyecto_validado-e.proyecto_hoy).toFixed(1)})</span></td><td>${pc(e.general_hoy)} → <b class="c-verde">${pc(e.general_validado)}</b> <span class="mut">(+${(e.general_validado-e.general_hoy).toFixed(1)})</span></td></tr>`});
 h+=`<tr><td><b>Grupo</b></td><td>${t.mobiliario}</td><td>${pc(t.proyecto)} → <b class="c-verde">${pc(t.proyecto_validado)}</b></td><td>${pc(t.general)} → <b class="c-verde">${pc(t.general_validado)}</b></td></tr></tbody></table></div>
 <p class="mini">Allende sólo cuenta si se acepta el mobiliario de la oficina móvil.</p>`;
 h+=`<h2>Puntos a validar en la visita</h2><div class="grid g2" style="margin-top:10px">`+DS.flatMap(d=>d.observaciones.map(o=>`<div class="card"><div class="tag">${E(d.corto)}</div><p style="margin:6px 0 0">${E(o.texto)}</p></div>`)).join('')+'</div>';
 return h}

function filtra(p){return (!ST.fInd||p.ind===ST.fInd)&&(!ST.fEst||p.estado===ST.fEst)&&(!ST.q||(p.titulo+' '+p.area+' '+p.texto).toLowerCase().includes(ST.q.toLowerCase()))}
function itemHTML(p){const pt=planTxt(p.id);
 return `<button class="item" data-pend="${p.id}"><span class="t">${sem(p.estado,'')}${E(p.titulo)}${p.estado==='amarillo'&&p.falta.startsWith('Parcial')?` <span class="c-amarillo" style="font-size:12px">${E(p.falta.replace('Parcial: ',''))}</span>`:''}${p.foto?' <span class="mut" style="font-size:11px">· foto</span>':''}</span>
  ${p.texto?`<span class="x">${E(p.texto)}</span>`:(p.ind!=='proyecto'&&!p.falta.startsWith('No cumple')&&!p.falta.startsWith('Parcial')?`<span class="x">${E(p.falta)}</span>`:'')}${pt?`<span class="pl">Plan: ${E(pt)}</span>`:''}</button>`}

function vDist(d){
 let h=`<section class="hero" style="background-image:url('${fotoSrc(d.fachada)}')">${d.fachada&&F[d.fachada].no_reciente?'<span class="badge">Foto no reciente</span>':''}
 <div class="in"><h1>${E(d.nombre.toUpperCase())}</h1><div class="row" style="margin-top:8px;font-size:17px;font-weight:700">
 <span>General <span class="c-${d.estado_general}">${pc(d.general)}</span></span><span>Proyecto <span class="c-${d.proyecto.estado}">${pc(d.proyecto.valor)}</span></span><span>${d.n_pendientes} pendientes</span></div>
 <div class="row" style="margin-top:6px">${sem(d.estado_general,'General: '+TXT[d.estado_general])}${sem(d.proyecto.estado,'Proyecto: '+TXT[d.proyecto.estado])}</div></div></section>`;
 if(d.mensaje)h+=`<div class="msg"><b>Clave:</b> ${E(d.mensaje)}</div>`;
 const ok=d.resumen.filter(r=>r.estado==='verde'),na=d.resumen.filter(r=>r.estado==='na'),pend=d.resumen.filter(r=>r.falta);
 h+=`<div class="okstrip"><b>Cumplen al 100%:</b> ${ok.map(r=>E(r.nombre)).join(' · ')||'—'}${na.length?` <span class="mut">· No aplican: ${na.map(r=>E(r.nombre)).join(', ')} (${E(d.motivo_na)})</span>`:''}</div>`;
 h+=`<div class="icards">`+pend.map(r=>`<button class="icard" data-ind="${r.id}" aria-pressed="${ST.fInd===r.id}">${sem(r.estado,'')}<b style="margin-left:4px">${E(r.nombre)}</b><span class="val c-${r.estado}">${pc(r.valor)}</span>
  <p><b class="c-azul" style="color:var(--cian)">Qué falta:</b> ${E(r.falta)}</p>${r.necesita?`<p>${E(r.necesita)}</p>`:''}</button>`).join('')+'</div>';
 // filtros
 const inds=[...new Set(d.pendientes.map(p=>p.ind))];
 h+=`<div class="filtros"><b>Pendientes</b><select id="fInd" aria-label="Filtrar por indicador"><option value="">Todos los indicadores</option>${inds.map(i=>`<option value="${i}" ${ST.fInd===i?'selected':''}>${E(d.pendientes.find(p=>p.ind===i).indicador)}</option>`).join('')}</select>
  ${[['','Todos'],['rojo','No cumple'],['amarillo','En proceso (parcial)']].map(([v,t])=>`<button class="chip" data-est="${v}" aria-pressed="${ST.fEst===v}">${t}</button>`).join('')}
  <input id="q" type="search" placeholder="Buscar…" value="${E(ST.q)}" aria-label="Buscar pendiente"><span class="mut" id="cuenta"></span></div>`;
 let n=0;
 // proyecto por área
 if(!ST.fInd||ST.fInd==='proyecto'){
  const pp=d.proyecto;h+=`<h2>Desarrollo de proyecto por área</h2><p class="sub">${pc(pp.valor)} · ${pp.cumple} de ${pp.n} elementos completos · faltan ${pp.faltan}</p>`;
  const def=d.pendientes.filter(p=>p.area==='Proyecto de la agencia'&&filtra(p));n+=def.length;
  if(def.length)h+=`<div class="items">${def.map(itemHTML).join('')}</div>`;
  d.areas.forEach(a=>{const its=a.items.map(i=>PEND[i]).filter(filtra);n+=its.length;
   if(!a.faltan){h+=`<div class="area" style="padding:10px 16px"><div class="hd">${sem('verde','')}<b>${E(a.nombre)}</b><span class="mut">cumple los ${a.n} elementos</span><span class="pc c-verde">100%</span></div></div>`;return}
   if(!its.length&&(ST.fEst||ST.q))return;
   h+=`<div class="area"><div class="hd">${sem(a.estado,'')}<h3>${E(a.nombre)}</h3><span class="mut">${a.n} elementos · cumple ${a.cumple} · faltan ${a.faltan}</span><span class="pc c-${a.estado}">${pc(a.valor)}</span></div>${bar(a.valor,a.estado)}
    ${a.mobiliario&&d.escenario?`<div class="esc"><b>Si se valida el mobiliario actual como OK:</b> proyecto ${pc(d.escenario.proyecto_hoy)} → <b>${pc(d.escenario.proyecto_validado)}</b> · general ${pc(d.escenario.general_hoy)} → <b>${pc(d.escenario.general_validado)}</b>${d.mensaje?' (sólo si se acepta el mobiliario de la oficina móvil)':''}</div>`:''}
    <div class="items">${its.map(itemHTML).join('')}</div></div>`});}
 // demás indicadores
 const otros=[...new Set(d.pendientes.filter(p=>p.ind!=='proyecto').map(p=>p.ind))].filter(i=>!ST.fInd||ST.fInd===i);
 otros.forEach(i=>{const its=d.pendientes.filter(p=>p.ind===i&&filtra(p));n+=its.length;if(!its.length)return;const r=d.resumen.find(x=>x.id===i);
  h+=`<div class="area"><div class="hd">${sem(r.estado,'')}<h3>${E(r.nombre)}</h3><span class="mut">${E(r.falta)}</span><span class="pc c-${r.estado}">${pc(r.valor)}</span></div>${bar(r.valor,r.estado)}`;
  if(i==='personal'&&d.personal){h+=`<div class="vacs" style="margin:10px 0">${d.personal.areas.map(a=>`<div class="pbar"><span>${E(a.area)}</span>${bar(a.pct,a.pct>=80?'verde':a.pct>=50?'amarillo':'rojo')}<span>${a.cub} de ${a.req}</span></div>`).join('')}</div>`}
  h+=`<div class="items">${its.map(itemHTML).join('')}</div></div>`});
 if(!n)h+=`<div class="vacio">No hay pendientes con este filtro.</div>`;
 // evidencia
 const fts=d.pendientes.filter(p=>p.foto).map(p=>({k:p.foto,t:p.titulo+': '+(F[p.foto].pie||''),id:p.id})).concat(d.observaciones.filter(o=>o.foto).map(o=>({k:o.foto,t:'Para validar: '+(F[o.foto].pie||'')})));
 if(fts.length&&!ST.fInd)h+=`<h2>Evidencia</h2><p class="sub">Fotos de lo que falta o de lo que hay que validar</p><div class="fotos">${fts.map(f=>`<div class="foto">${F[f.k].no_reciente?'<span class="badge">Foto no reciente</span>':''}<img src="${F[f.k].src}" alt="${E(f.t)}"><p>${E(f.t)}</p></div>`).join('')}</div>`;
 if(d.observaciones.length&&!ST.fInd)h+=`<h2>Para validar en la visita</h2>`+d.observaciones.map(o=>`<div class="card" style="margin-top:8px">${E(o.texto)}</div>`).join('');
 setTimeout(()=>{const c=$('#cuenta');if(c)c.textContent=n+' de '+d.pendientes.length},0);
 return h}

function vPlan(){
 const ps=DS.flatMap(d=>d.pendientes).filter(p=>(!ST.pDist||p.dist===ST.pDist)&&(!ST.pInd||p.ind===ST.pInd)&&(!ST.pEst||p.estado===ST.pEst)&&(!ST.pPlan||(ST.pPlan==='Sin asignar'?!plan(p.id).estatus:plan(p.id).estatus===ST.pPlan)));
 const inds=[...new Map(DS.flatMap(d=>d.pendientes).map(p=>[p.ind,p.indicador])).entries()];
 let h=`<h1>Plan de trabajo</h1><p class="sub">Todo lo que no está al 100% es prioridad alta. Lo que escribas aquí se guarda en este navegador.</p>
 <div class="filtros"><select id="pDist" aria-label="Sucursal"><option value="">Todas las sucursales</option>${DS.map(d=>`<option value="${d.id}" ${ST.pDist===d.id?'selected':''}>${E(d.nombre)}</option>`).join('')}</select>
 <select id="pInd" aria-label="Indicador"><option value="">Todos los indicadores</option>${inds.map(([i,n])=>`<option value="${i}" ${ST.pInd===i?'selected':''}>${E(n)}</option>`).join('')}</select>
 <select id="pEst" aria-label="Estado"><option value="">Todo estado</option><option value="rojo" ${ST.pEst==='rojo'?'selected':''}>No cumple</option><option value="amarillo" ${ST.pEst==='amarillo'?'selected':''}>En proceso (parcial)</option></select>
 <select id="pPlan" aria-label="Estatus del plan"><option value="">Todo estatus</option>${['Sin asignar','Pendiente','En proceso','Cerrado'].map(s=>`<option ${ST.pPlan===s?'selected':''}>${s}</option>`).join('')}</select>
 <span class="mut">${ps.length} renglones</span><div class="acciones"><button class="btn" id="xls">Exportar a Excel</button><button class="btn sec" id="imp">Imprimir / PDF</button></div></div>
 <div class="tw"><table class="plan"><thead><tr><th>Sucursal</th><th>Indicador</th><th>Pendiente</th><th>Qué falta</th><th>Prioridad</th><th>Responsable</th><th>Fecha compromiso</th><th>Estatus</th><th>Comentarios</th></tr></thead><tbody>`;
 ps.forEach(p=>{const v=plan(p.id);h+=`<tr data-id="${p.id}"><td>${E(POR[p.dist].corto)}</td><td>${E(p.indicador)}${p.area&&p.area!==p.indicador?`<div class="mini">${E(p.area)}</div>`:''}</td>
  <td><button class="item" style="padding:4px 8px" data-pend="${p.id}"><span class="t">${sem(p.estado,'')}${E(p.titulo)}</span></button></td><td class="mut2">${E(p.falta)}${p.texto?`<div class="mini">${E(p.texto)}</div>`:''}</td><td class="prio">ALTA</td>
  <td><input class="ed" data-f="resp" value="${E(v.resp||'')}" aria-label="Responsable"></td><td><input class="ed" type="date" data-f="fecha" value="${E(v.fecha||'')}" aria-label="Fecha compromiso"></td>
  <td><select class="ed" data-f="estatus" aria-label="Estatus"><option value=""></option>${['Pendiente','En proceso','Cerrado'].map(s=>`<option ${v.estatus===s?'selected':''}>${s}</option>`).join('')}</select></td>
  <td><textarea class="ed" data-f="coment" rows="1" aria-label="Comentarios">${E(v.coment||'')}</textarea></td></tr>`});
 h+=`</tbody></table></div>${ps.length?'':'<div class="vacio">No hay renglones con este filtro.</div>'}`;
 return h}

// ------------------------------------------------------------------ detalle
let LISTA=[];
function abrir(id){const p=PEND[id];if(!p)return;const d=POR[p.dist],v=plan(id);
 LISTA=[...document.querySelectorAll('[data-pend]')].map(b=>b.dataset.pend).filter((x,i,a)=>a.indexOf(x)===i);const i=LISTA.indexOf(id);
 $('#det').innerHTML=`<button class="cerrar" id="cerrar">Cerrar</button><div class="tag">${E(d.nombre)} · ${E(p.indicador)}</div><h3>${E(p.titulo)}</h3>${sem(p.estado)}
  <dl><dt>Área</dt><dd>${E(p.area||'—')}</dd><dt>Qué falta</dt><dd>${E(p.falta)}</dd>${p.texto?`<dt>Qué implica</dt><dd>${E(p.texto)}</dd>`:''}<dt>Avance</dt><dd>${p.avance}%</dd><dt>Prioridad</dt><dd class="prio">ALTA</dd></dl>
  ${p.foto?`<div style="position:relative">${F[p.foto].no_reciente?'<span class="badge">Foto no reciente</span>':''}<img src="${F[p.foto].src}" alt="${E(F[p.foto].pie)}"><p class="mini">${E(F[p.foto].pie)}</p></div>`:''}
  <div data-id="${id}"><label>Responsable<input class="ed" data-f="resp" value="${E(v.resp||'')}"></label><label>Fecha compromiso<input class="ed" type="date" data-f="fecha" value="${E(v.fecha||'')}"></label>
  <label>Estatus<select class="ed" data-f="estatus"><option value=""></option>${['Pendiente','En proceso','Cerrado'].map(s=>`<option ${v.estatus===s?'selected':''}>${s}</option>`).join('')}</select></label>
  <label>Comentarios<textarea class="ed" data-f="coment" rows="3">${E(v.coment||'')}</textarea></label></div>
  <div class="nav"><button class="btn sec" data-pend="${LISTA[i-1]||''}" ${i>0?'':'disabled'}>← Anterior</button><span class="mini">${i+1} de ${LISTA.length}</span><button class="btn sec" data-pend="${LISTA[i+1]||''}" ${i<LISTA.length-1?'':'disabled'}>Siguiente →</button></div>`;
 document.body.classList.add('abierto');$('#cerrar').focus()}
function cerrar(){if(!document.body.classList.contains('abierto'))return;document.body.classList.remove('abierto');render(true)}

// ------------------------------------------------------------------ navegación
function render(conservar){const y=window.scrollY;
 $('#main').innerHTML=ST.vista==='resumen'?vResumen():ST.vista==='plan'?vPlan():vDist(POR[ST.vista]);
 document.querySelectorAll('nav [data-go]').forEach(b=>b.setAttribute('aria-current',b.dataset.go===ST.vista?'page':'false'));
 if(conservar)window.scrollTo(0,y)}
function ir(v,ind){ST.vista=v;ST.fInd=ind||'';ST.fEst='';ST.q='';try{history.replaceState(null,'','#'+v)}catch(e){}render();window.scrollTo(0,0)}
document.addEventListener('click',e=>{const t=e.target.closest('[data-go],[data-pend],[data-ind],[data-est],#xls,#imp,#cerrar,.velo,[data-sl]');if(!t)return;
 if(t.classList.contains('velo')||t.id==='cerrar')return cerrar();
 if(t.id==='xls')return exportar();if(t.id==='imp')return window.print();
 if(t.dataset.sl!==undefined)return mover(+t.dataset.sl);
 if(t.dataset.pend!==undefined){if(t.dataset.pend)abrir(t.dataset.pend);return}
 if(t.dataset.go){if(t.dataset.go==='show')return presentar();return ir(t.dataset.go,t.dataset.ind)}
 if(t.dataset.ind!==undefined){ST.fInd=ST.fInd===t.dataset.ind?'':t.dataset.ind;render(true);return}
 if(t.dataset.est!==undefined){ST.fEst=t.dataset.est;render(true)}});
document.addEventListener('change',e=>{const t=e.target;if(['fInd','pDist','pInd','pEst','pPlan'].includes(t.id)){ST[t.id]=t.value;render(true);return}
 const r=t.closest('[data-id]');if(r&&t.dataset.f){setPlan(r.dataset.id,t.dataset.f,t.value)}});
document.addEventListener('input',e=>{const t=e.target;if(t.id==='q'){ST.q=t.value;clearTimeout(window._q);window._q=setTimeout(()=>{render(true);const q=$('#q');if(q){q.focus();q.setSelectionRange(q.value.length,q.value.length)}},250);return}
 const r=t.closest('[data-id]');if(r&&t.dataset.f&&t.tagName!=='SELECT')setPlan(r.dataset.id,t.dataset.f,t.value)});
document.addEventListener('keydown',e=>{if($('#show').classList.contains('on')){if(['ArrowRight','PageDown',' '].includes(e.key)){e.preventDefault();mover(1)}else if(['ArrowLeft','PageUp'].includes(e.key)){e.preventDefault();mover(-1)}else if(e.key==='Escape')salir();return}
 if(e.key==='Escape')cerrar()});

// ------------------------------------------------------------------ presentación
let SL=[],SI=0;const LOGO=document.querySelector('.top .logo').src;
function slides(){const t=D.total,s=[];const lg=`<img class="logo" src="${LOGO}" alt="FOTON">`;
 s.push(`<div class="sl">${lg}<h1>Resumen ejecutivo</h1><p class="sub">${E(D.grupo)} · corte al ${E(D.corte_txt)}</p><div class="grid g4" style="flex:1">
  <div class="card"><div class="tag">Total del grupo</div><div class="lab">General</div><div class="big">${pc(t.general)}</div>${sem(t.estado_general)}<div class="lab">Proyecto</div><div class="big">${pc(t.proyecto)}</div>${sem(t.estado_proyecto)}</div>
  ${DS.map(d=>`<div class="card"><div class="tag">${E(d.nombre)}</div><div class="lab">General</div><div class="mid c-${d.estado_general}">${pc(d.general)}</div>${sem(d.estado_general)}<div class="lab">Proyecto</div><div class="mid c-${d.proyecto.estado}">${pc(d.proyecto.valor)}</div>${sem(d.proyecto.estado)}<div class="lab"><b style="font-size:26px;color:#fff">${d.n_pendientes}</b> pendientes</div>${d.mensaje?`<p class="c-amarillo" style="font-weight:700">${E(d.mensaje)}</p>`:''}</div>`).join('')}</div></div>`);
 DS.forEach(d=>{
  s.push(`<div class="sl foto-bg" style="background-image:url('${fotoSrc(d.fachada)}')">${lg}<h1>${E(d.nombre.toUpperCase())}</h1><div class="row" style="font-size:24px;font-weight:700"><span>General <span class="c-${d.estado_general}">${pc(d.general)}</span></span><span>Proyecto <span class="c-${d.proyecto.estado}">${pc(d.proyecto.valor)}</span></span><span>${d.n_pendientes} pendientes</span></div>${d.mensaje?`<p style="font-size:22px;font-weight:700;color:var(--cian)">${E(d.mensaje)}</p>`:''}</div>`);
  const pend=d.resumen.filter(r=>r.falta);
  s.push(`<div class="sl">${lg}<h1>${E(d.corto)}: qué falta por indicador</h1><p class="sub">Cumplen al 100%: ${d.resumen.filter(r=>r.estado==='verde').map(r=>E(r.nombre)).join(' · ')}</p><div class="grid ${pend.length>2?'g2':'g2'}">${pend.map(r=>`<div class="card"><div class="row" style="justify-content:space-between"><b style="font-size:22px">${E(r.nombre)}</b><span class="mid c-${r.estado}">${pc(r.valor)}</span></div><p style="font-size:17px"><b style="color:var(--cian)">Qué falta:</b> ${E(r.falta)}</p>${r.necesita?`<p style="font-size:15px" class="mut2">${E(r.necesita)}</p>`:''}</div>`).join('')}</div></div>`);
  s.push(`<div class="sl">${lg}<h1>${E(d.corto)}: desarrollo de proyecto por área</h1><p class="sub">${pc(d.proyecto.valor)} · ${d.proyecto.cumple} de ${d.proyecto.n} elementos completos · faltan ${d.proyecto.faltan}</p>
   ${d.areas.map(a=>`<div class="sbar"><span>${E(a.nombre)} <span class="mut" style="font-size:14px">${a.cumple}/${a.n}</span></span>${bar(a.valor,a.estado)}<b class="c-${a.estado}">${pc(a.valor)}</b></div>`).join('')}
   ${d.escenario?`<p style="font-size:17px;margin-top:18px"><b style="color:var(--cian)">Si se valida el mobiliario actual:</b> proyecto ${pc(d.escenario.proyecto_hoy)} → ${pc(d.escenario.proyecto_validado)} · general ${pc(d.escenario.general_hoy)} → ${pc(d.escenario.general_validado)}</p>`:''}</div>`);
  const grandes=d.areas.filter(a=>a.faltan&&!a.mobiliario);
  const grupos=[];let suelto=null;grandes.forEach(a=>{if(a.items.length>8){grupos.push([a]);return}if(suelto){suelto.push(a);suelto=null}else{suelto=[a];grupos.push(suelto)}});
  for(const par of grupos){
   s.push(`<div class="sl">${lg}<h1>${E(d.corto)}: pendientes de proyecto</h1><p class="sub">${par.map(a=>E(a.nombre)).join(' · ')}</p><div class="grid ${par.length===1?'':'g2'}" style="align-items:start">${par.map(a=>`<div class="card"><div class="row" style="justify-content:space-between"><b style="font-size:20px">${E(a.nombre)}</b><b class="c-${a.estado}" style="font-size:22px">${pc(a.valor)}</b></div><p class="mut" style="margin:2px 0 8px">${a.n} elementos · cumple ${a.cumple} · faltan ${a.faltan}</p><div style="${par.length===1&&a.items.length>8?'columns:2;column-gap:28px':''}">${a.items.map(i=>PEND[i]).map(p=>`<div style="font-size:15px;margin:5px 0;break-inside:avoid">${sem(p.estado,'')} <b>${E(p.titulo)}</b>${p.estado==='amarillo'?` <span class="c-amarillo">(${E(p.falta.replace('Parcial: ',''))})</span>`:''}${par.length===1?`<div class="mut2" style="font-size:13px;margin-left:18px">${E(p.texto)}</div>`:''}</div>`).join('')}</div></div>`).join('')}</div></div>`)}
  if(d.personal&&d.personal.faltan)s.push(`<div class="sl">${lg}<h1>${E(d.corto)}: personal</h1><p class="sub">${d.personal.cub} de ${d.personal.req} plazas cubiertas (${pc(d.personal.valor)}) · ${d.personal.faltan} vacantes</p><div class="grid g2"><div>${d.personal.areas.map(a=>`<div class="sbar" style="grid-template-columns:160px 1fr 90px"><span>${E(a.area)}</span>${bar(a.pct,a.pct>=80?'verde':a.pct>=50?'amarillo':'rojo')}<b>${a.cub} de ${a.req}</b></div>`).join('')}</div><div class="card">${d.personal.vacantes.map(v=>`<div class="row" style="justify-content:space-between;font-size:16px;margin:3px 0"><span>${E(v.puesto)}</span><b>${v.faltan}</b></div>`).join('')}</div></div></div>`);
 });
 s.push(`<div class="sl">${lg}<h1>Puntos a validar en la visita</h1><p class="sub">Lo que se ve en campo no coincide con el KPI o falta información</p>${DS.flatMap(d=>d.observaciones.map(o=>`<div class="card" style="margin-bottom:10px;font-size:18px"><b style="color:var(--cian)">${E(d.corto)}.</b> ${E(o.texto)}</div>`)).join('')}</div>`);
 return s}
function escala(){const st=$('#stage'),k=Math.min(innerWidth/1280,innerHeight/720);st.style.transform=`scale(${k})`}
function mover(dx){SI=Math.max(0,Math.min(SL.length-1,SI+dx));$('#stage').innerHTML=SL[SI];$('#pnum').textContent=(SI+1)+' / '+SL.length}
function presentar(){SL=slides();SI=0;$('#show').classList.add('on');escala();mover(0);try{document.documentElement.requestFullscreen&&document.documentElement.requestFullscreen()}catch(e){}}
function salir(){$('#show').classList.remove('on');try{document.fullscreenElement&&document.exitFullscreen()}catch(e){}}
addEventListener('resize',()=>{if($('#show').classList.contains('on'))escala()});
document.addEventListener('fullscreenchange',()=>{if(!document.fullscreenElement&&$('#show').classList.contains('on'))salir()});
$('#salir').addEventListener('click',salir);

// ------------------------------------------------------------------ Excel
__XLSX__
function exportar(){const filas=DS.flatMap(d=>d.pendientes).filter(p=>(!ST.pDist||p.dist===ST.pDist)&&(!ST.pInd||p.ind===ST.pInd)&&(!ST.pEst||p.estado===ST.pEst)&&(!ST.pPlan||(ST.pPlan==='Sin asignar'?!plan(p.id).estatus:plan(p.id).estatus===ST.pPlan)));
 const hojas=DS.filter(d=>filas.some(p=>p.dist===d.id)).map(d=>{const f=[[D.grupo+' — Plan de trabajo '+d.nombre],['Fecha de corte: '+D.corte_txt],
  ['#','Indicador','Área','Pendiente','Qué falta','Qué implica','Prioridad','Responsable','Fecha compromiso','Estatus','Comentarios']];
  filas.filter(p=>p.dist===d.id).forEach((p,i)=>{const v=plan(p.id);f.push([i+1,p.indicador,p.area,p.titulo,p.falta,p.texto,'ALTA',v.resp||'',v.fecha||'',v.estatus||'',v.coment||''])});
  return [d.nombre.slice(0,31),hoja(f,[5,22,22,34,34,48,10,22,16,14,40])]});
 if(!hojas.length)return alert('No hay renglones para exportar.');
 const ct=`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>${hojas.map((h,i)=>`<Override PartName="/xl/worksheets/sheet${i+1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('')}</Types>`;
 const rels=`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`;
 const wb=`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>${hojas.map((h,i)=>`<sheet name="${X(h[0])}" sheetId="${i+1}" r:id="rId${i+1}"/>`).join('')}</sheets></workbook>`;
 const wbr=`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${hojas.map((h,i)=>`<Relationship Id="rId${i+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i+1}.xml"/>`).join('')}<Relationship Id="rId${hojas.length+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`;
 const st=`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="3"><font><sz val="11"/><name val="Arial"/></font><font><b/><color rgb="FFFFFFFF"/><sz val="11"/><name val="Arial"/></font><font><b/><sz val="14"/><name val="Arial"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF1C4587"/></patternFill></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf/></cellStyleXfs><cellXfs count="3"><xf/><xf fontId="1" fillId="2" applyFont="1" applyFill="1"/><xf fontId="2" applyFont="1"/></cellXfs></styleSheet>`;
 const files=[['[Content_Types].xml',ct],['_rels/.rels',rels],['xl/workbook.xml',wb],['xl/_rels/workbook.xml.rels',wbr],['xl/styles.xml',st],...hojas.map((h,i)=>[`xl/worksheets/sheet${i+1}.xml`,h[1]])];
 const a=document.createElement('a');a.href=URL.createObjectURL(zip(files));
 a.download=`Plan_de_trabajo_${D.grupo.replace(/\s+/g,'_')}${ST.pDist?'_'+POR[ST.pDist].corto:''}_${D.corte}.xlsx`;
 document.body.appendChild(a);a.click();setTimeout(()=>{URL.revokeObjectURL(a.href);a.remove()},1000)}

addEventListener('beforeprint',()=>document.querySelectorAll('.plan .ed').forEach(el=>{let v=el.nextElementSibling;if(!v||!v.classList.contains('pv')){v=document.createElement('span');v.className='pv';el.after(v)}
 v.textContent=el.type==='date'&&el.value?el.value.split('-').reverse().join('/'):el.value}));
const h0=(location.hash||'').slice(1);if(POR[h0]||h0==='plan')ST.vista=h0;render();
"""


def xlsx_js():
    """Reutiliza el escritor de .xlsx de la herramienta actual (ZIP + SpreadsheetML, sin librerías)."""
    js = g.JS
    a = js.index("const CRC=")
    b = js.index("function exportar(")
    return js[a:b]


def main():
    ctx = A.cargar()
    data = datos(ctx)
    logo = media_template("image2.png")
    js = JS.replace("__XLSX__", xlsx_js())
    nav = "".join(f'<button data-go="{d["id"]}">{html.escape(d["corto"])}</button>' for d in data["dists"])
    doc = f"""<!DOCTYPE html>
<html lang="es"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>{html.escape(data['grupo'])} · Estatus (propuesta)</title><style>{CSS}</style></head>
<body>
<header class="top"><div class="in"><img class="logo" src="{logo}" alt="FOTON">
<div class="grp"><b>{html.escape(data['grupo'])}</b><span>Estatus y plan de trabajo · corte al {html.escape(data['corte_txt'])}</span></div>
<nav aria-label="Secciones"><button data-go="resumen">Resumen</button>{nav}<button data-go="plan">Plan de trabajo</button><button class="pres" data-go="show">Presentar</button></nav></div></header>
<main id="main"></main>
<div class="velo"></div><aside class="det" id="det" aria-label="Detalle del pendiente"></aside>
<div id="show" role="dialog" aria-label="Modo presentación"><div id="stage"></div>
<div class="ctl"><button data-sl="-1" aria-label="Anterior">←</button><span id="pnum"></span><button data-sl="1" aria-label="Siguiente">→</button><button id="salir">Salir (Esc)</button></div></div>
<script type="application/json" id="datos">{json.dumps(data, ensure_ascii=False, default=str).replace('</', '<\\/')}</script>
<script>{js}</script>
</body></html>"""
    os.makedirs(os.path.dirname(SALIDA), exist_ok=True)
    with open(SALIDA, "w", encoding="utf-8") as fh:
        fh.write(doc)
    n = sum(len(d["pendientes"]) for d in data["dists"])
    print(f"  -> {os.path.relpath(SALIDA, RAIZ)} ({os.path.getsize(SALIDA) / 1024:.0f} KB, {n} pendientes)")


if __name__ == "__main__":
    main()
