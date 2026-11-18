#!/usr/bin/env python3
"""
PROPUESTA con fotos del checklist: propuesta/herramienta_estatus_PROPUESTA_fotos.html

Uso:
    python3 propuesta/generar_propuesta_fotos.py

Toma la propuesta actual (generar_propuesta.py, sin modificarla) y le agrega, por agencia,
TODOS los elementos del Desarrollo de Proyecto (los que ya cumplen y los pendientes) con la
foto del «Check list de seguimiento de proyecto» (datos/checklists/checklists.json):
  - Vista «Todos los elementos con foto»: tarjetas por área con foto, estatus y filtros
    (Cumple / Pendiente / Con foto / Sin foto).
  - Clic en un elemento: abre su ficha con la foto en grande, el % y el comentario del
    checklist, el estatus del KPI y, si está pendiente, los campos del plan de trabajo.
  - Clic en la foto: pantalla completa (flechas para cambiar de foto, Esc para cerrar).
  - En la vista de pendientes, cada pendiente con foto muestra su miniatura.
"""
import base64
import html
import io
import json
import os
import sys

from PIL import Image, ImageOps

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, RAIZ)
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import analisis as A  # noqa: E402
import checklist as C  # noqa: E402
import generar_propuesta as GP  # noqa: E402
import generar_reporte as g  # noqa: E402

SALIDA = os.path.join(RAIZ, "propuesta", "herramienta_estatus_PROPUESTA_fotos.html")


def jpg64(data, ancho, calidad):
    im = ImageOps.exif_transpose(Image.open(io.BytesIO(data))).convert("RGB")
    if im.width > ancho:
        im = im.resize((ancho, round(im.height * ancho / im.width)), Image.LANCZOS)
    buf = io.BytesIO()
    im.save(buf, "JPEG", quality=calidad, optimize=True)
    return "data:image/jpeg;base64," + base64.b64encode(buf.getvalue()).decode()


def agregar_checklist(ctx, data):
    cfg = C.config()
    exp = A.Explicaciones(os.path.join(RAIZ, "datos", "explicaciones.json"))
    elementos, n_foto = {}, 0
    reporte = []
    for dctx, d in zip(ctx["dists"], data["dists"]):
        d["chk"] = None
        info = cfg.get("agencias", {}).get(d["id"])
        pr = next((i for i in dctx["indicadores"] if i["id"] == "proyecto"), None)
        if not pr or not pr["conceptos"]:
            continue
        filas = C.leer(os.path.join(RAIZ, "datos", "checklists", info["archivo"])) if info else []
        empate = C.empatar([c["concepto"] for c in pr["conceptos"]], filas, cfg.get("equivalencias"))
        ids_pend = {p["id"] for p in d["pendientes"]}
        fechas = {}
        for actx, a in zip(dctx["areas"], d["areas"]):
            a["els"] = []
            for c in pr["conceptos"]:
                if c["subtema"] != actx["original"]:
                    continue
                e_exp = exp.c.get(g.norm(c["concepto"]))
                titulo = e_exp["titulo"] if e_exp else g.limpio(c["concepto"])
                clave = GP.slug(a["nombre"] + "-" + c["concepto"])
                pend = f"{d['id']}-proyecto-{clave}"
                f = empate.get(c["concepto"])
                fotos = []
                for data_foto in (f["fotos"] if f else []):
                    n_foto += 1
                    k = f"c{n_foto}"
                    data["fotos"][k] = {"src": jpg64(data_foto, 1000, 74), "thumb": jpg64(data_foto, 360, 62),
                                        "pie": titulo, "no_reciente": False}
                    fotos.append(k)
                if f and f.get("fecha"):
                    fechas[f["fecha"]] = fechas.get(f["fecha"], 0) + 1
                av = c["avance"]
                e = {"id": f"{d['id']}-el-{clave}", "dist": d["id"], "area": a["nombre"], "titulo": titulo,
                     "concepto": g.limpio(c["concepto"]), "avance": round(av * 100),
                     "estado": "verde" if av >= 1 else ("amarillo" if av > 0 else "rojo"),
                     "texto": (e_exp or {}).get("texto") if av < 1 else None,
                     "pend": pend if pend in ids_pend else None, "fotos": fotos,
                     "chk": ({"pct": None if f["pct"] is None else round(f["pct"] * 100), "fecha": f["fecha"],
                              "comentario": f["comentario"], "elemento": f["elemento"]} if f else None)}
                elementos[e["id"]] = e
                a["els"].append(e["id"])
        if info:
            con_foto = sum(1 for i in (x for a in d["areas"] for x in a["els"]) if elementos[i]["fotos"])
            d["chk"] = {"fuente": info.get("fuente", info["archivo"]), "n_fotos": sum(len(elementos[i]["fotos"]) for a in d["areas"] for i in a["els"]),
                        "con_foto": con_foto, "revision": max(fechas, key=fechas.get) if fechas else None,
                        "empatados": len(empate), "elementos": len(pr["conceptos"])}
            reporte.append(f"{d['nombre']}: {len(empate)} de {len(pr['conceptos'])} elementos ligados al checklist, "
                           f"{d['chk']['n_fotos']} fotos en {con_foto} elementos")
        else:
            reporte.append(f"{d['nombre']}: sin checklist con fotos cargado")
    data["elementos"] = elementos
    return reporte


CSS_EXTRA = r"""
.item-foto{width:100%;height:120px;object-fit:cover;border-radius:6px;margin-bottom:4px;display:block}
.egrid{display:grid;grid-template-columns:repeat(auto-fill,minmax(180px,1fr));gap:10px;margin-top:10px}
.ecard{background:var(--card2);border:1px solid transparent;border-radius:8px;padding:0;overflow:hidden;cursor:pointer;text-align:left;display:flex;flex-direction:column}
.ecard:hover,.ecard:focus-visible{border-color:var(--cian);outline:none}
.ecard img{width:100%;height:120px;object-fit:cover;display:block}.ecard .sinfoto{width:100%;height:120px}
.sinfoto{display:flex;align-items:center;justify-content:center;background:repeating-linear-gradient(45deg,#16233F,#16233F 8px,#1A2A4A 8px,#1A2A4A 16px);color:var(--mut);font-size:12px;font-weight:700}
.sinfoto.grande{height:120px;border-radius:8px;margin-top:10px}
.ebody{padding:8px 10px 10px;display:flex;flex-direction:column;gap:4px;font-size:13px}
.dfotos{display:grid;gap:8px;margin-top:10px}.dfoto{padding:0;border:0;background:none;cursor:zoom-in}.dfoto img{width:100%;border-radius:8px;display:block;margin:0}
.chkinfo{font-size:12px;color:var(--mut)}
#lb{position:fixed;inset:0;z-index:70;display:none;align-items:center;justify-content:center}
#lb.on{display:flex}
.lb-fondo{position:absolute;inset:0;background:rgba(0,0,0,.9)}
#lb img{position:relative;max-width:92vw;max-height:80vh;border-radius:6px}
#lb p{position:absolute;bottom:18px;left:16px;right:16px;text-align:center;color:#fff;font-weight:700;margin:0}
#lb button{position:absolute;background:rgba(255,255,255,.16);border:0;color:#fff;border-radius:8px;padding:10px 14px;cursor:pointer;font-weight:700}
#lb-prev{left:16px;top:50%}#lb-next{right:16px;top:50%}#lb-cerrar{top:16px;right:16px}
@media print{.egrid{grid-template-columns:repeat(4,1fr)}.ecard{background:#fff;border:1px solid #ccd3dc;color:#14202E}}
"""

JS_EXTRA = r"""
// ------------------------------------------------------------------ elementos del checklist con foto
const EL=D.elementos||{};const EL_DE_PEND={};Object.values(EL).forEach(e=>{if(e.pend)EL_DE_PEND[e.pend]=e});
const MES=['ene','feb','mar','abr','may','jun','jul','ago','sep','oct','nov','dic'];
const fechaTxt=s=>{if(!s)return '';const [a,m,d]=s.split('-');return +d+'-'+MES[m-1]+'-'+a};
const estTxt=e=>e.estado==='verde'?'Cumple':e.estado==='amarillo'?'Parcial '+e.avance+'%':'No cumple';
function vistaEls(d){const c=d.chk;
 return `<div class="filtros" style="position:static;margin-top:4px"><b>Ver</b>${[['pend','Sólo pendientes'],['todos','Todos los elementos con foto']].map(([v,t])=>`<button class="chip" data-vel="${v}" aria-pressed="${ST.vEl===v}">${t}</button>`).join('')}
 ${ST.vEl==='todos'?'<span style="width:12px"></span>'+[['','Todos'],['ok','Cumple'],['pend','Pendiente'],['foto','Con foto'],['sinfoto','Sin foto']].map(([v,t])=>`<button class="chip" data-fel="${v}" aria-pressed="${ST.fEl===v}">${t}</button>`).join(''):''}
 <span class="chkinfo">${c?`Checklist: ${c.n_fotos} fotos en ${c.con_foto} de ${c.elementos} elementos${c.revision?' · revisión '+fechaTxt(c.revision):''}`:'Aún no hay checklist con fotos para esta sucursal'}</span></div>`}
function filtraEl(e){return (!ST.fEl||(ST.fEl==='ok'&&e.estado==='verde')||(ST.fEl==='pend'&&e.estado!=='verde')||(ST.fEl==='foto'&&e.fotos.length)||(ST.fEl==='sinfoto'&&!e.fotos.length))
 &&(!ST.q||(e.titulo+' '+e.area+' '+(e.chk?e.chk.comentario:'')).toLowerCase().includes(ST.q.toLowerCase()))}
function tarjetaEl(e){const f=e.fotos[0];
 return `<button class="ecard" data-el="${e.id}">${f?`<img src="${F[f].thumb||F[f].src}" alt="${E(e.titulo)}">`:`<span class="sinfoto">Sin foto</span>`}
 <span class="ebody">${sem(e.estado,estTxt(e))}<b>${E(e.titulo)}</b>${e.fotos.length>1?`<span class="mini">${e.fotos.length} fotos</span>`:''}</span></button>`}
function areaTodos(d,a){const els=(a.els||[]).map(i=>EL[i]).filter(filtraEl);if(!els.length)return ['',0];
 return [`<div class="area"><div class="hd">${sem(a.estado,'')}<h3>${E(a.nombre)}</h3><span class="mut">${a.n} elementos · cumple ${a.cumple} · faltan ${a.faltan}</span><span class="pc c-${a.estado}">${pc(a.valor)}</span></div>${bar(a.valor,a.estado)}<div class="egrid">${els.map(tarjetaEl).join('')}</div></div>`,els.filter(e=>e.estado!=='verde').length]}
function abrirEl(id,via){const e=EL[id];if(!e)return;const d=POR[e.dist],p=e.pend?PEND[e.pend]:null,v=p?plan(p.id):{};
 const attr=via==='pend'?'data-pend':'data-el',key=via==='pend'?e.pend:id;
 LISTA=[...document.querySelectorAll('#main ['+attr+']')].map(b=>b.getAttribute(attr)).filter((x,i,a)=>x&&a.indexOf(x)===i);const i=LISTA.indexOf(key);
 $('#det').innerHTML=`<button class="cerrar" id="cerrar">Cerrar</button><div class="tag">${E(d.nombre)} · ${E(e.area)}</div><h3>${E(e.titulo)}</h3>${sem(e.estado,estTxt(e))}
  ${e.fotos.length?`<div class="dfotos">${e.fotos.map((f,j)=>`<button class="dfoto" data-foto="${id}|${j}" aria-label="Ver foto ${j+1} en grande"><img src="${F[f].src}" alt="${E(e.titulo)}"></button>`).join('')}</div><p class="mini">Toca la foto para verla en pantalla completa.</p>`:`<div class="sinfoto grande">El checklist no trae foto de este elemento</div>`}
  <dl><dt>Estatus KPI</dt><dd>${estTxt(e)}</dd>${e.chk?`<dt>Checklist</dt><dd>${e.chk.pct==null?'N/A':e.chk.pct+'%'}${e.chk.fecha?' · revisión '+fechaTxt(e.chk.fecha):''}</dd>${e.chk.comentario?`<dt>Comentario</dt><dd>${E(e.chk.comentario)}</dd>`:''}`:'<dt>Checklist</dt><dd>Sin renglón en el checklist</dd>'}${e.texto?`<dt>Qué implica</dt><dd>${E(e.texto)}</dd>`:''}${p?`<dt>Prioridad</dt><dd class="prio">ALTA</dd>`:''}</dl>
  ${p?`<div data-id="${p.id}"><label>Responsable<input class="ed" data-f="resp" value="${E(v.resp||'')}"></label><label>Fecha compromiso<input class="ed" type="date" data-f="fecha" value="${E(v.fecha||'')}"></label>
  <label>Estatus<select class="ed" data-f="estatus"><option value=""></option>${['Pendiente','En proceso','Cerrado'].map(s=>`<option ${v.estatus===s?'selected':''}>${s}</option>`).join('')}</select></label>
  <label>Comentarios<textarea class="ed" data-f="coment" rows="3">${E(v.coment||'')}</textarea></label></div>`:''}
  <div class="nav"><button class="btn sec" ${attr}="${LISTA[i-1]||''}" ${i>0?'':'disabled'}>← Anterior</button><span class="mini">${i+1} de ${LISTA.length}</span><button class="btn sec" ${attr}="${LISTA[i+1]||''}" ${i>=0&&i<LISTA.length-1?'':'disabled'}>Siguiente →</button></div>`;
 document.body.classList.add('abierto');$('#cerrar').focus()}
let LB=[],LBI=0,LBT='';
function lightbox(k){const [id,j]=k.split('|');LB=EL[id].fotos;LBI=+j;LBT=EL[id].titulo;pintaLb();$('#lb').classList.add('on');$('#lb-cerrar').focus()}
function pintaLb(){const f=F[LB[LBI]];$('#lb-img').src=f.src;$('#lb-img').alt=LBT;$('#lb-pie').textContent=LBT+(LB.length>1?` · foto ${LBI+1} de ${LB.length}`:'');$('#lb-prev').hidden=$('#lb-next').hidden=LB.length<2}
function moverLb(dx){LBI=(LBI+dx+LB.length)%LB.length;pintaLb()}
function cerrarLb(){$('#lb').classList.remove('on')}
"""

PARCHES = [
    # estado de la vista de elementos
    ("const ST={vista:'resumen',", "const ST={vEl:'pend',fEl:'',vista:'resumen',"),
    # miniatura en cada pendiente con foto del checklist
    ("function itemHTML(p){const pt=planTxt(p.id);\n return `<button class=\"item\" data-pend=\"${p.id}\"><span class=\"t\">",
     "function itemHTML(p){const pt=planTxt(p.id);const el=EL_DE_PEND[p.id],tf=el&&el.fotos.length?el.fotos[0]:null;\n"
     " return `<button class=\"item\" data-pend=\"${p.id}\">${tf?`<img class=\"item-foto\" src=\"${F[tf].thumb||F[tf].src}\" alt=\"\">`:''}<span class=\"t\">"),
    # selector de vista bajo el título del proyecto
    ("const pp=d.proyecto;h+=`<h2>Desarrollo de proyecto por área</h2><p class=\"sub\">${pc(pp.valor)} · ${pp.cumple} de ${pp.n} elementos completos · faltan ${pp.faltan}</p>`;",
     "const pp=d.proyecto;h+=`<h2>Desarrollo de proyecto por área</h2><p class=\"sub\">${pc(pp.valor)} · ${pp.cumple} de ${pp.n} elementos completos · faltan ${pp.faltan}</p>`;h+=vistaEls(d);"),
    # en «todos», cada área muestra todos sus elementos con foto
    ("d.areas.forEach(a=>{const its=a.items.map(i=>PEND[i]).filter(filtra);n+=its.length;",
     "d.areas.forEach(a=>{if(ST.vEl==='todos'){const r=areaTodos(d,a);h+=r[0];n+=r[1];return}const its=a.items.map(i=>PEND[i]).filter(filtra);n+=its.length;"),
    # un pendiente con elemento del checklist abre la ficha con foto
    ("function abrir(id){const p=PEND[id];", "function abrir(id){if(EL_DE_PEND[id])return abrirEl(EL_DE_PEND[id].id,'pend');const p=PEND[id];"),
    # clics nuevos
    ("document.addEventListener('click',e=>{const t=e.target.closest('[data-go],[data-pend],[data-ind],[data-est],#xls,#imp,#cerrar,.velo,[data-sl]');if(!t)return;\n if(t.classList.contains('velo')||t.id==='cerrar')return cerrar();",
     "document.addEventListener('click',e=>{const t=e.target.closest('[data-go],[data-pend],[data-ind],[data-est],#xls,#imp,#cerrar,.velo,[data-sl],[data-el],[data-vel],[data-fel],[data-foto],#lb-cerrar,#lb-prev,#lb-next,.lb-fondo');if(!t)return;\n"
     " if(t.id==='lb-cerrar'||t.classList.contains('lb-fondo'))return cerrarLb();if(t.id==='lb-prev')return moverLb(-1);if(t.id==='lb-next')return moverLb(1);\n"
     " if(t.dataset.foto!==undefined)return lightbox(t.dataset.foto);\n"
     " if(t.dataset.el!==undefined){if(t.dataset.el)abrirEl(t.dataset.el,'el');return}\n"
     " if(t.dataset.vel!==undefined){ST.vEl=t.dataset.vel;render(true);return}\n"
     " if(t.dataset.fel!==undefined){ST.fEl=t.dataset.fel;render(true);return}\n"
     " if(t.classList.contains('velo')||t.id==='cerrar')return cerrar();"),
    # teclado: la foto en pantalla completa va primero
    ("document.addEventListener('keydown',e=>{",
     "document.addEventListener('keydown',e=>{if($('#lb').classList.contains('on')){if(e.key==='Escape')cerrarLb();else if(e.key==='ArrowRight')moverLb(1);else if(e.key==='ArrowLeft')moverLb(-1);return}"),
]


def main():
    ctx = A.cargar()
    data = GP.datos(ctx)
    reporte = agregar_checklist(ctx, data)
    js = GP.JS.replace("__XLSX__", GP.xlsx_js())
    # EL y funciones nuevas antes de las vistas
    ancla = "// ------------------------------------------------------------------ vistas"
    assert js.count(ancla) == 1
    js = js.replace(ancla, JS_EXTRA + "\n" + ancla)
    for viejo, nuevo in PARCHES:
        assert js.count(viejo) == 1, viejo[:80]
        js = js.replace(viejo, nuevo)
    logo = GP.media_template("image2.png")
    nav = "".join(f'<button data-go="{d["id"]}">{html.escape(d["corto"])}</button>' for d in data["dists"])
    doc = f"""<!DOCTYPE html>
<html lang="es"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>{html.escape(data['grupo'])} · Estatus con fotos (propuesta)</title><style>{GP.CSS}{CSS_EXTRA}</style></head>
<body>
<header class="top"><div class="in"><img class="logo" src="{logo}" alt="FOTON">
<div class="grp"><b>{html.escape(data['grupo'])}</b><span>Estatus y plan de trabajo · corte al {html.escape(data['corte_txt'])}</span></div>
<nav aria-label="Secciones"><button data-go="resumen">Resumen</button>{nav}<button data-go="plan">Plan de trabajo</button><button class="pres" data-go="show">Presentar</button></nav></div></header>
<main id="main"></main>
<div class="velo"></div><aside class="det" id="det" aria-label="Detalle del elemento"></aside>
<div id="lb" role="dialog" aria-label="Foto en pantalla completa"><div class="lb-fondo"></div><img id="lb-img" alt=""><p id="lb-pie"></p><button id="lb-prev" aria-label="Foto anterior">←</button><button id="lb-next" aria-label="Foto siguiente">→</button><button id="lb-cerrar">Cerrar (Esc)</button></div>
<div id="show" role="dialog" aria-label="Modo presentación"><div id="stage"></div>
<div class="ctl"><button data-sl="-1" aria-label="Anterior">←</button><span id="pnum"></span><button data-sl="1" aria-label="Siguiente">→</button><button id="salir">Salir (Esc)</button></div></div>
<script type="application/json" id="datos">{json.dumps(data, ensure_ascii=False, default=str).replace('</', '<\\/')}</script>
<script>{js}</script>
</body></html>"""
    with open(SALIDA, "w", encoding="utf-8") as fh:
        fh.write(doc)
    print("\n".join("  " + r for r in reporte))
    print(f"  -> {os.path.relpath(SALIDA, RAIZ)} ({os.path.getsize(SALIDA) / 1e6:.1f} MB)")


if __name__ == "__main__":
    main()
