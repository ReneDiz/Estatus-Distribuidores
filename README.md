# Estatus de distribuidores FOTON

Herramientas para revisar el estatus de cada distribuidor y armar el plan de trabajo en la visita.
Hoy cubre **Grupo Velcen** (FOTON Apodaca, FOTON Escobedo y FOTON Allende). Corte vigente: **30 de
octubre de 2026**.

| Entregable | Archivo | Cómo se genera |
|---|---|---|
| Herramienta actual (HTML) | `salida/Grupo_Velcen.html` | `python3 generar_reporte.py` |
| Presentación ejecutiva (PowerPoint, template FOTON) | `presentacion/Grupo_Velcen_Estatus_nov2026.pptx` | `python3 presentacion/generar_presentacion.py` |
| Propuesta de herramienta nueva (HTML) | `propuesta/herramienta_estatus_PROPUESTA.html` | `python3 propuesta/generar_propuesta.py` |

Los tres leen los mismos datos y usan el mismo cálculo, así que siempre cuadran entre sí.

## Uso

```bash
pip install openpyxl pillow python-pptx lxml
```

```bash
python3 generar_reporte.py
python3 presentacion/generar_presentacion.py
python3 propuesta/generar_propuesta.py
```

Requiere Python 3.10 o superior. `generar_reporte.py` muestra en consola el resumen por distribuidor
(valor, semáforo y hoja/fila de donde salió cada dato). `python3 analisis.py` imprime el desglose por
área, el escenario de mobiliario y los renglones del plan.

Cuando llegue un Excel nuevo: guárdalo en `datos/`, mueve el anterior a `datos/anterior/`, cambia la
ruta y `fecha_corte` en `config.json` y vuelve a correr los tres scripts.

## Archivos de entrada (`datos/`)

| Archivo | Uso | Formato esperado |
|---|---|---|
| `KPIs_actualizado_nov2026.xlsx` | 10 indicadores (todos menos Personal) | Una hoja por indicador con encabezados: **Nombre Comercial**, **Fecha**, **Tema**, **Subtema**, **Concepto**, **Requerido**, **Adquirido**. Una fila por concepto y por mes. La pestaña *Personal* **no se usa** (obsoleta). |
| `anterior/KPIS.xlsx` | Sólo historial | Archivo de KPI's anterior (hasta ago-2026). |
| `Sistema_Personal_FOTON_v19_Regionales-Noviembre.xlsx` | Personal | Hoja `Evaluación_Red` (tabla *Detalle por Distribuidor*), hojas `Regional_*` (por área) y hoja oculta `Cálculo_Red` (por puesto). Se toma **literal**, sin recalcular. |
| `personal-activo-*.xlsx` | Sólo referencia local | Datos personales; **no se sube a git** (`.gitignore`) ni se muestra. |
| `ajustes.json` | Confirmaciones de campo | Ver «Ajustes manuales». |
| `explicaciones.json` | Textos de la presentación y la propuesta | Qué implica cada pendiente, nombres cortos de conceptos, áreas, puestos y requisitos Cummins. |
| `fotos/<distribuidor>/` + `fotos/fotos.json` | Fotos | Ver «Fotos». |

## Indicadores (11, mismo peso)

| # | Indicador | Hoja del Excel | Tipo | Semáforo |
|---|---|---|---|---|
| 1 | Desarrollo de Proyecto | `% DE DESARRO PROYEC` | Conceptos por área | `proyecto` |
| 2 | CRM | `Cum sofware oper`, concepto `CRM` | Sí / No | `si_no` |
| 3 | RIFOF | `RIFOF` | Sí / No | `si_no` |
| 4 | Cummins | `Cummins` | 7 requisitos | `cummins` |
| 5 | Equipo de Diagnóstico | `Equipo de diagnostico` | Sí / No | `si_no` |
| 6 | Fotonet | `Cum sofware oper`, concepto `FOTON NET` | Sí / No | `si_no` |
| 7 | Plan Piso | `Capacidad financiera` | Sí / No | `si_no` |
| 8 | Kit de Refacciones | `Stock de refacciones` | Sí / No | `si_no` |
| 9 | PLD | `PLD` | Sí / No | `si_no` |
| 10 | Anexos Operativos | `Anexos operativos 2026` | Sí / No | `si_no` |
| 11 | Personal | Sistema de Personal | Plazas cubiertas / requeridas | `personal` |

Cambios contra la versión anterior: *Software Operativo* se separó en **CRM** y **Fotonet**; *Stock de
refacciones* ahora se llama **Kit de Refacciones**; se agregaron **RIFOF**, **Plan Piso** y **PLD**.

**Allende** es showroom y se rige por Apodaca y Escobedo: CRM, Fotonet, Kit de Refacciones, Equipo de
Diagnóstico y Cummins **no aplican** y no entran a su promedio (`no_aplica` en `config.json`). Su
`mensaje_clave` («Es necesario definir el proyecto para la agencia.») se muestra en la presentación,
en la propuesta y como primer renglón de su plan.

## Cómo se calcula

- **Mes del dato:** el último mes disponible ≤ fecha de corte (oct-2026 en el archivo actual).
- **Valor por concepto** = Adquirido / Requerido (tope 100%). Un parcial 0.7 vale 70%; Cummins
  «certificación de técnicos» con 1 de 2 vale 50%.
- **% del indicador** = promedio de sus conceptos. Binarios: 100% o 0%.
- **% General de la agencia** = promedio simple de sus indicadores aplicables (11 en Apodaca y
  Escobedo, 6 en Allende).
- **Total del grupo** = promedio de **todos** los indicadores aplicables de las tres agencias
  (28, Allende incluido). Proyecto del grupo = suma del avance de los 234 elementos / 234.
- **% por área de proyecto** (Imagen corporativa, Showroom, Almacén, Taller, Áreas administrativas,
  Mobiliario, Generales): cada área es un bloque aparte con elementos, cumple, faltan y %.
- **Escenario de mobiliario:** cuánto subirían proyecto y general si el mobiliario actual se valida
  como OK (cada elemento de mobiliario al 100%, todo lo demás igual).
- **Pendientes:** todo lo que no está al 100% es prioridad **ALTA**, aunque el semáforo esté verde.

Resultado al 30-oct-2026:

| | General | Proyecto | Pendientes | Mobiliario validado (proyecto / general) |
|---|---|---|---|---|
| Apodaca | 73.5% amarillo | 53.7% | 70 | 72.9% / 75.3% |
| Escobedo | 92.2% verde | 66.7% | 47 | 85.9% / 94.0% |
| Allende | 78.9% amarillo | 30.6% | 29 | 47.2% / 81.7% (sólo si cuenta la oficina móvil) |
| **Grupo** | **82.0% verde** | **55.6%** | **146** | 74.4% / 84.0% |

## Semáforo (`config.json` → `semaforo`)

| Regla | Verde | Amarillo | Rojo |
|---|---|---|---|
| `si_no` (binarios) | 100% | — | < 100% |
| `proyecto` | ≥ 80% | 50–79% | < 50% |
| `cummins` | 100% | 50–99% | < 50% |
| `personal` | ≥ 80% | 50–79% | < 50% |
| `general` | ≥ 80% | 50–79% | < 50% |

Cada regla tiene `verde_min` y `amarillo_min` (`null` quita el amarillo).

## Ajustes manuales (`datos/ajustes.json`)

Confirmaciones de campo que todavía no están en el Excel. Cada ajuste reemplaza el *Adquirido* de un
concepto y, si se indica, el texto del pendiente (`falta`); el % y el semáforo se recalculan y aparece
la etiqueta «Ajuste: …». Cuando el Excel ya traiga el dato, borra el ajuste.

Vigente (nov-2026): sólo **Escobedo, Cummins**: el técnico ya está certificado; falta certificar al
jefe de taller (1 de 2). Los ajustes anteriores (anexos firmados, letrero Total Care y kit de Apodaca)
se quitaron porque el Excel nuevo ya los trae.

## Explicaciones (`datos/explicaciones.json`)

- `conceptos`: por cada concepto del Excel, `titulo` corto y `texto` (qué falta y qué se necesita).
- `indicadores`: `falta` y `necesita` de cada indicador binario.
- `cummins` / `cummins_corto`: acción y nombre corto de cada requisito Cummins.
- `puestos` y `areas`: nombres cortos para mostrar.

Si aparece un concepto nuevo sin explicación, los generadores avisan en consola
(«pendientes sin explicación…») y muestran el nombre del Excel tal cual.

## Fotos

- Sólo fotos del Drive de cada distribuidor o las que entrega el usuario. Se descargan a
  `datos/fotos/<id>/`; el Drive se usa **sólo en lectura**.
- `datos/fotos/fotos.json` asigna cada foto a un concepto. La foto debe ser congruente con el estatus:
  si el concepto cumple, que se vea cumplido; si está pendiente, que muestre lo que falta. Si una foto
  contradice el Excel va en `observaciones` («Para validar en la visita»); `texto_presentacion` es la
  versión corta, sin fechas ni rutas, que usan la presentación y la propuesta.
- Más reciente primero. Más de 6 meses antes del corte: etiqueta «Foto no reciente».
- En la presentación y la propuesta las fotos no muestran fechas ni rutas de Drive.

### Fotos de fachada editadas (`fotos/`)

- `fotos/referencia/`: originales (no se modifican).
- `fotos/editadas/`: versiones para portada y secciones, con su antes / después:
  - **Escobedo:** se quitó la camioneta roja frente a la fachada (relleno con el asfalto vecino).
  - **Apodaca:** cielo azul, piso más cálido y luminoso, charcos con reflejo azul claro.
  - Scripts en `fotos/editadas/scripts/` para repetir la edición.
- Allende usa la foto de la oficina móvil (no reciente) sin editar.

## Presentación (`presentacion/`)

`generar_presentacion.py` abre `template/Foton_Mexico_-_Template_Base-2026.pptx` y usa sus fondos,
logo, tipografías (Arial Black / Arial) y colores. Las diapositivas de muestra del template no se
incluyen. También se guarda en PDF (`Grupo_Velcen_Estatus_nov2026.pdf`) para abrirla directo.

Versión ejecutiva (28 diapositivas):

| # | Contenido |
|---|---|
| 1–4 | Portada, resumen ejecutivo (total del grupo y dos semáforos por agencia), matriz de 11 indicadores y escenario de mobiliario. |
| 5–11 | Apodaca: presentación de la agencia (calificación, foto de fachada y los 11 indicadores en tarjetas), proyecto por área (gráfica y tabla), pendientes por área con explicación (4) y personal. |
| 12–17 | Escobedo: lo mismo (pendientes en 3). |
| 18–22 | Allende: lo mismo (pendientes en 2), con «Es necesario definir el proyecto para la agencia». |
| 23 | Evidencia (fotos) y puntos a validar en la visita. |
| 24–27 | Plan de trabajo por agencia: pendiente, qué falta, prioridad, y Responsable, Fecha compromiso, Estatus y Comentarios en blanco para llenar en la visita. |
| 28 | Siguientes pasos. |

- **Pendientes:** cada área es un bloque a todo el ancho con su avance, conteo y barra; cada pendiente
  lleva su nombre y una explicación corta. Las alturas se calculan con la medida real de la letra para
  que nada se encime; un área chica sube a llenar el espacio libre de otra lámina.
- **Personal:** el área con menor cobertura como título, barras por área, vacantes por puesto, total
  de vacantes de esa área y personas extra (por arriba de lo requerido) que podrían cubrir vacantes.

Lo que cumple se muestra en una línea o en su tarjeta; el detalle es para lo pendiente. Cálculos y
fuentes van en las notas del orador. Gráficas y tablas son editables en PowerPoint.

Versión extendida (39 diapositivas, agrega por agencia qué falta por indicador, mobiliario, Cummins y evidencia):

```bash
python3 presentacion/generar_presentacion.py --completa
```

Para el PDF:

```bash
cd presentacion && soffice --headless --convert-to pdf Grupo_Velcen_Estatus_nov2026.pptx
```

## Herramienta actual (`salida/Grupo_Velcen.html`)

Se mantiene el diseño original; sólo se actualizaron los datos (11 indicadores, corte 30-oct-2026 y
tarjeta «Total del grupo»). Incluye resumen, desglose por agencia, fotos, observaciones y plan de
trabajo editable (localStorage), exportable a Excel e imprimible / PDF.

## Propuesta de herramienta nueva (`propuesta/`)

`herramienta_estatus_PROPUESTA.html`, un solo archivo que funciona sin internet, con el estilo del
template FOTON:

- **Navegación por sucursal:** Resumen, Apodaca, Escobedo, Allende y Plan de trabajo.
- **Filtros** por indicador, estado (no cumple / parcial) y búsqueda. Clic en una celda de la matriz
  abre la sucursal filtrada por ese indicador.
- **Detalle del pendiente** (clic): qué falta, qué implica, avance, foto si hay, y los campos del
  plan; anterior / siguiente.
- **Plan de trabajo:** 147 renglones (146 pendientes + definir el proyecto de Allende) con filtros por
  sucursal, indicador, estado y estatus; se guarda en el navegador; **Exportar a Excel** (una hoja por
  sucursal) e **Imprimir / PDF**. Cada pendiente tiene un id estable, así que lo capturado se
  conserva al regenerar con datos nuevos.
- **Presentar:** pantalla completa con flechas (← →) y Esc para salir.
- Funciona en celular (sin desbordes horizontales).

No reemplaza a la herramienta actual; es una propuesta para validar.

## Agregar otro distribuidor o grupo

En `config.json` → `grupos`:

```json
{ "id": "nuevo", "nombre": "FOTON Ciudad", "nombre_kpi": "FOTON_Ciudad",
  "nombre_personal": "Foton Ciudad", "no_aplica": [] }
```

- `nombre_kpi`: texto exacto de *Nombre Comercial* en el Excel de KPI's.
- `nombre_personal`: texto exacto de *Distribuidor* en el sistema de personal.
- `no_aplica`: ids de indicadores que no aplican (`proyecto`, `crm`, `rifof`, `cummins`,
  `diagnostico`, `fotonet`, `planpiso`, `kit`, `pld`, `anexos`, `personal`) y `motivo_no_aplica`.
- `mensaje_clave` (opcional): mensaje principal de la agencia.
- Para la presentación y la propuesta, agrega su foto de fachada en `FACHADAS` y sus fotos de evidencia
  en `EVIDENCIAS` (dentro de cada generador).

## Respaldo

`respaldo_2026-11-11/` guarda la herramienta y la presentación como estaban antes de la actualización
de noviembre (ver `LEEME.txt`).
