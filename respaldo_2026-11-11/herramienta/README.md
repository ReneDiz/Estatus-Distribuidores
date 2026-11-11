# Estatus de distribuidores FOTON

Genera un **HTML autocontenido** por grupo de distribuidores (funciona sin internet en laptop o
tablet) con: resumen del grupo, semáforo por indicador, desglose de conceptos, personal, fotos
de evidencia y **plan de trabajo editable** (exportable a Excel e imprimible / PDF).

Resultado actual: `salida/Grupo_Velcen.html` (FOTON Apodaca, FOTON Escobedo y FOTON Allende).

## Uso

```bash
pip install openpyxl pillow
python3 generar_reporte.py                  # genera todos los grupos de config.json
python3 generar_reporte.py --grupo velcen   # sólo un grupo
```

Requiere Python 3.10 o superior. La consola muestra el resumen por distribuidor (valor, semáforo y
hoja/fila de donde salió cada dato) y escribe `salida/<archivo_salida>.html`.

Cuando actualices los Excel, reemplázalos en `datos/` (mismo nombre o ajusta la ruta en
`config.json`), cambia `fecha_corte` si aplica y vuelve a correr el script.

## Archivos de entrada (`datos/`)

| Archivo | Uso | Formato esperado |
|---|---|---|
| `KPIS.xlsx` | Indicadores 1–6 | Una hoja por indicador con encabezados en la fila 1: **Nombre Comercial**, **Fecha**, **Tema**, **Subtema**, **Concepto**, **Requerido**, **Adquirido**. Una fila por concepto y por mes. La pestaña *Personal* de este archivo **no se usa** (obsoleta). |
| `Sistema_Personal_FOTON_…xlsx` | Indicador 7 (Personal) | Hoja `Evaluación_Red` (tabla *Detalle por Distribuidor*: plazas requeridas, cubiertas, faltan, extra, % cumplimiento), hojas `Regional_*` (detalle por área) y hoja oculta `Cálculo_Red` (detalle por puesto). Se toman **literal**, sin recalcular. |
| `Personal_FOTON_…pdf`, `Estructura_de_Personal_…pdf` | Sólo referencia | Impresiones del mismo sistema de personal; sirven para validar. |
| `personal-activo-*.xlsx` | Sólo referencia (local) | Lista de empleados con datos personales; **no se sube a git** (`.gitignore`) ni se usa en el HTML. |
| `fotos/<distribuidor>/` + `fotos/fotos.json` | Fotos | Ver «Fotos» abajo. |

## Cómo se calcula

- **Mes del dato:** para cada hoja se toma el **último mes disponible ≤ fecha de corte** para ese
  distribuidor. Si un indicador sólo tiene datos hasta un mes anterior (p. ej. Anexos y Cummins de
  Velcen están a ago-2026), la tarjeta lo indica con «dato a ago-2026».
- **Valor por concepto** = Adquirido / Requerido (tope 100%). Así un parcial 0.7 vale 70% y
  «Certificación técnicos» con 1 de 2 vale 50%.
- **% del indicador** = promedio de los conceptos (todos pesan igual). Software con 1 de 2 = 50%.
- **% General** = promedio simple de los 7 indicadores. Los que están en `no_aplica` (Allende:
  Software, Stock, Diagnóstico y Cummins, por ser showroom que se rige por Apodaca/Escobedo) se
  muestran como «No aplica» y **no** entran al promedio. Si un dato falta en el Excel se muestra
  «Sin información» y tampoco entra al promedio (se avisa en la nota del distribuidor).
- **Pendientes:** todo lo que no está al 100% va al plan de trabajo con prioridad ALTA (aunque el
  semáforo esté en verde). Cada concepto incompleto y cada puesto con vacantes es un pendiente.

## Ajustes manuales (`datos/ajustes.json`)

Para confirmaciones de campo que todavía no están en el Excel de KPI's (p. ej. un anexo ya firmado
o una certificación nueva). Cada ajuste reemplaza el *Adquirido* de un concepto y, opcionalmente, el
texto del pendiente (`falta`); el % y el semáforo se recalculan. En el HTML aparece la etiqueta
«Ajuste: …» junto al dato. Cuando el Excel ya traiga el dato, borra el ajuste.

```json
{ "distribuidor": "escobedo", "indicador": "cummins",
  "concepto": "Ceritificación Tecnicos o Gerente de Posventa o Gerente de Taller",
  "adquirido": 1, "fecha": "2026-11-10",
  "nota": "El técnico ya está certificado; falta certificar al jefe de taller",
  "falta": "Certificar al jefe de taller (1 de 2 acreditados)" }
```

Ajustes vigentes (10-nov-2026): anexos 2026 firmados en Apodaca, Escobedo y Allende; Escobedo Cummins
sigue en 1 de 2, ahora el pendiente es la certificación del jefe de taller; Apodaca ya tiene el letrero
Total Care y el kit inicial de partes al 100% (proyecto).

## Semáforo (`config.json` → `semaforo`)

| Regla | Verde | Amarillo | Rojo |
|---|---|---|---|
| `si_no` (Anexos, Software, Stock, Diagnóstico) | 100% | — (sin amarillo) | < 100% |
| `proyecto` | ≥ 80% | 50–79% | < 50% |
| `cummins` | 100% | 50–99% | < 50% |
| `personal` | ≥ 80% | 50–79% | < 50% |
| `general` | ≥ 80% | 50–79% | < 50% |

Cada regla tiene `verde_min` y `amarillo_min` (usar `null` para quitar el amarillo).

## Agregar otro distribuidor o grupo

En `config.json` → `grupos`, agrega un grupo (o un distribuidor dentro de uno existente):

```json
{ "id": "nuevo", "nombre": "FOTON Ciudad", "nombre_kpi": "FOTON_Ciudad",
  "nombre_personal": "Foton Ciudad", "no_aplica": [] }
```

- `nombre_kpi`: texto exacto de la columna *Nombre Comercial* del Excel de KPI's.
- `nombre_personal`: texto exacto de la columna *Distribuidor* del sistema de personal.
- `no_aplica`: ids de indicadores que no aplican (`anexos`, `proyecto`, `software`, `stock`,
  `diagnostico`, `cummins`, `personal`) y `motivo_no_aplica`.
- Cada grupo genera su propio archivo (`archivo_salida`).

## Fotos

Las fotos salen **sólo** del Drive de cada distribuidor (no de internet ni generadas). Flujo:

1. Descarga las fotos de la carpeta de visitas del distribuidor a `datos/fotos/<id_distribuidor>/`
   (el archivo `_drive_manifest.json` de cada carpeta guarda el ID de Drive de cada foto).
2. Asígnalas en `datos/fotos/fotos.json`:
   ```json
   { "distribuidor": "apodaca", "indicador": "proyecto", "concepto": "Totém",
     "archivo": "apodaca/WhatsApp Image 2026-07-02 at 6.14.53 PM.jpeg",
     "drive_carpeta": "1_Activos / Grupo Velcen (Foton_Apodaca) / Fotos Visitas / 20260715 Apodaca Avances",
     "descripcion": "Acceso desde la calle: hay lona temporal, no tótem" }
   ```
   `concepto` debe coincidir con la columna *Concepto* del Excel (no importan acentos ni mayúsculas).
3. Reglas que aplica el script:
   - **Congruencia:** la foto se rotula según el estatus del concepto: «Evidencia: cumple» (borde verde)
     o «Ejemplo de lo pendiente» (borde rojo). Al asignar, elige fotos que muestren lo que dice el
     estatus; si una foto contradice el Excel, ponla en `observaciones` (sección «Para verificar en la
     visita») en vez de asignarla al concepto.
   - Si hay varias fotos para un concepto, usa la **más reciente**. Fecha: EXIF de la foto → fecha del
     nombre «WhatsApp Image AAAA-MM-DD» → campo `fecha` → fecha de modificación del archivo.
   - Más de `meses_foto_reciente` (6) antes de la fecha de corte → etiqueta **«Foto no reciente»**.
   - Concepto sin foto → sólo un ícono simple.
   - Se reducen (`ancho_max_foto_px`, `calidad_jpeg`) y se incrustan en el HTML.
   - Debajo de cada foto se muestra su fecha y la ruta/archivo de Drive de donde salió.

## El HTML

- **Resumen del grupo:** agencias lado a lado, % general con semáforo, pendientes y matriz
  indicador × agencia.
- **Por agencia:** tarjetas por indicador (valor, semáforo, brecha contra 100%), desglose de
  Desarrollo de Proyecto, Cummins y Software (primero lo que **no** cumple), Personal (requeridas vs.
  cubiertas, por área y vacantes por puesto), fotos, observaciones y plan de trabajo.
- **Plan de trabajo:** Responsable, Fecha compromiso, Estatus y Comentarios se editan en pantalla y se
  guardan en el navegador del equipo donde se abre el archivo (localStorage). Para conservarlos o
  compartirlos, usa **Exportar a Excel** (.xlsx con una hoja por agencia).
- **Imprimir / PDF:** botón general (todo el grupo, cada agencia en página nueva) o por agencia.
  La fecha de corte aparece en cada página.
