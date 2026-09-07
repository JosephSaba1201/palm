# Prompt: Página web de Control de Obra (Presupuesto + Pagos)

> Plantilla reutilizable. Copia este texto completo, rellena los `[CORCHETES]` con los datos del nuevo proyecto y pégalo como primer mensaje a Claude para arrancar la página desde cero, siguiendo el mismo esquema de la página de PALM.

---

## 1. Contexto del proyecto (rellenar antes de enviar)

- **Nombre del proyecto:** [NOMBRE]
- **Ciudad / ubicación:** [CIUDAD]
- **Tipo:** [torres de departamentos / casas / mixto / etc.]
- **Unidades o torres:** [ej. Torre 1, Torre 2, Torre 3 — N unidades por torre]
- **Presupuesto total contratado:** [$ MXN]
- **Moneda(s) de pago:** [MXN / USD / mixto]
- **¿Hay comisiones de venta o de administración que rastrear?** [sí/no — quién cobra qué %]
- **¿Hay maestro de pagos ya existente (Excel/PDF) o se arranca desde cero?** [describir]

## 2. Objetivo principal (no negociable)

El objetivo central de la página es el **control financiero de la obra**, no solo mostrar avance o fotos. Todo lo demás es secundario a esto:

1. **Presupuesto de obra**: contratado vs. pagado, desglosado por partida, por torre/unidad y por contratista.
2. **Control de pagos**: cada pago individual registrado (fecha, beneficiario, concepto, partida, importe), con acumulados por proveedor, por partida y por periodo.
3. **División HARD / SOFT**: todo pago o partida debe poder clasificarse como HARD (obra física: estructura, acabados, instalaciones) o SOFT (honorarios, proyecto ejecutivo, gestión, permisos), más una categoría aparte para conceptos que **no** se suman al presupuesto de obra (ej. vigilancia, gastos no relacionados a construcción) — igual que en PALM, donde "Vigilancia" está fuera de Control de Pagos.
4. **Revisión semanal**: el sistema debe soportar una actualización periódica (semanal) donde se sube un PDF/reporte de pagos autorizados y se refleja en la página. Ver sección 5.

## 3. Esquema de secciones (mismo de la página que ya usamos)

Página de una sola vista (SPA) con navegación lateral/tabs entre estas secciones:

1. **Resumen** — panorama general: avance programado vs. real, temas que requieren atención, qué está pasando actualmente, próximos hitos, próximas 4 semanas.
2. **Avance de Obra** — avance programado vs. real (curva S), tabla de partidas de obra con programado/real/estado.
3. **Panorama Financiero** — composición del presupuesto, indicadores clave, flujo de gastos mensual, gasto acumulado vs. avance físico.
4. **Presupuesto** — consulta rápida por partida/torre/contratista, mezcla del presupuesto (HARD/SOFT), árbol de partidas con montos contratados.
5. **Control de Pagos** (sección más importante) — sub-tabs HARD / SOFT / (categoría "aparte", ej. vigilancia). Cada una con: consulta rápida por proveedor/partida/torre, pagos por mes, pagos por proveedor, pagos por partida, vista por semana/mes con detalle de todos los pagos, mayores pagos del periodo, proveedores nuevos.
6. **Comisiones** *(solo si aplica según sección 1)* — por unidad/torre, por receptor, detalle por unidad vendida.
7. **Ventas** *(solo si aplica)* — estatus de unidades, avance de ventas.
8. **Cronograma** — hitos del proyecto, próximas 4 semanas, actividad actual.
9. **Galería de Avance** — fotos organizadas cronológicamente, línea de tiempo, comparativo antes/ahora.
10. **Reporte Ejecutivo** — resumen imprimible: principales avances, pendientes, temas financieros que requieren atención, próximos hitos.

Mantener el mismo lenguaje visual: tarjetas (`card-head` + `h3` + hint), badges de estado, gráficas simples (barras/donas), todo en español, montos en MXN con formato `$X,XXX,XXX`.

## 4. División HARD / SOFT — reglas de clasificación

- **HARD**: cimentación, estructura, albañilería, acabados, instalaciones eléctricas/hidráulicas/sanitarias, elevadores, fachadas, amenidades construidas.
- **SOFT**: honorarios de arquitecto/ingeniero, proyecto ejecutivo, gerencia de obra, permisos y licencias, seguros de obra.
- **Aparte / no suma**: conceptos que se pagan desde el mismo flujo pero no son presupuesto de obra (ej. vigilancia, gastos administrativos ajenos a construcción) — mostrar por separado, nunca incluir en los totales de Control de Pagos.
- Cada pago debe llevar: **PARTIDA** (torre/frente) y **SUBPARTIDA** (oficio/concepto), para poder filtrar y agrupar.

## 5. Flujo de revisión semanal

1. Cada semana se recibe un PDF con los pagos autorizados de esa semana.
2. Ese PDF se clasifica por PARTIDA y SUBPARTIDA (por palabra clave en el concepto, con respaldo en el histórico del maestro para el mismo beneficiario) y se convierte a un Excel semanal con las mismas columnas del maestro.
3. Los renglones ambiguos (sin palabra clave ni precedente histórico claro) se marcan para revisión manual, nunca se adivinan en silencio.
4. El maestro (`[NOMBRE] YYYY SEM NN.xlsx`) es **acumulado** (todo lo pagado hasta esa semana, no solo el delta semanal) y vive en la carpeta de pagos (ver sección 6).
5. La actualización de `index.html` (los datos de pagos embebidos) se hace cuando el usuario lo pida explícitamente (ej. mensual o al cerrar el maestro) — **no automáticamente en cada PDF semanal**, salvo que el usuario indique lo contrario para este proyecto.
6. Reutilizar/adaptar el script `scripts/pdf_pagos_a_excel.py` de PALM como base para el parseo PDF → Excel.

## 6. Logo

- Si el proyecto **no tiene logo**, crearlo (versión simple, vectorial/SVG o PNG de alta resolución) coherente con el nombre y tipo de proyecto, en tonos neutros que combinen con el resto de la interfaz (fondo oscuro/claro).
- Guardarlo en `assets/logo-[nombre].png` (o `.svg`) y usarlo en el header de la página.

## 7. Estructura de carpetas a crear en el servidor local

Crear esta estructura dentro de una carpeta nueva del proyecto (ej. `Desktop/PAGINA [NOMBRE]/`):

```
PAGINA [NOMBRE]/
├── index.html                     ← página principal (SPA)
├── PRESUPUESTO BASE.xlsx          ← si existe, presupuesto contratado original
├── $/                             ← carpeta de pagos (fuente de verdad financiera)
│   ├── PRESUPUESTO BASE PAGO.xlsx        ← presupuesto base para cruzar contra pagos
│   └── [NOMBRE] YYYY SEM NN.xlsx         ← maestro acumulado de pagos (LISTA DE PAGOS)
├── assets/
│   ├── logo-[nombre].png/.svg     ← logo del proyecto
│   ├── planos/                    ← PDFs de planos/arquitectónicos completos (pesados, no tocar el código)
│   ├── sin_usar/                  ← archivos huérfanos detectados (duplicados exactos, capturas sin referenciar) — nunca se borran, solo se mueven aquí
│   └── [torre]_datos_reales.txt   ← notas de geometría/unidades si aplica corte interactivo
├── fotos/
│   ├── NN_YYYY-MM-DD.ext          ← una foto por archivo, numerada (01, 02...) + fecha real; el número fija el orden cronológico al ordenar alfabéticamente en Finder, sin depender de metadata
│   └── thumbs/
│       └── NN_YYYY-MM-DD.ext      ← miniatura (~700px de ancho) del archivo del mismo nombre en fotos/, generada con `sips --resampleWidth 700`
├── scripts/
│   ├── pdf_pagos_a_excel.py       ← conversión PDF semanal → Excel clasificado
│   └── sync_galeria.py            ← sincroniza fotos/ con el arreglo GALERIA_RAW de index.html: detecta fotos nuevas (fecha por EXIF o nombre de archivo, nunca inventada), renumera todo en orden cronológico, genera miniaturas faltantes y reescribe el bloque en el HTML. Se corre cada vez que se agregan fotos nuevas a la carpeta.
└── .backups/                      ← respaldo automático de index.html antes de cada cambio grande
```

### Qué hay que llenar en cada carpeta antes de empezar a construir la página

- **`$/`**: el presupuesto base contratado (Excel) y, si ya hay historial de pagos, el maestro acumulado con las columnas mínimas: SEMANA DE AÑO, FECHA, HARD/SOFT/FUERA DE OBRA, BENEFICIARIO, PAGADO SIN IVA, IVA, IMPORTE, PARTIDA, SUBPARTIDA, CONCEPTO, OBSERVACIONES.
- **`assets/`**: logo (o pedir que se cree), planos arquitectónicos en PDF/PNG dentro de `assets/planos/`, y cualquier dato de geometría de unidades si se quiere un corte interactivo tipo Torre 1 de PALM. Los archivos sin usar detectados durante limpieza van a `assets/sin_usar/`, nunca se eliminan directamente.
- **`fotos/`**: fotografías de avance de obra. El nombre de archivo es siempre `NN_YYYY-MM-DD.ext` (número de orden + fecha real, sin nombre original de cámara/WhatsApp) para que el orden alfabético del sistema de archivos coincida con el orden cronológico. Cada foto en `fotos/` debe tener su miniatura correspondiente (mismo nombre) en `fotos/thumbs/`. El código embebido en `index.html` (arreglo `GALERIA_RAW`) es la fuente de verdad de fecha/categoría/descripción de cada foto — al renombrar archivos hay que actualizar ese arreglo en el mismo cambio. Para eso existe `scripts/sync_galeria.py`: cualquier foto que se agregue a `fotos/` debe pasar por ese script (nunca agregarse a mano) para que aparezca en la página, quede bien numerada y tenga miniatura.
- **Línea de tiempo de Galería**: siempre ordenada de más reciente a más antigua (izquierda a derecha), incluyendo TODAS las categorías de foto (General, Estructura, Amenidades, Personal, etc.) — no debe filtrarse a una sola categoría, o las fotos más nuevas categorizadas distinto dejan de verse ahí.
- **`scripts/`**: no requiere nada previo, se genera durante el desarrollo.
- **Datos generales**: lista de unidades/departamentos con estatus de venta (si aplica Ventas/Comisiones), cronograma o lista de hitos del proyecto.

### Convención: mantener este documento al día

Cada vez que se haga un cambio de diseño o de estructura de carpetas/nomenclatura en la página de PALM (o en cualquier proyecto basado en esta plantilla), ese cambio debe reflejarse aquí mismo, en `PROMPT_NUEVA_OBRA.md`, para que la plantilla siga describiendo la convención real vigente y no una versión desactualizada.

## 8. Notas finales

- No usar frameworks pesados: página estática en HTML/CSS/JS embebido, igual que PALM, para que se pueda abrir localmente sin servidor.
- Todos los montos en pesos mexicanos salvo que se indique lo contrario; convertir pagos en USD usando el tipo de cambio impreso en el PDF de esa semana.
- Confirmar siempre con el usuario el número de semana de cada corte de pagos — nunca inferirlo del nombre del archivo ni del texto interno del PDF, ya que puede haber pagos cargados fuera de orden (backfill).
