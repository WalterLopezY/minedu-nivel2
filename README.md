# Validador de Tickets para Entregable — MINEDU

Herramienta web (HTML + Bootstrap + SheetJS + Chart.js + ExcelJS) para validar, muestrear y
analizar reportes de tickets exportados a Excel. No requiere backend ni instalación: todo el
procesamiento ocurre en el navegador del usuario.

## Cómo usarlo (3 pasos)

1. **Archivos**: suelta en una sola zona, juntos o uno por uno: el Excel de tickets (GLPI), el TdR y la
   Orden de Servicio (PDF), los Word del Entregable, Conformidad, Eleva Producto y Manual del mes, y las
   imágenes de la Ficha y de la Categorización. Cada archivo se identifica solo.
2. **Datos**: porcentaje, entregable (cuota de la OS), enlace del manual y números/fechas de
   Conformidad y Eleva Producto (se sugieren solos; se pueden cambiar).
3. **Generar**: **GENERAR LOS 3 DOCUMENTOS (ZIP)** crea los Word y los empaqueta en
   `Documentos - <Mes> <año>.zip`. También se pueden generar uno por uno o solo ver los cuadros y gráficos.
   Para el PDF: abrir cada Word → Archivo → Guardar como → PDF.

La barra de pasos de arriba muestra en verde lo que ya está listo.

## Informe en Word (TdR + Orden de Servicio + Word)

En la sección **DOCUMENTOS DEL INFORME** se suben, juntos o por separado: el Excel de tickets,
el TdR (PDF), la Orden de Servicio (PDF), el informe (Word .docx) y opcionalmente las imágenes
de la Ficha de evaluación y la Gráfica 11. Cada archivo se identifica solo: los PDF por su
título ("TÉRMINOS DE REFERENCIA" / "ORDEN DE SERVICIO") y las imágenes por su nombre de archivo
(debe contener "ficha" o "grafica 11").

Al presionar **PROCESAR INFORME** primero se generan los cuadros y gráficos con el % elegido
(igual que el botón GENERAR CUADROS Y GRÁFICOS) y luego se reemplaza en el Word, conservando su formato:

- Imágenes de **Gráfica 16: Ficha de evaluación** y **Gráfica 11**: la imagen nueva se coloca dentro
  del mismo recuadro de la anterior, centrada y sin deformarse (si no se suben, quedan las actuales).

- Portada, párrafo debajo de **SERVICIO:** → "DENOMINACIÓN DE LA CONTRATACIÓN" del TdR.
- Portada, **ENTREGABLE N°** → número de la cuota de la OS elegida (se sugiere según las
  fechas de apertura de los tickets del Excel).
- **4. OBJETIVO ESPECÍFICO** → todos los puntos de "CARACTERÍSTICAS Y CONDICIONES DEL
  SERVICIO" del TdR, incluidos los de su apartado "ACTIVIDADES".

- **Textos con datos del Excel** (requiere haber cargado el Excel): resumen del mes, análisis de
  las Tablas/Gráficos 1 a 6 (con las categorías más solicitadas), párrafos y títulos del análisis
  al % elegido en la página, títulos de tablas/gráficas con % (también en el índice), mes del
  manual y conclusiones. Solo se cambian las palabras distintas, así se conserva el formato.
  Las redacciones están en `textosBloqueA()` y los nombres cortos de categorías en
  `CATEGORIA_NOMBRE_FIJO` / `CATEGORIA_OBJETO` dentro de `informe.js`.

- **Tablas y gráficos del Word** (Tablas 1, 3–6 y Gráficos 1–6, 12–15): se reemplazan por los cuadros y
  gráficos que genera la página con el Excel y el % elegido, en alta resolución. Se ubican por su
  título (no por su número); la lista está en `FIGURAS_BLOQUE_B` dentro de `informe.js`.
  La Tabla 2, la ficha, GLPI, los catálogos y el procedimiento no se tocan.
- **Tabla de anexos**: se rehace con los tickets del Excel (encabezado + tickets al % elegido,
  encabezado + todos los tickets al 100%), con los mismos datos y orden que los Excel de la pestaña
  Anexos; también se actualiza el texto "Total de incidencias y solicitudes al X%".
- **Manual del mes (Word)**: tiene su propia casilla; un .docx se reconoce como informe si su portada
  tiene "SERVICIO:" / "ENTREGABLE N°", y como manual en otro caso o si su nombre contiene "manual".
- **Contenido del manual**: se borra el manual anterior (entre el enlace del manual y el siguiente
  objetivo) y se inserta el Word del manual completo en una página nueva: textos, numeración,
  imágenes, tablas e hipervínculos. Sus estilos se copian con el nombre "Manual - …" (no cambian
  los del informe y sus títulos no entran al índice).
- **Enlace del manual**: al subir el manual aparecen dos campos: la dirección (link) y el texto
  (se sugiere la primera línea del manual). Se coloca como hipervínculo debajo de
  "Link de acceso al manual – Mes de …", reemplazando el enlace anterior.
- **Orden del documento**: figuras renumeradas en el orden en que aparecen, todas como "Gráfico N"
  (y "Tabla N"), corrigiendo títulos, índice y menciones en el texto; tablas y gráficos se mantienen
  completos en una página junto a su título y su "Fuente" (excepto en ANEXOS); los subtítulos no
  quedan solos al pie de página. El índice se actualiza al abrir el Word: aceptar el mensaje
  "¿Desea actualizar los campos?". El nombre "Gráfico" se cambia en `NOMBRE_FIGURA` (`informe.js`).

Antes de descargar se muestra una lista de chequeo y una vista previa con el texto actual y el nuevo.
El archivo se descarga como `ENTREGABLE_<año> - <MES> <año>.docx` (ej. `ENTREGABLE_2026 - OCT 2026.docx`).
La lógica está en `informe.js`. Los PDF deben tener texto seleccionable (no escaneados).

## Archivos del repositorio

| Archivo / carpeta | Para qué sirve |
|---|---|
| `index.html` | Página principal: carga del Excel, cuadros, gráficos y anexos |
| `informe.js` | Lectura del TdR/OS y edición de los Word (entregable, conformidad, eleva producto) |
| `logo.png` | Logo del Ministerio de Educación que se muestra en la cabecera |
| `libs/` | Librerías (Bootstrap, SheetJS, ExcelJS, Chart.js, pdf.js, JSZip) incluidas en el repositorio: la página funciona aunque la red bloquee sitios externos |
| `README.md` | Este documento |

Todo debe estar en la **raíz** del repositorio (la carpeta `libs/` con su subcarpeta `fonts/`).

## Detalles útiles

- **Imágenes**: el nombre del archivo debe contener `ficha` (Ficha de evaluación) o
  `categorizacion` (Categorización en GLPI).
- **Aviso de mes**: si el manual o su título mencionan un mes distinto al del entregable
  (por ejemplo "SET 2026" en un entregable de octubre), la página lo avisa antes de descargar.
- **Porcentajes**: enteros en textos y tablas; incidencias + solicitudes siempre suman 100%.
- **Tema claro/oscuro**: sigue la configuración de Windows/Mac; con el interruptor ☀/🌙 de la esquina
  superior derecha se elige manualmente (se recuerda en ese navegador).

## Publicarlo con GitHub Pages (gratis, sin servidor)

1. Sube este repositorio a GitHub (puede ser público o privado — Pages en repos privados
   requiere una cuenta de pago de GitHub; en repos públicos es gratis).
2. Entra a **Settings → Pages** del repositorio.
3. En "Source", elige la rama `main` y la carpeta `/ (root)`.
4. Guarda. En un par de minutos GitHub te da una URL tipo
   `https://tu-usuario.github.io/nombre-repo/` donde la herramienta queda funcionando en línea.

## Nota sobre el código

Este es un sitio estático: todo el HTML/CSS/JS se ejecuta en el navegador del usuario, así que
el código siempre es visible vía "Ver código fuente" o las herramientas de desarrollador — es
una limitación inherente a cualquier app web del lado del cliente, no algo particular de este
proyecto.

## Librerías (carpeta libs/, no requieren instalación)

- (incluidas en `libs/`) Bootstrap 5.3.3 + Bootstrap Icons 1.11.3
- SheetJS (xlsx) 0.18.5 — lectura de Excel
- Chart.js 4.5.1 + chartjs-plugin-datalabels 2.2.0 — gráfico de Pareto
- ExcelJS 4.4.0 — generación de los Excel de Anexos con estilo
- pdf.js 3.11.174 — lectura del TdR y la Orden de Servicio
- JSZip 3.10.1 — apertura y guardado del Word (.docx)
