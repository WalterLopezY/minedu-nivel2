# Validador de Tickets para Entregable — MINEDU

Herramienta web (HTML + Bootstrap + SheetJS + Chart.js + ExcelJS) para validar, muestrear y
analizar reportes de tickets exportados a Excel. No requiere backend ni instalación: todo el
procesamiento ocurre en el navegador del usuario.

## Cómo usarlo

1. Abre `index.html` en cualquier navegador (o publícalo con GitHub Pages, ver abajo).
2. Sube un archivo Excel (`.xlsx` / `.xls`) que contenga las columnas:
   - `Título`
   - `Estado`
   - `Categoría`
   - `Tipo`
   - `Estadísticas - Hora de resolución`
   - `ID`
   - `Asignado a - Técnico`
   - `Tiempo para resolver excedido`

   Al cargar el archivo se descartan automáticamente las filas cuyo `ID` no es un número
   (filas basura que GLPI genera cuando una descripción muy larga se desborda a la fila
   siguiente). La herramienta avisa cuántas filas se descartaron.
3. Elige un porcentaje en la lista desplegable y presiona **GENERAR CUADROS Y GRÁFICOS**.
4. Navega entre las pestañas: Cuadros al 100%, Cuadros al % asignado, Pareto,
   Gráficos al 100%, Tablas y Gráficos de %, y Anexos (descarga de Excel).

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

| Archivo | Para qué sirve |
|---|---|
| `index.html` | Página principal: carga del Excel, cuadros, gráficos y anexos |
| `informe.js` | Lectura del TdR/OS y edición del informe Word |
| `logo.png` | Logo del Ministerio de Educación que se muestra en la cabecera |
| `README.md` | Este documento |

Los tres primeros deben estar en la **misma carpeta** (la raíz del repositorio).

## Informe de Conformidad y Eleva Producto

En la sección **DOCUMENTOS** hay tres pestañas: **Entregable · Informe de Conformidad · Eleva Producto**.
El TdR y la OS se suben una sola vez y sirven para las tres. Cada Word se reconoce por su título
("INFORME TÉCNICO N.º …" = Conformidad, "INFORME DE SERVICIO Nº …" = Eleva Producto).

| Dato | Conformidad | Eleva Producto | Origen |
|---|---|---|---|
| N.º del documento | Informe Técnico N.º | Informe de Servicio Nº …-año/WJLY | Siguiente al del Word subido (editable) |
| Referencia / Asunto | Informe de Servicio N.º + OS | OS | Campo de la página + N.º de la OS (PDF) |
| Fecha | "29 de octubre de 2026" | "San Borja, 29 de octubre de 2026" | Fecha de hoy (editable) |
| Nombre del servicio | En mayúsculas | En mayúsculas, entre comillas | Denominación del TdR |
| Entregable | — | Texto del entregable elegido | RESULTADOS ESPERADOS del TdR |
| Actividades | Tabla (10 puntos) | Viñetas (10 puntos) | Características + Actividades del TdR |

El N.º de Informe de Servicio es el mismo en ambas pestañas (se sincroniza). Archivos de salida:
`Informe de conformidad - <Mes> <año>.docx` y `ELEVA PRODUCTO - <Mes> <año>.docx`.

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

## Dependencias externas (CDN — no requieren instalación)

- Bootstrap 5.3.3 + Bootstrap Icons 1.11.3
- SheetJS (xlsx) 0.18.5 — lectura de Excel
- Chart.js 4.5.1 + chartjs-plugin-datalabels 2.2.0 — gráfico de Pareto
- ExcelJS 4.4.0 — generación de los Excel de Anexos con estilo
- pdf.js 3.11.174 — lectura del TdR y la Orden de Servicio
- JSZip 3.10.1 — apertura y guardado del Word (.docx)
- Google Fonts (Poppins, Inter)
