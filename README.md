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
3. Elige un porcentaje en la lista desplegable y presiona **GENERAR CUADROS Y GRÁFICOS**.
4. Navega entre las pestañas: Cuadros al 100%, Cuadros al % asignado, Pareto,
   Gráficos al 100%, Tablas y Gráficos de %, y Anexos (descarga de Excel).

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
- Google Fonts (Poppins, Inter)
