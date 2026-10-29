/* =====================================================================
   informe.js — Generación del INFORME (Word) a partir de:
     • TdR  (PDF Términos de Referencia)
     • OS   (PDF Orden de Servicio)
     • Word (plantilla del entregable)
   Todo se procesa en el navegador; nada se sube a ningún servidor.

   ORGANIZACIÓN DEL ARCHIVO
     1. Utilidades de texto
     2. Lectura de PDF (pdf.js) -> líneas de texto con su posición
     3. Extracción de datos del TdR
     4. Extracción de datos de la OS
     5. Edición del Word (document.xml)
     6. Interfaz de la página (solo en el navegador)
   ===================================================================== */

/* ---------- 1. Utilidades de texto ---------- */

// quita tildes y pasa a mayúsculas: "Denominación" -> "DENOMINACION"
function normTxt(s){
  return String(s || '').normalize('NFD').replace(/[̀-ͯ]/g, '')
    .replace(/\s+/g, ' ').trim().toUpperCase();
}

// limpia espacios repetidos y espacios antes de signos de puntuación
function cleanText(s){
  return String(s || '').replace(/\s+/g, ' ').replace(/\s+([,.;:)])/g, '$1').trim();
}

/* ---------- 2. Lectura de PDF -> líneas ----------
   pdf.js entrega "pedazos" de texto con su posición (x, y). Aquí se juntan
   los pedazos que están a la misma altura para formar líneas. Cada página se
   desplaza 1000 puntos hacia abajo para poder tratar el documento como uno
   solo (y mayor = más arriba en el documento). */

async function pdfToItems(pdfjsLib, data){
  const doc = await pdfjsLib.getDocument({ data }).promise;
  const items = [];
  for(let p = 1; p <= doc.numPages; p++){
    const page = await doc.getPage(p);
    const tc = await page.getTextContent();
    tc.items.forEach(it => {
      if(!it.str || !it.str.trim()) return;
      items.push({
        str: it.str,
        x: it.transform[4],
        y: it.transform[5] - (p - 1) * 1000,
        w: it.width || 0,
        page: p
      });
    });
  }
  return items;
}

// agrupa pedazos en líneas (misma altura ±1.5 pt), ordenadas de arriba a abajo
function itemsToLines(items){
  const sorted = items.slice().sort((a, b) => b.y - a.y || a.x - b.x);
  const lines = [];
  sorted.forEach(it => {
    let line = lines.find(l => Math.abs(l.y - it.y) <= 1.5);
    if(!line){ line = { y: it.y, parts: [] }; lines.push(line); }
    line.parts.push(it);
  });
  lines.forEach(l => {
    l.parts.sort((a, b) => a.x - b.x);
    l.x = l.parts[0].x;
    let text = '';
    l.parts.forEach((p, i) => {
      if(i > 0){
        const prev = l.parts[i - 1];
        const gap = p.x - (prev.x + prev.w);
        if(gap > 1 && !text.endsWith(' ') && !p.str.startsWith(' ')) text += ' ';
      }
      text += p.str;
    });
    l.text = cleanText(text);
  });
  lines.sort((a, b) => b.y - a.y);
  return lines;
}

/* ---------- 3. Extracción del TdR ----------
   El TdR es una tabla de 2 columnas: a la izquierda la etiqueta
   ("DENOMINACIÓN DE LA CONTRATACIÓN", "ACTIVIDADES", ...) y a la derecha el
   contenido. Se separan ambas columnas por su posición horizontal. */

function parseTdr(items){
  const all = itemsToLines(items);
  const titulo = all.find(l => normTxt(l.text).includes('TERMINOS DE REFERENCIA'));

  // etiqueta "DENOMINACIÓN" -> define dónde empieza la columna de etiquetas
  const denomItem = items.find(it => normTxt(it.str).startsWith('DENOMINACION'));
  if(!denomItem) throw new Error('No se encontró "DENOMINACIÓN DE LA CONTRATACIÓN" en el TdR');
  const labelX = denomItem.x;

  // la columna de contenido: posición horizontal donde empiezan la mayoría de los
  // textos largos (las etiquetas como "DE", "LA" son cortas y no cuentan)
  const conteo = {};
  items.filter(it => it.page === denomItem.page && it.x > labelX + 40 && it.str.trim().length > 25)
       .forEach(it => { const k = Math.round(it.x); conteo[k] = (conteo[k] || 0) + 1; });
  const modas = Object.entries(conteo).sort((a, b) => b[1] - a[1]);
  if(!modas.length) throw new Error('No se pudo ubicar la columna de contenido del TdR');
  const valueX = Number(modas[0][0]);

  const labelLines = itemsToLines(items.filter(it => it.x >= labelX - 2 && it.x < valueX - 2));
  const valueLines = itemsToLines(items.filter(it => it.x >= valueX - 2));

  // separación típica entre renglones de un mismo párrafo (mediana)
  const gaps = [];
  for(let i = 1; i < valueLines.length; i++){
    const g = valueLines[i - 1].y - valueLines[i].y;
    if(g > 0 && g < 15) gaps.push(g);
  }
  gaps.sort((a, b) => a - b);
  const lineGap = gaps.length ? gaps[Math.floor(gaps.length / 2)] : 8.6;
  const sameParaMax = lineGap * 1.03; // más que esto = empieza otro bloque

  const findLabel = (prefix) => labelLines.find(l => normTxt(l.text).startsWith(prefix));

  // --- DENOMINACIÓN DE LA CONTRATACIÓN: renglones seguidos desde la altura de la etiqueta
  const denomLabel = findLabel('DENOMINACION');
  let idx = valueLines.findIndex(l => Math.abs(l.y - denomLabel.y) <= 2);
  if(idx < 0) throw new Error('No se encontró el texto de la denominación en el TdR');
  const denomParts = [valueLines[idx].text];
  for(let i = idx + 1; i < valueLines.length; i++){
    if(valueLines[i - 1].y - valueLines[i].y > sameParaMax) break;
    denomParts.push(valueLines[i].text);
  }
  const denominacion = cleanText(denomParts.join(' '));

  // --- lista de puntos dentro de una franja vertical (de "top" hacia abajo hasta "bottom")
  // un punto nuevo empieza cuando: hay más espacio que entre renglones, o el renglón
  // empieza con un verbo en infinitivo con mayúscula (Resolver, Analizar, Velar...)
  const esInfinitivo = (t) => /^[A-ZÁÉÍÓÚÑ][a-záéíóúñ]+(ar|er|ir)\b/.test(t);
  function puntosEntre(top, bottom){
    const lines = valueLines.filter(l => l.y <= top && l.y > bottom);
    const puntos = [];
    lines.forEach((l, i) => {
      const gap = i > 0 ? lines[i - 1].y - l.y : Infinity;
      if(i === 0 || gap > sameParaMax || esInfinitivo(l.text)) puntos.push(l.text);
      else puntos[puntos.length - 1] += ' ' + l.text;
    });
    return puntos.map(cleanText).filter(Boolean);
  }

  const actLabel = findLabel('ACTIVIDADES');
  const caracLabel = findLabel('CARACTERISTICAS Y CONDICIONES');
  // la etiqueta que viene después de ACTIVIDADES marca el final de esa sección
  const afterAct = actLabel ? labelLines.find(l => l.y < actLabel.y - 2) : null;

  const actividades = actLabel
    ? puntosEntre(actLabel.y + 5, afterAct ? afterAct.y + 3 : actLabel.y - 400)
    : [];
  const caracteristicas = (caracLabel && actLabel)
    ? puntosEntre(caracLabel.y - 3, actLabel.y + 5)
    : [];

  // RESULTADOS ESPERADOS O ENTREGABLES: "Del 03/10/2026 hasta 01/11/2026 Primer entregable:" + su texto
  const entregables = [];
  let actual = null;
  valueLines.forEach(l => {
    const m = l.text.match(/^Del\s+(\d{2}\/\d{2}\/\d{4})\s+hasta\s+(\d{2}\/\d{2}\/\d{4})\s+(\S+)\s+entregable\s*:?\s*(.*)$/i);
    if(m){ actual = { inicio: m[1], fin: m[2], nombre: m[3], partes: m[4] ? [m[4]] : [] }; entregables.push(actual); return; }
    if(!actual) return;
    if(/^(Nota|LUGAR|DIRECCI)/i.test(l.text)){ actual = null; return; }
    actual.partes.push(l.text);
    if(/entregable\.\s*$/i.test(l.text)) actual = null;   // fin del texto del entregable
  });
  entregables.forEach(e => { e.texto = cleanText(e.partes.join(' ')); delete e.partes; });

  return { titulo: !!titulo, denominacion, actividades, caracteristicas, entregables };
}

/* ---------- 4. Extracción de la OS ----------
   Tabla de cuotas: "001  30  03/10/2026  01/11/2026  3,475.00" */

function parseOs(items){
  const lines = itemsToLines(items);
  const re = /\b(\d{3})\s+(\d{1,3})\s+(\d{2})\/(\d{2})\/(\d{4})\s+(\d{2})\/(\d{2})\/(\d{4})\s+([\d.,]+)/;
  const cuotas = [];
  lines.forEach(l => {
    const m = l.text.match(re);
    if(!m) return;
    const numero = parseInt(m[1], 10);
    if(cuotas.some(c => c.numero === numero)) return;
    cuotas.push({
      numero,
      plazo: parseInt(m[2], 10),
      inicio: new Date(+m[5], +m[4] - 1, +m[3]),
      fin: new Date(+m[8], +m[7] - 1, +m[6]),
      inicioTxt: `${m[3]}/${m[4]}/${m[5]}`,
      finTxt: `${m[6]}/${m[7]}/${m[8]}`,
      monto: m[9]
    });
  });
  cuotas.sort((a, b) => a.numero - b.numero);
  // N.º de la orden: "ORDEN DE SERVICIO Nº 0008060"
  const lnum = lines.find(l => /ORDEN DE SERVICIO\s*N/i.test(normTxt(l.text)) && /\d{5,}/.test(l.text));
  const numeroOs = lnum ? lnum.text.match(/(\d{5,})/)[1] : '';
  return { cuotas, numeroOs };
}

// identifica un PDF por su título: 'tdr' | 'os' | null
function clasificarPdf(items){
  const lines = itemsToLines(items.filter(it => it.page === 1));
  for(const l of lines){              // de arriba hacia abajo: gana el primero que aparezca
    const t = normTxt(l.text);
    if(t.includes('TERMINOS DE REFERENCIA')) return 'tdr';
    if(t.includes('ORDEN DE SERVICIO')) return 'os';
  }
  return null;
}

const MESES = ['enero','febrero','marzo','abril','mayo','junio','julio','agosto','setiembre','octubre','noviembre','diciembre'];
const MESES_ABR = ['ENE','FEB','MAR','ABR','MAY','JUN','JUL','AGO','SET','OCT','NOV','DIC'];

// nombre del Word descargado, con el formato de siempre: "ENTREGABLE_2026 - OCT 2026.docx"
function nombreArchivoInforme(mesIndice, anio){
  return `ENTREGABLE_${anio} - ${MESES_ABR[mesIndice]} ${anio}.docx`;
}

// elige la cuota cuyo periodo contiene más fechas de tickets del Excel
function sugerirCuota(cuotas, fechas){
  if(!cuotas.length) return null;
  if(fechas && fechas.length){
    let best = null, bestN = -1;
    cuotas.forEach(c => {
      const fin = new Date(c.fin); fin.setHours(23, 59, 59);
      const n = fechas.filter(f => f >= c.inicio && f <= fin).length;
      if(n > bestN){ bestN = n; best = c; }
    });
    if(bestN > 0) return best;
  }
  const hoy = new Date();
  return cuotas.find(c => hoy >= c.inicio && hoy <= c.fin) || cuotas[0];
}

/* ---------- 5. Edición del Word ----------
   Un .docx es un ZIP; el texto está en word/document.xml. Se edita ese XML
   reutilizando el formato (fuente, tamaño, viñetas) de los párrafos existentes,
   así el diseño del documento no cambia. */

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

function closestParagraph(node){
  let n = node.parentNode;
  while(n && !(n.namespaceURI === W_NS && n.localName === 'p')) n = n.parentNode;
  return n;
}

// texto propio del párrafo (sin el de cuadros de texto que pueda contener)
function paraText(p){
  const ts = p.getElementsByTagNameNS(W_NS, 't');
  let s = '';
  for(let i = 0; i < ts.length; i++) if(closestParagraph(ts[i]) === p) s += ts[i].textContent;
  return s;
}

function directRuns(p){
  const rs = p.getElementsByTagNameNS(W_NS, 'r');
  const out = [];
  for(let i = 0; i < rs.length; i++) if(closestParagraph(rs[i]) === p) out.push(rs[i]);
  return out;
}

// reemplaza el texto del párrafo conservando el formato de su primer tramo de texto
function setParagraphText(p, text){
  const runs = directRuns(p).filter(r => r.getElementsByTagNameNS(W_NS, 't').length);
  if(!runs.length) throw new Error('Párrafo sin texto para reemplazar');
  const first = runs[0];
  const ts = first.getElementsByTagNameNS(W_NS, 't');
  ts[0].textContent = text;
  ts[0].setAttributeNS('http://www.w3.org/XML/1998/namespace', 'xml:space', 'preserve');
  for(let i = ts.length - 1; i > 0; i--) ts[i].parentNode.removeChild(ts[i]);
  runs.slice(1).forEach(r => r.parentNode.removeChild(r));
  // quita marcas de ortografía sueltas que quedaron en el párrafo
  const pe = p.getElementsByTagNameNS(W_NS, 'proofErr');
  for(let i = pe.length - 1; i >= 0; i--) if(closestParagraph(pe[i]) === p) pe[i].parentNode.removeChild(pe[i]);
}

function allParagraphs(doc){
  return Array.from(doc.getElementsByTagNameNS(W_NS, 'p'));
}

function nextParagraphSibling(p){
  let n = p.nextSibling;
  while(n && !(n.nodeType === 1 && n.namespaceURI === W_NS && n.localName === 'p')) n = n.nextSibling;
  return n;
}

/* Aplica los cambios. Devuelve la lista de cambios para la vista previa.
   datos = { servicio, entregable (número), objetivos (array de textos) } */
function editarDocumento(doc, datos){
  const cambios = [];

  // (a) Portada: párrafo debajo de "SERVICIO:"
  //     (la portada está en un cuadro de texto que Word guarda 2 veces: versión
  //      moderna y versión de compatibilidad; se cambian ambas)
  let nServ = 0;
  allParagraphs(doc).forEach(p => {
    if(normTxt(paraText(p)) !== 'SERVICIO:') return;
    let sig = nextParagraphSibling(p);
    while(sig && !paraText(sig).trim()) sig = nextParagraphSibling(sig);
    if(!sig) return;
    const antes = paraText(sig);
    setParagraphText(sig, datos.servicio);
    if(nServ++ === 0) cambios.push({ campo: 'Portada — SERVICIO', antes, despues: datos.servicio });
  });
  if(!nServ) throw new Error('No se encontró "SERVICIO:" en la portada del Word');

  // (b) Portada: ENTREGABLE N° XX
  let nEnt = 0;
  const entTxt = 'ENTREGABLE N° ' + String(datos.entregable).padStart(2, '0');
  allParagraphs(doc).forEach(p => {
    const t = paraText(p);
    if(!/^\s*ENTREGABLE\s*N\s*[°º]\s*\d+\s*$/i.test(t)) return;
    setParagraphText(p, entTxt);
    if(nEnt++ === 0) cambios.push({ campo: 'Portada — ENTREGABLE N°', antes: t.trim(), despues: entTxt });
  });
  if(!nEnt) throw new Error('No se encontró "ENTREGABLE N°" en la portada del Word');

  // (c) Punto 4. OBJETIVO ESPECÍFICO: se reemplaza toda la lista
  const titulo = allParagraphs(doc).find(p => normTxt(paraText(p)) === 'OBJETIVO ESPECIFICO');
  if(!titulo) throw new Error('No se encontró el título "OBJETIVO ESPECÍFICO" en el Word');

  const viejos = [];
  let n = nextParagraphSibling(titulo);
  while(n && normTxt(paraText(n)) !== 'DETALLE DE ACTIVIDADES'){
    viejos.push(n);
    n = nextParagraphSibling(n);
  }
  const esLista = p => p.getElementsByTagNameNS(W_NS, 'numPr').length > 0 && paraText(p).trim();
  const modelo = viejos.find(esLista);
  if(!modelo) throw new Error('No se encontró la lista de puntos del OBJETIVO ESPECÍFICO');

  // se conserva el último párrafo vacío (separación antes del siguiente título)
  const ultimo = viejos[viejos.length - 1];
  const conservar = (ultimo && !paraText(ultimo).trim() && !esLista(ultimo)) ? ultimo : null;
  const antesLista = viejos.filter(esLista).map(paraText);
  const ancla = conservar || n;

  datos.objetivos.forEach(texto => {
    const nuevo = modelo.cloneNode(true);
    setParagraphText(nuevo, texto);
    ancla.parentNode.insertBefore(nuevo, ancla);
  });
  viejos.forEach(p => { if(p !== conservar) p.parentNode.removeChild(p); });

  cambios.push({ campo: '4. OBJETIVO ESPECÍFICO', antes: antesLista, despues: datos.objetivos });

  // (d) Bloque A: párrafos con datos del Excel
  if(datos.textos) cambios.push(...aplicarTextos(doc, datos.textos));
  return cambios;
}

/* ---------- 5b. Reemplazo "fino" dentro de un párrafo ----------
   En vez de borrar el párrafo y escribirlo de nuevo, se compara el texto
   viejo con el nuevo palabra por palabra y solo se cambian las partes
   distintas. Así se conservan negritas, cursivas, etc. de lo que no cambia. */

// palabras, números y signos por separado: así "FECHA:02" cambia solo el "02"
function tokenizar(s){ return s.match(/\s+|[0-9A-Za-zÁÉÍÓÚÜÑáéíóúüñ]+|[^\s0-9A-Za-zÁÉÍÓÚÜÑáéíóúüñ]/g) || []; }

// devuelve los tramos a cambiar: [{ini, fin, texto}] en posiciones del texto viejo
function diferencias(viejo, nuevo){
  const A = tokenizar(viejo), B = tokenizar(nuevo);
  const n = A.length, m = B.length;
  const dp = Array.from({ length: n + 1 }, () => new Int32Array(m + 1));
  for(let i = n - 1; i >= 0; i--)
    for(let j = m - 1; j >= 0; j--)
      dp[i][j] = A[i] === B[j] ? dp[i + 1][j + 1] + 1 : Math.max(dp[i + 1][j], dp[i][j + 1]);
  const tramos = [];
  let i = 0, j = 0, pos = 0, actual = null;
  const cerrar = () => { if(actual){ tramos.push(actual); actual = null; } };
  while(i < n || j < m){
    if(i < n && j < m && A[i] === B[j]){ cerrar(); pos += A[i].length; i++; j++; }
    else if(j < m && (i >= n || dp[i][j + 1] >= dp[i + 1][j])){
      if(!actual) actual = { ini: pos, fin: pos, texto: '' };
      actual.texto += B[j]; j++;
    } else {
      if(!actual) actual = { ini: pos, fin: pos, texto: '' };
      pos += A[i].length; actual.fin = pos; i++;
    }
  }
  cerrar();
  return tramos;
}

function setTextoNodo(t, valor){
  t.textContent = valor;
  t.setAttributeNS('http://www.w3.org/XML/1998/namespace', 'xml:space', 'preserve');
}

function reemplazarRango(nodos, ini, fin, texto){
  if(ini === fin){ // solo insertar
    const nd = nodos.find(x => ini > x.ini && ini <= x.ini + x.len) || nodos.find(x => ini >= x.ini && ini <= x.ini + x.len);
    if(!nd) return;
    const v = nd.t.textContent, o = ini - nd.ini;
    setTextoNodo(nd.t, v.slice(0, o) + texto + v.slice(o));
    return;
  }
  let puesto = false;
  nodos.forEach(nd => {
    const a = Math.max(ini - nd.ini, 0), b = Math.min(fin - nd.ini, nd.len);
    if(b <= 0 || a >= nd.len) return;
    const v = nd.t.textContent;
    setTextoNodo(nd.t, v.slice(0, a) + (puesto ? '' : texto) + v.slice(b));
    puesto = true;
  });
}

function reemplazarTextoFino(p, nuevo){
  const ts = Array.from(p.getElementsByTagNameNS(W_NS, 't')).filter(t => closestParagraph(t) === p);
  let acc = 0;
  const nodos = ts.map(t => { const nd = { t, ini: acc, len: t.textContent.length }; acc += nd.len; return nd; });
  const viejo = ts.map(t => t.textContent).join('');
  // se aplica de atrás hacia adelante para que las posiciones anteriores no se muevan
  diferencias(viejo, nuevo).reverse().forEach(d => reemplazarRango(nodos, d.ini, d.fin, d.texto));
}

/* textos = [{ campo, busca: RegExp (inicio del párrafo), nuevo: string }] */
function aplicarTextos(doc, textos){
  const ps = allParagraphs(doc);
  const cambios = [];
  textos.forEach(tx => {
    // modo "todos": cambia cada párrafo que coincida (ej. títulos de tablas en el cuerpo y en el índice)
    if(tx.todos){
      const antes = [], despues = [];
      ps.filter(p => tx.busca.test(paraText(p).trim())).forEach(p => {
        const viejo = paraText(p), nuevo = tx.transformar(viejo);
        if(nuevo === viejo) return;
        reemplazarTextoFino(p, nuevo);
        antes.push(viejo.trim()); despues.push(nuevo.trim());
      });
      if(antes.length) cambios.push({ campo: tx.campo, antes, despues });
      return;
    }
    const p = ps.find(p => tx.busca.test(paraText(p).trim()));
    if(!p){ cambios.push({ campo: tx.campo, antes: '(no se encontró este párrafo en el Word)', despues: tx.nuevo, falta: true }); return; }
    const antes = paraText(p).trim();
    if(antes === tx.nuevo.trim()) return;
    reemplazarTextoFino(p, tx.nuevo);
    cambios.push({ campo: tx.campo, antes, despues: tx.nuevo });
  });
  return cambios;
}

/* ---------- 5b'. Reemplazo de imágenes (Ficha de evaluación, Gráfica 11) ----------
   Se busca el título ("Gráfica 16: Ficha de evaluación…") y la imagen que está en ese
   párrafo o en los siguientes. La imagen nueva se coloca DENTRO del mismo recuadro
   (mismo ancho y alto máximo, centrada), así no se mueve nada del documento. */

const NS_WP  = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
const NS_A   = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const NS_R   = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const NS_REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const NS_CT  = 'http://schemas.openxmlformats.org/package/2006/content-types';

// qué imagen es cada una (por el nombre del archivo) y dónde va en el Word
const IMAGENES_INFORME = {
  ficha: {
    label: 'Ficha de evaluación', titulo: /^Gr[aá]fic[ao]\s*\d+\s*:\s*Ficha de evaluaci/i,
    reconocer: n => n.includes('FICHA')
  },
  grafica11: {
    label: 'Categorización en GLPI (imagen "grafica 11")', titulo: /^Gr[aá]fic[ao]\s*\d+\s*:\s*Atenci[oó]n correctamente categorizada/i,
    reconocer: n => /GRAFIC[AO]?[\s_\-.]*11\b/.test(n) || /\bG[\s_\-]*11\b/.test(n)
  }
};

function tipoImagenPorNombre(nombre){
  const n = normTxt(nombre.replace(/\.[^.]+$/, ''));
  for(const [k, def] of Object.entries(IMAGENES_INFORME)) if(def.reconocer(n)) return k;
  return null;
}

/* zip: JSZip del Word · doc: document.xml ya leído · img: {bytes, ext:'png'|'jpeg', ancho, alto} */
async function reemplazarImagen(zip, doc, clave, img){
  const def = IMAGENES_INFORME[clave];
  await colocarImagenEnTitulo(zip, doc, t => def.titulo.test(t), def.label, clave, img, 'contener');
  return { campo: def.label, antes: '(imagen anterior)', despues: '(imagen nueva subida)', imagen: true };
}

/* Busca el título (no el del índice) y reemplaza la imagen que lo acompaña.
   modo 'contener': la imagen nueva entra en el recuadro anterior (para fotos/capturas subidas).
   modo 'ancho'   : mismo ancho que la anterior y alto según la imagen nueva (para cuadros y gráficos). */
async function colocarImagenEnTitulo(zip, doc, coincide, label, clave, img, modo){
  const ps = allParagraphs(doc);
  const titulos = ps.map((p, i) => (!esIndice(p) && coincide(paraText(p).trim())) ? i : -1).filter(i => i >= 0);
  if(!titulos.length) throw new Error(`no se encontró el título "${label}" en el Word`);
  // la imagen: en el mismo párrafo del título o en los 3 siguientes (sin pasar al siguiente título)
  let blip = null;
  for(const iTit of titulos){
    for(let k = iTit; k <= iTit + 3 && k < ps.length && !blip; k++){
      if(k > iTit && tipoTitulo(paraText(ps[k]))) break;
      blip = ps[k].getElementsByTagNameNS(NS_A, 'blip')[0] || null;
    }
    if(blip) break;
  }
  if(!blip) throw new Error(`no se encontró la imagen de "${label}" junto a su título`);

  // 1) guardar el archivo nuevo dentro del Word y enlazarlo
  const nombreMedia = `media/claude_${clave}.${img.ext}`;
  zip.file('word/' + nombreMedia, img.bytes);

  const relsPath = 'word/_rels/document.xml.rels';
  const rels = new DOMParser().parseFromString(await zip.file(relsPath).async('string'), 'application/xml');
  const raiz = rels.documentElement;
  const rid = 'rIdClaude_' + clave;
  Array.from(raiz.getElementsByTagNameNS(NS_REL, 'Relationship'))
    .filter(r => r.getAttribute('Id') === rid).forEach(r => raiz.removeChild(r));
  const rel = rels.createElementNS(NS_REL, 'Relationship');
  rel.setAttribute('Id', rid);
  rel.setAttribute('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image');
  rel.setAttribute('Target', nombreMedia);
  raiz.appendChild(rel);
  zip.file(relsPath, new XMLSerializer().serializeToString(rels));

  const ctPath = '[Content_Types].xml';
  const ct = new DOMParser().parseFromString(await zip.file(ctPath).async('string'), 'application/xml');
  const yaExiste = Array.from(ct.getElementsByTagNameNS(NS_CT, 'Default'))
    .some(d => (d.getAttribute('Extension') || '').toLowerCase() === img.ext);
  if(!yaExiste){
    const d = ct.createElementNS(NS_CT, 'Default');
    d.setAttribute('Extension', img.ext);
    d.setAttribute('ContentType', 'image/' + img.ext);
    ct.documentElement.insertBefore(d, ct.documentElement.firstChild);
    zip.file(ctPath, new XMLSerializer().serializeToString(ct));
  }
  blip.setAttributeNS(NS_R, 'r:embed', rid);

  // 2) ajustar tamaño: la imagen nueva entra en el recuadro original sin deformarse
  let dibujo = blip.parentNode;
  while(dibujo && !(dibujo.namespaceURI === NS_WP && (dibujo.localName === 'anchor' || dibujo.localName === 'inline'))) dibujo = dibujo.parentNode;
  const extent = dibujo && dibujo.getElementsByTagNameNS(NS_WP, 'extent')[0];
  if(extent){
    const cx = +extent.getAttribute('cx'), cy = +extent.getAttribute('cy');
    const r = img.ancho / img.alto;
    let ncx = cx, ncy = Math.round(cx / r);
    const altoMax = modo === 'ancho' ? 8280000 : cy;          // 'ancho': hasta 23 cm de alto
    if(ncy > altoMax){ ncy = altoMax; ncx = Math.round(altoMax * r); }
    extent.setAttribute('cx', ncx); extent.setAttribute('cy', ncy);
    Array.from(dibujo.getElementsByTagNameNS(NS_A, 'ext'))
      .filter(e => e.parentNode.localName === 'xfrm')
      .forEach(e => { e.setAttribute('cx', ncx); e.setAttribute('cy', ncy); });
    // centrar dentro del espacio original (si la imagen es flotante con posición fija)
    const posH = dibujo.getElementsByTagNameNS(NS_WP, 'positionH')[0];
    const off = posH && posH.getElementsByTagNameNS(NS_WP, 'posOffset')[0];
    if(off) off.textContent = String(Math.round(+off.textContent + (cx - ncx) / 2));
  }
  // quitar el recorte que tenía la imagen anterior
  const recorte = blip.parentNode.getElementsByTagNameNS(NS_A, 'srcRect')[0];
  if(recorte) recorte.parentNode.removeChild(recorte);
}

/* ---------- 5b'''. Bloque B: tablas y gráficos del Excel dentro del Word ----------
   Cada imagen del Word se reemplaza por el cuadro o gráfico que genera la página,
   identificándola por su título (no por su número, que puede cambiar).
   >>> Para cambiar qué imagen va en cada título, edita esta lista. */

const FIGURAS_BLOQUE_B = [
  { clave: 't1',  tipo: 'tabla', re: /Tickets\s+Por\s+Mesa/i,                  fuente: () => tablaAtenciones100() },
  { clave: 'g1',  tipo: 'fig',   re: /Diagrama de Pareto/i,                    fuente: () => imagenPareto() },
  { clave: 'g2',  tipo: 'fig',   re: /Cantidad de Peticiones por Estado/i,     canvas: 'chartEstado' },
  { clave: 'g3',  tipo: 'fig',   re: /categor[ií]as m[aá]s solicitadas/i,      canvas: 'chartCategoria' },
  { clave: 'g4',  tipo: 'fig',   re: /por tipo de atenci[oó]n/i,               canvas: 'chartTipo' },
  { clave: 'g5',  tipo: 'fig',   re: /por estado de atenci[oó]n/i,             canvas: 'chartEstado' },
  { clave: 'g6',  tipo: 'fig',   re: /Tiempo de resoluci[oó]n/i,               canvas: 'chartTiempo' },
  { clave: 't3',  tipo: 'tabla', re: /Resumen de servicios atendidos/i,        fuente: () => tablaTipoMuestra(false) },
  { clave: 'g12', tipo: 'fig',   re: /Distribuci[oó]n de atenciones por tipo/i, canvas: 'chartTipoPct' },
  { clave: 't4',  tipo: 'tabla', re: /\d+\s*% de atenciones gestionadas/i,     fuente: () => tablaTipoMuestra(true) },
  { clave: 'g13', tipo: 'fig',   re: /Representaci[oó]n en barras/i,           canvas: 'chartTipoPct' },
  { clave: 't5',  tipo: 'tabla', re: /\d+\s*% de solicitudes gestionadas/i,    fuente: () => tablaCategorias('solicitud', 'Solicitudes') },
  { clave: 'g14', tipo: 'fig',   re: /gesti[oó]n del \d+\s*% de solicitudes/i, canvas: 'chartSolicitudesCat' },
  { clave: 't6',  tipo: 'tabla', re: /\d+\s*% de incidencias gestionadas/i,    fuente: () => tablaCategorias('incidencia', 'Incidentes') },
  { clave: 'g15', tipo: 'fig',   re: /gesti[oó]n del \d+\s*% de atenciones/i,  canvas: 'chartIncidentesCat' }
];

function canvasABlob(cv){ return new Promise(r => cv.toBlob(r, 'image/png')); }
async function imagenDeCanvas(cv){
  const blob = await canvasABlob(cv);
  return { bytes: new Uint8Array(await blob.arrayBuffer()), ext: 'png', ancho: cv.width, alto: cv.height, blob };
}

/* Dibuja una tabla con el mismo estilo de la página (encabezado azul, filas celeste/blanco). */
function dibujarTabla(encabezados, filas, opts = {}){
  const ESC = (typeof ESCALA_GRAFICOS !== 'undefined') ? ESCALA_GRAFICOS : 2;
  const W = opts.ancho || 900, padX = 10, lineaH = 16, padY = 6, fuente = '12px Arial, sans-serif';
  const numCols = opts.numCols || encabezados.map((_, i) => i > 0);
  const anchosFijos = opts.anchos || encabezados.map((_, i) => i === 0 ? 0 : 130);
  const medir = document.createElement('canvas').getContext('2d');
  medir.font = fuente;
  const anchos = anchosFijos.slice();
  anchos[0] = W - anchosFijos.slice(1).reduce((a, b) => a + b, 0);
  const partir = (texto, ancho) => {   // corta el texto en varias líneas si no entra
    const palabras = String(texto).split(' '), lineas = [];
    let l = '';
    palabras.forEach(p => {
      const prueba = l ? l + ' ' + p : p;
      if(medir.measureText(prueba).width > ancho - 2 * padX && l){ lineas.push(l); l = p; } else l = prueba;
    });
    lineas.push(l);
    return lineas;
  };
  const filasLineas = [encabezados, ...filas].map(f => f.map((c, i) => partir(c, anchos[i])));
  const altos = filasLineas.map(f => Math.max(...f.map(ls => ls.length)) * lineaH + 2 * padY);
  const H = altos.reduce((a, b) => a + b, 0) + 1;
  const cv = document.createElement('canvas');
  cv.width = W * ESC; cv.height = H * ESC;
  const ctx = cv.getContext('2d');
  ctx.scale(ESC, ESC);
  let y = 0;
  filasLineas.forEach((f, r) => {
    const esEnc = r === 0, esTotal = opts.total && r === filasLineas.length - 1;
    ctx.fillStyle = esEnc ? '#156082' : (r % 2 === 1 ? '#C0E6F5' : '#FFFFFF');
    ctx.fillRect(0, y, W, altos[r]);
    let x = 0;
    f.forEach((lineas, c) => {
      ctx.strokeStyle = '#A9D4E8'; ctx.lineWidth = 1;
      ctx.strokeRect(x + 0.5, y + 0.5, anchos[c], altos[r]);
      ctx.fillStyle = esEnc ? '#FFFFFF' : '#1a1a1a';
      ctx.font = (esEnc || esTotal ? 'bold ' : '') + fuente;
      ctx.textBaseline = 'top';
      const der = numCols[c] && !esEnc;
      ctx.textAlign = der ? 'right' : 'left';
      lineas.forEach((l, k) => ctx.fillText(l, der ? x + anchos[c] - padX : x + padX, y + padY + k * lineaH + 2));
      x += anchos[c];
    });
    y += altos[r];
  });
  return cv;
}

function tablaAtenciones100(){
  const c = getTipoCounts(excelRows);
  return dibujarTabla(['ATENCIONES REALIZADAS', 'CANTIDAD'],
    [['INCIDENCIAS', c.incidencias], ['SOLICITUDES', c.solicitudes], ['TOTAL', c.total]],
    { total: true, anchos: [0, 260] });
}
function tablaTipoMuestra(conPorcentaje){
  const { sample } = getSeleccionPorPorcentaje(Number(percentSelect.value));
  const c = getTipoCounts(sample);
  if(!conPorcentaje)
    return dibujarTabla(['SOLICITUDES POR TIPO', 'TOTAL'],
      [['INCIDENCIAS', c.incidencias], ['SOLICITUDES', c.solicitudes], ['TOTAL', c.total]], { total: true, anchos: [0, 260] });
  return dibujarTabla(['SOLICITUDES POR TIPO', 'TOTAL', 'PORCENTAJE'],
    [['INCIDENCIAS', c.incidencias, pct(c.incidencias, c.total) + '%'],
     ['SOLICITUDES', c.solicitudes, pct(c.solicitudes, c.total) + '%'],
     ['TOTAL', c.total, '100.0%']], { total: true, anchos: [0, 220, 220] });
}
function tablaCategorias(tipo, etiqueta){
  const { sample } = getSeleccionPorPorcentaje(Number(percentSelect.value));
  const filas = Object.entries(getCategoryCountsByTipo(sample, tipo)).sort((a, b) => b[1] - a[1]);
  return dibujarTabla(['Categoría', etiqueta], filas, { anchos: [0, 110] });
}

// el Pareto de la página está en una pestaña oculta; para el Word se dibuja aparte a tamaño fijo
async function imagenPareto(){
  const counts = {};
  excelRows.forEach(r => { const k = String(r['Categoría']).trim() || '(sin categoría)'; counts[k] = (counts[k] || 0) + 1; });
  const total = Object.values(counts).reduce((a, b) => a + b, 0);
  const orden = Object.entries(counts).sort((a, b) => b[1] - a[1]);
  let acum = 0;
  const cums = orden.map(([, v]) => { acum += v / total * 100; return +acum.toFixed(1); });
  const corto = s => s.length > 45 ? s.slice(0, 44) + '…' : s;
  const caja = document.createElement('div');
  caja.style.cssText = 'position:fixed; left:-20000px; top:0; width:1100px; height:650px;';
  const cv = document.createElement('canvas');
  cv.width = 1100; cv.height = 650;
  caja.appendChild(cv); document.body.appendChild(caja);
  const chart = new Chart(cv.getContext('2d'), {
    data: { labels: orden.map(([k]) => corto(k)), datasets: [
      { type: 'bar', label: 'Frecuencia', data: orden.map(([, v]) => v), backgroundColor: '#46D55A', order: 2, yAxisID: 'y' },
      { type: 'line', label: '% Acumulado', data: cums, borderColor: '#FFFF00', backgroundColor: '#FFFF00', tension: 0.25, order: 1, yAxisID: 'y1', pointRadius: 3 }
    ]},
    options: {
      responsive: false, animation: false, devicePixelRatio: (typeof ESCALA_GRAFICOS !== 'undefined') ? ESCALA_GRAFICOS : 2,
      plugins: { customCanvasBackgroundColor: { color: '#156082' }, legend: { labels: { color: '#FFFFFF' } }, datalabels: { display: false } },
      scales: {
        x: { ticks: { color: '#FFFFFF', font: { size: 9 }, maxRotation: 90, minRotation: 90, autoSkip: false }, grid: { color: '#E77032' } },
        y: { position: 'left', beginAtZero: true, max: total, title: { display: true, text: 'Frecuencia', color: '#FFFFFF' }, ticks: { color: '#FFFFFF', precision: 0 }, grid: { color: '#E77032' } },
        y1: { position: 'right', min: 0, max: 100, title: { display: true, text: '% Acumulado', color: '#FFFFFF' }, ticks: { color: '#FFFFFF' }, grid: { drawOnChartArea: false } }
      }
    }
  });
  const img = await imagenDeCanvas(cv);
  chart.destroy(); caja.remove();
  return img;
}

/* Reemplaza las 15 imágenes. Requiere que antes se hayan generado los cuadros y gráficos. */
async function insertarBloqueB(zip, doc){
  const resultados = [];
  for(const f of FIGURAS_BLOQUE_B){
    const coincide = t => { const tt = tipoTitulo(t); return tt && tt.tipo === f.tipo && f.re.test(t.replace(RE_TITULO, '')); };
    const titulo = (allParagraphs(doc).map(paraText).find(t => !/\d+\s*$/.test(t) && coincide(t.trim())) ||
                    allParagraphs(doc).map(paraText).find(t => coincide(t.trim())) || f.clave).trim();
    try{
      let img;
      if(f.canvas){
        const cv = document.getElementById(f.canvas);
        if(!cv || !cv.width) throw new Error('el gráfico de la página aún no está generado');
        img = await imagenDeCanvas(cv);
      } else {
        const res = await f.fuente();
        img = res instanceof HTMLCanvasElement ? await imagenDeCanvas(res) : res;
      }
      await colocarImagenEnTitulo(zip, doc, coincide, titulo, 'b_' + f.clave, img, 'ancho');
      resultados.push({ titulo, url: URL.createObjectURL(img.blob) });
    }catch(e){
      resultados.push({ titulo, error: e.message });
    }
  }
  return resultados;
}

/* ---------- 5b''. Orden del documento ----------
   (1) Renumera figuras y tablas en el orden en que aparecen, todas como "Gráfico N" / "Tabla N",
       y corrige el índice y las menciones en el texto ("En la Gráfica N°3" -> "En el Gráfico N°3").
   (2) Evita que una tabla o un gráfico quede partido entre dos páginas: título + imagen + fuente
       se mantienen juntos (excepto en ANEXOS, donde las tablas son muy grandes).
   (3) El índice se actualiza solo al abrir el Word (ver activarActualizarIndice). */

const NOMBRE_FIGURA = 'Gráfico';   // >>> nombre único para todas las figuras

function pPrDe(p, crear){
  let pPr = Array.from(p.childNodes).find(n => n.namespaceURI === W_NS && n.localName === 'pPr');
  if(!pPr && crear){ pPr = p.ownerDocument.createElementNS(W_NS, 'w:pPr'); p.insertBefore(pPr, p.firstChild); }
  return pPr;
}
function ponerKeepNext(p){
  const pPr = pPrDe(p, true);
  if(Array.from(pPr.childNodes).some(n => n.localName === 'keepNext')) return;
  const kn = p.ownerDocument.createElementNS(W_NS, 'w:keepNext');
  const estilo = Array.from(pPr.childNodes).find(n => n.localName === 'pStyle');
  pPr.insertBefore(kn, estilo ? estilo.nextSibling : pPr.firstChild);
}
function esIndice(p){
  const pPr = pPrDe(p, false);
  const st = pPr && Array.from(pPr.childNodes).find(n => n.localName === 'pStyle');
  if(st && /^(TDC|TOC)/i.test(st.getAttributeNS(W_NS, 'val') || st.getAttribute('w:val') || '')) return true;
  return Array.from(p.getElementsByTagNameNS(W_NS, 'instrText')).some(t => /PAGEREF/.test(t.textContent));
}
const RE_TITULO = /^\s*(Tabla|Gr[aá]fic[ao])\s*(\d+)\s*:/i;
function tipoTitulo(texto){
  const m = texto.match(RE_TITULO);
  return m ? { tipo: /^tabla/i.test(m[1]) ? 'tabla' : 'fig', num: +m[2] } : null;
}
function cuerpoDe(doc){ return doc.getElementsByTagNameNS(W_NS, 'body')[0]; }
function hijosCuerpo(doc){
  return Array.from(cuerpoDe(doc).childNodes).filter(n => n.nodeType === 1 && (n.localName === 'p' || n.localName === 'tbl'));
}

// quita los campos automáticos (SEQ) del título y deja su número como texto fijo
function camposATexto(p){
  let modo = null;
  Array.from(p.childNodes).forEach(n => {
    if(n.nodeType !== 1 || n.localName !== 'r') return;
    const fc = n.getElementsByTagNameNS(W_NS, 'fldChar')[0];
    const tipo = fc && (fc.getAttributeNS(W_NS, 'fldCharType') || fc.getAttribute('w:fldCharType'));
    if(tipo === 'begin'){ modo = 'codigo'; p.removeChild(n); return; }
    if(tipo === 'separate'){ modo = 'resultado'; p.removeChild(n); return; }
    if(tipo === 'end'){ modo = null; p.removeChild(n); return; }
    if(modo === 'codigo') p.removeChild(n);
  });
}

function renumerar(doc){
  const ps = allParagraphs(doc);
  const mapa = { fig: {}, tabla: {} }, cont = { fig: 0, tabla: 0 };
  const titulos = ps.filter(p => !esIndice(p) && tipoTitulo(paraText(p)));
  titulos.forEach(p => {
    const t = tipoTitulo(paraText(p));
    mapa[t.tipo][t.num] = ++cont[t.tipo];
  });
  const nombre = t => t.tipo === 'tabla' ? 'Tabla' : NOMBRE_FIGURA;
  const cambiados = [];
  // títulos y entradas del índice
  ps.filter(p => tipoTitulo(paraText(p))).forEach(p => {
    const t = tipoTitulo(paraText(p));
    const nuevoNum = mapa[t.tipo][t.num];
    if(!nuevoNum) return;
    if(!esIndice(p)) camposATexto(p);
    const viejo = paraText(p);
    const nuevo = viejo.replace(/^(\s*)(Tabla|Gr[aá]fic[ao])\s*\d+/i, `$1${nombre(t)} ${nuevoNum}`);
    if(nuevo !== viejo){
      reemplazarTextoFino(p, nuevo);
      if(!esIndice(p)) cambiados.push(`${viejo.match(RE_TITULO)[0].replace(':','')} → ${nombre(t)} ${nuevoNum}`);
    }
  });
  // menciones dentro del texto
  let menciones = 0;
  const reFig = /(\b[lLeE][al]\s+)?([gG])r[aá]fic[ao](\s*(?:N\s*[.°º]+\s*)?)(\d+)/g;
  const reTab = /(\b[tT])abla(\s*(?:N\s*[.°º]+\s*)?)(\d+)/g;
  ps.filter(p => !esIndice(p) && !tipoTitulo(paraText(p))).forEach(p => {
    const viejo = paraText(p);
    let nuevo = viejo.replace(reFig, (m, art, g, sep, num) => {
      const n = mapa.fig[+num]; if(!n) return m;
      const artNuevo = art ? (art[0] === art[0].toUpperCase() ? 'El ' : 'el ') : '';
      const palabra = (g === 'G' ? NOMBRE_FIGURA : NOMBRE_FIGURA.toLowerCase());
      return artNuevo + palabra + sep + n;
    });
    nuevo = nuevo.replace(reTab, (m, t, sep, num) => { const n = mapa.tabla[+num]; return n ? `${t}abla${sep}${n}` : m; });
    if(nuevo !== viejo){ reemplazarTextoFino(p, nuevo); menciones++; }
  });
  return { cambiados, menciones, figuras: cont.fig, tablas: cont.tabla };
}

// convierte una imagen "flotante" en imagen "en línea con el texto" (no puede partirse ni quedar fuera de página)
function anclaAEnLinea(anchor, anchoMaxEmu){
  const d = anchor.ownerDocument;
  const hijo = name => Array.from(anchor.childNodes).find(n => n.localName === name);
  const ext = hijo('extent');
  let cx = +ext.getAttribute('cx'), cy = +ext.getAttribute('cy');
  if(cx > anchoMaxEmu){ cy = Math.round(cy * anchoMaxEmu / cx); cx = anchoMaxEmu; }
  const inline = d.createElementNS(NS_WP, 'wp:inline');
  ['distT','distB','distL','distR'].forEach(a => inline.setAttribute(a, '0'));
  const e = d.createElementNS(NS_WP, 'wp:extent'); e.setAttribute('cx', cx); e.setAttribute('cy', cy);
  inline.appendChild(e);
  const ee = d.createElementNS(NS_WP, 'wp:effectExtent'); ['l','t','r','b'].forEach(a => ee.setAttribute(a, '0'));
  inline.appendChild(ee);
  ['docPr', 'cNvGraphicFramePr'].forEach(n => { const h = hijo(n); if(h) inline.appendChild(h); });
  const graf = Array.from(anchor.childNodes).find(n => n.localName === 'graphic');
  inline.appendChild(graf);
  Array.from(graf.getElementsByTagNameNS(NS_A, 'ext'))
    .filter(x => x.parentNode.localName === 'xfrm')
    .forEach(x => { x.setAttribute('cx', cx); x.setAttribute('cy', cy); });
  anchor.parentNode.replaceChild(inline, anchor);
}

// subtítulo = título de sección (nivel de esquema / estilo de título) o línea corta en negrita sin punto final
function esSubtitulo(p){
  const texto = paraText(p).trim();
  if(!texto || tipoTitulo(texto) || /^\s*Fuente/i.test(texto)) return false;
  const pPr = pPrDe(p, false);
  if(pPr && Array.from(pPr.childNodes).some(x => x.localName === 'outlineLvl')) return true;
  const st = pPr && Array.from(pPr.childNodes).find(x => x.localName === 'pStyle');
  const estilo = st ? (st.getAttributeNS(W_NS, 'val') || st.getAttribute('w:val') || '') : '';
  if(/^(T[ií]tulo|Heading|NUEVO|Objetivos)/i.test(estilo)) return true;
  if(texto.length > 100 || /[.:;]$/.test(texto)) return false;
  const runs = directRuns(p).filter(r => r.getElementsByTagNameNS(W_NS, 't').length);
  const negrita = r => Array.from(r.getElementsByTagNameNS(W_NS, 'b')).some(b => {
    const v = b.getAttributeNS(W_NS, 'val') || b.getAttribute('w:val'); return !v || v === 'true' || v === '1';
  });
  return runs.length > 0 && runs.every(negrita);
}

function esParrafoVacio(p){
  if(p.localName !== 'p' || paraText(p).trim()) return false;
  const tiene = n => p.getElementsByTagNameNS(W_NS, n).length > 0;
  return !(tiene('drawing') || tiene('pict') || tiene('object') || tiene('sectPr') || tiene('fldChar') ||
           Array.from(p.getElementsByTagNameNS(W_NS, 'br')).some(b => (b.getAttribute('w:type') || b.getAttributeNS(W_NS, 'type')) === 'page'));
}

function mantenerJuntos(doc){
  const hijos = hijosCuerpo(doc);
  const iAnexos = hijos.findIndex(n => n.localName === 'p' && normTxt(paraText(n)) === 'ANEXOS' && !esIndice(n));
  const limite = iAnexos >= 0 ? iAnexos : hijos.length;
  // ancho útil (en EMU) de la sección donde está cada elemento: el Word tiene secciones con
  // márgenes distintos, y una imagen más ancha que su sección no se puede centrar
  const at = (el, a) => el ? +(el.getAttributeNS(W_NS, a) || el.getAttribute('w:' + a) || 0) : 0;
  const anchoDeSeccion = sect => {
    const pgSz = sect && sect.getElementsByTagNameNS(W_NS, 'pgSz')[0], pgMar = sect && sect.getElementsByTagNameNS(W_NS, 'pgMar')[0];
    return pgSz ? (at(pgSz, 'w') - at(pgMar, 'left') - at(pgMar, 'right')) * 635 : 5400000;
  };
  // la sección de un párrafo termina en el primer párrafo (desde él en adelante) que lleva <w:sectPr>;
  // la última sección es la del <w:sectPr> final del cuerpo
  const sectFinal = Array.from(cuerpoDe(doc).childNodes).find(n => n.localName === 'sectPr');
  const anchos = new Array(hijos.length);
  let actual = anchoDeSeccion(sectFinal);
  for(let k = hijos.length - 1; k >= 0; k--){
    const pPr = hijos[k].localName === 'p' ? pPrDe(hijos[k], false) : null;
    const s = pPr && Array.from(pPr.childNodes).find(x => x.localName === 'sectPr');
    if(s) actual = anchoDeSeccion(s);
    anchos[k] = actual;
  }

  let figuras = 0, tablas = 0;
  for(let i = 0; i < limite; i++){
    const n = hijos[i];
    // --- subtítulos: nunca quedan solos al pie de una página (van con el párrafo siguiente)
    if(n.localName === 'p' && !esIndice(n) && esSubtitulo(n)){
      ponerKeepNext(n);
      const pPr = pPrDe(n, true);
      if(!Array.from(pPr.childNodes).some(x => x.localName === 'keepLines')){
        const kn = Array.from(pPr.childNodes).find(x => x.localName === 'keepNext');
        pPr.insertBefore(doc.createElementNS(W_NS, 'w:keepLines'), kn.nextSibling);
      }
    }
    // --- tablas reales: ninguna fila se parte y la tabla entera va junta con su título
    if(n.localName === 'tbl'){
      Array.from(n.getElementsByTagNameNS(W_NS, 'tr')).forEach(tr => {
        let trPr = Array.from(tr.childNodes).find(x => x.localName === 'trPr');
        if(!trPr){
          trPr = doc.createElementNS(W_NS, 'w:trPr');
          const prex = Array.from(tr.childNodes).find(x => x.localName === 'tblPrEx');
          tr.insertBefore(trPr, prex ? prex.nextSibling : tr.firstChild);
        }
        if(!Array.from(trPr.childNodes).some(x => x.localName === 'cantSplit'))
          trPr.insertBefore(doc.createElementNS(W_NS, 'w:cantSplit'), trPr.firstChild);
        Array.from(tr.getElementsByTagNameNS(W_NS, 'p')).forEach(ponerKeepNext);
      });
      for(let k = i - 1; k >= 0 && k >= i - 3; k--){
        if(hijos[k].localName === 'p' && tipoTitulo(paraText(hijos[k]))){ ponerKeepNext(hijos[k]); break; }
      }
      tablas++;
      continue;
    }
    // --- figuras: título -> imagen(es) -> "Fuente: ..."
    if(n.localName !== 'p' || esIndice(n) || !tipoTitulo(paraText(n))) continue;
    let iF = -1;
    for(let k = i; k < Math.min(limite, i + 30); k++){
      const h = hijos[k];
      if(h.localName === 'tbl') break;
      if(k > i && tipoTitulo(paraText(h))) break;
      if(/^\s*Fuente/i.test(paraText(h))){ iF = k; break; }
    }
    if(iF < 0) continue;
    const tramo = hijos.slice(i, iF + 1);
    const runsImagen = [];
    tramo.forEach(p => Array.from(p.childNodes).forEach(r => {
      if(r.localName === 'r' && r.getElementsByTagNameNS(W_NS, 'drawing').length &&
         r.getElementsByTagNameNS(NS_A, 'graphic').length && !r.getElementsByTagName('mc:AlternateContent').length) runsImagen.push(r);
    }));
    if(!runsImagen.length) continue;
    const fuente = hijos[iF];
    const anchoMax = anchos[i];
    // líneas en blanco que solo reservaban espacio para la imagen flotante
    hijos.slice(i + 1, iF).forEach(p => { if(esParrafoVacio(p)) p.parentNode.removeChild(p); });
    runsImagen.forEach(r => {
      r.parentNode.removeChild(r);
      Array.from(r.getElementsByTagNameNS(NS_WP, 'anchor')).forEach(a => anclaAEnLinea(a, anchoMax));
      Array.from(r.getElementsByTagNameNS(NS_WP, 'inline')).forEach(inl => {
        const e = Array.from(inl.childNodes).find(x => x.localName === 'extent');
        const cx = +e.getAttribute('cx');
        if(cx > anchoMax){ e.setAttribute('cy', Math.round(+e.getAttribute('cy') * anchoMax / cx)); e.setAttribute('cx', anchoMax); }
      });
      const p = doc.createElementNS(W_NS, 'w:p');
      const pPr = doc.createElementNS(W_NS, 'w:pPr');
      pPr.appendChild(doc.createElementNS(W_NS, 'w:keepNext'));
      const sp = doc.createElementNS(W_NS, 'w:spacing'); sp.setAttribute('w:before', '60'); sp.setAttribute('w:after', '60');
      pPr.appendChild(sp);
      const jc = doc.createElementNS(W_NS, 'w:jc'); jc.setAttribute('w:val', 'center');
      pPr.appendChild(jc);
      p.appendChild(pPr); p.appendChild(r);
      fuente.parentNode.insertBefore(p, fuente);
    });
    // el párrafo donde estaba anclada la imagen puede haber quedado vacío: se quita para no cortar la unión
    hijos.slice(i + 1, iF).forEach(p => { if(p.parentNode && esParrafoVacio(p)) p.parentNode.removeChild(p); });
    // si "Fuente" quedó vacío de texto propio no se toca; el título se une a la imagen
    ponerKeepNext(n);
    figuras++;
  }
  return { figuras, tablas };
}

// pide a Word que actualice el índice (números de página) al abrir el documento
async function activarActualizarIndice(zip){
  const ruta = 'word/settings.xml';
  const f = zip.file(ruta);
  if(!f) return;
  const s = new DOMParser().parseFromString(await f.async('string'), 'application/xml');
  const raiz = s.documentElement;
  if(raiz.getElementsByTagNameNS(W_NS, 'updateFields').length) return;
  const uf = s.createElementNS(W_NS, 'w:updateFields'); uf.setAttribute('w:val', 'true');
  const despues = ['hdrShapeDefaults','footnotePr','endnotePr','compat','docVars','rsids','mathPr','attachedSchema',
    'themeFontLang','clrSchemeMapping','doNotIncludeSubdocsInStats','doNotAutoCompressPictures','forceUpgrade','captions',
    'readModeInkLockDown','smartTagType','schemaLibrary','shapeDefaults','doNotEmbedSmartTags','decimalSymbol','listSeparator'];
  const ref = Array.from(raiz.childNodes).find(n => n.nodeType === 1 && despues.includes(n.localName));
  raiz.insertBefore(uf, ref || null);
  zip.file(ruta, new XMLSerializer().serializeToString(s));
}

function ordenarDocumento(doc){
  const r = renumerar(doc);
  const j = mantenerJuntos(doc);
  return { campo: 'Orden del documento', orden: true, antes: '',
    despues: [
      `Numeración en orden: ${r.figuras} gráficos y ${r.tablas} tablas` + (r.cambiados.length ? ` (${r.cambiados.length} renumerados: ${r.cambiados.join(', ')})` : ''),
      `${r.menciones} párrafo(s) con menciones a gráficos/tablas corregidos`,
      `${j.figuras} gráficos y ${j.tablas} tabla(s) se mantienen completos en una página (anexos excluidos)`,
      'El índice se actualizará al abrir el Word (acepta el mensaje "¿Desea actualizar los campos?")'
    ] };
}

/* ---------- 5b''''. Tabla de anexos ----------
   La tabla que sigue al título ANEXOS se rehace con los tickets del Excel:
   encabezado + tickets al % elegido, encabezado + todos los tickets (100%).
   Son los mismos datos y el mismo orden que los Excel de la pestaña "Anexos".
   Se reutiliza el formato de las filas existentes (encabezado azul, filas celeste/blanco). */

function llenarAnexos(doc, pctElegido){
  const hijos = hijosCuerpo(doc);
  const iAnexos = hijos.findIndex(n => n.localName === 'p' && normTxt(paraText(n)) === 'ANEXOS' && !esIndice(n));
  if(iAnexos < 0) throw new Error('no se encontró el título ANEXOS en el Word');
  const textoFila = tr => Array.from(tr.getElementsByTagNameNS(W_NS, 'tc')).map(tc => paraText(tc.getElementsByTagNameNS(W_NS, 'p')[0] || tc).trim());
  const esEncabezado = tr => normTxt(textoFila(tr)[0] || '') === 'ITEM';
  // tablas de anexos (puede haber 1 o 2 si el Word ya fue generado antes con esta herramienta)
  const tablasAnexo = hijos.slice(iAnexos + 1).filter(n => n.localName === 'tbl' &&
    Array.from(n.childNodes).some(tr => tr.localName === 'tr' && esEncabezado(tr)));
  if(!tablasAnexo.length) throw new Error('no se encontró la tabla de anexos');
  const tabla = tablasAnexo[0];
  // si había una segunda tabla (de una generación anterior) se quita junto con el espacio entre ambas
  if(tablasAnexo.length > 1){
    const iT1 = hijos.indexOf(tabla), iTn = hijos.indexOf(tablasAnexo[tablasAnexo.length - 1]);
    hijos.slice(iT1 + 1, iTn).forEach(n => { if(n.localName === 'p' && esParrafoVacio(n)) n.parentNode.removeChild(n); });
  }

  const filas = tablasAnexo.flatMap(t => Array.from(t.childNodes).filter(n => n.localName === 'tr'));
  tablasAnexo.slice(1).forEach(t => t.parentNode.removeChild(t));
  const modeloEnc = filas.find(esEncabezado);
  const datos = filas.filter(tr => !esEncabezado(tr));
  if(!modeloEnc || datos.length < 2) throw new Error('la tabla de anexos no tiene el formato esperado (ITEM, ID, Título…)');
  const modeloImpar = datos[0], modeloPar = datos[1];   // fila celeste y fila blanca

  const ponerTextos = (tr, valores) => {
    Array.from(tr.getElementsByTagNameNS(W_NS, 'tc')).forEach((tc, i) => {
      const p = tc.getElementsByTagNameNS(W_NS, 'p')[0];
      if(!p) return;
      const v = valores[i] == null ? '' : String(valores[i]);
      if(directRuns(p).some(r => r.getElementsByTagNameNS(W_NS, 't').length)) setParagraphText(p, v);
      else {
        const r = doc.createElementNS(W_NS, 'w:r'), t = doc.createElementNS(W_NS, 'w:t');
        t.textContent = v; r.appendChild(t); p.appendChild(r);
      }
    });
  };
  const filaTicket = (r, i) => {
    const tr = (i % 2 === 0 ? modeloImpar : modeloPar).cloneNode(true);
    ponerTextos(tr, [i + 1, r['ID'], r['Título'], r['Estado'], r['Asignado a - Técnico'], r['Tipo']]);
    return tr;
  };

  // encabezado: se repite en cada página y nunca queda solo al pie (va unido a la primera fila)
  const prepararEncabezado = tr => {
    let trPr = Array.from(tr.childNodes).find(x => x.localName === 'trPr');
    if(!trPr){ trPr = doc.createElementNS(W_NS, 'w:trPr'); tr.insertBefore(trPr, tr.firstChild); }
    if(!Array.from(trPr.childNodes).some(x => x.localName === 'tblHeader'))
      trPr.appendChild(doc.createElementNS(W_NS, 'w:tblHeader'));
    if(!Array.from(trPr.childNodes).some(x => x.localName === 'cantSplit'))
      trPr.insertBefore(doc.createElementNS(W_NS, 'w:cantSplit'), trPr.firstChild);
    Array.from(tr.getElementsByTagNameNS(W_NS, 'p')).forEach(ponerKeepNext);
    return tr;
  };
  const encabezado = () => prepararEncabezado(modeloEnc.cloneNode(true));

  const { sample } = getSeleccionPorPorcentaje(pctElegido);
  const filasAntes = filas.length;
  filas.forEach(tr => tr.parentNode && tr.parentNode.removeChild(tr));

  // tabla 1: tickets al % elegido
  [encabezado(), ...sample.map(filaTicket)].forEach(tr => tabla.appendChild(tr));

  // separación visible entre las dos tablas (~1.5 cm)
  const espacio = doc.createElementNS(W_NS, 'w:p');
  const ePr = doc.createElementNS(W_NS, 'w:pPr');
  const eSp = doc.createElementNS(W_NS, 'w:spacing');
  eSp.setAttribute('w:before', '0'); eSp.setAttribute('w:after', '840');
  ePr.appendChild(doc.createElementNS(W_NS, 'w:keepNext'));
  ePr.appendChild(eSp); espacio.appendChild(ePr);

  // tabla 2: todos los tickets (100%), con el mismo formato de la tabla 1
  const tabla2 = doc.createElementNS(W_NS, 'w:tbl');
  Array.from(tabla.childNodes).filter(n => n.localName === 'tblPr' || n.localName === 'tblGrid')
    .forEach(n => tabla2.appendChild(n.cloneNode(true)));
  [encabezado(), ...excelRows.map(filaTicket)].forEach(tr => tabla2.appendChild(tr));

  tabla.parentNode.insertBefore(espacio, tabla.nextSibling);
  tabla.parentNode.insertBefore(tabla2, espacio.nextSibling);
  return { filasAntes, muestra: sample.length, total: excelRows.length, pct: pctElegido };
}

/* ---------- 5b'''''. Enlace del manual del mes ----------
   Debajo de "Link de acceso al manual – Mes de …" va el título del manual como hipervínculo.
   Se quita el enlace anterior (que puede ser un campo HYPERLINK repartido en varios párrafos)
   y se pone uno nuevo con la dirección y el texto indicados en la página. */

async function actualizarEnlaceManual(zip, doc, url, texto){
  const ps = allParagraphs(doc);
  const iL = ps.findIndex(p => !esIndice(p) && /^Link de acceso al manual/i.test(paraText(p).trim()));
  if(iL < 0) throw new Error('no se encontró "Link de acceso al manual" en el Word');
  let iM = -1;
  for(let k = iL + 1; k <= iL + 4 && k < ps.length; k++) if(paraText(ps[k]).trim()){ iM = k; break; }
  if(iM < 0) throw new Error('no se encontró el enlace debajo de "Link de acceso al manual"');
  const M = ps[iM];
  const antes = paraText(M).trim();

  // formato del texto del enlace actual (negrita, tamaño, fuente) para conservarlo
  // (se toma el tramo con más texto, no el del ícono 🛠)
  const conTexto = Array.from(M.getElementsByTagNameNS(W_NS, 'r')).filter(r => r.getElementsByTagNameNS(W_NS, 't').length);
  const largo = r => Array.from(r.getElementsByTagNameNS(W_NS, 't')).map(t => t.textContent).join('').trim().length;
  const runTexto = conTexto.sort((a, b) => largo(b) - largo(a))[0];
  const rPrModelo = runTexto && Array.from(runTexto.childNodes).find(n => n.localName === 'rPr');

  // 1) quitar el campo HYPERLINK anterior (begin / código / separate / end), que puede seguir en los párrafos siguientes
  let abiertos = 0;
  for(let k = iM; k < Math.min(ps.length, iM + 4); k++){
    const p = ps[k];
    Array.from(p.getElementsByTagNameNS(W_NS, 'r')).forEach(r => {
      const fc = r.getElementsByTagNameNS(W_NS, 'fldChar')[0];
      const tipo = fc && (fc.getAttributeNS(W_NS, 'fldCharType') || fc.getAttribute('w:fldCharType'));
      const esCodigo = r.getElementsByTagNameNS(W_NS, 'instrText').length > 0;
      if(tipo === 'begin') abiertos++;
      if(tipo === 'end') abiertos--;
      if(tipo || esCodigo) r.parentNode.removeChild(r);
    });
    Array.from(p.getElementsByTagNameNS(W_NS, 'hyperlink'))
      .filter(h => !paraText(h).trim()).forEach(h => h.parentNode.removeChild(h));
    if(k > iM && abiertos <= 0) break;
  }

  // 2) dejar el párrafo solo con el nuevo hipervínculo
  Array.from(M.childNodes).forEach(n => { if(n.localName !== 'pPr') M.removeChild(n); });
  const rid = 'rIdClaude_manual';
  const relsPath = 'word/_rels/document.xml.rels';
  const rels = new DOMParser().parseFromString(await zip.file(relsPath).async('string'), 'application/xml');
  Array.from(rels.documentElement.getElementsByTagNameNS(NS_REL, 'Relationship'))
    .filter(r => r.getAttribute('Id') === rid).forEach(r => r.parentNode.removeChild(r));
  const rel = rels.createElementNS(NS_REL, 'Relationship');
  rel.setAttribute('Id', rid);
  rel.setAttribute('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink');
  rel.setAttribute('Target', url);
  rel.setAttribute('TargetMode', 'External');
  rels.documentElement.appendChild(rel);
  zip.file(relsPath, new XMLSerializer().serializeToString(rels));

  const h = doc.createElementNS(W_NS, 'w:hyperlink');
  h.setAttributeNS(NS_R, 'r:id', rid);
  h.setAttributeNS(W_NS, 'w:history', '1');
  const r = doc.createElementNS(W_NS, 'w:r');
  const rPr = rPrModelo ? rPrModelo.cloneNode(true) : doc.createElementNS(W_NS, 'w:rPr');
  // estilo "Hipervínculo" del documento (azul subrayado)
  Array.from(rPr.childNodes).filter(n => n.localName === 'rStyle').forEach(n => rPr.removeChild(n));
  const rs = doc.createElementNS(W_NS, 'w:rStyle'); rs.setAttributeNS(W_NS, 'w:val', 'Hipervnculo');
  rPr.insertBefore(rs, rPr.firstChild);
  const t = doc.createElementNS(W_NS, 'w:t');
  t.setAttributeNS('http://www.w3.org/XML/1998/namespace', 'xml:space', 'preserve');
  t.textContent = texto;
  r.appendChild(rPr); r.appendChild(t); h.appendChild(r); M.appendChild(h);

  return { campo: 'Enlace del manual', enlace: true, antes, despues: `${texto} → ${url}` };
}

/* ---------- 5f. Informe de Conformidad y Eleva Producto ----------
   Plantillas cortas (1 página). Se cambian: números correlativos, N.º de OS, fecha,
   nombre del servicio (TdR), descripción del entregable (TdR) y lista de actividades (TdR). */

// identifica el tipo de Word por su primer título
function tipoWord(xmlDoc){
  const primeras = allParagraphs(xmlDoc).map(p => normTxt(paraText(p))).filter(Boolean).slice(0, 4).join(' | ');
  if(/INFORME TECNICO N/.test(primeras)) return 'conf';
  if(/INFORME DE SERVICIO N/.test(primeras)) return 'eleva';
  return null;
}

// números que ya tiene la plantilla (para sugerir el siguiente)
function leerNumeros(xmlDoc){
  const t = allParagraphs(xmlDoc).map(p => paraText(p)).join('\n');
  const m1 = t.match(/INFORME T[ÉE]CNICO N\.?\s*[º°]\s*(\d+)/i);
  const m2 = t.match(/INFORME DE SERVICIO N\s*[º°]\s*(\d+)/i) || t.match(/Informe de Servicio N\s*[º°]\s*(\d+)/i);
  return { tecnico: m1 ? m1[1] : null, servicio: m2 ? m2[1] : null };
}

// fecha "02 de setiembre de 2026"
function fechaLarga(d){
  return `${String(d.getDate()).padStart(2, '0')} de ${MESES[d.getMonth()]} de ${d.getFullYear()}`;
}

// cambia en un párrafo lo que coincide con la expresión, conservando el formato del resto
function cambiarEnParrafo(doc, buscaParrafo, re, reemplazo){
  const p = allParagraphs(doc).find(p => buscaParrafo.test(paraText(p)));
  if(!p) return null;
  const antes = paraText(p);
  const despues = antes.replace(re, reemplazo);
  if(despues !== antes) reemplazarTextoFino(p, despues);
  return { antes: antes.trim(), despues: despues.trim(), ok: re.test(antes) };
}

// reemplaza una lista con viñetas (desde su primer hasta su último ítem) por los textos dados;
// si entre los ítems había una línea en blanco, se mantiene esa separación
function reemplazarLista(doc, primero, ultimo, textos){
  const padre = primero.parentNode;
  const region = [];
  for(let n = primero; n; n = n.nextSibling){
    if(n.nodeType === 1 && n.localName === 'p') region.push(n);
    if(n === ultimo) break;
  }
  const esItem = p => p.getElementsByTagNameNS(W_NS, 'numPr').length > 0 && paraText(p).trim();
  const modelo = region.find(esItem);
  // línea en blanco entre ítems: solo si el original la tenía entre la mayoría de sus ítems
  const items = region.filter(esItem);
  const huecos = items.slice(0, -1).filter(it => { const sig = region[region.indexOf(it) + 1]; return sig && !esItem(sig) && !paraText(sig).trim(); }).length;
  const separador = huecos >= Math.ceil((items.length - 1) / 2)
    ? region.find(p => !paraText(p).trim() && !esItem(p) && region.indexOf(p) > region.indexOf(modelo)) : null;
  const ancla = ultimo.nextSibling;
  const antes = region.filter(esItem).map(p => paraText(p).trim());
  region.forEach(p => padre.removeChild(p));
  textos.forEach((t, i) => {
    const item = modelo.cloneNode(true);
    setParagraphText(item, t);
    padre.insertBefore(item, ancla);
    if(separador && i < textos.length - 1) padre.insertBefore(separador.cloneNode(true), ancla);
  });
  return antes;
}

function itemsDeLista(doc, desdeTexto){
  // ítems con viñeta que siguen a un párrafo de referencia (o todos si no se indica)
  const ps = allParagraphs(doc);
  let inicio = 0;
  if(desdeTexto){ const i = ps.findIndex(p => desdeTexto.test(paraText(p))); inicio = i >= 0 ? i + 1 : 0; }
  const items = [];
  for(let k = inicio; k < ps.length; k++){
    const p = ps[k], txt = paraText(p).trim();
    const esItem = p.getElementsByTagNameNS(W_NS, 'numPr').length > 0 && txt;
    if(esItem) items.push(p);
    else if(items.length && txt) break;   // terminó la lista
  }
  return items;
}

/* d = { numTecnico, numServicio, fecha (Date), numeroOs, servicio, actividades[] } */
function editarConformidad(doc, d){
  const cambios = [];
  const anio = d.fecha.getFullYear();
  const reg = (campo, r, ok = true) => cambios.push(r ? { campo, antes: r.antes, despues: r.despues, falta: !r.ok && ok } :
                                                       { campo, antes: '(no se encontró)', despues: '', falta: true });
  reg('N.º de Informe Técnico', cambiarEnParrafo(doc, /INFORME T[ÉE]CNICO N/i, /(N\.?\s*[º°]\s*)(\d+)/i,
      (m, a, n) => a + String(d.numTecnico).padStart(n.length, '0')));
  reg('Referencia: Informe de Servicio', cambiarEnParrafo(doc, /Informe de Servicio N/i, /(Informe de Servicio N\s*[º°]\s*)(\d+)(\s*-\s*)(\d{4})/i,
      (m, a, n, g) => a + String(d.numServicio).padStart(n.length, '0') + g + anio));
  if(d.numeroOs) reg('Referencia: Orden de Servicio', cambiarEnParrafo(doc, /Orden de Servicio N/i, /(Orden de Servicio N\s*[º°]\s*)(\d+)/i,
      (m, a) => a + d.numeroOs));
  reg('Fecha', cambiarEnParrafo(doc, /^\s*FECHA/i, /\d{1,2}\s+de\s+[a-záéíóú]+\s+de\s+\d{4}/i, fechaLarga(d.fecha)));
  reg('Nombre del servicio (TdR)', cambiarEnParrafo(doc, /términos de referencia de la orden del/i,
      /(orden del\s+)([\s\S]*?)(,\s*se ha revisado)/i, (m, a, s, b) => a + d.servicio.toUpperCase() + b));
  // actividades: lista dentro de la tabla (columna "Actividades")
  const items = itemsDeLista(doc, /^\s*Situaci[oó]n\s*$/i);
  if(items.length){
    const antes = reemplazarLista(doc, items[0], items[items.length - 1], d.actividades);
    cambios.push({ campo: 'Actividades (TdR)', antes, despues: d.actividades });
  } else cambios.push({ campo: 'Actividades (TdR)', antes: '(no se encontró la lista)', despues: '', falta: true });
  return cambios;
}

/* d = { numServicio, fecha (Date), numeroOs, servicio, entregableTexto, actividades[] } */
function editarEleva(doc, d){
  const cambios = [];
  const anio = d.fecha.getFullYear();
  const reg = (campo, r) => cambios.push(r ? { campo, antes: r.antes, despues: r.despues, falta: !r.ok } :
                                             { campo, antes: '(no se encontró)', despues: '', falta: true });
  reg('N.º de Informe de Servicio', cambiarEnParrafo(doc, /INFORME DE SERVICIO N/i, /(N\s*[º°]\s*)(\d+)(\s*-\s*)(\d{4})/i,
      (m, a, n, g) => a + String(d.numServicio).padStart(n.length, '0') + g + anio));
  if(d.numeroOs) reg('Asunto: Orden de Servicio', cambiarEnParrafo(doc, /Orden de Servicio N/i, /(Orden de Servicio N\s*[º°]\s*)(\d+)/i,
      (m, a) => a + d.numeroOs));
  reg('Fecha', cambiarEnParrafo(doc, /^\s*FECHA/i, /\d{1,2}\s+de\s+[a-záéíóú]+\s+de\s+\d{4}/i, fechaLarga(d.fecha)));
  reg('Nombre del servicio (TdR)', cambiarEnParrafo(doc, /brindar un/i,
      /(brindar un\s*[“"])([\s\S]*?)([”"])/i, (m, a, s, b) => a + d.servicio.toUpperCase() + b));

  // descripción del entregable: ítem(s) entre el párrafo de presentación y "Que comprende…"
  const ps = allParagraphs(doc);
  const iIntro = ps.findIndex(p => /elevo a continuaci[oó]n/i.test(paraText(p)));
  const iComp = ps.findIndex(p => /Que comprende las siguientes actividades/i.test(paraText(p)));
  if(d.entregableTexto && iIntro >= 0 && iComp > iIntro){
    const bloque = ps.slice(iIntro + 1, iComp).filter(p => paraText(p).trim());
    if(bloque.length){
      const antes = bloque.map(p => paraText(p).trim()).join(' ');
      setParagraphText(bloque[0], d.entregableTexto);
      bloque.slice(1).forEach(p => p.parentNode.removeChild(p));   // el texto anterior estaba partido en varios párrafos
      cambios.push({ campo: 'Descripción del entregable (TdR)', antes, despues: d.entregableTexto });
    }
  } else cambios.push({ campo: 'Descripción del entregable (TdR)', antes: '(no se encontró)', despues: '', falta: true });

  const items = itemsDeLista(doc, /Que comprende las siguientes actividades/i);
  if(items.length){
    const antes = reemplazarLista(doc, items[0], items[items.length - 1], d.actividades);
    cambios.push({ campo: 'Actividades (TdR)', antes, despues: d.actividades });
  } else cambios.push({ campo: 'Actividades (TdR)', antes: '(no se encontró la lista)', despues: '', falta: true });
  return cambios;
}

/* ---------- 5c. Nombres "legibles" de las categorías de GLPI ----------
   GLPI guarda la categoría como ruta: "SOFTWARE ESPECIALIZADO > Instalaciones > RENIEC - Certificado digital".
   Para los párrafos se arma un nombre corto: "Instalación de certificado digital".
   >>> Si un nombre no queda bien, agrégalo en CATEGORIA_NOMBRE_FIJO (ruta completa o último tramo). */

// nombre COMPLETO para una categoría (por ruta completa o por su último tramo)
const CATEGORIA_NOMBRE_FIJO = {
  'Habilitar Plugin': 'Habilitación de plugins'
};
// cómo se escribe el último tramo dentro del nombre ("Instalación de" + esto)
const CATEGORIA_OBJETO = {
  'CPU- Small Form Factor': 'CPU',
  'RENIEC - Certificado digital': 'certificado digital',
  'Perufirma': 'PeruFirma',
  'JAVA': 'Java',
  'Ms Office': 'MS Office'
};
// categorías que terminan en un verbo suelto ("Office 365 App > Habilitar")
const CATEGORIA_VERBO = {
  'HABILITAR': 'Habilitación de', 'CONFIGURAR': 'Configuración de', 'CONFIGURACION': 'Configuración de',
  'INSTALAR': 'Instalación de', 'ACTUALIZAR': 'Actualización de', 'ERROR': 'Error en', 'AVERIADO': 'Avería de'
};
// tramo intermedio (acción) -> cómo empieza el nombre
const CATEGORIA_ACCION = {
  'INSTALACIONES': 'Instalación de', 'ACTUALIZACIONES': 'Actualización de',
  'CONFIGURAR': 'Configuración de', 'FORMATEAR Y CONFIGURAR': 'Formateo y configuración de',
  'LICENCIAMIENTO': 'Licenciamiento de', 'ACCESOS': 'Accesos a', 'EVALUACION': 'Evaluación de'
};
// palabras que no se pasan a minúscula
const NOMBRES_PROPIOS = ['Windows','Office','Teams','Java','JAVA','Adobe','Refirma','REFIRMA','PeruFirma','Perufirma','MS','Ms','PDF','SPIJ','CPU','AIO','Forticlient','RENIEC','WLAN_CORP','LAN','VPN'];

function minusculaInicial(s){
  const primera = s.split(/\s+/)[0];
  if(NOMBRES_PROPIOS.includes(primera) || primera === primera.toUpperCase()) return s;
  return s.charAt(0).toLowerCase() + s.slice(1);
}

function nombreCategoria(ruta){
  const r = String(ruta || '').trim();
  if(CATEGORIA_NOMBRE_FIJO[r]) return CATEGORIA_NOMBRE_FIJO[r];
  const tramos = r.split('>').map(s => s.trim()).filter(Boolean);
  if(!tramos.length) return '(sin categoría)';
  const ultimo = tramos[tramos.length - 1];
  if(CATEGORIA_NOMBRE_FIJO[ultimo]) return CATEGORIA_NOMBRE_FIJO[ultimo];
  const verbo = CATEGORIA_VERBO[normTxt(ultimo)];
  if(verbo && tramos.length > 1){
    const prev = tramos[tramos.length - 2];
    return `${verbo} ${CATEGORIA_OBJETO[prev] || minusculaInicial(prev)}`;
  }
  const objeto = CATEGORIA_OBJETO[ultimo] || minusculaInicial(ultimo);
  const accion = tramos.length > 1 ? CATEGORIA_ACCION[normTxt(tramos[tramos.length - 2])] : null;
  const nombre = accion ? `${accion} ${objeto}` : objeto;
  return nombre.charAt(0).toUpperCase() + nombre.slice(1);
}

function unirLista(arr){
  if(arr.length <= 1) return arr.join('');
  return arr.slice(0, -1).join(', ') + ' y ' + arr[arr.length - 1];
}

/* ---------- 5d. Datos del Excel para el informe ----------
   Usa las mismas funciones que las pestañas de la página (definidas en index.html),
   así los números del Word coinciden con los cuadros y gráficos. */

function calcularDatosInforme(pctElegido, cuota){
  const filas = excelRows;
  const c = getTipoCounts(filas);
  let cerrados = 0, resueltos = 0;
  filas.forEach(r => {
    const e = normalizeEstado(r['Estado']);
    if(e.startsWith('cerrado')) cerrados++; else if(e.startsWith('resuelto')) resueltos++;
  });
  const sel = getSeleccionPorPorcentaje(pctElegido);
  const cm = getTipoCounts(sel.sample);
  const pctSolMuestra = cm.total ? Math.round(cm.solicitudes / cm.total * 100) : 0;

  // categorías más frecuentes entre las solicitudes (al 100%), mismo orden que el Gráfico 3
  const catSol = {};
  filas.forEach(r => {
    if(classifyTipo(r['Tipo']) !== 'solicitud') return;
    const k = String(r['Categoría']).trim() || '(sin categoría)';
    catSol[k] = (catSol[k] || 0) + 1;
  });
  const topSol = Object.entries(catSol).sort((a, b) => b[1] - a[1]).map(([k]) => nombreCategoria(k));

  // rango de horas con más tickets (Gráfico 6)
  const horas = {};
  filas.forEach(r => {
    const b = bucketHoras(parseHorasResolucion(r['Estadísticas - Hora de resolución']));
    if(b) horas[b] = (horas[b] || 0) + 1;
  });
  const rangoTop = Object.entries(horas).sort((a, b) => b[1] - a[1])[0];

  return {
    mes: MESES[cuota.inicio.getMonth()], anio: cuota.inicio.getFullYear(),
    total: c.total, incidencias: c.incidencias, solicitudes: c.solicitudes,
    cerrados, resueltos,
    pct: pctElegido, totalGeneral: sel.totalGeneral, siCount: sel.siCount,
    noExcedido: sel.noExcedidoRows.length, muestra: sel.sampleSize,
    muestraInc: cm.incidencias, muestraSol: cm.solicitudes,
    pctSolMuestra, pctIncMuestra: cm.total ? 100 - pctSolMuestra : 0,
    topSol, rangoTop: rangoTop ? rangoTop[0] : '0 a 3 horas'
  };
}

/* ---------- 5e. Textos del bloque A ----------
   "busca" reconoce el párrafo por cómo empieza; "nuevo" es el texto completo.
   >>> Para cambiar la redacción de un párrafo, edita aquí su texto "nuevo". */

function textosBloqueA(d){
  const plural = (n, s, p) => n === 1 ? s : p;
  const masSol = d.solicitudes >= d.incidencias;
  const top = d.topSol.length ? d.topSol : ['(sin solicitudes)'];
  const t3 = top.slice(0, 3).map(minusculaInicial);
  const t4 = top.slice(0, 4);
  const lugares = ['En primer lugar, se realizaron ', '. En segundo lugar, ', ', en tercer lugar, ', ', y último lugar, '];
  let parrafoLugares = '';
  t4.forEach((nombre, i) => {
    const pref = (i === t4.length - 1 && i > 0) ? lugares[3] : lugares[i];
    parrafoLugares += pref + minusculaInicial(nombre);
  });
  parrafoLugares += '.';
  const rango = d.rangoTop === '12 horas a más' ? 'en 12 horas a más' : 'de ' + d.rangoTop;
  const mayorSol = d.muestraSol >= d.muestraInc;

  return [
    { campo: '5. Resumen del mes', busca: /^Durante el mes de/i,
      nuevo: `Durante el mes de ${d.mes} de ${d.anio}, en la Mesa de servicios Institucional (MSI) se gestionó ${d.total} ${plural(d.total,'ticket','tickets')}, siendo ${d.incidencias} incidencias y ${d.solicitudes} peticiones registradas por los usuarios finales de diferentes unidades orgánicas del Ministerio de Educación.` },
    { campo: 'Análisis de la Tabla 1', busca: /^En la tabla N.?\s*1/i,
      nuevo: `En la tabla N°1, se observa que las ${masSol ? 'solicitudes superan en cantidad a las incidencias' : 'incidencias superan en cantidad a las solicitudes'}, lo cual se debe a que los usuarios enfrentan problemas relacionados con ${unirLista(t3)}.` },
    { campo: 'Análisis del Gráfico 3', busca: /^En (la|el) Gr[aá]fic[ao] N.?\s*3/i,
      nuevo: `En el Gráfico N°3, se presenta el top 10 de las solicitudes más frecuentes, destacando las cuatro principales categorías. Las solicitudes más recurrentes son: ${unirLista(top.slice(0, 3))}.` },
    { campo: 'Orden de las principales solicitudes', busca: /^En primer lugar/i, nuevo: parrafoLugares },
    { campo: 'Análisis del Gráfico 4', busca: /^En (la|el) Gr[aá]fic[ao] N.?\s*4/i,
      nuevo: `En el Gráfico N°4 se observa que el porcentaje de ${masSol ? 'solicitudes es superior al de incidencias' : 'incidencias es superior al de solicitudes'}. Esto se debe a que los usuarios solicitaron de manera masiva ${minusculaInicial(top[0])}.` },
    { campo: 'Análisis del Gráfico 5', busca: /^En (la|el) Gr[aá]fic[ao] N.?\s*5/i,
      nuevo: `En el Gráfico N°5 se muestra que la mayoría de las atenciones atendidas están en estado ${d.cerrados >= d.resueltos ? 'cerrado' : 'resuelto'}, lo que indica que, según la gráfica, las solicitudes de los usuarios se gestionan de manera efectiva.` },
    { campo: 'Análisis del Gráfico 6', busca: /^En (la|el) Gr[aá]fic[ao] N.?\s*6/i,
      nuevo: `En el Gráfico N°6 se aprecia que la mayoría de las atenciones son atendidos ${rango}, lo que refleja un esfuerzo por brindar soluciones rápidas a las solicitudes o incidencias, permitiendo a los usuarios retomar sus actividades sin interrupciones prolongadas` +
        (d.siCount ? `, sin embargo, ${d.siCount} ${plural(d.siCount,'ticket superó','tickets superaron')} el tiempo de resolución establecido.` : '.') },
    { campo: 'Manual del mes', busca: /^Link de acceso al manual/i,
      nuevo: `Link de acceso al manual – Mes de ${d.mes} ${d.anio}` },
    { campo: 'Categorización correcta', busca: /^De un total de \d+ tickets, se ha logrado/i,
      nuevo: `De un total de ${d.total} tickets, se ha logrado una correcta categorización del ${d.pct}% de los mismos. Este desempeño refleja una precisión destacable en el proceso de clasificación, lo cual contribuye a la correcta asignación de las categorías correspondientes.` },
    { campo: 'Título: atenciones al %', busca: /^Atenciones gestionadas del/i,
      nuevo: `Atenciones gestionadas del ${d.pct}% según lo solicitado` },
    { campo: 'Atenciones al %', busca: /^De un total de \d+ atenciones gestionadas/i,
      nuevo: `De un total de ${d.totalGeneral} atenciones gestionadas, se excluyen ${d.siCount} marcados con “Sí” en “Tiempo para resolver excedido” quedando ${d.noExcedido}, se ha seleccionado el ${d.pct}%, equivalente a ${d.muestra} tickets, los cuales han sido identificados para un análisis detallado. Estos tickets han sido gestionados conforme a los requisitos establecidos y representan el conjunto designado para su revisión y seguimiento, asegurando el cumplimiento de los estándares de calidad y eficiencia definidos en el proceso. Esta selección permite evaluar el desempeño del servicio y detectar posibles mejoras.` },
    { campo: 'Título: solicitudes al %', busca: /^Gesti[oó]n del \d+% de solicitudes/i,
      nuevo: `Gestión del ${d.pctSolMuestra}% de solicitudes dentro del ${d.pct}% requerido` },
    { campo: 'Solicitudes al %', busca: /^Del total de atenciones gestionados/i,
      nuevo: `Del total de atenciones gestionados según lo requerido, se ha identificado que el ${d.pctSolMuestra}% corresponde a solicitudes que han sido atendidas. Esto indica que la mayoría de los tickets dentro del ${d.pct}% seleccionado han sido procesados y resueltos, cumpliendo con los objetivos establecidos para su gestión y atención.` },
    { campo: 'Título: incidencias al %', busca: /^Gesti[oó]n del \d+% de incidencias/i,
      nuevo: `Gestión del ${d.pctIncMuestra}% de incidencias dentro del ${d.pct}% requerido` },
    { campo: 'Incidencias al %', busca: /^De un total de \d+ tickets, el \d+%/i,
      nuevo: `De un total de ${d.totalGeneral} tickets, el ${d.pct}% fue categorizado correctamente, equivalente a ${d.muestra} tickets. De estos, el ${d.pctIncMuestra}% corresponde a incidencias, es decir, ${d.muestraInc} ${plural(d.muestraInc,'ticket','tickets')}. Esto sugiere que, aunque una parte significativa de los tickets ha sido clasificada de manera adecuada, la proporción de incidencias dentro de esa categorización fue ${mayorSol ? 'menor' : 'mayor'} a la de solicitud, lo que podría indicar una ${mayorSol ? 'menor' : 'mayor'} presencia de problemas críticos.` },
    { campo: 'Conclusión 1', busca: /^Se ha gestionado un total de/i,
      nuevo: `Se ha gestionado un total de ${d.total} tickets, de los cuales ${d.cerrados} han sido cerrados y ${d.resueltos} se encuentran en estado "Resuelto". Estas atenciones se han llevado a cabo con el objetivo de brindar soluciones de manera eficiente y en el menor tiempo posible, garantizando la continuidad operativa.` },
    { campo: 'Anexo: lista al %', busca: /^Total de incidencias y solicitudes al \d+\s*%/i,
      nuevo: `Total de incidencias y solicitudes al ${d.pct}% reportados durante el periodo de atención.` },
    { campo: 'Títulos de tablas y gráficas con %', todos: true,
      busca: /^(Tabla|Gr[aá]fic[ao])\s+\d+\s*:.*\d+\s*%/i,
      transformar: t => t.replace(/\b\d+(\s*%)/, d.pct + '$1') },
    { campo: 'Conclusión 2', busca: /^Del \d+% de tickets seleccionados/i,
      nuevo: `Del ${d.pct}% de tickets seleccionados para una categorización precisa, el ${d.pctSolMuestra}% corresponde a solicitudes y el ${d.pctIncMuestra}% a incidencias, lo que indica una ${mayorSol ? 'menor' : 'mayor'} demanda de incidencias en comparación con la cantidad de solicitudes registradas.` }
  ];
}

// texto "actual" del Word (para mostrar antes de procesar)
function leerDatosWord(doc){
  const ps = allParagraphs(doc);
  const iServ = ps.findIndex(p => normTxt(paraText(p)) === 'SERVICIO:');
  let servicio = '';
  if(iServ >= 0){
    let sig = nextParagraphSibling(ps[iServ]);
    while(sig && !paraText(sig).trim()) sig = nextParagraphSibling(sig);
    servicio = sig ? paraText(sig) : '';
  }
  const ent = ps.map(paraText).find(t => /^\s*ENTREGABLE\s*N\s*[°º]\s*\d+\s*$/i.test(t)) || '';
  return { servicio, entregable: ent.trim() };
}

/* ---------- 6. Interfaz de la página ---------- */

if(typeof document !== 'undefined' && document.getElementById('docsCard')){
  (function(){
    if(window.pdfjsLib){
      // window.PDF_WORKER_SRC solo se define en la copia de prueba dentro de Claude
      pdfjsLib.GlobalWorkerOptions.workerSrc = window.PDF_WORKER_SRC ||
        'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
    }

    // estado de los documentos. "opcional": si no se sube, el Word conserva lo que tiene.
    // "tabs": en qué pestaña se muestra cada casilla (ent = Entregable, conf = Conformidad, eleva = Eleva Producto)
    const docs = {
      tdr:       { label: 'Términos de Referencia', icon: 'bi-file-earmark-pdf', file: null, detalle: '', tabs: ['ent', 'conf', 'eleva'] },
      os:        { label: 'Orden de Servicio',      icon: 'bi-file-earmark-pdf', file: null, detalle: '', tabs: ['ent', 'conf', 'eleva'] },
      word:      { label: 'Informe (Word)',         icon: 'bi-file-earmark-word', file: null, buffer: null, detalle: '', tabs: ['ent'] },
      manual:    { label: 'Manual del mes (Word)',  icon: 'bi-journal-text', file: null, buffer: null, opcional: true, textoOpcional: 'Opcional — si no lo subes, se mantiene el manual actual', tabs: ['ent'] },
      ficha:     { label: IMAGENES_INFORME.ficha.label,     icon: 'bi-file-earmark-image', file: null, img: null, opcional: true, tabs: ['ent'] },
      grafica11: { label: IMAGENES_INFORME.grafica11.label, icon: 'bi-file-earmark-image', file: null, img: null, opcional: true, tabs: ['ent'] },
      conf:      { label: 'Informe de Conformidad (Word)', icon: 'bi-file-earmark-word', file: null, buffer: null, detalle: '', tabs: ['conf'] },
      eleva:     { label: 'Eleva Producto (Word)',  icon: 'bi-file-earmark-word', file: null, buffer: null, detalle: '', tabs: ['eleva'] }
    };
    let tabActiva = 'ent';
    const resultados = { ent: null, conf: null, eleva: null };   // { blob, nombre } del último Word generado en cada pestaña
    let cuotaElegidaManual = false;

    const $ = id => document.getElementById(id);
    const dz = $('dropZoneDocs'), input = $('docsInput'), slots = $('docSlots'), msg = $('docsMsg');
    const controls = $('docsControls'), cuotaSel = $('cuotaSelect'), cuotaHint = $('cuotaHint');
    const preview = $('informePreview');
    // oculta la vista previa y el botón flotante de descarga (cuando cambian los archivos)
    function ocultarPreview(){
      ['informePreview', 'confPreview', 'elevaPreview'].forEach(id => { const el = document.getElementById(id); if(el){ el.style.display = 'none'; el.dataset.listo = ''; } });
      resultados.ent = resultados.conf = resultados.eleva = null;
      ponerFab();
    }
    // botón flotante de descarga: muestra el Word generado en la pestaña activa
    function ponerFab(){
      const r = resultados[tabActiva];
      let fab = document.getElementById('mnFabDescarga');
      if(!r){ if(fab) fab.remove(); return; }
      if(!fab){ fab = document.createElement('div'); fab.id = 'mnFabDescarga'; fab.className = 'mn-fab'; document.body.appendChild(fab); }
      fab.innerHTML = `<span><i class="bi bi-file-earmark-word"></i> ${esc(r.nombre)}</span>
        <button class="btn btn-mn btn-sm" type="button"><i class="bi bi-download"></i> Descargar</button>`;
      fab.querySelector('button').addEventListener('click', () => descargarArchivo(r.blob, r.nombre));
    }

    const esc = s => String(s).replace(/[&<>"]/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
    const hayExcel = () => typeof excelRows !== 'undefined' && excelRows.length > 0;

    function mensaje(tipo, html){
      if(!html){ msg.innerHTML = ''; return; }
      const color = tipo === 'err' ? 'var(--mn-red)' : '#9A6700';
      const icon = tipo === 'err' ? 'bi-x-circle-fill' : 'bi-exclamation-triangle';
      msg.innerHTML = `<div class="small mt-2" style="color:${color};"><i class="bi ${icon}"></i> ${html}</div>`;
    }

    const yaCargados = new Set();   // para animar solo las casillas que acaban de recibir archivo
    function pintarSlots(){
      const recien = Object.keys(docs).filter(k => docs[k].file && !yaCargados.has(k));
      Object.keys(docs).forEach(k => docs[k].file ? yaCargados.add(k) : yaCargados.delete(k));
      slots.innerHTML = Object.entries(docs).filter(([, d]) => d.tabs.includes(tabActiva)).map(([k, d]) => {
        let ok, nombre, detalle = d.detalle || '', quitar = !!d.file;
        ok = !!d.file;
        nombre = ok ? esc(d.file.name) : (d.opcional ? (d.textoOpcional || 'Opcional — si no lo subes, se mantiene la imagen actual') : 'Pendiente');
        const thumb = d.url ? `<img src="${d.url}" alt="" style="max-height:60px; margin-top:6px; border:1px solid var(--mn-line); border-radius:4px;">` : '';
        return `
        <div class="mn-doc-slot ${ok ? 'ok' : ''} ${recien.includes(k) ? 'recien' : ''}">
          <div class="mn-doc-name">
            <i class="bi ${ok ? 'bi-check-circle-fill' : d.icon}" style="color:${ok ? '#1E7B45' : 'var(--mn-muted)'}"></i>
            ${d.label}
          </div>
          <div class="mn-doc-file">${nombre}</div>
          ${detalle ? `<div class="mn-doc-file">${detalle}</div>` : ''}
          ${thumb}
          ${quitar ? `<div><button data-quitar="${k}"><i class="bi bi-x-lg"></i> quitar</button></div>` : ''}
        </div>`;
      }).join('');
      // campos del enlace del manual: aparecen solo cuando hay un manual cargado
      $('manualEnlace').style.display = (docs.manual.file && tabActiva === 'ent') ? 'block' : 'none';
      // estado del Excel: se valida arriba; aquí solo se informa en una línea
      $('docsExcelEstado').innerHTML = hayExcel()
        ? `<span style="color:var(--mn-ok);"><i class="bi bi-check-circle-fill"></i> Excel de tickets validado arriba: ${esc($('fileRowText').textContent)}</span>`
        : `<span style="color:var(--mn-warn);"><i class="bi bi-info-circle"></i> Falta el Excel de tickets: súbelo arriba (o suéltalo aquí); se valida en la parte superior.</span>`;
      slots.querySelectorAll('[data-quitar]').forEach(b => b.addEventListener('click', () => {
        const k = b.dataset.quitar;
        if(docs[k].url) URL.revokeObjectURL(docs[k].url);
        Object.assign(docs[k], { file: null, buffer: null, detalle: '', parsed: null, img: null, url: null });
        if(k === 'os') cuotaElegidaManual = false;
        ocultarPreview();
        pintarSlots(); actualizarControles();
      }));
    }

    // fechas de apertura de los tickets del Excel (si ya se cargó)
    function fechasExcel(){
      if(!hayExcel()) return [];
      return excelRows.map(r => {
        const v = r['Fecha de apertura'];
        if(typeof v === 'number') return new Date(Math.round((v - 25569) * 864e5) + new Date().getTimezoneOffset() * 6e4);
        const m = String(v || '').match(/(\d{2})[-\/](\d{2})[-\/](\d{4})/);
        return m ? new Date(+m[3], +m[2] - 1, +m[1]) : null;
      }).filter(Boolean);
    }

    function actualizarControles(){
      const base = docs.tdr.parsed && docs.os.parsed;
      controls.style.display = (tabActiva === 'ent' && base && docs.word.buffer) ? 'block' : 'none';
      $('confControls').style.display = (tabActiva === 'conf' && base && docs.conf.buffer) ? 'block' : 'none';
      $('elevaControls').style.display = (tabActiva === 'eleva' && base && docs.eleva.buffer) ? 'block' : 'none';
      if(!docs.os.parsed) return;
      const cuotas = docs.os.parsed.cuotas;
      const previo = cuotaSel.value;
      cuotaSel.innerHTML = cuotas.map(c =>
        `<option value="${c.numero}">Entregable ${String(c.numero).padStart(2,'0')} — cuota ${String(c.numero).padStart(3,'0')}: ${c.inicioTxt} al ${c.finTxt} (${MESES[c.inicio.getMonth()]})</option>`
      ).join('');
      const fechas = fechasExcel();
      const sug = sugerirCuota(cuotas, fechas);
      if(cuotaElegidaManual && previo) cuotaSel.value = previo;
      else if(sug) cuotaSel.value = sug.numero;
      // misma lista y mismo valor en la pestaña Eleva Producto
      const cuotaEleva = $('cuotaEleva');
      const prevEleva = cuotaEleva.value;
      cuotaEleva.innerHTML = cuotaSel.innerHTML;
      cuotaEleva.value = prevEleva && cuotaEleva.querySelector(`option[value="${prevEleva}"]`) ? prevEleva : cuotaSel.value;
      cuotaHint.innerHTML = cuotaElegidaManual ? 'Elegido manualmente.'
        : fechas.length ? '<i class="bi bi-lightbulb"></i> Sugerido según las fechas de apertura de los tickets del Excel cargado.'
        : '<i class="bi bi-lightbulb"></i> Sugerido según la fecha de hoy (carga el Excel para sugerirlo según los tickets).';
    }
    // index.html llama a esta función cuando se valida o cambia el Excel
    window.informeActualizarSugerencia = () => { pintarSlots(); actualizarControles(); };
    cuotaSel.addEventListener('change', () => { cuotaElegidaManual = true; cuotaHint.textContent = 'Elegido manualmente.'; });

    // selector de % de esta sección = mismo valor que el de arriba (percentSelect, en index.html)
    const pctInforme = $('pctInforme');
    pctInforme.innerHTML = percentSelect.innerHTML;
    pctInforme.value = percentSelect.value;
    pctInforme.addEventListener('change', () => {
      percentSelect.value = pctInforme.value;
      percentSelect.dispatchEvent(new Event('change'));
    });
    percentSelect.addEventListener('change', () => { pctInforme.value = percentSelect.value; });

    // lee una imagen y la deja lista para el Word (PNG o JPEG; otros formatos se pasan a PNG)
    async function prepararImagen(file){
      const bmp = await createImageBitmap(file);
      const ancho = bmp.width, alto = bmp.height;
      let ext = (file.name.split('.').pop() || '').toLowerCase();
      let blob = file;
      if(ext === 'jpg') ext = 'jpeg';
      if(ext !== 'png' && ext !== 'jpeg'){
        const cv = document.createElement('canvas');
        cv.width = ancho; cv.height = alto;
        cv.getContext('2d').drawImage(bmp, 0, 0);
        blob = await new Promise(r => cv.toBlob(r, 'image/png'));
        ext = 'png';
      }
      bmp.close && bmp.close();
      return { bytes: new Uint8Array(await blob.arrayBuffer()), ext, ancho, alto };
    }

    async function recibirArchivos(fileList){
      mensaje(null);
      ocultarPreview();
      const avisos = [];
      for(const file of Array.from(fileList)){
        const ext = (file.name.split('.').pop() || '').toLowerCase();
        try{
          if(ext === 'xlsx' || ext === 'xls'){
            handleFile(file);   // misma validación del Excel de arriba (index.html)
          } else if(['png','jpg','jpeg','webp','gif','bmp'].includes(ext)){
            const clave = tipoImagenPorNombre(file.name);
            if(!clave){
              avisos.push(`"${esc(file.name)}": no sé qué imagen es. El nombre debe contener "ficha" (Ficha de evaluación) o "grafica 11".`);
              continue;
            }
            const img = await prepararImagen(file);
            if(docs[clave].url) URL.revokeObjectURL(docs[clave].url);
            Object.assign(docs[clave], { file, img, url: URL.createObjectURL(file), detalle: `${img.ancho} × ${img.alto} px` });
          } else if(ext === 'docx'){
            const buffer = await file.arrayBuffer();
            const zip = await JSZip.loadAsync(buffer);
            const f = zip.file('word/document.xml');
            if(!f) throw new Error('no parece un Word válido');
            const xmlDoc = new DOMParser().parseFromString(await f.async('string'), 'application/xml');
            const tw = tipoWord(xmlDoc);
            if(tw === 'conf' || tw === 'eleva'){
              const nums = leerNumeros(xmlDoc);
              docs[tw] = Object.assign(docs[tw], { file, buffer, nums,
                detalle: tw === 'conf' ? `Actual: Informe Técnico N.º ${nums.tecnico || '—'}` : `Actual: Informe de Servicio N.º ${nums.servicio || '—'}` });
              sugerirNumeros();
              continue;
            }
            const actual = leerDatosWord(xmlDoc);
            // ¿informe o manual? El informe tiene "SERVICIO:" / "ENTREGABLE N°" en la portada
            const esInforme = !!(actual.entregable || actual.servicio) && !/manual/i.test(file.name);
            if(esInforme){
              Object.assign(docs.word, { file, buffer,
                detalle: actual.entregable ? `Actualmente: ${esc(actual.entregable)}` : '' });
            } else {
              const n = allParagraphs(xmlDoc).filter(p => paraText(p).trim()).length;
              const imgs = xmlDoc.getElementsByTagNameNS(NS_A, 'blip').length;
              Object.assign(docs.manual, { file, buffer, xmlDoc, detalle: `${n} párrafos · ${imgs} imagen(es)` });
              // título sugerido para el enlace: primera línea con texto del manual
              const primera = allParagraphs(xmlDoc).map(p => paraText(p).trim()).find(Boolean) || file.name.replace(/\.docx$/i, '');
              $('manualTexto').value = '🛠 ' + primera.replace(/^🛠\s*/, '');
            }
          } else if(ext === 'pdf'){
            const data = new Uint8Array(await file.arrayBuffer());
            const items = await pdfToItems(pdfjsLib, data.slice());
            if(!items.length){ avisos.push(`"${esc(file.name)}" no tiene texto seleccionable (¿es un escaneo?).`); continue; }
            const tipo = clasificarPdf(items);
            if(tipo === 'tdr'){
              const parsed = parseTdr(items);
              Object.assign(docs.tdr, { file, parsed,
                detalle: `${parsed.caracteristicas.length + parsed.actividades.length} puntos encontrados ` +
                         `(${parsed.caracteristicas.length} características + ${parsed.actividades.length} actividades)` });
            } else if(tipo === 'os'){
              const parsed = parseOs(items);
              if(!parsed.cuotas.length) throw new Error('no se encontró la tabla de CUOTAS');
              cuotaElegidaManual = false;
              Object.assign(docs.os, { file, parsed,
                detalle: `${parsed.cuotas.length} cuota(s) encontradas` });
            } else {
              avisos.push(`"${esc(file.name)}" no se reconoce como TdR ni como Orden de Servicio.`);
            }
          } else if(ext === 'doc'){
            avisos.push(`"${esc(file.name)}" es un Word antiguo (.doc). Ábrelo en Word y guárdalo como .docx.`);
          } else {
            avisos.push(`"${esc(file.name)}" no es un tipo de archivo permitido.`);
          }
        }catch(e){
          avisos.push(`"${esc(file.name)}": ${esc(e.message)}`);
        }
      }
      if(avisos.length){
        mensaje('warn', avisos.join('<br>'));
        dz.classList.remove('mn-sacude'); void dz.offsetWidth; dz.classList.add('mn-sacude');
      }
      pintarSlots(); actualizarControles();
    }

    dz.addEventListener('click', () => input.click());
    dz.addEventListener('dragover', e => { e.preventDefault(); dz.classList.add('over'); });
    dz.addEventListener('dragleave', () => dz.classList.remove('over'));
    dz.addEventListener('drop', e => { e.preventDefault(); dz.classList.remove('over'); recibirArchivos(e.dataTransfer.files); });
    input.addEventListener('change', e => { recibirArchivos(e.target.files); input.value = ''; });


    const btnProc = $('btnProcesarInforme');
    const htmlBtnProc = btnProc.innerHTML;
    btnProc.addEventListener('click', async () => {
      mensaje(null);
      btnProc.classList.add('cargando');
      btnProc.innerHTML = '<span class="mn-spin"></span> PROCESANDO…';
      await new Promise(r => setTimeout(r, 30));   // deja que se pinte el estado "procesando"
      try{
        if(!hayExcel())
          throw new Error('falta el Excel de tickets: súbelo junto con los demás documentos');

        // 1) primero se generan los cuadros y gráficos con el % elegido (igual que el botón de arriba)
        $('btnGenerate').click();

        const tdr = docs.tdr.parsed;
        const cuota = docs.os.parsed.cuotas.find(c => String(c.numero) === cuotaSel.value);
        const datosExcel = calcularDatosInforme(Number(percentSelect.value), cuota);
        const conPunto = s => /[.;:]$/.test(s) ? s : s + '.';
        // 4. OBJETIVO ESPECÍFICO = todos los puntos de la sección del TdR
        // "CARACTERÍSTICAS Y CONDICIONES DEL SERVICIO": primero sus puntos propios
        // y luego los de su sub-apartado ACTIVIDADES (mismo orden que en el TdR)
        const objetivos = tdr.caracteristicas.concat(tdr.actividades);
        if(!tdr.caracteristicas.length) throw new Error('No se encontraron los puntos de CARACTERÍSTICAS Y CONDICIONES DEL SERVICIO en el TdR');
        if(!tdr.actividades.length) throw new Error('No se encontraron los puntos de ACTIVIDADES en el TdR');

        // 2) se parte siempre del Word original, así se puede procesar varias veces
        const zip = await JSZip.loadAsync(docs.word.buffer);
        const xmlDoc = new DOMParser().parseFromString(await zip.file('word/document.xml').async('string'), 'application/xml');
        const cambios = editarDocumento(xmlDoc, {
          servicio: tdr.denominacion,
          entregable: cuota.numero,
          objetivos: objetivos.map(conPunto),
          textos: textosBloqueA(datosExcel)
        });

        // 3) imágenes subidas (Ficha de evaluación, Gráfica 11)
        for(const clave of Object.keys(IMAGENES_INFORME)){
          if(!docs[clave].img) continue;
          try{
            const c = await reemplazarImagen(zip, xmlDoc, clave, docs[clave].img);
            c.url = docs[clave].url;
            cambios.push(c);
          }catch(e){
            cambios.push({ campo: IMAGENES_INFORME[clave].label, antes: '(sin cambios)', despues: e.message, falta: true });
          }
        }

        // 4) bloque B: tablas y gráficos generados con el Excel, en lugar de las imágenes anteriores
        const bloqueB = await insertarBloqueB(zip, xmlDoc);
        cambios.push({ campo: 'Tablas y gráficos del Excel', bloqueB, antes: '', despues: '' });

        // 4b) enlace del manual del mes (si se subió el manual)
        if(docs.manual.file){
          const url = $('manualUrl').value.trim(), texto = $('manualTexto').value.trim();
          if(!/^https?:\/\/\S+$/i.test(url)){
            cambios.push({ campo: 'Enlace del manual', enlace: true, antes: '(sin cambios)',
              despues: 'falta el enlace del manual o no empieza con http:// o https://', falta: true });
          } else {
            try{ cambios.push(await actualizarEnlaceManual(zip, xmlDoc, url, texto || '🛠 MANUAL DEL MES')); }
            catch(e){ cambios.push({ campo: 'Enlace del manual', enlace: true, antes: '(sin cambios)', despues: e.message, falta: true }); }
          }
        }

        // 5) tabla de anexos con los tickets del Excel (al % elegido y al 100%)
        try{
          const an = llenarAnexos(xmlDoc, Number(percentSelect.value));
          cambios.push({ campo: 'Tabla de anexos', anexos: an, antes: `${an.filasAntes} filas`,
            despues: `${an.muestra} tickets al ${an.pct}% + ${an.total} al 100%` });
        }catch(e){
          cambios.push({ campo: 'Tabla de anexos', anexos: {}, antes: '(sin cambios)', despues: e.message, falta: true });
        }

        // 6) orden: numeración, tablas/gráficos sin partirse e índice actualizado al abrir
        cambios.push(ordenarDocumento(xmlDoc));
        await activarActualizarIndice(zip);

        let xml = new XMLSerializer().serializeToString(xmlDoc);
        if(!xml.startsWith('<?xml')) xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' + xml;
        zip.file('word/document.xml', xml);
        resultados.ent = { blob: await zip.generateAsync({ type: 'blob', compression: 'DEFLATE',
          mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' }),
          nombre: nombreArchivoInforme(cuota.inicio.getMonth(), cuota.inicio.getFullYear()) };
        mostrarVistaPrevia(cambios, datosExcel);
      }catch(e){
        ocultarPreview();
        mensaje('err', 'No se pudo procesar el informe: ' + esc(e.message));
      }finally{
        btnProc.classList.remove('cargando');
        btnProc.innerHTML = htmlBtnProc;
      }
    });

    function celda(v){
      return Array.isArray(v) ? `<ul>${v.map(x => `<li>${esc(x)}</li>`).join('')}</ul>` : esc(v);
    }
    /* Lista de chequeo: resumen de qué quedó listo y qué falta en el Word.
       estado: 'ok' (listo) · 'warn' (se mantiene lo anterior / revisar) · 'pend' (aún no automatizado) */
    function listaChequeo(cambios){
      const textos = cambios.filter(c => !c.imagen && !c.orden && !c.bloqueB && !c.anexos && !c.enlace && !/^Portada|OBJETIVO/.test(c.campo));
      const bB = (cambios.find(c => c.bloqueB) || {}).bloqueB || [];
      const bBok = bB.filter(x => !x.error).length;
      const textosFalta = textos.filter(c => c.falta).length;
      const tiene = re => cambios.some(c => re.test(c.campo) && !c.falta);
      const item = (estado, texto) => ({ estado, texto });
      const items = [
        item(tiene(/SERVICIO/) && tiene(/ENTREGABLE/) ? 'ok' : 'warn', 'Portada: servicio y N° de entregable'),
        item(tiene(/OBJETIVO/) ? 'ok' : 'warn', '4. Objetivo específico (TdR)'),
        item(textosFalta ? 'warn' : 'ok', (textos.length - textosFalta
            ? `Textos con datos del Excel: ${textos.length - textosFalta} actualizados`
            : 'Textos con datos del Excel: ya estaban al día') + (textosFalta ? `, ${textosFalta} no encontrados` : '')),
        ...Object.keys(IMAGENES_INFORME).map(k => {
          const c = cambios.find(x => x.campo === IMAGENES_INFORME[k].label);
          if(!c) return item('warn', `${IMAGENES_INFORME[k].label}: no se subió, queda la imagen anterior`);
          return c.falta ? item('warn', `${IMAGENES_INFORME[k].label}: ${c.despues}`) : item('ok', `${IMAGENES_INFORME[k].label}: imagen nueva`);
        }),
        item('ok', 'Cuadros y gráficos de la página generados'),
        item(cambios.some(c => c.orden) ? 'ok' : 'warn', 'Numeración en orden y tablas/gráficos sin partirse entre páginas'),
        item('ok', 'Índice: se actualiza al abrir el Word (acepta el mensaje)'),
        item(bB.length && bBok === bB.length ? 'ok' : 'warn',
             `Tablas y gráficos del Word con datos del Excel: ${bBok} de ${bB.length}` +
             (bBok < bB.length ? ` (no se pudo: ${bB.filter(x => x.error).map(x => x.titulo.split(':')[0]).join(', ')})` : '')),
        docs.manual.file
          ? item('pend', 'Manual del mes: archivo recibido — su inserción se activa al revisar tu ejemplo')
          : item('warn', 'Manual del mes: no se subió, queda el manual anterior'),
        ...(docs.manual.file ? [(() => { const e = cambios.find(c => c.enlace);
            return e && !e.falta ? item('ok', 'Enlace del manual: hipervínculo actualizado') : item('warn', 'Enlace del manual: ' + (e ? e.despues : 'sin cambios')); })()] : []),
        (() => { const a = cambios.find(c => c.anexos);
                 return a && !a.falta ? item('ok', `Tabla de anexos: ${a.anexos.muestra} tickets al ${a.anexos.pct}% + ${a.anexos.total} al 100%`)
                                      : item('warn', 'Tabla de anexos: ' + (a ? a.despues : 'no se encontró')); })()
      ];
      const icono = { ok: ['bi-check-circle-fill', '#1E7B45'], warn: ['bi-exclamation-triangle-fill', '#9A6700'], pend: ['bi-dash-circle', 'var(--mn-muted)'] };
      return `<ul class="mn-checklist">${items.map((i, n) =>
        `<li style="--i:${n}"><i class="bi ${icono[i.estado][0]}" style="color:${icono[i.estado][1]}"></i> ${esc(i.texto)}</li>`).join('')}</ul>`;
    }

    function mostrarVistaPrevia(cambios, d){
      const faltan = cambios.filter(c => c.falta);
      const fila = c => c.bloqueB
        ? `<tr><td><strong>${esc(c.campo)}</strong></td><td colspan="2"><div class="mn-thumbs">${c.bloqueB.map(x => x.error
            ? `<figure><figcaption>${esc(x.titulo)}</figcaption><div class="small" style="color:#9A6700;">${esc(x.error)}</div></figure>`
            : `<figure><img src="${x.url}" alt=""><figcaption>${esc(x.titulo)}</figcaption></figure>`).join('')}</div></td></tr>`
        : c.imagen
        ? `<tr><td><strong>${esc(c.campo)}</strong></td><td class="antes">${esc(c.antes)}</td>
             <td><img src="${c.url}" alt="" style="max-height:120px; max-width:100%; border:1px solid #A9D4E8;"></td></tr>`
        : `<tr><td><strong>${esc(c.campo)}</strong></td><td class="antes">${celda(c.antes)}</td><td>${celda(c.despues)}</td></tr>`;
      preview.innerHTML = `
        <p class="mn-table-title">LISTA DE CHEQUEO</p>
        ${listaChequeo(cambios)}
        <div class="mn-resultado-acciones">
          <button id="btnDescargarInforme" class="btn btn-mn px-4 py-2">
            <i class="bi bi-download"></i> Descargar ${esc(resultados.ent.nombre)}
          </button>
          ${faltan.length ? `<p class="small mb-0" style="color:#9A6700;"><i class="bi bi-exclamation-triangle"></i>
            No se pudo aplicar: ${faltan.map(f => esc(f.campo)).join(', ')}.
            Revisa que el Word tenga ese texto o título igual que en la plantilla.</p>` : ''}
          <p class="text-muted small mb-0">Datos usados: ${d.total} tickets del Excel, porcentaje <strong>${d.pct}%</strong>
            (muestra de ${d.muestra}), mes <strong>${d.mes} ${d.anio}</strong>. ${cambios.length - faltan.length} cambios.</p>

          <!-- vista previa: oculta por defecto, se abre con el botón -->
          <button id="btnVerCambios" class="btn btn-outline-secondary btn-sm" aria-expanded="false" aria-controls="detalleCambios">
            <i class="bi bi-eye"></i> <span>Mostrar vista previa de los cambios</span>
          </button>
        </div>
        <div id="detalleCambios" class="mn-desplegable"><div>
          <p class="mn-table-title">VISTA PREVIA DE LOS CAMBIOS</p>
          <div class="mn-img-table-wrap">
            <table class="mn-prev-table">
              <thead><tr><th style="width:18%;">Campo</th><th style="width:41%;">Texto actual</th><th>Texto nuevo</th></tr></thead>
              <tbody>${cambios.map(fila).join('')}</tbody>
            </table>
          </div>
        </div></div>`;
      preview.style.display = 'block';
      preview.classList.remove('entrando'); void preview.offsetWidth; preview.classList.add('entrando');

      preview.dataset.listo = '1';
      ponerFab();

      const btnVer = $('btnVerCambios'), detalle = $('detalleCambios');
      btnVer.addEventListener('click', () => {
        const abrir = !detalle.classList.contains('abierto');
        detalle.classList.toggle('abierto', abrir);
        btnVer.setAttribute('aria-expanded', abrir);
        btnVer.querySelector('i').className = 'bi ' + (abrir ? 'bi-chevron-up' : 'bi-eye');
        btnVer.querySelector('span').textContent = (abrir ? 'Ocultar' : 'Mostrar') + ' vista previa de los cambios';
      });
      // descargarArchivo() está definida en index.html (sirve en GitHub y dentro de Claude)
      $('btnDescargarInforme').addEventListener('click', () => descargarArchivo(resultados.ent.blob, resultados.ent.nombre));
    }

    /* ===== Pestañas: Entregable · Informe de Conformidad · Eleva Producto ===== */
    const AYUDA = {
      ent: 'Sube juntos o uno por uno: el <strong>TdR</strong> y la <strong>Orden de Servicio</strong> (PDF), el <strong>informe</strong> y el <strong>manual del mes</strong> (Word) y las imágenes de la <strong>Ficha de evaluación</strong> y la <strong>Gráfica 11</strong> (el nombre debe contener "ficha" o "grafica 11"). El Excel de tickets también se puede soltar aquí.',
      conf: 'Sube el <strong>TdR</strong> y la <strong>Orden de Servicio</strong> (PDF) y el Word del <strong>Informe de Conformidad</strong> anterior. Se actualizan el N.º, la referencia, la fecha, el nombre del servicio y las actividades.',
      eleva: 'Sube el <strong>TdR</strong> y la <strong>Orden de Servicio</strong> (PDF) y el Word del <strong>Eleva Producto</strong> anterior. Se actualizan el N.º, la OS, la fecha, el nombre del servicio, el entregable y las actividades.'
    };
    const segInd = document.querySelector('#docTabs .mn-seg-ind');
    function moverIndicadorDocs(){
      const act = document.querySelector('#docTabs .mn-seg-btn.active');
      if(!act || !act.offsetWidth) return;
      const a = act.getBoundingClientRect(), t = $('docTabs').getBoundingClientRect();
      Object.assign(segInd.style, { left: (a.left - t.left) + 'px', top: (a.top - t.top) + 'px', width: a.width + 'px', height: a.height + 'px' });
    }
    function mostrarTab(t){
      tabActiva = t;
      document.querySelectorAll('#docTabs .mn-seg-btn').forEach(b => {
        const on = b.dataset.dtab === t;
        b.classList.toggle('active', on); b.setAttribute('aria-selected', on);
      });
      // elementos de cada pestaña; las vistas previas solo se muestran si ya hay un resultado
      document.querySelectorAll('#docsCard [data-solo]').forEach(el => {
        if(el.dataset.solo !== t){ el.style.display = 'none'; return; }
        if(/Preview$/.test(el.id)) el.style.display = el.dataset.listo ? 'block' : 'none';
      });
      $('docsAyuda').innerHTML = AYUDA[t];
      mensaje(null);
      pintarSlots(); actualizarControles(); ponerFab();
      requestAnimationFrame(moverIndicadorDocs);
    }
    document.querySelectorAll('#docTabs .mn-seg-btn').forEach(b => b.addEventListener('click', () => mostrarTab(b.dataset.dtab)));
    window.addEventListener('resize', moverIndicadorDocs);

    /* ===== Números correlativos y fechas ===== */
    const numEditado = { tecnico: false, servicio: false };
    const siguiente = (n, ancho) => n == null ? '' : String(parseInt(n, 10) + 1).padStart(ancho || String(n).length, '0');
    function sugerirNumeros(){
      if(!numEditado.tecnico && docs.conf.nums && docs.conf.nums.tecnico) $('numTecnico').value = siguiente(docs.conf.nums.tecnico);
      if(!numEditado.servicio){
        const vistos = [docs.conf.nums && docs.conf.nums.servicio, docs.eleva.nums && docs.eleva.nums.servicio].filter(Boolean).map(Number);
        if(vistos.length){ const v = siguiente(Math.max(...vistos), 3); $('numServConf').value = $('numServEleva').value = v; }
      }
    }
    $('numTecnico').addEventListener('input', () => { numEditado.tecnico = true; });
    // el N.º de Informe de Servicio es el mismo en ambas pestañas
    ['numServConf', 'numServEleva'].forEach(id => $(id).addEventListener('input', e => {
      numEditado.servicio = true;
      $('numServConf').value = $('numServEleva').value = e.target.value;
    }));
    const hoyISO = (() => { const d = new Date(); return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`; })();
    $('fechaConf').value = $('fechaEleva').value = hoyISO;
    const leerFecha = id => { const v = $(id).value; if(!v) return null; const [y, m, d] = v.split('-').map(Number); return new Date(y, m - 1, d); };
    const MesCap = d => MESES[d.getMonth()].charAt(0).toUpperCase() + MESES[d.getMonth()].slice(1);

    /* ===== Generar Informe de Conformidad / Eleva Producto ===== */
    async function generarDocCorto(tipo){
      const btn = $(tipo === 'conf' ? 'btnConf' : 'btnEleva');
      const htmlBtn = btn.innerHTML;
      mensaje(null);
      btn.classList.add('cargando'); btn.innerHTML = '<span class="mn-spin"></span> GENERANDO…';
      await new Promise(r => setTimeout(r, 30));
      try{
        const tdr = docs.tdr.parsed, os = docs.os.parsed;
        const fecha = leerFecha(tipo === 'conf' ? 'fechaConf' : 'fechaEleva');
        if(!fecha) throw new Error('falta la fecha');
        const numServicio = parseInt(($(tipo === 'conf' ? 'numServConf' : 'numServEleva').value || '').replace(/\D/g, ''), 10);
        if(!numServicio) throw new Error('falta el N.º de Informe de Servicio');
        const actividades = tdr.caracteristicas.concat(tdr.actividades).map(t => /[.;:]$/.test(t) ? t : t + '.');
        const zip = await JSZip.loadAsync(docs[tipo].buffer);
        const xmlDoc = new DOMParser().parseFromString(await zip.file('word/document.xml').async('string'), 'application/xml');
        let cambios, nombre;
        if(tipo === 'conf'){
          const numTecnico = parseInt(($('numTecnico').value || '').replace(/\D/g, ''), 10);
          if(!numTecnico) throw new Error('falta el N.º de Informe Técnico');
          cambios = editarConformidad(xmlDoc, { numTecnico, numServicio, fecha, numeroOs: os.numeroOs, servicio: tdr.denominacion, actividades });
          nombre = `Informe de conformidad - ${MesCap(fecha)} ${fecha.getFullYear()}.docx`;
        } else {
          const nCuota = Number($('cuotaEleva').value);
          const ent = (tdr.entregables || [])[nCuota - 1];
          cambios = editarEleva(xmlDoc, { numServicio, fecha, numeroOs: os.numeroOs, servicio: tdr.denominacion,
            entregableTexto: ent ? ent.texto : '', actividades });
          nombre = `ELEVA PRODUCTO - ${MesCap(fecha)} ${fecha.getFullYear()}.docx`;
        }
        let xml = new XMLSerializer().serializeToString(xmlDoc);
        if(!xml.startsWith('<?xml')) xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' + xml;
        zip.file('word/document.xml', xml);
        resultados[tipo] = { blob: await zip.generateAsync({ type: 'blob', compression: 'DEFLATE',
          mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' }), nombre };
        mostrarResultadoCorto(tipo, cambios);
      }catch(e){
        mensaje('err', 'No se pudo generar el documento: ' + esc(e.message));
      }finally{
        btn.classList.remove('cargando'); btn.innerHTML = htmlBtn;
      }
    }
    $('btnConf').addEventListener('click', () => generarDocCorto('conf'));
    $('btnEleva').addEventListener('click', () => generarDocCorto('eleva'));

    function mostrarResultadoCorto(tipo, cambios){
      const cont = $(tipo === 'conf' ? 'confPreview' : 'elevaPreview');
      const r = resultados[tipo];
      const icono = c => c.falta ? '<i class="bi bi-exclamation-triangle-fill" style="color:#9A6700"></i>' : '<i class="bi bi-check-circle-fill" style="color:#1E7B45"></i>';
      const resumen = c => Array.isArray(c.despues) ? `${c.despues.length} actividades` : (c.falta ? 'no se encontró en el Word' : 'actualizado');
      const idDet = 'det_' + tipo, idBtn = 'ver_' + tipo, idDesc = 'desc_' + tipo;
      cont.innerHTML = `
        <p class="mn-table-title">LISTA DE CHEQUEO</p>
        <ul class="mn-checklist">${cambios.map((c, n) => `<li style="--i:${n}">${icono(c)} ${esc(c.campo)}: ${esc(resumen(c))}</li>`).join('')}</ul>
        <div class="mn-resultado-acciones">
          <button id="${idDesc}" class="btn btn-mn px-4 py-2"><i class="bi bi-download"></i> Descargar ${esc(r.nombre)}</button>
          <button id="${idBtn}" class="btn btn-outline-secondary btn-sm" aria-expanded="false"><i class="bi bi-eye"></i> <span>Mostrar vista previa de los cambios</span></button>
        </div>
        <div id="${idDet}" class="mn-desplegable"><div>
          <p class="mn-table-title">VISTA PREVIA DE LOS CAMBIOS</p>
          <div class="mn-img-table-wrap"><table class="mn-prev-table">
            <thead><tr><th style="width:18%;">Campo</th><th style="width:41%;">Texto actual</th><th>Texto nuevo</th></tr></thead>
            <tbody>${cambios.map(c => `<tr><td><strong>${esc(c.campo)}</strong></td><td class="antes">${celda(c.antes)}</td><td>${celda(c.despues)}</td></tr>`).join('')}</tbody>
          </table></div>
        </div></div>`;
      cont.dataset.listo = '1';
      cont.style.display = 'block';
      cont.classList.remove('entrando'); void cont.offsetWidth; cont.classList.add('entrando');
      $(idDesc).addEventListener('click', () => descargarArchivo(r.blob, r.nombre));
      $(idBtn).addEventListener('click', () => {
        const det = $(idDet), abrir = !det.classList.contains('abierto');
        det.classList.toggle('abierto', abrir);
        $(idBtn).setAttribute('aria-expanded', abrir);
        $(idBtn).querySelector('i').className = 'bi ' + (abrir ? 'bi-chevron-up' : 'bi-eye');
        $(idBtn).querySelector('span').textContent = (abrir ? 'Ocultar' : 'Mostrar') + ' vista previa de los cambios';
      });
      ponerFab();
    }

    mostrarTab('ent');
  })();
}

if(typeof module !== 'undefined'){
  module.exports = { pdfToItems, itemsToLines, parseTdr, parseOs, clasificarPdf, sugerirCuota,
    editarDocumento, leerDatosWord, paraText, MESES, diferencias, nombreCategoria, textosBloqueA,
    reemplazarImagen, tipoImagenPorNombre, ordenarDocumento, activarActualizarIndice, insertarBloqueB, FIGURAS_BLOQUE_B, llenarAnexos, actualizarEnlaceManual,
    tipoWord, leerNumeros, editarConformidad, editarEleva, fechaLarga };
}
