/* =====================================================================
   GENERADOR DE TICKETS ("tickets Jason") — MINEDU / OTIC
   ---------------------------------------------------------------------
   Qué hace:
     1. Con los datos del usuario y el N.º de ticket padre, arma un correo
        por cada atención marcada (Windows Update, Office…).
     2. Cada correo se abre listo en Outlook web (cualquier PC, sin instalar
        nada) o en Outlook de escritorio. Opcional: archivo .cmd para enviar
        todos de golpe desde una PC con Outlook clásico.
     3. Genera el "speech" de cierre de cada ticket en tarjetas para copiar
        y pegar en GLPI.
     4. Mi base de datos: guarda usuarios (se autocompletan), atenciones
        propias e historial de bloques (para recuperar los speeches al
        cerrar). Se guarda en el navegador y se puede exportar/importar
        como archivo .json para llevarla a otra PC.

   Dónde cambiar cosas (sin saber programar):
     - CATALOGO: atenciones por grupo. t = título del asunto,
       c = lo que se encontró (va en "CASO ENCONTRADO" después del motivo),
       a = texto de "ACCIONES REALIZADAS", h = true → botón "Habituales".
     - MOTIVOS: textos de "CASO ENCONTRADO" ({nombre} = nombre del usuario).
     - armarAsunto() y armarCuerpo(): formato del correo.
     - MAX_HISTORIAL: cuántos bloques se guardan en el historial.
   ===================================================================== */
(function(){
  'use strict';

  // ---------------- CATÁLOGO DE ATENCIONES ----------------
  const CATALOGO = [
    { grupo: 'Actualizaciones', icono: 'bi-arrow-repeat', items: [
      { t: 'Actualización de Windows Update', c: 'se encontraron actualizaciones de Windows pendientes de instalación', h: true,
        a: 'Se instalaron las actualizaciones pendientes de Windows Update (acumulativas y de seguridad) y se reinició el equipo, lo que corrige fallas del sistema operativo, mejora su estabilidad y lo protege frente a vulnerabilidades.' },
      { t: 'Actualización de Office', c: 'el paquete Microsoft Office no contaba con las últimas actualizaciones', h: true,
        a: 'Se actualizó el paquete Microsoft Office a la última versión disponible, lo que corrige errores, mejora el rendimiento de las aplicaciones de ofimática y refuerza su seguridad.' },
      { t: 'Actualización de Java', c: 'se encontró una versión desactualizada de Java (JRE)', h: true,
        a: 'Se actualizó Java (JRE) a la versión vigente, lo que garantiza la compatibilidad con los sistemas institucionales que lo requieren y corrige vulnerabilidades de seguridad.' },
      { t: 'Actualización de navegadores', c: 'los navegadores Google Chrome y Microsoft Edge se encontraban desactualizados',
        a: 'Se actualizaron los navegadores Google Chrome y Microsoft Edge a su última versión, lo que mejora la velocidad de navegación, la compatibilidad con los sistemas web institucionales y la seguridad.' },
      { t: 'Actualización de Adobe Acrobat Reader', c: 'Adobe Acrobat Reader se encontraba desactualizado',
        a: 'Se actualizó Adobe Acrobat Reader a su última versión, lo que mejora la apertura y visualización de documentos PDF y corrige vulnerabilidades de seguridad.' },
      { t: 'Actualización de controladores', c: 'se encontraron controladores desactualizados en el equipo',
        a: 'Se actualizaron los controladores del equipo, lo que ayuda a que los componentes del equipo puedan funcionar de manera eficiente y estable.' },
      { t: 'Actualización de .NET Framework y Visual C++', c: 'se encontraron componentes .NET Framework y Visual C++ desactualizados',
        a: 'Se actualizaron los componentes .NET Framework y Visual C++ Redistributable, necesarios para el correcto funcionamiento de diversas aplicaciones instaladas en el equipo.' }
    ]},
    { grupo: 'Limpieza y rendimiento', icono: 'bi-speedometer2', items: [
      { t: 'Limpieza de temporales', c: 'se encontró una gran acumulación de archivos temporales en el equipo', h: true,
        a: 'Se eliminaron los archivos temporales del usuario y del sistema (%TEMP%, C:\\Windows\\Temp y Prefetch), lo que libera espacio en disco y mejora el tiempo de respuesta del equipo.' },
      { t: 'Limpieza de caché de navegadores', c: 'los navegadores tenían acumulada una gran cantidad de caché y datos temporales',
        a: 'Se limpió la caché, cookies e historial temporal de los navegadores, lo que resuelve problemas de carga de páginas web y mejora la velocidad de navegación.' },
      { t: 'Liberación de espacio en disco', c: 'el disco contaba con poco espacio libre disponible',
        a: 'Se ejecutó el Liberador de espacio en disco eliminando archivos innecesarios del sistema (restos de actualizaciones, papelera y registros), lo que aumenta el espacio libre y mejora el rendimiento del equipo.' },
      { t: 'Optimización de programas de inicio', c: 'se encontraron varios programas innecesarios configurados para ejecutarse al iniciar Windows',
        a: 'Se deshabilitaron los programas innecesarios que se ejecutaban al iniciar Windows, lo que reduce el tiempo de arranque y el consumo de memoria del equipo.' },
      { t: 'Optimización de unidad de disco', c: 'la unidad de disco requería optimización',
        a: 'Se ejecutó la optimización de la unidad de disco (desfragmentación en disco HDD / TRIM en disco SSD), lo que mejora la velocidad de lectura y escritura de la información.' },
      { t: 'Ajuste de plan de energía y efectos visuales', c: 'el plan de energía y los efectos visuales no estaban configurados para priorizar el rendimiento',
        a: 'Se configuró el plan de energía y se ajustaron los efectos visuales de Windows para priorizar el rendimiento, lo que mejora la fluidez del equipo.' }
    ]},
    { grupo: 'Sistema operativo', icono: 'bi-windows', items: [
      { t: 'Reparación de archivos del sistema (SFC / DISM)', c: 'se identificaron posibles archivos del sistema operativo dañados',
        a: 'Se ejecutaron las herramientas SFC y DISM para verificar y reparar archivos dañados del sistema operativo, lo que mejora la estabilidad de Windows y evita errores inesperados.' },
      { t: 'Verificación del estado del disco (CHKDSK)', c: 'se requería verificar la integridad del sistema de archivos del disco',
        a: 'Se verificó el estado del disco con la herramienta CHKDSK, comprobando la integridad del sistema de archivos y corrigiendo errores lógicos que pueden causar lentitud o pérdida de información.' },
      { t: 'Actualización de políticas de dominio (GPUPDATE)', c: 'el equipo no tenía aplicadas las políticas de dominio vigentes',
        a: 'Se actualizaron las políticas de grupo del dominio (gpupdate /force), lo que asegura que el equipo aplique la configuración institucional vigente.' },
      { t: 'Revisión del Visor de eventos', c: 'se requería revisar los registros de errores del sistema y de las aplicaciones',
        a: 'Se revisó el Visor de eventos de Windows para identificar errores recurrentes del sistema y de las aplicaciones, descartando fallas que afecten el funcionamiento del equipo.' },
      { t: 'Desinstalación de programas innecesarios', c: 'se encontraron programas innecesarios o no autorizados instalados',
        a: 'Se desinstalaron programas innecesarios o no autorizados, lo que libera recursos del equipo y reduce riesgos de seguridad.' }
    ]},
    { grupo: 'Seguridad', icono: 'bi-shield-check', items: [
      { t: 'Actualización de antivirus', c: 'el antivirus institucional no contaba con las firmas de protección actualizadas',
        a: 'Se actualizaron las firmas y el motor del antivirus institucional, lo que asegura la protección del equipo frente a las amenazas más recientes.' },
      { t: 'Escaneo completo de antivirus', c: 'se requería descartar la presencia de software malicioso',
        a: 'Se ejecutó un análisis completo con el antivirus institucional, verificando que el equipo se encuentre libre de software malicioso que afecte su rendimiento o la seguridad de la información.' }
    ]},
    { grupo: 'Office y Outlook', icono: 'bi-envelope-paper', items: [
      { t: 'Optimización de Outlook', c: 'Microsoft Outlook presentaba lentitud al abrir y buscar correos',
        a: 'Se optimizó Microsoft Outlook (reparación y compactación del archivo de datos), lo que mejora la velocidad de apertura, búsqueda y sincronización del correo.' },
      { t: 'Reparación de Office', c: 'las aplicaciones de Microsoft Office presentaban errores o cierres inesperados',
        a: 'Se ejecutó la reparación de Microsoft Office, lo que corrige errores y cierres inesperados en las aplicaciones de ofimática.' }
    ]}
  ];
  const GRUPO_PERSONAL = 'Agregadas por mí';

  // ---------------- MOTIVOS (CASO ENCONTRADO) ----------------
  const MOTIVOS = [
    { id: 'lentitud',  txt: 'Lentitud del equipo',           caso: 'Se detectó lentitud en el equipo del usuario {nombre}.' },
    { id: 'revision',  txt: 'Revisión de equipo',            caso: 'Se solicitó la revisión del equipo del usuario {nombre}.' },
    { id: 'congela',   txt: 'Equipo se congela / no responde', caso: 'Se reportó que el equipo del usuario {nombre} se congela o deja de responder.' },
    { id: 'preventivo',txt: 'Mantenimiento preventivo de software', caso: 'Se realizó mantenimiento preventivo de software en el equipo del usuario {nombre}.' },
    { id: 'otro',      txt: 'Otro (escribir)',               caso: '' }
  ];

  // ---------------- BASE DE DATOS EN ESTE NAVEGADOR ----------------
  const K_AJUSTES = 'mn-gen-ajustes', K_PERSONAL = 'mn-gen-personalizadas', K_USUARIOS = 'mn-gen-usuarios', K_HIST = 'mn-gen-historial', K_SPEECH = 'mn-gen-speeches';
  const MAX_HISTORIAL = 300;
  const AJUSTES_BASE = { cuenta: 'operador_usau07@minedu.gob.pe', destino: 'mesadeservicios@minedu.gob.pe', web: 'https://outlook.office.com' };
  const leerJSON = (k, def) => { try{ return JSON.parse(localStorage.getItem(k)) || def; }catch(e){ return def; } };
  const guardarJSON = (k, v) => { try{ localStorage.setItem(k, JSON.stringify(v)); return true; }catch(e){ return false; } };
  let ajustes = Object.assign({}, AJUSTES_BASE, leerJSON(K_AJUSTES, {}));
  let personales = leerJSON(K_PERSONAL, []);   // [{t, a}]
  let usuarios = leerJSON(K_USUARIOS, {});     // { USUARIO: {nombre, usuario, sede, piso, area, contacto, actualizado} }
  let historial = leerJSON(K_HIST, []);        // [{id, fecha, padre, usuario, nombre, motivo, items:[…]}] (más reciente primero)
  let misSpeeches = leerJSON(K_SPEECH, {});    // { 'motivo|Título': {caso, accion, fecha} } — {nombre} = nombre del usuario

  // ---------------- utilidades ----------------
  const $ = id => document.getElementById(id);
  const esc = s => String(s == null ? '' : s).replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
  const limpio = s => String(s || '').trim().replace(/\s+/g, ' ');
  const aliasCuenta = () => (ajustes.cuenta.split('@')[0] || '').trim();
  const fechaCorta = iso => { const d = new Date(iso); return isNaN(d) ? '' : d.toLocaleString('es-PE', { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' }); };

  function todasLasAtenciones(){
    const lista = [];
    CATALOGO.forEach(g => g.items.forEach(it => lista.push({ ...it, grupo: g.grupo, personal: false })));
    personales.forEach((it, i) => lista.push({ ...it, grupo: GRUPO_PERSONAL, personal: true, idx: i }));
    return lista.map((it, n) => ({ ...it, id: 'at' + n }));
  }

  async function copiar(texto, boton){
    try{ await navigator.clipboard.writeText(texto); }
    catch(e){
      const ta = document.createElement('textarea'); ta.value = texto; document.body.appendChild(ta);
      ta.select(); document.execCommand('copy'); ta.remove();
    }
    if(boton){
      const antes = boton.innerHTML;
      boton.innerHTML = '<i class="bi bi-check2"></i> Copiado';
      boton.classList.add('copiado');
      setTimeout(() => { boton.innerHTML = antes; boton.classList.remove('copiado'); }, 1400);
    }
  }

  // ---------------- datos del formulario ----------------
  const CAMPOS_USUARIO = { nombre: 'genNombre', usuario: 'genUsuario', sede: 'genSede', piso: 'genPiso', area: 'genArea', contacto: 'genContacto' };

  function datos(){
    const motivo = MOTIVOS.find(m => m.id === $('genMotivo').value) || MOTIVOS[0];
    const nombre = limpio($('genNombre').value);
    let caso = motivo.id === 'otro' ? limpio($('genMotivoOtro').value) : motivo.caso;
    if(caso && !/[.!?]$/.test(caso)) caso += '.';
    return {
      motivo: motivo.id,
      motivoTxt: motivo.id === 'otro' ? 'Otro' : motivo.txt,
      padre: limpio($('genPadre').value).replace(/^#/, ''),
      nombre,
      usuario: limpio($('genUsuario').value).toUpperCase(),
      sede: limpio($('genSede').value),
      piso: limpio($('genPiso').value),
      area: limpio($('genArea').value),
      contacto: limpio($('genContacto').value),
      caso
    };
  }

  // ---------------- formato del correo ----------------
  function armarAsunto(titulo, d){
    return `${titulo} - Ticket Padre: ${d.padre} - Sin Notificar - Por Regularizar`;
  }
  function armarCuerpo(d){
    const lineas = [
      `Nombres completos: ${d.nombre}`,
      `Usuario: ${d.usuario}`,
      `Sede: ${d.sede}`,
      `Piso: ${d.piso}`,
      `Área: ${d.area}`
    ];
    if(d.contacto) lineas.push(`Anexo / Celular: ${d.contacto}`);
    lineas.push('', `Regularizar con ${aliasCuenta()}`);
    return lineas.join('\n');
  }

  // ---------------- speech por atención ----------------
  // CASO ENCONTRADO = frase del motivo + "Al revisar, " + lo encontrado (c) de la atención.
  // Si editas el caso o la acción de una tarjeta, se guarda como TU speech para esa atención y ese motivo.
  const minuscula = t => t.charAt(0).toLowerCase() + t.slice(1);
  const hallazgoDe = it => it.c || `se requirió realizar la ${minuscula(it.t)}`;
  const accionDe = it => it.a || `Se realizó la ${minuscula(it.t)} en el equipo, lo que contribuye a su correcto funcionamiento.`;
  const claveSpeech = (motivo, titulo) => `${motivo}|${titulo}`;
  const conNombre = (txt, nombre) => String(txt || '').replace(/\{nombre\}/g, nombre || '(nombre del usuario)');
  const sinNombre = (txt, nombre) => nombre ? String(txt).split(nombre).join('{nombre}') : String(txt);

  function speechOriginal(it, d){
    return { caso: conNombre(`${d.caso} Al revisar, ${hallazgoDe(it)}.`, d.nombre), accion: accionDe(it) };
  }
  function speechPara(it, d){
    const orig = speechOriginal(it, d);
    const mio = misSpeeches[claveSpeech(d.motivo, it.t)];
    return mio ? { caso: conNombre(mio.caso, d.nombre), accion: mio.accion, propio: true, orig } : { ...orig, propio: false, orig };
  }
  function guardarMiSpeech(b, item){
    misSpeeches[claveSpeech(b.motivo || 'lentitud', item.titulo)] = {
      caso: sinNombre(item.caso, b.nombre), accion: item.accion, fecha: new Date().toISOString(), motivoTxt: b.motivoTxt || ''
    };
    guardarJSON(K_SPEECH, misSpeeches);
    item.propio = true;
  }

  // ---------------- usuarios: guardar y autocompletar ----------------
  function guardarUsuario(d){
    if(!d.usuario) return;
    usuarios[d.usuario] = { nombre: d.nombre, usuario: d.usuario, sede: d.sede, piso: d.piso, area: d.area, contacto: d.contacto, actualizado: new Date().toISOString() };
    guardarJSON(K_USUARIOS, usuarios);
    pintarDatalists();
  }
  function cargarUsuario(u){
    Object.entries(CAMPOS_USUARIO).forEach(([k, id]) => { $(id).value = u[k] || ''; });
    const aviso = $('genUsuarioAviso');
    aviso.innerHTML = `<i class="bi bi-person-check-fill"></i> Datos completados desde tu base (puedes corregirlos; se actualizan al generar).`;
    aviso.style.display = '';
  }
  function buscarUsuario(){
    const usu = limpio($('genUsuario').value).toUpperCase();
    const nom = limpio($('genNombre').value).toLowerCase();
    let u = usuarios[usu];
    if(!u && nom) u = Object.values(usuarios).find(x => (x.nombre || '').toLowerCase() === nom);
    if(u) cargarUsuario(u);
  }
  function pintarDatalists(){
    const lista = Object.values(usuarios).sort((a, b) => (a.nombre || '').localeCompare(b.nombre || ''));
    const unicos = k => [...new Set(lista.map(u => u[k]).filter(Boolean))].sort((a, b) => a.localeCompare(b, 'es', { numeric: true }));
    $('genDlUsuarios').innerHTML = lista.map(u => `<option value="${esc(u.usuario)}" label="${esc(u.nombre)}"></option>`).join('');
    $('genDlNombres').innerHTML = lista.map(u => `<option value="${esc(u.nombre)}" label="${esc(u.usuario)}"></option>`).join('');
    $('genDlSedes').innerHTML = unicos('sede').map(v => `<option value="${esc(v)}">`).join('');
    $('genDlPisos').innerHTML = unicos('piso').map(v => `<option value="${esc(v)}">`).join('');
    $('genDlAreas').innerHTML = unicos('area').map(v => `<option value="${esc(v)}">`).join('');
  }

  // ---------------- lista de atenciones ----------------
  function pintarLista(){
    const marcadas = new Set([...document.querySelectorAll('#genLista input.gen-at:checked')].map(c => c.dataset.t));
    const lista = todasLasAtenciones();
    const grupos = CATALOGO.map(g => ({ grupo: g.grupo, icono: g.icono }));
    if(personales.length) grupos.push({ grupo: GRUPO_PERSONAL, icono: 'bi-person-plus' });

    $('genLista').innerHTML = grupos.map(g => {
      const items = lista.filter(it => it.grupo === g.grupo);
      return `<div class="gen-grupo">
        <label class="gen-grupo-tit"><input type="checkbox" class="form-check-input gen-grupo-chk">
          <i class="bi ${g.icono}"></i> ${esc(g.grupo)}</label>
        ${items.map(it => `<label class="gen-item" title="${esc(it.a)}">
            <input type="checkbox" class="form-check-input gen-at" data-id="${it.id}" data-t="${esc(it.t)}" ${marcadas.has(it.t) ? 'checked' : ''}>
            <span>${esc(it.t)}${it.h ? ' <i class="bi bi-star-fill gen-estrella" title="Habitual"></i>' : ''}</span>
            ${it.personal ? `<button type="button" class="gen-borrar" data-idx="${it.idx}" title="Quitar de la lista"><i class="bi bi-x-lg"></i></button>` : ''}
          </label>`).join('')}
      </div>`;
    }).join('');
    actualizarContador();
  }

  function actualizarContador(){
    const todos = [...document.querySelectorAll('#genLista input.gen-at')];
    const n = todos.filter(c => c.checked).length;
    $('genContador').textContent = n === 1 ? '1 ticket seleccionado' : `${n} tickets seleccionados`;
    $('genTodo').checked = n > 0 && n === todos.length;
    $('genTodo').indeterminate = n > 0 && n < todos.length;
    document.querySelectorAll('#genLista .gen-grupo').forEach(g => {
      const cs = [...g.querySelectorAll('input.gen-at')], k = cs.filter(c => c.checked).length;
      const gc = g.querySelector('.gen-grupo-chk');
      gc.checked = k > 0 && k === cs.length; gc.indeterminate = k > 0 && k < cs.length;
    });
  }

  function seleccionadas(){
    const ids = new Set([...document.querySelectorAll('#genLista input.gen-at:checked')].map(c => c.dataset.id));
    return todasLasAtenciones().filter(it => ids.has(it.id));
  }

  // ---------------- generar correos y speeches ----------------
  let bloque = null;   // bloque mostrado: {id, fecha, padre, usuario, nombre, items:[{titulo, asunto, cuerpo, caso, accion, sino, evid, enviado}]}

  function validar(d, sel){
    const faltan = [];
    if(!d.padre) faltan.push('N.º de ticket padre');
    if(!d.nombre) faltan.push('nombres completos');
    if(!d.usuario) faltan.push('usuario');
    if(!d.sede) faltan.push('sede');
    if(!d.piso) faltan.push('piso');
    if(!d.area) faltan.push('área');
    if($('genMotivo').value === 'otro' && !d.caso) faltan.push('caso encontrado');
    if(!sel.length) faltan.push('al menos una atención de la lista');
    return faltan;
  }

  function generar(){
    const d = datos(), sel = seleccionadas();
    const faltan = validar(d, sel);
    if(faltan.length){
      $('genMsg').innerHTML = `<div class="mn-validation err"><i class="bi bi-exclamation-circle-fill"></i> Falta completar: ${esc(faltan.join(', '))}.</div>`;
      return;
    }
    $('genMsg').innerHTML = '';
    const cuerpo = armarCuerpo(d);
    bloque = {
      id: Date.now().toString(36), fecha: new Date().toISOString(),
      padre: d.padre, usuario: d.usuario, nombre: d.nombre, motivo: d.motivo, motivoTxt: d.motivoTxt,
      items: sel.map(it => {
        const sp = speechPara(it, d);
        return { titulo: it.t, asunto: armarAsunto(it.t, d), cuerpo, caso: sp.caso, accion: sp.accion, sino: 'SI', evid: sp.accion,
                 propio: sp.propio, casoOrig: sp.orig.caso, accionOrig: sp.orig.accion, enviado: false };
      })
    };
    guardarUsuario(d);
    historial.unshift(bloque);
    historial = historial.slice(0, MAX_HISTORIAL);
    guardarHistorial();
    mostrarBloque();
  }

  function guardarHistorial(){
    if(!guardarJSON(K_HIST, historial)){
      // si el navegador se queda sin espacio, se borran los bloques más antiguos
      historial = historial.slice(0, Math.max(20, Math.floor(historial.length / 2)));
      guardarJSON(K_HIST, historial);
    }
    pintarBase();
  }

  function mostrarBloque(){
    $('genBloqueInfo').innerHTML = `Ticket padre <strong>${esc(bloque.padre)}</strong> · ${esc(bloque.nombre)} (${esc(bloque.usuario)}) · ${esc(fechaCorta(bloque.fecha))}`;
    pintarCorreos();
    pintarSpeeches();
    $('genResultados').style.display = '';
    $('genResultados').scrollIntoView({ behavior: 'smooth', block: 'start' });
  }

  // enlaces para abrir el correo listo
  function urlWeb(c){
    const base = (ajustes.web || AJUSTES_BASE.web).replace(/\/+$/, '');
    return `${base}/mail/deeplink/compose?to=${encodeURIComponent(ajustes.destino)}&subject=${encodeURIComponent(c.asunto)}&body=${encodeURIComponent(c.cuerpo)}`;
  }
  function urlEscritorio(c){
    return `mailto:${encodeURIComponent(ajustes.destino)}?subject=${encodeURIComponent(c.asunto)}&body=${encodeURIComponent(c.cuerpo.replace(/\n/g, '\r\n'))}`;
  }

  function pintarCorreos(){
    const items = bloque.items;
    $('genCorreos').innerHTML = items.map((c, i) => `
      <div class="gen-correo ${c.enviado ? 'abierto' : ''}" id="genCorreo${i}">
        <div class="gen-correo-num">${c.enviado ? '<i class="bi bi-check-lg"></i>' : i + 1}</div>
        <div class="gen-correo-txt">
          <div class="gen-correo-asunto">${esc(c.asunto)}</div>
          <div class="small text-muted">Para: ${esc(ajustes.destino)}${c.enviado ? ' · <span class="gen-ok">abierto</span>' : ''}</div>
        </div>
        <div class="gen-correo-btns">
          <button type="button" class="btn btn-outline-secondary btn-sm" data-copiar-asunto="${i}" title="Copiar asunto"><i class="bi bi-clipboard"></i></button>
          <a class="btn btn-outline-secondary btn-sm" href="${esc(urlEscritorio(c))}" data-abrir="${i}" title="Abrir en Outlook de escritorio"><i class="bi bi-pc-display"></i> Escritorio</a>
          <a class="btn btn-azul btn-sm" href="${esc(urlWeb(c))}" target="_blank" rel="noopener" data-abrir="${i}"><i class="bi bi-globe2"></i> Outlook web</a>
        </div>
      </div>`).join('') +
      `<details class="gen-cuerpo-prev"><summary>Ver cuerpo del correo (igual para todos)</summary><pre>${esc(items[0].cuerpo)}</pre>
        <button type="button" class="btn btn-outline-secondary btn-sm" id="genCopiarCuerpo"><i class="bi bi-clipboard"></i> Copiar cuerpo</button></details>`;
    const pend = items.filter(c => !c.enviado).length;
    $('genSiguiente').disabled = pend === 0;
    $('genSiguienteTxt').textContent = pend ? `ABRIR SIGUIENTE CORREO (${items.length - pend + 1} de ${items.length})` : 'TODOS LOS CORREOS FUERON ABIERTOS';
  }

  function marcarEnviado(i){
    if(!bloque || !bloque.items[i]) return;
    bloque.items[i].enviado = true;
    guardarHistorial();
    setTimeout(pintarCorreos, 250);
  }

  function abrirSiguiente(){
    const i = bloque.items.findIndex(c => !c.enviado);
    if(i < 0) return;
    window.open(urlWeb(bloque.items[i]), '_blank', 'noopener');
    marcarEnviado(i);
  }

  function pintarSpeeches(){
    $('genSpeeches').innerHTML = bloque.items.map((c, i) => `
      <div class="gen-speech" data-i="${i}">
        <div class="gen-speech-head">
          <span class="gen-correo-num">${i + 1}</span>
          <strong>${esc(c.titulo)}</strong>
        </div>
        <div class="gen-propio" id="genPropio${i}">${etiquetaPropio(c, i)}</div>
        ${campoSpeech(i, 'caso', 'CASO ENCONTRADO', c.caso, 2)}
        ${campoSpeech(i, 'accion', 'ACCIONES REALIZADAS', c.accion, 3)}
        <div class="gen-campo">
          <div class="gen-campo-tit">SE SOLUCIONÓ O ATENDIÓ EL CASO (SI/NO)</div>
          <select class="form-select mn-select gen-sino gen-edita" data-campo="sino" data-i="${i}">
            <option ${c.sino !== 'NO' ? 'selected' : ''}>SI</option><option ${c.sino === 'NO' ? 'selected' : ''}>NO</option>
          </select>
        </div>
        ${campoSpeech(i, 'evid', 'SE ADJUNTA EVIDENCIAS', c.evid, 2)}
        <div class="small text-muted mt-1"><i class="bi bi-image"></i> Si tienes una imagen, después de pegar reemplaza este texto por la imagen en GLPI.</div>
        <button type="button" class="btn btn-azul w-100" data-copiar-todo="${i}"><i class="bi bi-clipboard-check"></i> Copiar speech</button>
      </div>`).join('');
  }

  function etiquetaPropio(c, i){
    const modif = c.casoOrig != null && (c.caso !== c.casoOrig || c.accion !== c.accionOrig);
    if(!c.propio && !modif) return '<span class="small text-muted"><i class="bi bi-stars"></i> Speech sugerido · si lo editas, se guarda como tuyo</span>';
    return `<span class="small gen-ok"><i class="bi bi-bookmark-check-fill"></i> Tu speech guardado para «${esc(bloque.motivoTxt || 'este motivo')}»</span>` +
      (c.casoOrig != null ? ` <button type="button" class="gen-link" data-restaurar="${i}">Restaurar sugerido</button>` : '');
  }

  function campoSpeech(i, campo, titulo, texto, filas){
    return `<div class="gen-campo">
      <div class="gen-campo-tit">${titulo}</div>
      <textarea class="form-control mn-input gen-txt gen-edita" data-campo="${campo}" data-i="${i}" rows="${filas}">${esc(texto)}</textarea>
    </div>`;
  }

  // copia con formato (títulos azules, negrita y subrayado, como la plantilla de GLPI) y también como texto simple
  async function copiarSpeech(i, boton){
    const c = bloque.items[i];
    const secciones = [
      ['CASO ENCONTRADO:', c.caso], ['ACCIONES REALIZADAS:', c.accion],
      ['SE SOLUCIONÓ O ATENDIÓ EL CASO (SI/NO)', c.sino], ['SE ADJUNTA EVIDENCIAS:', c.evid]
    ].map(([t, v]) => [t, String(v || '').trim()]);
    const texto = secciones.map(([t, v]) => `${t}\n${v}`).join('\n\n');
    const html = secciones.map(([t, v]) =>
      `<p><span style="color:#2E74B5;"><strong><u>${esc(t)}</u></strong></span></p><p>${esc(v).replace(/\n/g, '<br>')}</p>`).join('<p>&nbsp;</p>');
    try{
      await navigator.clipboard.write([new ClipboardItem({
        'text/html': new Blob([html], { type: 'text/html' }),
        'text/plain': new Blob([texto], { type: 'text/plain' })
      })]);
      if(boton){ const antes = boton.innerHTML; boton.innerHTML = '<i class="bi bi-check2"></i> Copiado — en GLPI: Ctrl+A y Ctrl+V'; boton.classList.add('copiado'); setTimeout(() => { boton.innerHTML = antes; boton.classList.remove('copiado'); }, 2200); }
    }catch(e){ copiar(texto, boton); }
  }

  const valorCampo = (i, campo) => String(bloque.items[i][campo] || '').trim();

  // ---------------- archivo .cmd (opcional, PC con Outlook clásico) ----------------
  // Windows lo abre como .cmd, este llama a PowerShell, que lee el mismo archivo
  // y usa el Outlook abierto en la PC. No pide ni guarda contraseñas.
  function ps(s){
    return "'" + String(s).replace(/[‘’‚‛]/g, "'").replace(/'/g, "''") + "'";
  }
  function armarScript(){
    const lista = bloque.items.map(c =>
      `  @{ Asunto = ${ps(c.asunto)}; Cuerpo = (@(${c.cuerpo.split('\n').map(ps).join(', ')}) -join "\`r\`n") }`
    ).join(",\r\n");
    return [
      '<# : batch',
      '@echo off',
      'chcp 65001 >nul',
      'powershell -NoProfile -ExecutionPolicy Bypass -Command "$s=[IO.File]::ReadAllText(\'%~f0\',[Text.Encoding]::UTF8); iex $s"',
      'echo.',
      'pause',
      'exit /b',
      '#>',
      '# ===== Envío de tickets a Mesa de Servicios (generado por el Validador de Tickets) =====',
      '[Console]::OutputEncoding = [Text.Encoding]::UTF8',
      `$cuenta = ${ps(ajustes.cuenta)}`,
      `$para   = ${ps(ajustes.destino)}`,
      '$correos = @(',
      lista,
      ')',
      '$n = $correos.Count',
      'Write-Host ""',
      'Write-Host "Se van a generar $n correos desde $cuenta para $para :" -ForegroundColor Cyan',
      '$i = 0; foreach ($c in $correos) { $i++; Write-Host "  $i. $($c.Asunto)" }',
      'Write-Host ""',
      '$op = Read-Host "Escribe S para ENVIAR, B para guardar como BORRADORES, o Enter para cancelar"',
      'if ($op -notmatch "^[SsBb]$") { Write-Host "Cancelado. No se envio nada." -ForegroundColor Yellow; return }',
      'try { $ol = [Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application") } catch { try { $ol = New-Object -ComObject Outlook.Application } catch { $ol = $null } }',
      'if (-not $ol) { Write-Host "No se pudo abrir Outlook de escritorio (clasico). Abrelo e intenta de nuevo, o usa los botones de la pagina." -ForegroundColor Red; return }',
      '$acc = $null',
      'foreach ($a in $ol.Session.Accounts) { if ($a.SmtpAddress -ieq $cuenta) { $acc = $a } }',
      'if (-not $acc) {',
      '  Write-Host "No se encontro la cuenta $cuenta en Outlook. Se usara la cuenta predeterminada de Outlook." -ForegroundColor Yellow',
      '  $sigue = Read-Host "Continuar? (S/N)"',
      '  if ($sigue -notmatch "^[Ss]$") { Write-Host "Cancelado." -ForegroundColor Yellow; return }',
      '}',
      '$i = 0; $ok = 0',
      'foreach ($c in $correos) {',
      '  $i++',
      '  try {',
      '    $m = $ol.CreateItem(0)',
      '    $m.To = $para',
      '    $m.Subject = $c.Asunto',
      '    $m.Body = $c.Cuerpo',
      '    if ($acc) { try { $m.SendUsingAccount = $acc } catch { [void]$m.GetType().InvokeMember("SendUsingAccount","SetProperty",$null,$m,@($acc)) } }',
      '    if ($op -match "^[Ss]$") { $m.Send(); Write-Host "[$i/$n] Enviado:  $($c.Asunto)" -ForegroundColor Green; Start-Sleep -Seconds 2 }',
      '    else { $m.Save(); Write-Host "[$i/$n] Borrador: $($c.Asunto)" -ForegroundColor Green }',
      '    $ok++',
      '  } catch { Write-Host "[$i/$n] ERROR en: $($c.Asunto) -> $($_.Exception.Message)" -ForegroundColor Red }',
      '}',
      'Write-Host ""',
      'Write-Host "Listo: $ok de $n correos procesados." -ForegroundColor Cyan',
      ''
    ].join('\r\n');
  }
  function descargar(nombre, contenido, tipo){
    const blob = new Blob([contenido], { type: tipo });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob); a.download = nombre;
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(a.href), 2000);
  }
  function descargarScript(){
    if(!bloque) return;
    descargar(`enviar_tickets_${(bloque.padre || 'bloque').replace(/[^\w-]/g, '')}.cmd`, armarScript(), 'application/octet-stream');
  }

  // ---------------- atenciones propias ----------------
  function agregarPersonal(){
    const t = limpio($('genNuevoTitulo').value), a = limpio($('genNuevaAccion').value), c = limpio($('genNuevoHallazgo').value).replace(/\.$/, '');
    const aviso = $('genNuevoMsg');
    if(!t){ aviso.textContent = 'Escribe al menos el título.'; return; }
    if(todasLasAtenciones().some(x => x.t.toLowerCase() === t.toLowerCase())){ aviso.textContent = 'Ya existe una atención con ese título.'; return; }
    const nueva = { t };
    if(c) nueva.c = minuscula(c);
    if(a) nueva.a = /[.!?]$/.test(a) ? a : a + '.';
    personales.push(nueva);
    guardarJSON(K_PERSONAL, personales);
    $('genNuevoTitulo').value = ''; $('genNuevaAccion').value = ''; $('genNuevoHallazgo').value = ''; aviso.textContent = '';
    pintarLista(); pintarBase();
    const nuevo = document.querySelector(`#genLista input.gen-at[data-t="${CSS.escape(t)}"]`);
    if(nuevo){ nuevo.checked = true; actualizarContador(); }
  }

  // ---------------- MI BASE DE DATOS (usuarios, historial, respaldo) ----------------
  function pintarBase(){
    const nu = Object.keys(usuarios).length, nh = historial.length, np = personales.length, ns = Object.keys(misSpeeches).length;
    $('genBaseResumen').textContent = `${nu} usuario${nu === 1 ? '' : 's'} · ${nh} bloque${nh === 1 ? '' : 's'} en historial · ${np} atenci${np === 1 ? 'ón propia' : 'ones propias'} · ${ns} speech${ns === 1 ? '' : 'es'} propio${ns === 1 ? '' : 's'}`;

    const sps = Object.entries(misSpeeches).sort((a, b) => a[0].localeCompare(b[0]));
    $('genMisSpeeches').innerHTML = sps.length ? `<table class="gen-tabla"><thead><tr><th>Atención</th><th>Motivo</th><th>Caso encontrado / Acción realizada</th><th></th></tr></thead><tbody>${
      sps.map(([k, v]) => { const [mot, ...tt] = k.split('|'); const m = MOTIVOS.find(x => x.id === mot);
        return `<tr><td><strong>${esc(tt.join('|'))}</strong></td><td>${esc(m ? m.txt : mot)}</td>
          <td class="small">${esc(v.caso)}<br><span class="text-muted">${esc(v.accion)}</span></td>
          <td class="text-end"><button type="button" class="gen-borrar" data-sp-borrar="${esc(k)}" title="Borrar (vuelve al sugerido)"><i class="bi bi-trash3"></i></button></td></tr>`; }).join('')}</tbody></table>`
      : '<p class="small text-muted mb-0">Aún no hay. Cuando edites el caso o la acción de una tarjeta, se guarda aquí.</p>';

    const filtroH = limpio($('genBuscarHist').value).toLowerCase();
    const hs = historial.filter(h => !filtroH || `${h.padre} ${h.usuario} ${h.nombre}`.toLowerCase().includes(filtroH));
    $('genHistorial').innerHTML = hs.length ? `<table class="gen-tabla"><thead><tr><th>Fecha</th><th>Ticket padre</th><th>Usuario</th><th>Tickets</th><th></th></tr></thead><tbody>${
      hs.slice(0, 100).map(h => {
        const env = h.items.filter(x => x.enviado).length;
        return `<tr><td>${esc(fechaCorta(h.fecha))}</td><td><strong>${esc(h.padre)}</strong></td><td>${esc(h.nombre)}<br><span class="small text-muted">${esc(h.usuario)}</span></td>
          <td>${h.items.length} <span class="small text-muted">(${env} abiertos)</span></td>
          <td class="text-end"><button type="button" class="btn btn-outline-secondary btn-sm" data-hist-ver="${esc(h.id)}"><i class="bi bi-chat-square-text"></i> Ver speeches</button>
          <button type="button" class="gen-borrar" data-hist-borrar="${esc(h.id)}" title="Borrar del historial"><i class="bi bi-trash3"></i></button></td></tr>`;
      }).join('')}</tbody></table>` : '<p class="small text-muted mb-0">Aún no hay bloques. Se guardan solos al generar.</p>';

    const filtroU = limpio($('genBuscarUsu').value).toLowerCase();
    const us = Object.values(usuarios).filter(u => !filtroU || `${u.usuario} ${u.nombre} ${u.area} ${u.sede}`.toLowerCase().includes(filtroU))
      .sort((a, b) => (a.nombre || '').localeCompare(b.nombre || ''));
    $('genUsuariosTabla').innerHTML = us.length ? `<table class="gen-tabla"><thead><tr><th>Usuario</th><th>Nombres</th><th>Sede / Piso</th><th>Área</th><th>Anexo</th><th></th></tr></thead><tbody>${
      us.slice(0, 200).map(u => `<tr><td><strong>${esc(u.usuario)}</strong></td><td>${esc(u.nombre)}</td><td>${esc(u.sede)} · ${esc(u.piso)}</td><td>${esc(u.area)}</td><td>${esc(u.contacto)}</td>
        <td class="text-end"><button type="button" class="btn btn-outline-secondary btn-sm" data-usu-usar="${esc(u.usuario)}"><i class="bi bi-box-arrow-in-up"></i> Usar</button>
        <button type="button" class="gen-borrar" data-usu-borrar="${esc(u.usuario)}" title="Borrar usuario"><i class="bi bi-trash3"></i></button></td></tr>`).join('')}</tbody></table>`
      : '<p class="small text-muted mb-0">Aún no hay usuarios. Se guardan solos al generar.</p>';
  }

  function exportar(){
    const data = { app: 'generador-tickets-minedu', version: 2, exportado: new Date().toISOString(), ajustes, personales, usuarios, historial, speeches: misSpeeches };
    const f = new Date().toISOString().slice(0, 10);
    descargar(`respaldo_generador_tickets_${f}.json`, JSON.stringify(data, null, 1), 'application/json');
  }

  function importar(archivo){
    const lector = new FileReader();
    lector.onload = () => {
      let data;
      try{ data = JSON.parse(lector.result); }catch(e){ data = null; }
      const msg = $('genBaseMsg');
      if(!data || data.app !== 'generador-tickets-minedu'){ msg.innerHTML = '<span style="color:var(--mn-red)">El archivo no es un respaldo del Generador de Tickets.</span>'; return; }
      // se combinan: no se pierde nada de lo que ya hay en esta PC
      let nU = 0, nP = 0, nH = 0;
      Object.values(data.usuarios || {}).forEach(u => {
        const a = usuarios[u.usuario];
        if(!a || (u.actualizado || '') > (a.actualizado || '')){ usuarios[u.usuario] = u; nU++; }
      });
      (data.personales || []).forEach(p => { if(!personales.some(x => x.t.toLowerCase() === p.t.toLowerCase())){ personales.push(p); nP++; } });
      let nS = 0;
      Object.entries(data.speeches || {}).forEach(([k, v]) => {
        if(!misSpeeches[k] || (v.fecha || '') > (misSpeeches[k].fecha || '')){ misSpeeches[k] = v; nS++; }
      });
      guardarJSON(K_SPEECH, misSpeeches);
      // atenciones propias: si llegan con speech mejorado (c / a), se actualizan
      (data.personales || []).forEach(p => { const x = personales.find(y => y.t.toLowerCase() === p.t.toLowerCase()); if(x){ if(p.c) x.c = p.c; if(p.a) x.a = p.a; } });
      const ids = new Set(historial.map(h => h.id));
      (data.historial || []).forEach(h => { if(!ids.has(h.id)){ historial.push(h); nH++; } });
      historial.sort((a, b) => (b.fecha || '').localeCompare(a.fecha || ''));
      historial = historial.slice(0, MAX_HISTORIAL);
      if(data.ajustes && !localStorage.getItem(K_AJUSTES)){ ajustes = Object.assign({}, AJUSTES_BASE, data.ajustes); guardarJSON(K_AJUSTES, ajustes); pintarAjustes(); }
      guardarJSON(K_USUARIOS, usuarios); guardarJSON(K_PERSONAL, personales); guardarHistorial();
      pintarLista(); pintarDatalists(); pintarBase();
      msg.innerHTML = `<span class="gen-ok"><i class="bi bi-check-circle-fill"></i> Respaldo importado: ${nU} usuarios, ${nP} atenciones propias, ${nS} speeches y ${nH} bloques nuevos o actualizados.</span>`;
    };
    lector.readAsText(archivo);
  }

  // ---------------- cambio de modo: Entregable / Generador ----------------
  function modo(m){
    const gen = m === 'gen';
    document.body.classList.toggle('modo-gen', gen);
    document.querySelectorAll('#mnModo .mn-seg-btn').forEach(b => b.classList.toggle('active', b.dataset.modo === m));
    moverIndicadorModo();
    if(gen){ history.replaceState(null, '', '#generador'); }
    else if(location.hash === '#generador'){ history.replaceState(null, '', location.pathname + location.search); }
  }
  function moverIndicadorModo(){
    const seg = $('mnModo'), ind = seg.querySelector('.mn-seg-ind'), act = seg.querySelector('.mn-seg-btn.active');
    if(!act || !act.offsetWidth) return;
    const a = act.getBoundingClientRect(), s = seg.getBoundingClientRect();
    Object.assign(ind.style, { left: (a.left - s.left - seg.clientLeft) + 'px', top: (a.top - s.top - seg.clientTop) + 'px', width: a.width + 'px', height: a.height + 'px' });
  }

  function pintarAjustes(){
    $('genCuenta').value = ajustes.cuenta;
    $('genDestino').value = ajustes.destino;
    $('genWeb').value = ajustes.web;
    $('genAlias').textContent = aliasCuenta();
  }

  // ---------------- inicio ----------------
  function iniciar(){
    $('genMotivo').innerHTML = MOTIVOS.map(m => `<option value="${m.id}">${esc(m.txt)}</option>`).join('');
    pintarAjustes();
    pintarLista();
    pintarDatalists();
    pintarBase();

    document.querySelectorAll('#mnModo .mn-seg-btn').forEach(b => b.addEventListener('click', () => modo(b.dataset.modo)));
    window.addEventListener('resize', moverIndicadorModo);
    modo(location.hash === '#generador' ? 'gen' : 'ent');

    $('genMotivo').addEventListener('change', () => { $('genMotivoOtroWrap').style.display = $('genMotivo').value === 'otro' ? '' : 'none'; });
    $('genUsuario').addEventListener('input', e => { const p = e.target.selectionStart; e.target.value = e.target.value.toUpperCase(); e.target.setSelectionRange(p, p); });
    $('genUsuario').addEventListener('change', buscarUsuario);
    $('genNombre').addEventListener('change', () => { if(!limpio($('genUsuario').value)) buscarUsuario(); });

    ['genCuenta', 'genDestino', 'genWeb'].forEach(id => $(id).addEventListener('change', () => {
      ajustes = {
        cuenta: limpio($('genCuenta').value) || AJUSTES_BASE.cuenta,
        destino: limpio($('genDestino').value) || AJUSTES_BASE.destino,
        web: limpio($('genWeb').value) || AJUSTES_BASE.web
      };
      guardarJSON(K_AJUSTES, ajustes);
      pintarAjustes();
    }));

    // casillas de atenciones
    $('genLista').addEventListener('change', e => {
      if(e.target.classList.contains('gen-grupo-chk')){
        e.target.closest('.gen-grupo').querySelectorAll('input.gen-at').forEach(c => c.checked = e.target.checked);
      }
      actualizarContador();
    });
    $('genLista').addEventListener('click', e => {
      const b = e.target.closest('.gen-borrar'); if(!b) return;
      e.preventDefault();
      const it = personales[+b.dataset.idx];
      if(it && confirm(`¿Quitar "${it.t}" de la lista?`)){ personales.splice(+b.dataset.idx, 1); guardarJSON(K_PERSONAL, personales); pintarLista(); pintarBase(); }
    });
    $('genTodo').addEventListener('change', e => { document.querySelectorAll('#genLista input.gen-at').forEach(c => c.checked = e.target.checked); actualizarContador(); });
    $('genHabituales').addEventListener('click', () => {
      const hab = new Set(todasLasAtenciones().filter(x => x.h).map(x => x.id));
      document.querySelectorAll('#genLista input.gen-at').forEach(c => c.checked = hab.has(c.dataset.id));
      actualizarContador();
    });
    $('genAgregar').addEventListener('click', agregarPersonal);

    $('genBtnGenerar').addEventListener('click', generar);
    $('genBtnScript').addEventListener('click', descargarScript);
    $('genSiguiente').addEventListener('click', abrirSiguiente);
    $('genBtnLimpiar').addEventListener('click', () => {
      if(!confirm('¿Limpiar el formulario y empezar un nuevo bloque? (lo generado queda en el historial)')) return;
      ['genPadre','genNombre','genUsuario','genSede','genPiso','genArea','genContacto','genMotivoOtro'].forEach(id => $(id).value = '');
      $('genUsuarioAviso').style.display = 'none';
      document.querySelectorAll('#genLista input.gen-at').forEach(c => c.checked = false);
      actualizarContador();
      bloque = null; $('genResultados').style.display = 'none'; $('genMsg').innerHTML = '';
      window.scrollTo({ top: 0, behavior: 'smooth' });
    });

    // copiar / abrir / editar speeches (lo editado se guarda en el historial)
    $('genResultados').addEventListener('click', e => {
      const b = e.target.closest('button, a'); if(!b || !bloque) return;
      if(b.dataset.copiarAsunto != null) copiar(bloque.items[+b.dataset.copiarAsunto].asunto, b);
      else if(b.id === 'genCopiarCuerpo') copiar(bloque.items[0].cuerpo, b);
      else if(b.dataset.copiarCampo) copiar(valorCampo(+b.dataset.i, b.dataset.copiarCampo), b);
      else if(b.dataset.copiarTodo != null) copiarSpeech(+b.dataset.copiarTodo, b);
      else if(b.dataset.restaurar != null){
        const i = +b.dataset.restaurar, c = bloque.items[i];
        if(!confirm(`¿Volver al speech sugerido para "${c.titulo}"? Se borra tu versión guardada.`)) return;
        delete misSpeeches[claveSpeech(bloque.motivo || 'lentitud', c.titulo)];
        guardarJSON(K_SPEECH, misSpeeches);
        if(c.evid === c.accion) c.evid = c.accionOrig;
        c.caso = c.casoOrig; c.accion = c.accionOrig; c.propio = false;
        guardarHistorial(); pintarSpeeches();
      }
      else if(b.dataset.abrir != null) marcarEnviado(+b.dataset.abrir);
    });
    let tGuardar;
    $('genResultados').addEventListener('input', e => {
      const el = e.target.closest('.gen-edita'); if(!el || !bloque) return;
      const i = +el.dataset.i, c = bloque.items[i], campo = el.dataset.campo;
      if(campo === 'accion' && c.evid === c.accion){            // evidencias sigue a la acción mientras no la cambies
        c.evid = el.value;
        const ev = document.querySelector(`.gen-txt[data-campo="evid"][data-i="${i}"]`); if(ev) ev.value = el.value;
      }
      c[campo] = el.value;
      clearTimeout(tGuardar);
      tGuardar = setTimeout(() => {
        if(campo === 'caso' || campo === 'accion'){ guardarMiSpeech(bloque, c); $('genPropio' + i).innerHTML = etiquetaPropio(c, i); }
        guardarHistorial();
      }, 700);
    });

    // base de datos
    $('genBuscarHist').addEventListener('input', pintarBase);
    $('genBuscarUsu').addEventListener('input', pintarBase);
    $('genBase').addEventListener('click', e => {
      const b = e.target.closest('button'); if(!b) return;
      if(b.dataset.histVer){
        const h = historial.find(x => x.id === b.dataset.histVer);
        if(h){ bloque = h; mostrarBloque(); }
      } else if(b.dataset.histBorrar){
        if(confirm('¿Borrar este bloque del historial?')){
          historial = historial.filter(x => x.id !== b.dataset.histBorrar);
          if(bloque && bloque.id === b.dataset.histBorrar){ bloque = null; $('genResultados').style.display = 'none'; }
          guardarHistorial();
        }
      } else if(b.dataset.spBorrar){
        if(confirm('¿Borrar este speech? La próxima vez se usará el sugerido.')){ delete misSpeeches[b.dataset.spBorrar]; guardarJSON(K_SPEECH, misSpeeches); pintarBase(); }
      } else if(b.dataset.usuUsar){
        const u = usuarios[b.dataset.usuUsar];
        if(u){ cargarUsuario(u); window.scrollTo({ top: 0, behavior: 'smooth' }); }
      } else if(b.dataset.usuBorrar){
        if(confirm(`¿Borrar al usuario ${b.dataset.usuBorrar} de tu base?`)){
          delete usuarios[b.dataset.usuBorrar]; guardarJSON(K_USUARIOS, usuarios); pintarDatalists(); pintarBase();
        }
      }
    });
    $('genExportar').addEventListener('click', exportar);
    $('genImportar').addEventListener('click', () => $('genImportarInput').click());
    $('genImportarInput').addEventListener('change', e => { if(e.target.files[0]) importar(e.target.files[0]); e.target.value = ''; });
  }

  if(document.readyState === 'loading') document.addEventListener('DOMContentLoaded', iniciar);
  else iniciar();
})();
