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

  // ---------------- CATÁLOGO DE TICKETS HIJO ----------------
  // t = título (asunto) · c = qué se encontró (CASO ENCONTRADO, después de "Al revisar,")
  // a = ACCIONES REALIZADAS: acción técnica + beneficio + resultado · h = habitual (⭐)
  const CATALOGO = [
    { grupo: 'Actualizaciones', icono: 'bi-arrow-repeat', items: [
      { t: 'Actualización de Windows Update', h: true,
        c: 'se encontraron actualizaciones acumulativas y de seguridad de Windows pendientes de instalación',
        a: 'Se instalaron las actualizaciones acumulativas y de seguridad pendientes de Windows Update y se reinició el equipo para completar su aplicación. Esto corrige fallas del sistema operativo, mejora su estabilidad y reduce la exposición a vulnerabilidades conocidas. El equipo quedó actualizado y operativo, garantizando la continuidad de las labores.' },
      { t: 'Actualización de Office', h: true,
        c: 'el paquete Microsoft Office no contaba con las últimas actualizaciones publicadas',
        a: 'Se actualizó el paquete Microsoft Office a la última versión disponible mediante su servicio de actualización. Esto corrige errores conocidos, mejora el rendimiento de las aplicaciones de ofimática y refuerza su seguridad. Las aplicaciones quedaron actualizadas y funcionando correctamente, garantizando la continuidad de las labores.' },
      { t: 'Actualización de Java', h: true,
        c: 'se encontró una versión desactualizada de Java (JRE)',
        a: 'Se actualizó Java (JRE) a la versión vigente, desinstalando previamente las versiones anteriores. Esto asegura la compatibilidad con los sistemas y aplicativos institucionales que lo requieren y corrige vulnerabilidades de seguridad. Java quedó actualizado y operativo en el equipo, garantizando la continuidad de las labores.' },
      { t: 'Actualización de navegadores',
        c: 'los navegadores Google Chrome y Microsoft Edge se encontraban desactualizados',
        a: 'Se actualizaron los navegadores Google Chrome y Microsoft Edge a su última versión estable. Esto mejora la velocidad de navegación, la compatibilidad con los sistemas web institucionales y la protección frente a sitios maliciosos. Los navegadores quedaron actualizados y operativos, garantizando la continuidad de las labores.' },
      { t: 'Actualización de Adobe Acrobat Reader',
        c: 'Adobe Acrobat Reader se encontraba desactualizado',
        a: 'Se actualizó Adobe Acrobat Reader a su última versión disponible. Esto mejora la apertura, visualización y firma de documentos PDF y corrige vulnerabilidades de seguridad. La aplicación quedó actualizada y operativa, garantizando la continuidad de las labores.' },
      { t: 'Actualización de controladores',
        c: 'se encontraron controladores desactualizados en el equipo',
        a: 'Se actualizaron los controladores del equipo con las versiones vigentes publicadas por el fabricante. Esto permite que los componentes del equipo funcionen de manera eficiente y estable y corrige fallas de compatibilidad con el sistema operativo. El equipo quedó con sus controladores actualizados y operativo, garantizando la continuidad de las labores.' },
      { t: 'Actualización de .NET Framework y Visual C++',
        c: 'se encontraron componentes .NET Framework y Visual C++ Redistributable desactualizados',
        a: 'Se actualizaron los componentes .NET Framework y Visual C++ Redistributable a sus versiones vigentes. Estos componentes son necesarios para el correcto funcionamiento de diversas aplicaciones y su actualización evita errores de ejecución. Las aplicaciones dependientes quedaron funcionando correctamente, garantizando la continuidad de las labores.' },
      { t: 'Actualización de BIOS/UEFI',
        c: 'el BIOS/UEFI del equipo se encontraba desactualizado',
        a: 'Se actualizó el BIOS/UEFI a la última versión publicada por el fabricante, verificando previamente la compatibilidad con el modelo del equipo y la estabilidad de la alimentación eléctrica. La actualización corrige fallas de firmware, mejora la compatibilidad con el sistema operativo y cierra vulnerabilidades de seguridad a nivel de arranque. El equipo quedó operativo, iniciando correctamente y alineado a las políticas de seguridad de la información de la entidad.' }
    ]},
    { grupo: 'Limpieza y rendimiento', icono: 'bi-speedometer2', items: [
      { t: 'Limpieza de temporales', h: true,
        c: 'se encontró una gran acumulación de archivos temporales en el equipo',
        a: 'Se eliminaron los archivos temporales del usuario y del sistema (%TEMP%, C:\\Windows\\Temp y Prefetch). Esto libera espacio en disco y reduce la carga de lectura del sistema, mejorando el tiempo de respuesta del equipo. El equipo quedó con mayor espacio disponible y un mejor rendimiento, garantizando la continuidad de las labores.' },
      { t: 'Limpieza de caché de navegadores',
        c: 'los navegadores tenían acumulada una gran cantidad de caché y datos temporales',
        a: 'Se limpió la caché, las cookies y los datos temporales de los navegadores. Esto resuelve problemas de carga de páginas web y de acceso a los sistemas institucionales en línea. Los navegadores quedaron funcionando con normalidad, garantizando la continuidad de las labores.' },
      { t: 'Liberación de espacio en disco',
        c: 'el disco contaba con poco espacio libre disponible',
        a: 'Se ejecutó el Liberador de espacio en disco, eliminando archivos innecesarios del sistema como restos de actualizaciones, papelera y registros. Esto aumenta el espacio libre disponible y evita la degradación del rendimiento por falta de espacio. El equipo quedó con espacio suficiente para su operación, garantizando la continuidad de las labores.' },
      { t: 'Optimización de programas de inicio',
        c: 'se encontraron varios programas innecesarios configurados para ejecutarse al iniciar Windows',
        a: 'Se deshabilitaron los programas innecesarios que se ejecutaban al iniciar Windows, conservando los requeridos por la entidad. Esto reduce el tiempo de arranque y el consumo de memoria del equipo. El equipo quedó iniciando con mayor rapidez, garantizando la continuidad de las labores.' },
      { t: 'Optimización de unidad de disco',
        c: 'la unidad de disco requería optimización',
        a: 'Se ejecutó la optimización de la unidad de disco (desfragmentación en disco HDD / TRIM en disco SSD). Esto mejora la velocidad de lectura y escritura de la información. La unidad quedó optimizada y el equipo con mejor tiempo de respuesta, garantizando la continuidad de las labores.' },
      { t: 'Ajuste de plan de energía y efectos visuales',
        c: 'el plan de energía y los efectos visuales no estaban configurados para priorizar el rendimiento',
        a: 'Se configuró el plan de energía y se ajustaron los efectos visuales de Windows para priorizar el rendimiento. Esto reduce el consumo innecesario de recursos del procesador y la memoria. El equipo quedó funcionando con mayor fluidez, garantizando la continuidad de las labores.' }
    ]},
    { grupo: 'Sistema operativo', icono: 'bi-windows', items: [
      { t: 'Reparación de archivos del sistema (SFC / DISM)',
        c: 'se identificaron posibles archivos del sistema operativo dañados',
        a: 'Se ejecutaron las herramientas SFC y DISM para verificar y reparar los archivos dañados del sistema operativo. Esto restablece la integridad de Windows y evita errores y cierres inesperados. El sistema operativo quedó íntegro y estable, garantizando la continuidad de las labores.' },
      { t: 'Verificación del estado del disco (CHKDSK)',
        c: 'se requería verificar la integridad del sistema de archivos del disco',
        a: 'Se verificó el estado del disco con la herramienta CHKDSK, comprobando la integridad del sistema de archivos y corrigiendo errores lógicos. Esto previene la lentitud y la pérdida de información por sectores o índices dañados. El disco quedó verificado y operativo, garantizando la continuidad de las labores.' },
      { t: 'Actualización de políticas de dominio (GPUPDATE)',
        c: 'el equipo no tenía aplicadas las políticas de dominio vigentes',
        a: 'Se actualizaron las políticas de grupo del dominio mediante el comando gpupdate /force. Esto asegura que el equipo aplique la configuración y las restricciones de seguridad institucionales vigentes. El equipo quedó alineado a las políticas de la entidad, garantizando la continuidad de las labores.' },
      { t: 'Revisión del Visor de eventos',
        c: 'se requería revisar los registros de errores del sistema y de las aplicaciones',
        a: 'Se revisó el Visor de eventos de Windows para identificar errores recurrentes del sistema y de las aplicaciones. Esto permite descartar o atender fallas que afecten el funcionamiento del equipo. No se identificaron fallas críticas pendientes y el equipo quedó operativo, garantizando la continuidad de las labores.' },
      { t: 'Desinstalación de programas innecesarios',
        c: 'se encontraron programas innecesarios o no autorizados instalados',
        a: 'Se desinstalaron los programas innecesarios o no autorizados por la entidad. Esto libera recursos del equipo y reduce los riesgos de seguridad de la información. El equipo quedó únicamente con el software autorizado, garantizando la continuidad de las labores.' }
    ]},
    { grupo: 'Seguridad', icono: 'bi-shield-check', items: [
      { t: 'Actualización de antivirus',
        c: 'el antivirus institucional no contaba con las firmas de protección actualizadas',
        a: 'Se actualizaron las firmas y el motor del antivirus institucional. Esto asegura la protección del equipo frente a las amenazas más recientes. El antivirus quedó actualizado y activo, garantizando la continuidad de las labores de forma segura.' },
      { t: 'Escaneo completo de antivirus',
        c: 'se requería descartar la presencia de software malicioso',
        a: 'Se ejecutó un análisis completo con el antivirus institucional. Esto permite verificar que el equipo se encuentre libre de software malicioso que afecte su rendimiento o la seguridad de la información. El análisis concluyó sin amenazas activas y el equipo quedó operativo, garantizando la continuidad de las labores.' }
    ]},
    { grupo: 'Office y Outlook', icono: 'bi-envelope-paper', items: [
      { t: 'Optimización de Outlook',
        c: 'Microsoft Outlook presentaba lentitud al abrir y buscar correos',
        a: 'Se optimizó Microsoft Outlook mediante la reparación y compactación de su archivo de datos. Esto mejora la velocidad de apertura, búsqueda y sincronización del correo institucional. Outlook quedó funcionando con normalidad, garantizando la continuidad de las labores.' },
      { t: 'Reparación de Office',
        c: 'las aplicaciones de Microsoft Office presentaban errores o cierres inesperados',
        a: 'Se ejecutó la reparación de Microsoft Office desde el panel de aplicaciones. Esto corrige archivos dañados de la instalación y elimina los errores y cierres inesperados. Las aplicaciones de ofimática quedaron funcionando correctamente, garantizando la continuidad de las labores.' }
    ]},
    { grupo: 'Firma digital', icono: 'bi-pen', items: [
      { t: 'Instalación de ReFirma',
        c: 'el equipo no contaba con el software ReFirma, necesario para la firma digital con certificado',
        a: 'Se instaló y configuró el software ReFirma, verificando su integración con el certificado digital. Esto permite firmar digitalmente documentos con validez legal conforme a la normativa de firma digital del Estado. La aplicación quedó operativa, garantizando la continuidad de las labores.' },
      { t: 'Instalación de Perú Firma',
        c: 'el equipo no contaba con la aplicación Perú Firma',
        a: 'Se instaló y configuró la aplicación Perú Firma, verificando el reconocimiento del certificado digital. Esto habilita la firma digital de documentos en los sistemas del Estado que la requieren. La aplicación quedó operativa, garantizando la continuidad de las labores.' },
      { t: 'Cambio de cargo en Perú Firma',
        c: 'el cargo registrado en Perú Firma no correspondía al cargo actual',
        a: 'Se actualizó el cargo en la configuración de Perú Firma de acuerdo con la designación vigente. Esto asegura que los documentos firmados digitalmente muestren el cargo correcto. La firma quedó configurada correctamente, garantizando la continuidad de las labores.' }
    ]},
    { grupo: 'Impresión y escaneo', icono: 'bi-printer', items: [
      { t: 'Configuración de impresora',
        c: 'el equipo no tenía configurada la impresora asignada',
        a: 'Se instaló el controlador y se configuró la impresora asignada en el equipo, realizando una impresión de prueba. Esto permite la impresión de documentos desde el puesto de trabajo. La impresora quedó operativa, garantizando la continuidad de las labores.' },
      { t: 'Configuración de escáner',
        c: 'el equipo no tenía configurada la función de escaneo',
        a: 'Se configuró la función de escaneo del equipo multifuncional, incluyendo la carpeta o el destino de digitalización, y se realizó un escaneo de prueba. Esto permite digitalizar documentos desde el puesto de trabajo. El escáner quedó operativo, garantizando la continuidad de las labores.' }
    ]}
  ];
  const GRUPO_PERSONAL = 'Agregados por mí';

  // ---------------- MOTIVOS (CASO ENCONTRADO) ----------------
  // Siempre escribir "del usuario {nombre}" o "el usuario {nombre}": se cambia solo a "usuaria" cuando corresponde.
  const MOTIVOS = [
    { id: 'lentitud',    txt: 'Lentitud del equipo',
      caso: 'Se detectó lentitud en el equipo del usuario {nombre}, lo que afectaba el normal desarrollo de sus labores.' },
    { id: 'revision',    txt: 'Revisión de equipo',
      caso: 'Se solicitó la revisión técnica del equipo del usuario {nombre} para asegurar su correcto funcionamiento.' },
    { id: 'congela',     txt: 'Equipo se congela / no responde',
      caso: 'Se reportó que el equipo del usuario {nombre} se congelaba o dejaba de responder, lo que afectaba el normal desarrollo de sus labores.' },
    { id: 'preventivo',  txt: 'Mantenimiento preventivo de software',
      caso: 'Se programó el mantenimiento preventivo de software del equipo del usuario {nombre}, con el fin de asegurar su óptimo funcionamiento.' },
    { id: 'certificado', txt: 'Instalación de certificado digital',
      caso: 'Se solicitó la instalación y configuración del certificado digital para el usuario {nombre}, necesario para la firma digital de documentos institucionales.' },
    { id: 'impresora',   txt: 'Configuración de impresora',
      caso: 'Se solicitó la configuración de impresión en el equipo del usuario {nombre}, necesaria para el desarrollo de sus labores.' },
    { id: 'otro',        txt: 'Otro (escribir)', caso: '' }
  ];

  // ---------------- PLANTILLAS (según el ticket padre) ----------------
  // Un clic elige el motivo y marca los tickets hijo de ese caso.
  const PLANTILLAS = [
    { n: 'Lentitud del equipo', icono: 'bi-speedometer', motivo: 'lentitud',
      items: ['Actualización de Windows Update', 'Actualización de Office', 'Actualización de Java', 'Limpieza de temporales'] },
    { n: 'Certificado digital', icono: 'bi-pen', motivo: 'certificado',
      items: ['Actualización de Java', 'Instalación de ReFirma', 'Instalación de Perú Firma', 'Cambio de cargo en Perú Firma'] },
    { n: 'Impresora y escáner', icono: 'bi-printer', motivo: 'impresora',
      items: ['Configuración de impresora', 'Configuración de escáner'] }
  ];

  // ---------------- BASE DE DATOS EN ESTE NAVEGADOR ----------------
  const K_AJUSTES = 'mn-gen-ajustes', K_PERSONAL = 'mn-gen-personalizadas', K_USUARIOS = 'mn-gen-usuarios', K_HIST = 'mn-gen-historial', K_SPEECH = 'mn-gen-speeches';
  const MAX_HISTORIAL = 300;
  const AJUSTES_BASE = { cuenta: 'operador_usau07@minedu.gob.pe', destino: 'mesadeservicios@minedu.gob.pe', web: 'https://outlook.office.com' };
  const leerJSON = (k, def) => { try{ return JSON.parse(localStorage.getItem(k)) || def; }catch(e){ return def; } };
  const CLAVES_DATOS = new Set(['mn-gen-personalizadas', 'mn-gen-usuarios', 'mn-gen-historial', 'mn-gen-speeches', 'mn-gen-plantillas', 'mn-gen-borrados']);
  const guardarJSON = (k, v) => {
    let ok = true;
    try{ localStorage.setItem(k, JSON.stringify(v)); }catch(e){ ok = false; }
    if(CLAVES_DATOS.has(k)) programarGuardadoArchivo();   // copia automática al archivo de base (si está conectado)
    return ok;
  };
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
  const CAMPOS_USUARIO = { nombre: 'genNombre', usuario: 'genUsuario', trato: 'genTrato', sede: 'genSede', piso: 'genPiso', area: 'genArea', contacto: 'genContacto' };

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
      trato: $('genTrato').value === 'Usuaria' ? 'Usuaria' : 'Usuario',
      sede: limpio($('genSede').value),
      piso: limpio($('genPiso').value).replace(/^piso\s*/i, ''),
      area: limpio($('genArea').value),
      contacto: limpio($('genContacto').value),
      caso
    };
  }

  // ---------------- formato del correo ----------------
  function armarAsunto(titulo, d){
    return `${titulo} - Ticket Padre: ${d.padre} - Sin Notificacion - Por Regularizar.`;
  }
  // El cuerpo empieza repitiendo el asunto; luego usuario, lugar (sede - área - piso) y anexo.
  function armarCuerpo(d, asunto){
    const lineas = [
      asunto,
      '',
      `${d.trato}: ${d.nombre} (${d.usuario})`,
      `Lugar: ${[d.sede, d.area, d.piso ? 'Piso ' + d.piso : ''].filter(Boolean).join(' - ')}`
    ];
    if(d.contacto) lineas.push(`Anexo: ${d.contacto}`);
    lineas.push('', `Por Regularizar con ${aliasCuenta()}`);
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

  // los textos se guardan con "del usuario {nombre}" y se cambian a "de la usuaria" si corresponde
  const aTrato = (txt, trato) => trato === 'Usuaria'
    ? String(txt).replace(/\bdel usuario \{nombre\}/g, 'de la usuaria {nombre}').replace(/\bel usuario \{nombre\}/g, 'la usuaria {nombre}')
    : String(txt).replace(/\bde la usuaria \{nombre\}/g, 'del usuario {nombre}').replace(/\bla usuaria \{nombre\}/g, 'el usuario {nombre}');
  function speechOriginal(it, d){
    return { caso: conNombre(aTrato(`${d.caso} Al revisar, ${hallazgoDe(it)}.`, d.trato), d.nombre), accion: accionDe(it) };
  }
  function speechPara(it, d){
    const orig = speechOriginal(it, d);
    const mio = misSpeeches[claveSpeech(d.motivo, it.t)];
    return mio ? { caso: conNombre(aTrato(mio.caso, d.trato), d.nombre), accion: mio.accion, propio: true, orig } : { ...orig, propio: false, orig };
  }
  function guardarMiSpeech(b, item){
    misSpeeches[claveSpeech(b.motivo || 'lentitud', item.titulo)] = {
      caso: aTrato(sinNombre(item.caso, b.nombre), 'Usuario'), accion: item.accion, fecha: new Date().toISOString(), motivoTxt: b.motivoTxt || ''
    };
    guardarJSON(K_SPEECH, misSpeeches);
    item.propio = true;
  }

  // ---------------- usuarios: guardar y autocompletar ----------------
  function guardarUsuario(d){
    if(!d.usuario) return;
    quitarBorrado('u', d.usuario);
    usuarios[d.usuario] = { nombre: d.nombre, usuario: d.usuario, trato: d.trato, sede: d.sede, piso: d.piso, area: d.area, contacto: d.contacto, actualizado: new Date().toISOString() };
    guardarJSON(K_USUARIOS, usuarios);
    pintarDatalists();
  }
  function cargarUsuario(u){
    Object.entries(CAMPOS_USUARIO).forEach(([k, id]) => { $(id).value = u[k] || (k === 'trato' ? 'Usuario' : ''); });
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
    $('genMsg').innerHTML = `<div class="gen-listo"><i class="bi bi-check-circle-fill"></i> ${sel.length} correos preparados. Todavía NO se envió nada: envíalos en el paso 3 con <strong>ABRIR SIGUIENTE CORREO</strong>.</div>`;
    bloque = {
      id: Date.now().toString(36), fecha: new Date().toISOString(),
      padre: d.padre, usuario: d.usuario, nombre: d.nombre, trato: d.trato, motivo: d.motivo, motivoTxt: d.motivoTxt,
      items: sel.map(it => {
        const sp = speechPara(it, d), asunto = armarAsunto(it.t, d);
        return { titulo: it.t, asunto, cuerpo: armarCuerpo(d, asunto), caso: sp.caso, accion: sp.accion, sino: 'SI', evid: sp.accion,
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
    const prueba = ajustes.destino.toLowerCase() !== AJUSTES_BASE.destino;
    $('genModoPrueba').style.display = prueba ? '' : 'none';
    $('genModoPrueba').innerHTML = `<i class="bi bi-cone-striped"></i> MODO PRUEBA: los correos irán a ${esc(ajustes.destino)} y no a Mesa de Servicios. Para el uso real, vuelve a poner ${esc(AJUSTES_BASE.destino)} en "Cuenta de operador y destino".`;
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
      `<details class="gen-cuerpo-prev"><summary>Ver cuerpo del primer correo</summary><pre>${esc(items[0].cuerpo)}</pre>
        <button type="button" class="btn btn-outline-secondary btn-sm" id="genCopiarCuerpo"><i class="bi bi-clipboard"></i> Copiar cuerpo</button></details>`;
    const pend = items.filter(c => !c.enviado).length;
    const sig = items.find(c => !c.enviado), btn = $('genSiguiente');
    btn.href = sig ? urlWeb(sig) : '#';
    btn.classList.toggle('disabled', !sig);
    btn.setAttribute('aria-disabled', String(!sig));
    $('genSiguienteTxt').textContent = pend ? `ABRIR SIGUIENTE CORREO (${items.length - pend + 1} de ${items.length})` : 'TODOS LOS CORREOS FUERON ABIERTOS';
  }

  function marcarEnviado(i){
    if(!bloque || !bloque.items[i]) return;
    bloque.items[i].enviado = true;
    bloque.modificado = new Date().toISOString();
    guardarHistorial();
    setTimeout(pintarCorreos, 250);
  }

  // el botón es un enlace: el navegador lo abre en otra pestaña sin bloquearlo; aquí solo se marca
  function abrirSiguiente(e){
    const i = bloque ? bloque.items.findIndex(c => !c.enviado) : -1;
    if(i < 0){ e.preventDefault(); return; }
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
        <button type="button" class="gen-link mb-2 align-self-start" data-claude="${i}"><i class="bi bi-magic"></i> Mejorar con Claude (copia la instrucción para pegarla en un chat)</button>
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
      '# ============================================================================',
      '#  Envio de tickets a Mesa de Servicios - Generador de Tickets (MINEDU/OTIC)',
      '# ----------------------------------------------------------------------------',
      '#  Que hace: crea un correo por cada ticket y los ENVIA o los deja como',
      '#  BORRADORES usando el Outlook de escritorio que ya tienes abierto con tu',
      '#  sesion. NO pide ni guarda contrasenas y NO se conecta a ningun otro lugar.',
      '#  Puedes abrir este archivo con el Bloc de notas para leerlo antes de usarlo.',
      '#  Como ejecutarlo: clic derecho sobre el archivo > Ejecutar con PowerShell.',
      '# ============================================================================',
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
      'if ($op -notmatch "^[SsBb]$") { Write-Host "Cancelado. No se envio nada." -ForegroundColor Yellow; Read-Host "Presiona Enter para cerrar"; return }',
      'try { $ol = [Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application") } catch { try { $ol = New-Object -ComObject Outlook.Application } catch { $ol = $null } }',
      'if (-not $ol) { Write-Host "No se pudo abrir Outlook de escritorio (clasico). Abrelo e intenta de nuevo, o usa los botones de la pagina." -ForegroundColor Red; Read-Host "Presiona Enter para cerrar"; return }',
      '$acc = $null',
      'foreach ($a in $ol.Session.Accounts) { if ($a.SmtpAddress -ieq $cuenta) { $acc = $a } }',
      'if (-not $acc) {',
      '  Write-Host "No se encontro la cuenta $cuenta en Outlook. Se usara la cuenta predeterminada de Outlook." -ForegroundColor Yellow',
      '  $sigue = Read-Host "Continuar? (S/N)"',
      '  if ($sigue -notmatch "^[Ss]$") { Write-Host "Cancelado." -ForegroundColor Yellow; Read-Host "Presiona Enter para cerrar"; return }',
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
      'Write-Host ""',
      'Read-Host "Presiona Enter para cerrar"',
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
    descargar(`enviar_tickets_${(bloque.padre || 'bloque').replace(/[^\w-]/g, '')}.ps1`, armarScript(), 'text/plain');
  }

  // ---------------- atenciones propias ----------------
  // Speech automático para un ticket hijo nuevo (sin IA: se arma con reglas a partir del título,
  // el tipo y tu nota). Luego lo puedes pulir en su tarjeta o con el botón "Mejorar con Claude".
  const sinTilde = t => t.normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase();
  function articulo(t){
    const w = sinTilde(t.split(/\s+/)[0] || '');
    return /(ion|dad|a)$/.test(w) ? 'la' : 'el';
  }
  function frase(t){ t = limpio(t); if(!t) return ''; t = t.charAt(0).toUpperCase() + t.slice(1); return /[.!?]$/.test(t) ? t : t + '.'; }
  function speechAutomatico(t, tipo, nota){
    const art = articulo(t), obj = `${art} ${minuscula(t)}`;
    const c = `se requería ${obj}`;
    const accion = nota ? frase(nota) : `Se realizó ${obj} en el equipo, verificando su correcto funcionamiento.`;
    const beneficio = tipo === 'Incidente'
      ? 'Esto restablece el normal funcionamiento del equipo y de los servicios afectados.'
      : 'Esto permite atender el requerimiento y contar con las herramientas necesarias en el puesto de trabajo.';
    return { c, a: `${accion} ${beneficio} La atención quedó completada y verificada, garantizando la continuidad de las labores.` };
  }

  function agregarPersonal(){
    const t = limpio($('genNuevoTitulo').value), tipo = $('genNuevoTipo').value, nota = limpio($('genNuevaNota').value);
    const aviso = $('genNuevoMsg');
    if(!t){ aviso.innerHTML = '<span style="color:var(--mn-red)">Escribe el título del ticket hijo.</span>'; return; }
    if(todasLasAtenciones().some(x => x.t.toLowerCase() === t.toLowerCase())){ aviso.innerHTML = '<span style="color:var(--mn-red)">Ya existe un ticket hijo con ese título.</span>'; return; }
    const sp = speechAutomatico(t, tipo, nota);
    quitarBorrado('p', t.toLowerCase());
    personales.push({ t, tipo, nota, c: sp.c, a: sp.a, auto: true, fecha: new Date().toISOString() });
    guardarJSON(K_PERSONAL, personales);
    $('genNuevoTitulo').value = ''; $('genNuevaNota').value = '';
    aviso.innerHTML = `<span class="gen-ok"><i class="bi bi-check-circle-fill"></i> Agregado con speech automático.</span> Lo verás en su tarjeta al preparar; ahí puedes ajustarlo o usar «Mejorar con Claude».`;
    pintarLista(); pintarBase();
    const nuevo = document.querySelector(`#genLista input.gen-at[data-t="${CSS.escape(t)}"]`);
    if(nuevo){ nuevo.checked = true; actualizarContador(); }
  }

  // Texto listo para pegar en un chat con Claude y obtener un speech más formal
  function promptClaude(i){
    const c = bloque.items[i];
    const it = todasLasAtenciones().find(x => x.t === c.titulo) || {};
    return [
      'Redacta el speech técnico de cierre de un ticket GLPI de soporte en sitio del MINEDU (entidad del Estado peruano).',
      `Ticket hijo: "${c.titulo}"${it.tipo ? ` (${it.tipo})` : ''}.`,
      `Ticket padre / motivo: ${MOTIVOS.find(m => m.id === bloque.motivo)?.txt || 'Otro'}.`,
      `Caso encontrado actual: ${c.caso}`,
      `Acción realizada (nota del técnico): ${it.nota || c.accion}`,
      '',
      'Devuélveme solo dos textos, en tono técnico e institucional y sin inventar datos:',
      'CASO ENCONTRADO: 1 o 2 oraciones (motivo, afectación y qué se encontró al revisar).',
      'ACCIONES REALIZADAS: 3 oraciones (acción técnica concreta, beneficio y resultado), terminando en "garantizando la continuidad de las labores".'
    ].join('\n');
  }

  // ---------------- plantillas ----------------
  const K_PLANT = 'mn-gen-plantillas';
  let misPlantillas = leerJSON(K_PLANT, []);   // [{n, motivo, items:[títulos]}]
  function pintarPlantillas(){
    const todas = [...PLANTILLAS.map(p => ({ ...p, mia: false })), ...misPlantillas.map((p, k) => ({ ...p, icono: 'bi-bookmark-star', mia: true, k }))];
    $('genPlantillas').innerHTML = todas.map((p, k) =>
      `<span class="gen-plant"><button type="button" class="btn btn-outline-secondary btn-sm" data-plant="${k}"><i class="bi ${p.icono}"></i> ${esc(p.n)}</button>${
        p.mia ? `<button type="button" class="gen-borrar" data-plant-borrar="${p.k}" title="Borrar plantilla"><i class="bi bi-x-lg"></i></button>` : ''}</span>`).join('') +
      `<button type="button" class="btn btn-sm gen-link" id="genGuardarPlant"><i class="bi bi-plus-lg"></i> Guardar selección como plantilla</button>`;
    pintarPlantillas.todas = todas;
  }
  function aplicarPlantilla(p){
    const tit = new Set(p.items);
    document.querySelectorAll('#genLista input.gen-at').forEach(c => c.checked = tit.has(c.dataset.t));
    if(p.motivo && MOTIVOS.some(m => m.id === p.motivo)){ $('genMotivo').value = p.motivo; $('genMotivo').dispatchEvent(new Event('change')); }
    actualizarContador();
  }
  function guardarPlantilla(){
    const items = seleccionadas().map(x => x.t);
    if(!items.length){ alert('Primero marca los tickets hijo que tendrá la plantilla.'); return; }
    const n = limpio(prompt('Nombre de la plantilla (por ejemplo, el título del ticket padre):', '') || '');
    if(!n) return;
    quitarBorrado('pl', n);
    misPlantillas = misPlantillas.filter(p => p.n !== n);
    misPlantillas.push({ n, motivo: $('genMotivo').value, items });
    guardarJSON(K_PLANT, misPlantillas); pintarPlantillas();
  }

  // ---------------- MI BASE DE DATOS (usuarios, historial, respaldo) ----------------
  function pintarBase(){
    const nu = Object.keys(usuarios).length, nh = historial.length, np = personales.length, ns = Object.keys(misSpeeches).length;
    $('genBaseResumen').textContent = `${nu} usuario${nu === 1 ? '' : 's'} · ${nh} bloque${nh === 1 ? '' : 's'} en historial · ${np} ticket${np === 1 ? '' : 's'} hijo propio${np === 1 ? '' : 's'} · ${ns} speech${ns === 1 ? '' : 'es'} propio${ns === 1 ? '' : 's'}`;

    const nd = directorio.lista.length;
    $('genDirEstado').innerHTML = nd
      ? `<span class="gen-ok"><i class="bi bi-journal-check"></i> Directorio cargado: ${nd.toLocaleString('es-PE')} personas</span> <span class="small text-muted">(importado ${esc(fechaCorta(directorio.fecha))})</span>`
      : '<span class="small text-muted"><i class="bi bi-journal-x"></i> Aún no hay directorio cargado.</span>';

    const sps = Object.entries(misSpeeches).sort((a, b) => a[0].localeCompare(b[0]));
    $('genMisSpeeches').innerHTML = sps.length ? `<table class="gen-tabla"><thead><tr><th>Ticket hijo</th><th>Motivo</th><th>Caso encontrado / Acción realizada</th><th></th></tr></thead><tbody>${
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

  // =====================================================================
  //  BASE DE DATOS EN UN ARCHIVO (OneDrive) + DIRECTORIO + EXCEL
  // =====================================================================
  // Todo lo que guardas (usuarios, historial, speeches, tickets hijo propios,
  // plantillas y el directorio) se copia solo a un archivo .json que eliges una vez,
  // por ejemplo en tu carpeta de OneDrive. En otra PC eliges ese mismo archivo y
  // se combina con lo que haya. Si un navegador no permite esto, queda el
  // respaldo manual (Exportar / Importar).

  // ---- pequeña base IndexedDB (guarda el "enlace" al archivo y el directorio) ----
  const idb = {
    abrir(){
      if(this._p) return this._p;
      this._p = new Promise((ok, mal) => {
        const r = indexedDB.open('mn-generador', 1);
        r.onupgradeneeded = () => r.result.createObjectStore('kv');
        r.onsuccess = () => ok(r.result); r.onerror = () => mal(r.error);
      });
      return this._p;
    },
    async get(k){ const db = await this.abrir(); return new Promise(ok => { const r = db.transaction('kv').objectStore('kv').get(k); r.onsuccess = () => ok(r.result); r.onerror = () => ok(undefined); }); },
    async set(k, v){ const db = await this.abrir(); return new Promise(ok => { const t = db.transaction('kv', 'readwrite'); t.objectStore('kv').put(v, k); t.oncomplete = () => ok(true); t.onerror = () => ok(false); }); },
    async del(k){ const db = await this.abrir(); return new Promise(ok => { const t = db.transaction('kv', 'readwrite'); t.objectStore('kv').delete(k); t.oncomplete = () => ok(true); }); }
  };

  // ---- borrados (para que lo que borras no "reviva" al combinar con otra PC) ----
  const K_BORR = 'mn-gen-borrados';
  let borrados = Object.assign({ u: {}, h: {}, s: {}, p: {}, pl: {} }, leerJSON(K_BORR, {}));
  function marcarBorrado(tipo, clave){ borrados[tipo][clave] = new Date().toISOString(); guardarJSON(K_BORR, borrados); }
  function quitarBorrado(tipo, clave){ if(borrados[tipo][clave]){ delete borrados[tipo][clave]; guardarJSON(K_BORR, borrados); } }

  // ---- directorio MINEDU (exportado de la libreta de direcciones de Outlook) ----
  let directorio = { fecha: '', lista: [] };   // lista: [{n: nombre, u: usuario, a: área, o: oficina, t: teléfono, m: correo}]
  const dirIndice = [];                         // texto normalizado para buscar rápido
  function indexarDirectorio(){
    dirIndice.length = 0;
    directorio.lista.forEach(p => dirIndice.push(sinTilde(`${p.n} ${p.u}`)));
  }

  function tituloNombre(t){
    t = limpio(t);
    if(t !== t.toUpperCase()) return t;
    const menores = new Set(['de', 'del', 'la', 'las', 'los', 'y', 'e']);
    return t.toLowerCase().split(' ').map((w, k) => (k > 0 && menores.has(w)) ? w : w.charAt(0).toUpperCase() + w.slice(1)).join(' ');
  }

  function parsearCSV(txt){
    txt = txt.replace(/^﻿/, '');
    const sep = (txt.split('\n')[0].match(/;/g) || []).length > (txt.split('\n')[0].match(/,/g) || []).length ? ';' : ',';
    const filas = []; let fila = [], campo = '', comillas = false;
    for(let i = 0; i < txt.length; i++){
      const ch = txt[i];
      if(comillas){
        if(ch === '"'){ if(txt[i + 1] === '"'){ campo += '"'; i++; } else comillas = false; }
        else campo += ch;
      } else if(ch === '"') comillas = true;
      else if(ch === sep){ fila.push(campo); campo = ''; }
      else if(ch === '\n'){ fila.push(campo.replace(/\r$/, '')); filas.push(fila); fila = []; campo = ''; }
      else campo += ch;
    }
    if(campo || fila.length){ fila.push(campo); filas.push(fila); }
    return filas.filter(f => f.some(x => x.trim()));
  }

  async function importarDirectorio(archivo){
    const msg = $('genDirMsg');
    const filas = parsearCSV(await archivo.text());
    const cab = (filas.shift() || []).map(h => sinTilde(h.trim()));
    const col = n => cab.indexOf(n);
    if(col('nombre') < 0 || (col('usuario') < 0 && col('correo') < 0)){
      msg.innerHTML = '<span style="color:var(--mn-red)">El archivo no tiene el formato del exportador (faltan las columnas Nombre y Usuario).</span>'; return;
    }
    const v = (f, n) => (col(n) >= 0 ? limpio(f[col(n)]) : '');
    const vistos = new Set(), lista = [];
    filas.forEach(f => {
      const correo = v(f, 'correo');
      const usuario = (v(f, 'usuario') || correo.split('@')[0] || '').toUpperCase();
      const nombres = v(f, 'nombres'), apellidos = v(f, 'apellidos');
      const nombre = tituloNombre(nombres && apellidos ? `${nombres} ${apellidos}` : v(f, 'nombre'));
      if(!nombre || !usuario || vistos.has(usuario)) return;
      vistos.add(usuario);
      lista.push({ n: nombre, u: usuario, a: v(f, 'area'), o: v(f, 'oficina'), t: v(f, 'telefono') || v(f, 'celular'), m: correo });
    });
    if(!lista.length){ msg.innerHTML = '<span style="color:var(--mn-red)">No se encontraron personas en el archivo.</span>'; return; }
    lista.sort((a, b) => a.n.localeCompare(b.n, 'es'));
    directorio = { fecha: new Date().toISOString(), lista };
    await idb.set('directorio', directorio);
    indexarDirectorio(); pintarBase(); programarGuardadoArchivo();
    msg.innerHTML = `<span class="gen-ok"><i class="bi bi-check-circle-fill"></i> Directorio importado: ${lista.length.toLocaleString('es-PE')} personas.</span>`;
  }

  // ---- buscador de personas (directorio + tus usuarios registrados) ----
  let sugeridos = [], sugSel = -1;
  function buscarPersonas(texto){
    const tokens = sinTilde(limpio(texto)).split(' ').filter(Boolean);
    if(!tokens.length) return [];
    const res = [], ya = new Set();
    Object.values(usuarios).forEach(u => {
      const h = sinTilde(`${u.nombre} ${u.usuario}`);
      if(tokens.every(t => h.includes(t))){ res.push({ n: u.nombre, u: u.usuario, a: u.area, registrado: true }); ya.add(u.usuario); }
    });
    for(let i = 0; i < dirIndice.length && res.length < 40; i++){
      if(tokens.every(t => dirIndice[i].includes(t))){ const p = directorio.lista[i]; if(!ya.has(p.u)) res.push(p); }
    }
    return res.slice(0, 8);
  }
  function mostrarSugerencias(campo){
    const caja = $('genSugerencias');
    sugeridos = buscarPersonas(campo.value); sugSel = -1;
    if(!sugeridos.length || limpio(campo.value).length < 2){ caja.style.display = 'none'; return; }
    caja.innerHTML = sugeridos.map((p, k) => `<button type="button" class="gen-sug" data-k="${k}">
        <span><strong>${esc(p.n)}</strong> <span class="gen-sug-usu">${esc(p.u)}</span>${p.registrado ? ' <i class="bi bi-person-check-fill gen-ok" title="Ya registrado por ti"></i>' : ''}</span>
        <span class="small text-muted">${esc(p.a || '')}</span></button>`).join('');
    const r = campo.getBoundingClientRect(), base = $('genTickets').getBoundingClientRect();
    Object.assign(caja.style, { display: 'block', left: (r.left - base.left) + 'px', top: (r.bottom - base.top + 4) + 'px', width: Math.max(r.width, 360) + 'px' });
  }
  function elegirPersona(p){
    $('genSugerencias').style.display = 'none';
    const mio = usuarios[p.u];
    if(mio){ cargarUsuario(mio); return; }
    $('genNombre').value = p.n; $('genUsuario').value = p.u;
    if(!$('genArea').value && p.a) $('genArea').value = p.a;
    if(!$('genContacto').value && p.t) $('genContacto').value = p.t;
    const aviso = $('genUsuarioAviso');
    aviso.innerHTML = '<i class="bi bi-journal-check"></i> Nombre y usuario tomados del directorio. Revisa sede, área, piso y anexo.';
    aviso.style.display = '';
  }
  function teclasSugerencias(e){
    const caja = $('genSugerencias');
    if(caja.style.display !== 'block') return;
    const bs = [...caja.querySelectorAll('.gen-sug')];
    if(e.key === 'ArrowDown' || e.key === 'ArrowUp'){
      e.preventDefault(); sugSel = (sugSel + (e.key === 'ArrowDown' ? 1 : -1) + bs.length) % bs.length;
      bs.forEach((b, k) => b.classList.toggle('activo', k === sugSel));
    } else if(e.key === 'Enter' && sugSel >= 0){ e.preventDefault(); elegirPersona(sugeridos[sugSel]); }
    else if(e.key === 'Escape') caja.style.display = 'none';
  }

  // ---- exportador del directorio (.ps1 limpio, para una PC con Outlook clásico) ----
  function scriptDirectorio(){
    return [
      '# ============================================================================',
      '#  Exporta la libreta de direcciones de Outlook a directorio_minedu.csv',
      '# ----------------------------------------------------------------------------',
      '#  Solo LEE del Outlook de escritorio ya abierto: nombre, usuario, correo,',
      '#  area, oficina, telefono y cargo. NO envia nada a ningun lado, NO borra nada',
      '#  y NO pide contrasenas. Puedes abrirlo con el Bloc de notas para leerlo.',
      '#  Como ejecutarlo: clic derecho sobre el archivo > Ejecutar con PowerShell.',
      '# ============================================================================',
      '[Console]::OutputEncoding = [Text.Encoding]::UTF8',
      '$DESTINO = if ($PSScriptRoot) { $PSScriptRoot } else { [Environment]::GetFolderPath("Desktop") }',
      'Write-Host "Leyendo la libreta de direcciones de Outlook. Puede tardar varios minutos..." -ForegroundColor Cyan',
      'try { $ol = [Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application") } catch { try { $ol = New-Object -ComObject Outlook.Application } catch { $ol = $null } }',
      'if (-not $ol) { Write-Host "No se pudo abrir Outlook de escritorio (clasico). Abrelo e intenta de nuevo." -ForegroundColor Red; Read-Host "Presiona Enter para cerrar"; return }',
      '$entradas = $ol.Session.GetGlobalAddressList().AddressEntries',
      '$total = $entradas.Count',
      '$lista = New-Object System.Collections.Generic.List[object]',
      '$i = 0',
      'foreach ($e in $entradas) {',
      '  $i++',
      '  if ($i % 200 -eq 0) { Write-Progress -Activity "Leyendo directorio" -Status "$i de $total" -PercentComplete ([int](100 * $i / [Math]::Max($total,1))) }',
      '  try {',
      '    if ($e.AddressEntryUserType -ne 0 -and $e.AddressEntryUserType -ne 5) { continue }',
      '    $u = $e.GetExchangeUser()',
      '    if (-not $u) { continue }',
      '    $lista.Add([pscustomobject]@{ Nombre = $u.Name; Nombres = $u.FirstName; Apellidos = $u.LastName; Usuario = $u.Alias; Correo = $u.PrimarySmtpAddress; Area = $u.Department; Oficina = $u.OfficeLocation; Telefono = $u.BusinessTelephoneNumber; Celular = $u.MobileTelephoneNumber; Cargo = $u.JobTitle })',
      '  } catch { }',
      '}',
      'Write-Progress -Activity "Leyendo directorio" -Completed',
      '$dest = Join-Path $DESTINO "directorio_minedu.csv"',
      '$lista | Export-Csv -Path $dest -NoTypeInformation -Encoding UTF8',
      'Write-Host ""',
      'Write-Host "Listo: $($lista.Count) personas exportadas en:" -ForegroundColor Green',
      'Write-Host "  $dest"',
      'Write-Host "Ahora, en el Generador de Tickets usa Importar directorio y elige ese archivo." -ForegroundColor Cyan',
      'try { Start-Process explorer.exe "/select,`"$dest`"" } catch { }',
      'Write-Host ""',
      'Read-Host "Presiona Enter para cerrar"',
      ''
    ].join('\r\n');
  }

  // ---- respaldo manual y combinación de datos ----
  function datosCompletos(){
    return { app: 'generador-tickets-minedu', version: 3, guardado: new Date().toISOString(),
      ajustes, personales, usuarios, historial, speeches: misSpeeches, plantillas: misPlantillas, borrados, directorio };
  }
  function exportar(){
    descargar(`respaldo_generador_tickets_${new Date().toISOString().slice(0, 10)}.json`, JSON.stringify(datosCompletos()), 'application/json');
  }

  // combina "data" con lo que hay en este navegador; no se pierde nada, gana lo más reciente
  let suspenderAuto = false;
  async function fusionar(data){
    const n = { u: 0, p: 0, s: 0, h: 0, pl: 0 };
    suspenderAuto = true;
    try{
      ['u', 'h', 's', 'p', 'pl'].forEach(t => Object.entries((data.borrados || {})[t] || {}).forEach(([k, f]) => {
        if(!borrados[t][k] || f > borrados[t][k]) borrados[t][k] = f;
      }));
      Object.values(data.usuarios || {}).forEach(u => {
        const a = usuarios[u.usuario];
        if(!a || (u.actualizado || '') > (a.actualizado || '')){ usuarios[u.usuario] = u; n.u++; }
      });
      (data.personales || []).forEach(p => {
        const x = personales.find(y => y.t.toLowerCase() === p.t.toLowerCase());
        if(!x){ personales.push(p); n.p++; }
        else if((p.fecha || '') > (x.fecha || '') || (!x.c && p.c)){ Object.assign(x, p); n.p++; }
      });
      Object.entries(data.speeches || {}).forEach(([k, v]) => {
        if(!misSpeeches[k] || (v.fecha || '') > (misSpeeches[k].fecha || '')){ misSpeeches[k] = v; n.s++; }
      });
      (data.historial || []).forEach(h => {
        const k = historial.findIndex(x => x.id === h.id);
        if(k < 0){ historial.push(h); n.h++; }
        else if((h.modificado || h.fecha || '') > (historial[k].modificado || historial[k].fecha || '')){ historial[k] = h; n.h++; }
      });
      (data.plantillas || []).forEach(p => { if(!misPlantillas.some(x => x.n === p.n)){ misPlantillas.push(p); n.pl++; } });
      if(data.directorio && (data.directorio.fecha || '') > (directorio.fecha || '')){
        directorio = data.directorio; await idb.set('directorio', directorio); indexarDirectorio();
      }
      if(data.ajustes && !localStorage.getItem(K_AJUSTES)){ ajustes = Object.assign({}, AJUSTES_BASE, data.ajustes); guardarJSON(K_AJUSTES, ajustes); pintarAjustes(); }

      // aplicar borrados
      Object.keys(borrados.u).forEach(k => { if(usuarios[k] && (usuarios[k].actualizado || '') <= borrados.u[k]) delete usuarios[k]; });
      Object.keys(borrados.s).forEach(k => { if(misSpeeches[k] && (misSpeeches[k].fecha || '') <= borrados.s[k]) delete misSpeeches[k]; });
      historial = historial.filter(h => !borrados.h[h.id]);
      personales = personales.filter(p => !(borrados.p[p.t.toLowerCase()] && (p.fecha || '') <= borrados.p[p.t.toLowerCase()]));
      misPlantillas = misPlantillas.filter(p => !borrados.pl[p.n]);
      historial.sort((a, b) => (b.fecha || '').localeCompare(a.fecha || ''));
      historial = historial.slice(0, MAX_HISTORIAL);

      guardarJSON(K_BORR, borrados); guardarJSON(K_USUARIOS, usuarios); guardarJSON(K_PERSONAL, personales);
      guardarJSON(K_SPEECH, misSpeeches); guardarJSON(K_PLANT, misPlantillas); guardarJSON(K_HIST, historial);
    } finally { suspenderAuto = false; }
    pintarLista(); pintarPlantillas(); pintarDatalists(); pintarBase();
    return n;
  }

  async function importar(archivo){
    const msg = $('genBaseMsg');
    let data = null;
    try{ data = JSON.parse(await archivo.text()); }catch(e){}
    if(!data || data.app !== 'generador-tickets-minedu'){ msg.innerHTML = '<span style="color:var(--mn-red)">El archivo no es un respaldo del Generador de Tickets.</span>'; return; }
    const n = await fusionar(data);
    programarGuardadoArchivo();
    msg.innerHTML = `<span class="gen-ok"><i class="bi bi-check-circle-fill"></i> Respaldo importado: ${n.u} usuarios, ${n.p} tickets hijo, ${n.s} speeches, ${n.h} bloques y ${n.pl} plantillas nuevos o actualizados.</span>`;
  }

  // ---- archivo de base con guardado automático (Chrome / Edge) ----
  const archivoSoportado = 'showSaveFilePicker' in window;
  let archivo = null, ultimoGuardado = '', estadoArchivo = '';
  async function leerArchivo(){
    const f = await archivo.getFile();
    const txt = await f.text();
    if(!txt.trim()) return null;
    try{ return JSON.parse(txt); }catch(e){ return null; }
  }
  async function escribirArchivo(){
    if(!archivo) return;
    try{
      const otro = await leerArchivo();                 // por si otra PC lo cambió (OneDrive)
      if(otro && otro.app === 'generador-tickets-minedu') await fusionar(otro);
      const w = await archivo.createWritable();
      await w.write(JSON.stringify(datosCompletos()));
      await w.close();
      ultimoGuardado = new Date().toISOString(); estadoArchivo = 'ok';
    }catch(e){ estadoArchivo = 'error'; console.warn('No se pudo guardar el archivo de base', e); }
    pintarEstadoArchivo();
  }
  let tArchivo;
  function programarGuardadoArchivo(){
    if(!archivo || suspenderAuto) return;
    estadoArchivo = 'pendiente'; pintarEstadoArchivo();
    clearTimeout(tArchivo); tArchivo = setTimeout(escribirArchivo, 1500);
  }
  async function usarArchivo(h, esNuevo){
    archivo = h;
    await idb.set('archivo', h);
    if(!esNuevo){
      const data = await leerArchivo();
      if(data && data.app === 'generador-tickets-minedu'){
        const n = await fusionar(data);
        $('genBaseMsg').innerHTML = `<span class="gen-ok"><i class="bi bi-check-circle-fill"></i> Base cargada: ${n.u} usuarios, ${n.h} bloques y ${n.s} speeches traídos del archivo.</span>`;
      }
    }
    await escribirArchivo();
  }
  const tiposArchivo = [{ description: 'Base del Generador de Tickets', accept: { 'application/json': ['.json'] } }];
  async function crearArchivo(){
    try{ await usarArchivo(await window.showSaveFilePicker({ suggestedName: 'base_generador_tickets.json', types: tiposArchivo }), true); }
    catch(e){ if(e.name !== 'AbortError') alert('No se pudo crear el archivo: ' + e.message); }
  }
  async function abrirArchivo(){
    try{ const [h] = await window.showOpenFilePicker({ types: tiposArchivo }); await usarArchivo(h, false); }
    catch(e){ if(e.name !== 'AbortError') alert('No se pudo abrir el archivo: ' + e.message); }
  }
  async function reconectarArchivo(){
    const h = await idb.get('archivo'); if(!h) return;
    if(await h.requestPermission({ mode: 'readwrite' }) === 'granted') await usarArchivo(h, false);
  }
  async function desconectarArchivo(){
    if(!confirm('¿Dejar de guardar en el archivo? Tus datos siguen en este navegador y en el archivo.')) return;
    archivo = null; await idb.del('archivo'); estadoArchivo = ''; pintarEstadoArchivo();
  }
  async function iniciarArchivo(){
    if(!archivoSoportado){ pintarEstadoArchivo(); return; }
    const h = await idb.get('archivo');
    if(h){
      try{
        if(await h.queryPermission({ mode: 'readwrite' }) === 'granted'){ await usarArchivo(h, false); return; }
        estadoArchivo = 'permiso'; archivo = null; pintarEstadoArchivo.nombre = h.name;
      }catch(e){}
    }
    pintarEstadoArchivo();
  }
  function pintarEstadoArchivo(){
    const el = $('genArchivoEstado'); if(!el) return;
    let html;
    if(!archivoSoportado){
      html = `<span class="gen-estado gris"><i class="bi bi-info-circle"></i> Este navegador no permite el guardado automático en archivo. Usa Chrome o Edge, o el respaldo manual.</span>`;
    } else if(archivo){
      const txt = estadoArchivo === 'error' ? '<span style="color:var(--mn-red)">no se pudo guardar (¿archivo abierto o sin conexión?)</span>'
        : estadoArchivo === 'pendiente' ? 'guardando…' : `guardado ${ultimoGuardado ? new Date(ultimoGuardado).toLocaleTimeString('es-PE', { hour: '2-digit', minute: '2-digit' }) : ''}`;
      html = `<span class="gen-estado verde"><i class="bi bi-cloud-check-fill"></i> Guardado automático en <strong>${esc(archivo.name)}</strong> · ${txt}</span>
        <button type="button" class="gen-link" id="genArchDesc">Desconectar</button>`;
    } else if(estadoArchivo === 'permiso'){
      html = `<span class="gen-estado ambar"><i class="bi bi-exclamation-triangle-fill"></i> Tu archivo de base <strong>${esc(pintarEstadoArchivo.nombre || '')}</strong> necesita permiso para seguir guardando.</span>
        <button type="button" class="btn btn-mn btn-sm" id="genArchReconectar"><i class="bi bi-plug-fill"></i> Reconectar (1 clic)</button>`;
    } else {
      html = `<span class="gen-estado gris"><i class="bi bi-cloud-slash"></i> Sin archivo de base: los datos solo están en este navegador.</span>
        <button type="button" class="btn btn-azul btn-sm" id="genArchAbrir"><i class="bi bi-folder2-open"></i> Abrir mi base existente</button>
        <button type="button" class="btn btn-outline-secondary btn-sm" id="genArchCrear"><i class="bi bi-file-earmark-plus"></i> Crear archivo de base</button>`;
    }
    el.innerHTML = html;
  }

  // ---- exportar a Excel (usuarios, historial, speeches, tickets hijo propios) ----
  async function exportarExcel(){
    if(!window.ExcelJS){ alert('No se cargó la librería de Excel.'); return; }
    const wb = new ExcelJS.Workbook();
    const hoja = (nombre, columnas, filas) => {
      const ws = wb.addWorksheet(nombre, { views: [{ state: 'frozen', ySplit: 1 }] });
      ws.columns = columnas.map(([header, key, width]) => ({ header, key, width }));
      filas.forEach(f => ws.addRow(f));
      ws.getRow(1).eachCell(c => {
        c.font = { bold: true, color: { argb: 'FFFFFFFF' } };
        c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF156082' } };
        c.alignment = { vertical: 'middle' };
      });
      ws.eachRow((r, k) => { if(k > 1){ r.alignment = { vertical: 'top', wrapText: true }; if(k % 2 === 1) r.eachCell(c => c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEAF5FB' } }); } });
      ws.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: columnas.length } };
    };
    hoja('Usuarios', [['Usuario', 'usuario', 14], ['Nombres completos', 'nombre', 34], ['Se escribe como', 'trato', 14], ['Sede', 'sede', 20], ['Área', 'area', 26], ['Piso', 'piso', 8], ['Anexo', 'contacto', 12], ['Actualizado', 'actualizado', 18]],
      Object.values(usuarios).sort((a, b) => (a.nombre || '').localeCompare(b.nombre || '')).map(u => ({ ...u, actualizado: fechaCorta(u.actualizado) })));
    const filasH = [];
    historial.forEach(h => h.items.forEach(c => filasH.push({ fecha: fechaCorta(h.fecha), padre: h.padre, usuario: h.usuario, nombre: h.nombre,
      motivo: h.motivoTxt || '', titulo: c.titulo, asunto: c.asunto, caso: c.caso, accion: c.accion, sino: c.sino, evid: c.evid, enviado: c.enviado ? 'Sí' : 'No' })));
    hoja('Historial', [['Fecha', 'fecha', 18], ['Ticket padre', 'padre', 13], ['Usuario', 'usuario', 13], ['Nombres', 'nombre', 28], ['Motivo', 'motivo', 22], ['Ticket hijo', 'titulo', 30],
      ['Asunto', 'asunto', 50], ['Caso encontrado', 'caso', 60], ['Acciones realizadas', 'accion', 70], ['¿Solucionado?', 'sino', 12], ['Evidencias', 'evid', 50], ['Abierto en Outlook', 'enviado', 12]], filasH);
    hoja('Mis speeches', [['Ticket hijo', 'titulo', 32], ['Motivo', 'motivo', 26], ['Caso encontrado', 'caso', 60], ['Acciones realizadas', 'accion', 70], ['Fecha', 'fecha', 18]],
      Object.entries(misSpeeches).map(([k, v]) => { const [m, ...t] = k.split('|'); return { titulo: t.join('|'), motivo: (MOTIVOS.find(x => x.id === m) || {}).txt || m, caso: v.caso, accion: v.accion, fecha: fechaCorta(v.fecha) }; }));
    hoja('Tickets hijo propios', [['Título', 'tit', 34], ['Tipo', 'tipo', 12], ['Nota', 'nota', 50], ['Qué se encontró', 'c', 50], ['Acciones realizadas', 'a', 70]],
      personales.map(p => ({ tit: p.t, tipo: p.tipo || '', nota: p.nota || '', c: p.c || '', a: p.a || '' })));
    const buf = await wb.xlsx.writeBuffer();
    descargar(`base_generador_tickets_${new Date().toISOString().slice(0, 10)}.xlsx`, buf, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
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
      if(it && confirm(`¿Quitar "${it.t}" de la lista?`)){ marcarBorrado('p', it.t.toLowerCase()); personales.splice(+b.dataset.idx, 1); guardarJSON(K_PERSONAL, personales); pintarLista(); pintarBase(); }
    });
    $('genTodo').addEventListener('change', e => { document.querySelectorAll('#genLista input.gen-at').forEach(c => c.checked = e.target.checked); actualizarContador(); });
    pintarPlantillas();
    $('genPlantillas').addEventListener('click', e => {
      const b = e.target.closest('button'); if(!b) return;
      if(b.id === 'genGuardarPlant') guardarPlantilla();
      else if(b.dataset.plantBorrar != null){
        const p = misPlantillas[+b.dataset.plantBorrar];
        if(p && confirm(`¿Borrar la plantilla "${p.n}"?`)){ marcarBorrado('pl', p.n); misPlantillas.splice(+b.dataset.plantBorrar, 1); guardarJSON(K_PLANT, misPlantillas); pintarPlantillas(); }
      }
      else if(b.dataset.plant != null) aplicarPlantilla(pintarPlantillas.todas[+b.dataset.plant]);
    });
    $('genAgregar').addEventListener('click', agregarPersonal);

    $('genBtnGenerar').addEventListener('click', generar);
    $('genBtnScript').addEventListener('click', descargarScript);
    $('genSiguiente').addEventListener('click', abrirSiguiente);
    $('genBtnLimpiar').addEventListener('click', () => {
      if(!confirm('¿Limpiar el formulario y empezar un nuevo bloque? (lo generado queda en el historial)')) return;
      ['genPadre','genNombre','genUsuario','genSede','genPiso','genArea','genContacto','genMotivoOtro'].forEach(id => $(id).value = '');
      $('genTrato').value = 'Usuario';
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
      else if(b.dataset.claude != null) copiar(promptClaude(+b.dataset.claude), b);
      else if(b.dataset.restaurar != null){
        const i = +b.dataset.restaurar, c = bloque.items[i];
        if(!confirm(`¿Volver al speech sugerido para "${c.titulo}"? Se borra tu versión guardada.`)) return;
        marcarBorrado('s', claveSpeech(bloque.motivo || 'lentitud', c.titulo));
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
      bloque.modificado = new Date().toISOString();
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
          marcarBorrado('h', b.dataset.histBorrar);
          historial = historial.filter(x => x.id !== b.dataset.histBorrar);
          if(bloque && bloque.id === b.dataset.histBorrar){ bloque = null; $('genResultados').style.display = 'none'; }
          guardarHistorial();
        }
      } else if(b.dataset.spBorrar){
        if(confirm('¿Borrar este speech? La próxima vez se usará el sugerido.')){ marcarBorrado('s', b.dataset.spBorrar); delete misSpeeches[b.dataset.spBorrar]; guardarJSON(K_SPEECH, misSpeeches); pintarBase(); }
      } else if(b.dataset.usuUsar){
        const u = usuarios[b.dataset.usuUsar];
        if(u){ cargarUsuario(u); window.scrollTo({ top: 0, behavior: 'smooth' }); }
      } else if(b.dataset.usuBorrar){
        if(confirm(`¿Borrar al usuario ${b.dataset.usuBorrar} de tu base?`)){
          marcarBorrado('u', b.dataset.usuBorrar); delete usuarios[b.dataset.usuBorrar]; guardarJSON(K_USUARIOS, usuarios); pintarDatalists(); pintarBase();
        }
      }
    });
    $('genExportar').addEventListener('click', exportar);
    $('genImportar').addEventListener('click', () => $('genImportarInput').click());
    $('genImportarInput').addEventListener('change', e => { if(e.target.files[0]) importar(e.target.files[0]); e.target.value = ''; });
    $('genExcel').addEventListener('click', exportarExcel);

    // archivo de base (OneDrive)
    $('genArchivoEstado').addEventListener('click', e => {
      const b = e.target.closest('button'); if(!b) return;
      if(b.id === 'genArchCrear') crearArchivo();
      else if(b.id === 'genArchAbrir') abrirArchivo();
      else if(b.id === 'genArchReconectar') reconectarArchivo();
      else if(b.id === 'genArchDesc') desconectarArchivo();
    });

    // directorio
    $('genDirScript').addEventListener('click', () => descargar('exportar_directorio_outlook.ps1', scriptDirectorio(), 'text/plain'));
    $('genDirImportar').addEventListener('click', () => $('genDirInput').click());
    $('genDirInput').addEventListener('change', e => { if(e.target.files[0]) importarDirectorio(e.target.files[0]); e.target.value = ''; });

    // buscador de personas en Nombres y Usuario
    ['genNombre', 'genUsuario'].forEach(id => {
      $(id).addEventListener('input', () => mostrarSugerencias($(id)));
      $(id).addEventListener('keydown', teclasSugerencias);
      $(id).addEventListener('blur', () => setTimeout(() => { $('genSugerencias').style.display = 'none'; }, 180));
    });
    $('genSugerencias').addEventListener('mousedown', e => {
      const b = e.target.closest('.gen-sug'); if(!b) return;
      e.preventDefault(); elegirPersona(sugeridos[+b.dataset.k]);
    });

    idb.get('directorio').then(d => { if(d && d.lista){ directorio = d; indexarDirectorio(); pintarBase(); } }).catch(() => {});
    iniciarArchivo();
  }

  if(document.readyState === 'loading') document.addEventListener('DOMContentLoaded', iniciar);
  else iniciar();
})();
