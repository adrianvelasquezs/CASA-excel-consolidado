/**
 * Cronograma 2027 Bloque Neón — permisos por rol y propiedad de filas en Google Sheets.
 *
 * Roles (hoja «Usuarios»): ADMIN (todo el libro), CASA, GESTOR y PROFESOR (estos dos, de su(s) programa(s)).
 * Qué rol edita cada columna lo define la hoja «Permisos» (columnas Profesores / Gestores / CASA con SI o NO).
 * Cada fila que registra un profesor queda solo para quien la registró (y los gestores, si así se configura).
 *
 * Cómo se aplica:
 *   1. Protecciones nativas de Google Sheets por correo (bloquean la edición en la interfaz).
 *   2. Disparador «alEditar»: cuando un profesor o gestor escribe en una fila libre, la registra a su
 *      nombre y la protege. Además revisa cada edición contra los roles: si alguien logra editar algo
 *      que no le corresponde (por ejemplo, porque quitó una protección), la revierte, lo anota en
 *      «_Auditoria», avisa a los ADMIN y restaura las protecciones de esa hoja.
 *   3. Disparador diario «restaurarPermisos»: vuelve a dejar todas las protecciones como deben estar.
 *
 * Instalar como script vinculado a la hoja (Extensiones → Apps Script) con una cuenta
 * @uniandes.edu.co, idealmente la cuenta dueña del archivo. Ver la guía de despliegue de la app.
 */

const CONFIG = {
  DOMINIO: 'uniandes.edu.co',
  HOJA_USUARIOS: 'Usuarios',
  HOJA_CONSOLIDADO: 'Consolidado',
  HOJA_INSTRUCCIONES: 'Instrucciones',
  HOJA_AUDITORIA: '_Auditoria',
  HOJA_PERMISOS: 'Permisos',
  FILA_ENCABEZADO: 3,
  PRIMERA_FILA: 4,
  COL_PROGRAMA: 'Programa',
  COL_REGISTRADO: 'Registrado por',
  COL_FECHA: 'Fecha de registro',
  PREFIJO: '[CASA] ',
  GESTOR_EDITA_FILAS_DE_PROFESORES: true,
  ALERTAS_POR_CORREO: true,
  HORA_RESTAURACION: 2,
  LIMITE_MS: 5 * 60 * 1000, // Apps Script corta a los 6 minutos
};

const ROLES = ['ADMIN', 'CASA', 'GESTOR', 'PROFESOR'];
const HOJAS_ESPECIALES = [CONFIG.HOJA_USUARIOS, CONFIG.HOJA_CONSOLIDADO, CONFIG.HOJA_INSTRUCCIONES, CONFIG.HOJA_PERMISOS,
  CONFIG.HOJA_AUDITORIA];
const SIN_EDITORES = '---'; // columna que solo edita ADMIN
const PREFIJO_ARCHIVO = '_Archivado ';  // las hojas que empiezan por «_» no cuentan como programas
// Respaldo para libros sin hoja «Permisos»: permisos según el color del encabezado.
const COLORES_PERMISO = { '#00b0f0': 'PG-', '#c6e0b4': '--C', '#ffe699': 'PGC' };
const INICIO = Date.now();

// ---------------------------------------------------------------------------
// Menú
// ---------------------------------------------------------------------------

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('CASA')
    .addItem('1. Configurar libro (una sola vez)', 'configurarLibro')
    .addItem('2. Compartir con los correos de «Usuarios»', 'compartirConUsuarios')
    .addSeparator()
    .addItem('Aplicar permisos ahora', 'aplicarPermisos')
    .addItem('Agregar programa…', 'agregarPrograma')
    .addItem('Quitar programa…', 'quitarPrograma')
    .addItem('Diagnosticar un correo…', 'diagnosticarCorreo')
    .addItem('Ver auditoría', 'verAuditoria')
    .addToUi();
}

function configurarLibro() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const u = exigirAdmin_(ss);
  if (!u) return;

  const yo = correoEfectivo_();
  const props = PropertiesService.getDocumentProperties();
  const instalador = props.getProperty('INSTALADO_POR');
  if (instalador && instalador !== yo) {
    ui.alert(`Los disparadores ya los instaló ${instalador}. Debe ejecutar esta opción esa misma cuenta ` +
      '(o eliminar sus disparadores primero) para no duplicar el procesamiento de cada edición.');
    return;
  }

  const programas = hojasPrograma_(ss);
  if (!programas.length) {
    ui.alert('No se encontraron hojas de programa (fila 3 con encabezado «Programa» en la columna A).');
    return;
  }
  hojaAuditoria_(ss);
  instalarDisparadores_(ss);
  props.setProperty('INSTALADO_POR', yo);

  const r = sincronizarTodo_(ss, u);
  ui.alert('Libro configurado', resumen_(r, u) +
    '\n\nSiguiente paso: CASA → «Compartir con los correos de Usuarios».', ui.ButtonSet.OK);
}

function aplicarPermisos() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const u = exigirAdmin_(ss);
  if (!u) return;
  const r = sincronizarTodo_(ss, u);
  SpreadsheetApp.getUi().alert('Permisos aplicados', resumen_(r, u), SpreadsheetApp.getUi().ButtonSet.OK);
}

function compartirConUsuarios() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const u = exigirAdmin_(ss);
  if (!u) return;
  const actuales = new Set(ss.getEditors().map(x => x.getEmail().toLowerCase()));
  const nuevos = u.todos.filter(c => !actuales.has(c));
  if (!nuevos.length) {
    ui.alert('Todos los correos de «Usuarios» ya tienen acceso de edición.');
    return;
  }
  const ok = ui.alert(`Se dará acceso de edición a ${nuevos.length} correo(s):\n\n${nuevos.join('\n')}\n\n` +
    'Las protecciones definen qué celdas puede modificar cada uno. ¿Continuar?', ui.ButtonSet.YES_NO);
  if (ok !== ui.Button.YES) return;
  const fallos = [];
  nuevos.forEach(c => {
    try { ss.addEditor(c); } catch (err) { fallos.push(`${c}: ${err.message}`); }
  });
  // Con el acceso ya otorgado, se vuelven a aplicar los permisos para que las personas nuevas
  // queden como editoras de sus columnas (antes de tener acceso, Google puede no aceptarlas).
  const r = sincronizarTodo_(ss, leerUsuarios_(ss));
  ui.alert('Compartir', `Acceso de edición otorgado a ${nuevos.length - fallos.length} correo(s).` +
    (fallos.length ? `\n\nNo se pudo con:\n${fallos.join('\n')}` : '') +
    '\n\nEl acceso de edición al archivo es necesario: las protecciones limitan qué celdas puede modificar cada uno.' +
    '\nComparta el enlace también por Teams (el aviso de Google puede no llegar a Outlook).' +
    '\n\nPermisos aplicados de nuevo:\n' + resumen_(r, leerUsuarios_(ss)), ui.ButtonSet.OK);
}

function agregarPrograma() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const u = exigirAdmin_(ss);
  if (!u) return;
  const programas = hojasPrograma_(ss);
  if (!programas.length) {
    ui.alert('No hay ninguna hoja de programa para usar como modelo.');
    return;
  }
  const respuesta = ui.prompt('Agregar programa',
    'Nombre del programa nuevo. Será el nombre de la hoja (máximo 31 caracteres), por ejemplo: MBA2', ui.ButtonSet.OK_CANCEL);
  if (respuesta.getSelectedButton() !== ui.Button.OK) return;
  const nombre = normalizar_(respuesta.getResponseText());
  const problema = validarNombrePrograma_(ss, nombre);
  if (problema) {
    ui.alert('No se pudo agregar el programa', problema, ui.ButtonSet.OK);
    return;
  }
  const hoja = crearHojaPrograma_(ss, nombre, programas);
  actualizarListaProgramas_(ss);
  const r = sincronizarTodo_(ss, leerUsuarios_(ss));
  ss.setActiveSheet(hoja);
  ui.alert('Programa agregado',
    `Se creó la hoja «${nombre}» con las mismas columnas que «${programas[0].getName()}», vacía y ya incluida en el ` +
    'Consolidado.\n\nSiguiente paso: en la hoja Usuarios agregue a sus gestores y profesores con el programa ' +
    `«${nombre}» y use CASA → 2. Compartir.\n\nPara que este programa también esté en los libros que se generen ` +
    'en el futuro, agréguelo en la app (pestaña Programas).\n\n' + resumen_(r, leerUsuarios_(ss)), ui.ButtonSet.OK);
}

function quitarPrograma() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const u = exigirAdmin_(ss);
  if (!u) return;
  const programas = hojasPrograma_(ss);
  const respuesta = ui.prompt('Quitar programa', 'Nombre del programa (hoja) que quiere quitar:', ui.ButtonSet.OK_CANCEL);
  if (respuesta.getSelectedButton() !== ui.Button.OK) return;
  const escrito = normalizar_(respuesta.getResponseText());
  const hoja = programas.find(h => claveTexto_(h.getName()) === claveTexto_(escrito));
  if (!hoja) {
    ui.alert('No se encontró el programa', `«${escrito}» no es una hoja de programa. Programas: ` +
      programas.map(h => h.getName()).join(', '), ui.ButtonSet.OK);
    return;
  }
  if (programas.length === 1) {
    ui.alert('No se puede quitar', 'Es el único programa del libro.', ui.ButtonSet.OK);
    return;
  }
  const nombre = hoja.getName();
  const cols = columnas_(hoja, leerPermisos_(ss));
  const filas = hoja.getMaxRows() - CONFIG.PRIMERA_FILA + 1;
  const registradas = hoja.getRange(CONFIG.PRIMERA_FILA, 1, filas, cols.ancho).getValues()
    .filter(f => cols.identifican.some(c => f[c - 1] !== '' && f[c - 1] !== null)).length;
  const personas = Object.keys(u.programas).includes(nombre)
    ? [...u.programas[nombre].gestores, ...u.programas[nombre].profesores] : [];

  const decision = ui.alert(`Quitar «${nombre}»`,
    `El programa tiene ${registradas} curso(s) registrado(s).\n\n` +
    'SÍ = ARCHIVAR (recomendado): la hoja se oculta y queda solo para ADMIN; sale del Consolidado, de los ' +
    'permisos y de la lista de programas, pero conserva los datos.\n\n' +
    'NO = ELIMINAR la hoja y sus datos (solo se recuperan con Archivo → Historial de versiones).\n\n' +
    'CANCELAR = no hacer nada.', ui.ButtonSet.YES_NO_CANCEL);
  if (decision !== ui.Button.YES && decision !== ui.Button.NO) return;
  if (decision === ui.Button.NO && ui.alert('Confirmar eliminación',
    `¿Eliminar definitivamente la hoja «${nombre}» con ${registradas} curso(s) registrado(s)?`,
    ui.ButtonSet.YES_NO) !== ui.Button.YES) return;

  let hecho;
  if (decision === ui.Button.YES) {
    const archivada = archivarHojaPrograma_(ss, hoja, u);
    hecho = `«${nombre}» quedó archivada en la hoja oculta «${archivada}» (solo ADMIN). Para recuperarla: muestre la ` +
      `hoja, quítele «${PREFIJO_ARCHIVO}» del nombre y use CASA → Aplicar permisos ahora.`;
  } else {
    ss.deleteSheet(hoja);
    hecho = `La hoja «${nombre}» se eliminó.`;
  }
  actualizarListaProgramas_(ss);
  const usuarios = leerUsuarios_(ss);
  const r = sincronizarTodo_(ss, usuarios);
  ui.alert('Programa quitado', hecho +
    (personas.length ? `\n\nEstas personas de la hoja Usuarios tenían el programa «${nombre}»; ajuste o borre sus ` +
      `filas:\n${personas.map(c => '  • ' + c).join('\n')}` : '') +
    '\n\nQuítelo también en la app (pestaña Programas) para los libros que se generen en el futuro.\n\n' +
    resumen_(r, usuarios), ui.ButtonSet.OK);
}

/** Oculta la hoja con un nombre que empieza por «_» (deja de contar como programa) y la deja solo para ADMIN. */
function archivarHojaPrograma_(ss, hoja, u) {
  let nombre = (PREFIJO_ARCHIVO + hoja.getName()).slice(0, 31);
  for (let i = 2; ss.getSheetByName(nombre); i++) nombre = (PREFIJO_ARCHIVO + i + ' ' + hoja.getName()).slice(0, 31);
  hoja.getProtections(SpreadsheetApp.ProtectionType.RANGE)
    .filter(p => tipoProteccion_(p))
    .forEach(p => p.remove());
  hoja.setName(nombre);
  protegerHoja_(hoja, [...u.admins], [], nuevoResumen_());
  hoja.hideSheet();
  return nombre;
}

/** Devuelve por qué el nombre no sirve para un programa nuevo, o null si sirve. */
function validarNombrePrograma_(ss, nombre) {
  if (!nombre) return 'Escriba el nombre del programa.';
  if (nombre.length > 31 || /[\[\]:*?\/\\']/.test(nombre)) {
    return `«${nombre}» no sirve como nombre de hoja: máximo 31 caracteres y sin [ ] : * ? / \\ '`;
  }
  if (nombre.startsWith('_') || HOJAS_ESPECIALES.some(h => claveTexto_(h) === claveTexto_(nombre))) {
    return `«${nombre}» es un nombre reservado del libro.`;
  }
  const existente = ss.getSheets().find(h => claveTexto_(h.getName()) === claveTexto_(nombre));
  return existente ? `Ya existe una hoja llamada «${existente.getName()}».` : null;
}

/** Copia una hoja de programa (mismas columnas, formatos y listas), la vacía y la adapta al programa nuevo. */
function crearHojaPrograma_(ss, nombre, programas) {
  const hoja = programas[0].copyTo(ss).setName(nombre);
  ss.setActiveSheet(hoja);
  ss.moveActiveSheet(programas[programas.length - 1].getIndex() + 1);  // junto a los demás programas
  const cols = columnas_(hoja, leerPermisos_(ss));
  hoja.getRange(CONFIG.PRIMERA_FILA, 1, hoja.getMaxRows() - CONFIG.PRIMERA_FILA + 1, cols.ancho).clearContent();
  restaurarFormulaPrograma_(hoja, cols, CONFIG.PRIMERA_FILA, hoja.getMaxRows());
  const titulo = hoja.getRange(1, 1, 2, hoja.getLastColumn()).getValues();
  titulo.forEach((fila, i) => fila.forEach((valor, j) => {
    if (String(valor).startsWith('Programa: ')) hoja.getRange(i + 1, j + 1).setValue('Programa: ' + nombre);
  }));
  return hoja;
}

/** Lista desplegable de la columna Programa(s) de Usuarios con las hojas de programa actuales. */
function actualizarListaProgramas_(ss) {
  const hoja = ss.getSheetByName(CONFIG.HOJA_USUARIOS);
  if (!hoja) return;
  const regla = SpreadsheetApp.newDataValidation()
    .requireValueInList(hojasPrograma_(ss).map(h => h.getName()), true)
    .setAllowInvalid(true)  // permite escribir varios programas separados por coma
    .setHelpText('Para varios programas, sepárelos con coma.')
    .build();
  hoja.getRange(2, 3, Math.max(hoja.getMaxRows() - 1, 1), 1).setDataValidation(regla);
}

function diagnosticarCorreo() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const u = exigirAdmin_(ss);
  if (!u) return;
  const respuesta = ui.prompt('Diagnosticar un correo', 'Correo de la persona, tal como entra a Google:', ui.ButtonSet.OK_CANCEL);
  if (respuesta.getSelectedButton() !== ui.Button.OK) return;
  const correo = respuesta.getResponseText().trim().toLowerCase();
  if (correo) ui.alert('Diagnóstico de ' + correo, diagnostico_(ss, u, correo), ui.ButtonSet.OK);
}

/** Revisa por qué una persona podría no poder editar: acceso al archivo, rol, columnas y protecciones. */
function diagnostico_(ss, u, correo) {
  const bien = [];
  const problemas = [];

  const editores = ss.getEditors().map(x => x.getEmail().toLowerCase());
  const lectores = ss.getViewers().map(x => x.getEmail().toLowerCase());
  if (editores.includes(correo) || correo === correoDueno_()) {
    bien.push('Acceso al archivo: Editor.');
  } else if (lectores.includes(correo)) {
    problemas.push('Solo tiene acceso de lectura o comentarios. Necesita acceso de Editor (menú CASA → 2. Compartir); ' +
      'las protecciones limitan qué celdas puede modificar.');
  } else {
    problemas.push('No tiene acceso al archivo con este correo. Use menú CASA → 2. Compartir.');
  }
  if (!correo.endsWith('@' + CONFIG.DOMINIO)) {
    problemas.push(`No es un correo @${CONFIG.DOMINIO}: el script no puede identificarlo y revierte sus ediciones.`);
  }

  const roles = [];
  if (u.admins.has(correo)) roles.push('ADMIN');
  if (u.casa.has(correo)) roles.push('CASA');
  Object.keys(u.programas).forEach(p => {
    if (u.programas[p].gestores.has(correo)) roles.push(`GESTOR de ${p}`);
    if (u.programas[p].profesores.has(correo)) roles.push(`PROFESOR de ${p}`);
  });
  if (roles.length) bien.push('Rol en Usuarios: ' + roles.join(', ') + '.');
  else problemas.push('No aparece en la hoja Usuarios con un rol y programa válidos (revise el correo exacto, el rol y el programa).');

  if (u.admins.has(correo)) {
    bien.push('Como ADMIN puede editar todo.');
  } else {
    const permisos = leerPermisos_(ss);
    hojasPrograma_(ss).forEach(hoja => {
      const r = rolesDe_(u, correo, hoja.getName());
      if (!r.casa && !r.gestor && !r.profesor) return;
      const cols = columnas_(hoja, permisos);
      const editables = cols.tipos
        .map((t, i) => (t && t !== 'sistema' && rolPuedeEditar_(t, r) ? cols.nombres[i] : null))
        .filter(Boolean);
      if (!editables.length) {
        problemas.push(`${hoja.getName()}: su rol no puede editar ninguna columna (revise la hoja Permisos).`);
        return;
      }
      bien.push(`${hoja.getName()}: puede editar ${editables.length} columna(s): ` +
        editables.slice(0, 5).join(', ') + (editables.length > 5 ? '…' : '.'));
      const sinAgregar = hoja.getProtections(SpreadsheetApp.ProtectionType.RANGE).filter(p => {
        const tipo = tipoProteccion_(p) || '';
        return tipo.startsWith('col') && rolPuedeEditar_(tipo.slice(3), r) &&
          !p.getEditors().map(x => x.getEmail().toLowerCase()).includes(correo);
      });
      if (sinAgregar.length) {
        problemas.push(`${hoja.getName()}: aún no figura en ${sinAgregar.length} protección(es) de sus columnas. ` +
          'Use menú CASA → Aplicar permisos ahora.');
      }
    });
  }

  const texto = problemas.length
    ? '⚠️ PROBLEMAS:\n' + problemas.map(p => '  • ' + p).join('\n')
    : '✓ No se encontraron problemas.';
  return texto + '\n\n' + bien.map(b => '✓ ' + b).join('\n') +
    '\n\nSi todo está bien y aun así no puede editar: confirme que entra a Google con este mismo correo ' +
    '(no con un alias ni con una cuenta personal).';
}

function verAuditoria() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!exigirAdmin_(ss)) return;
  const hoja = hojaAuditoria_(ss);
  hoja.showSheet();
  ss.setActiveSheet(hoja);
}

// ---------------------------------------------------------------------------
// Disparadores
// ---------------------------------------------------------------------------

function instalarDisparadores_(ss) {
  ScriptApp.getProjectTriggers()
    .filter(t => ['alEditar', 'restaurarPermisos'].includes(t.getHandlerFunction()))
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('alEditar').forSpreadsheet(ss).onEdit().create();
  ScriptApp.newTrigger('restaurarPermisos').timeBased().everyDays(1).atHour(CONFIG.HORA_RESTAURACION).create();
}

/** Disparador diario: deja todas las protecciones como indica la hoja «Usuarios». */
function restaurarPermisos() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const r = sincronizarTodo_(ss, leerUsuarios_(ss));
  console.log(resumen_(r));
}

/** Disparador instalable onEdit (corre con la cuenta que lo instaló). */
function alEditar(e) {
  if (!e || !e.range) return;
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) {
    console.warn('No se obtuvo el bloqueo; edición sin procesar en ' + e.range.getA1Notation());
    return;
  }
  try {
    procesarEdicion_(e);
  } catch (err) {
    console.error(err);
    registrarAuditoria_(e.source, correoDe_(e), e.range.getSheet().getName(), e.range.getA1Notation(),
      'Error del script: ' + err.message, 'No', e.value, e.oldValue);
  } finally {
    lock.releaseLock();
  }
}

function procesarEdicion_(e) {
  const ss = e.source;
  const hoja = e.range.getSheet();
  const nombre = hoja.getName();
  const correo = correoDe_(e);
  const u = leerUsuarios_(ss);
  const esAdmin = u.admins.has(correo);

  if (!esHojaPrograma_(hoja)) {
    if (esAdmin && (nombre === CONFIG.HOJA_USUARIOS || nombre === CONFIG.HOJA_PERMISOS)) {
      ss.toast(`Cuando termine de editar «${nombre}», use CASA → Aplicar permisos ahora.`, 'CASA', 8);
    } else if (!esAdmin && HOJAS_ESPECIALES.includes(nombre)) {
      incidente_(e, u, correo, null, 'Edición en hoja restringida');
    }
    return;
  }

  const cols = columnas_(hoja, leerPermisos_(ss));
  const r1 = e.range.getRow();
  const r2 = e.range.getLastRow();
  const tipos = new Set();
  for (let c = e.range.getColumn(); c <= e.range.getLastColumn(); c++) {
    tipos.add(cols.tipos[c - 1] || 'fuera');
  }

  if (!esAdmin) {
    const motivo = motivoRechazo_(u, correo, hoja, cols, r1, r2, tipos);
    if (motivo) {
      incidente_(e, u, correo, cols, motivo);
      return;
    }
  }
  if (r2 >= CONFIG.PRIMERA_FILA && [...tipos].some(esDelPrograma_)) {
    reclamarOLiberarFilas_(hoja, cols, Math.max(r1, CONFIG.PRIMERA_FILA), r2, correo, u);
  }
}

/** Devuelve por qué la edición no está permitida para ese correo, o null si lo está. */
function motivoRechazo_(u, correo, hoja, cols, r1, r2, tipos) {
  const roles = rolesDe_(u, correo, hoja.getName());
  if (!correo) return 'Usuario no identificado (cuenta fuera de @' + CONFIG.DOMINIO + ')';
  if (!u.todos.includes(correo) && !u.admins.has(correo)) {
    return `El correo ${correo} no está en la hoja Usuarios (¿entró con otra cuenta o con un alias de su correo?)`;
  }
  if (r1 < CONFIG.PRIMERA_FILA) return 'Edición de títulos o encabezados';
  if (tipos.has('sistema') || tipos.has('fuera')) return 'Edición de columnas del sistema o fuera del área de registro';
  for (const tipo of tipos) {
    if (!rolPuedeEditar_(tipo, roles)) return 'Columna que su rol no puede editar en ' + hoja.getName();
  }
  // Filas registradas por otra persona: solo CASA (en sus columnas) y los gestores (si así se configura).
  const exento = roles.casa || (roles.gestor && CONFIG.GESTOR_EDITA_FILAS_DE_PROFESORES);
  if (!exento && [...tipos].some(esDelPrograma_)) {
    const duenos = hoja.getRange(r1, cols.indice[CONFIG.COL_REGISTRADO], r2 - r1 + 1, 1).getValues();
    const ajena = duenos.some(d => {
      const dueno = String(d[0]).trim().toLowerCase();
      return dueno && dueno !== correo;
    });
    if (ajena) return 'Fila registrada por otra persona';
  }
  return null;
}

/**
 * Fila libre con datos en columnas del programa, escritos por un profesor o gestor → queda a su nombre y se protege.
 * Fila registrada sin datos en columnas del programa → se libera.
 */
function reclamarOLiberarFilas_(hoja, cols, r1, r2, correo, u) {
  const roles = rolesDe_(u, correo, hoja.getName());
  const puedeReclamar = !!correo && !roles.admin && (roles.profesor || roles.gestor);
  const n = r2 - r1 + 1;
  const valores = hoja.getRange(r1, 1, n, cols.ancho).getValues();
  const colReg = cols.indice[CONFIG.COL_REGISTRADO];
  const colFecha = cols.indice[CONFIG.COL_FECHA];
  const delPrograma = cols.tipos.map((t, i) => (esDelPrograma_(t) ? i : -1)).filter(i => i >= 0);

  valores.forEach((fila, i) => {
    const r = r1 + i;
    const dueno = String(fila[colReg - 1]).trim().toLowerCase();
    const tieneDatos = delPrograma.some(c => fila[c] !== '' && fila[c] !== null);
    if (!dueno && tieneDatos && puedeReclamar) {
      hoja.getRange(r, colReg).setValue(correo);
      hoja.getRange(r, colFecha).setValue(new Date());
      crearProteccion_(rangoFila_(hoja, cols, r), 'fila', correo, editoresFila_(u, hoja.getName(), correo));
    } else if (dueno && !tieneDatos) {
      hoja.getRange(r, colReg, 1, 1).clearContent();
      hoja.getRange(r, colFecha, 1, 1).clearContent();
      hoja.getProtections(SpreadsheetApp.ProtectionType.RANGE)
        .filter(p => tipoProteccion_(p) === 'fila' && p.getRange().getRow() === r)
        .forEach(p => p.remove());
    }
  });
}

/** Edición no permitida: revertir si se puede, registrar, avisar y restaurar protecciones. */
function incidente_(e, u, correo, cols, motivo) {
  const ss = e.source;
  const hoja = e.range.getSheet();
  const unaCelda = e.range.getNumRows() === 1 && e.range.getNumColumns() === 1;
  let revertido = 'No (use Archivo → Historial de versiones)';
  if (hoja.getName() === CONFIG.HOJA_CONSOLIDADO) {
    // Restaurar el valor de una celda rompería la fórmula: se reescribe completa.
    escribirConsolidado_(ss, hojasPrograma_(ss));
    revertido = 'Sí (fórmula reescrita)';
  } else if (unaCelda) {
    e.range.setValue(e.oldValue === undefined ? '' : e.oldValue);
    revertido = 'Sí';
  }
  if (cols) restaurarFormulaPrograma_(hoja, cols, e.range.getRow(), e.range.getLastRow());

  registrarAuditoria_(ss, correo, hoja.getName(), e.range.getA1Notation(), motivo, revertido, e.value, e.oldValue);

  if (cols) {
    sincronizarHoja_(ss, hoja, u, nuevoResumen_());
  } else {
    protegerHoja_(hoja, [...u.admins], [], nuevoResumen_());
  }

  if (CONFIG.ALERTAS_POR_CORREO && u.admins.size) {
    try {
      MailApp.sendEmail({
        to: [...u.admins].join(','),
        subject: `[Cronograma Bloque Neón] Edición no permitida en ${hoja.getName()}`,
        body: `Usuario: ${correo || '(no identificado)'}\nHoja: ${hoja.getName()}\nRango: ${e.range.getA1Notation()}\n` +
          `Motivo: ${motivo}\nRevertido: ${revertido}\n\nLas protecciones de la hoja se restauraron.\n${ss.getUrl()}`,
      });
    } catch (err) {
      console.warn('No se pudo enviar la alerta: ' + err.message);
    }
  }
}

// ---------------------------------------------------------------------------
// Sincronización de protecciones
// ---------------------------------------------------------------------------

function sincronizarTodo_(ss, u) {
  const r = nuevoResumen_();
  const programas = hojasPrograma_(ss);

  if (programas.length && ss.getSheetByName(CONFIG.HOJA_CONSOLIDADO)) {
    const problema = escribirConsolidado_(ss, programas);
    if (problema) r.fallos.push(problema);
  }
  try {
    HOJAS_ESPECIALES.forEach(nombre => {
      const hoja = ss.getSheetByName(nombre);
      if (hoja) protegerHoja_(hoja, [...u.admins], [], r);
    });
    ss.getSheets().filter(h => h.getName().startsWith(PREFIJO_ARCHIVO))  // programas archivados: solo ADMIN
      .forEach(h => protegerHoja_(h, [...u.admins], [], r));
    const permisos = leerPermisos_(ss);
    permisos.errores.forEach(e => r.fallos.push(e));
    programas.forEach(hoja => sincronizarHoja_(ss, hoja, u, r, permisos));
  } catch (err) {
    if (err.message !== 'TIEMPO') throw err;
    r.completo = false;
  }
  return r;
}

function sincronizarHoja_(ss, hoja, u, r, permisos) {
  const nombre = hoja.getName();
  const cols = columnas_(hoja, permisos || leerPermisos_(ss));
  cols.desconocidas.forEach(c => r.fallos.push(`La columna «${c}» no está en la hoja Permisos: solo ADMIN puede editarla`));
  const n = hoja.getMaxRows() - CONFIG.PRIMERA_FILA + 1;
  const g = u.programas[nombre] || { gestores: new Set(), profesores: new Set() };
  const admins = [...u.admins];
  const deseadas = [];
  const desprotegidos = [];

  // Un bloque por cada grupo contiguo de columnas con los mismos editores (según la hoja Permisos).
  const combinaciones = [...new Set(cols.tipos.filter(t => t && t !== 'sistema' && t !== SIN_EDITORES))];
  combinaciones.forEach(tipo => {
    const editores = [...admins];
    if (tipo[0] === 'P') editores.push(...g.profesores);
    if (tipo[1] === 'G') editores.push(...g.gestores);
    if (tipo[2] === 'C') editores.push(...u.casa);
    bloques_(cols.tipos, tipo).forEach(([a, b]) => {
      const rango = hoja.getRange(CONFIG.PRIMERA_FILA, a, n, b - a + 1);
      desprotegidos.push(rango);
      deseadas.push({ tipo: 'col' + tipo, detalle: nombre, rango, editores });
    });
  });
  const registros = hoja.getRange(CONFIG.PRIMERA_FILA, cols.indice[CONFIG.COL_REGISTRADO], n, 1).getValues();
  registros.forEach((v, i) => {
    const dueno = String(v[0]).trim().toLowerCase();
    if (!dueno) return;
    const fila = CONFIG.PRIMERA_FILA + i;
    deseadas.push({ tipo: 'fila', detalle: dueno, rango: rangoFila_(hoja, cols, fila), editores: editoresFila_(u, nombre, dueno) });
  });

  // Toda la hoja queda solo para ADMIN, salvo los bloques de columnas editables, que tienen su propia protección.
  protegerHoja_(hoja, admins, desprotegidos, r);

  const existentes = {};
  hoja.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(p => {
    const tipo = tipoProteccion_(p);
    const clave = tipo + '|' + p.getRange().getA1Notation();
    if (!tipo || existentes[clave]) {
      // Protecciones creadas a mano (o duplicadas) pueden bloquear a quien sí tiene permiso.
      quitarProteccion_(p, r);
    } else {
      existentes[clave] = p;
    }
  });

  deseadas.forEach(d => {
    controlarTiempo_();
    const clave = d.tipo + '|' + d.rango.getA1Notation();
    const p = existentes[clave];
    if (p) {
      delete existentes[clave];
      p.setDescription(CONFIG.PREFIJO + d.tipo + '|' + d.detalle);
      if (ajustarEditores_(p, d.editores, r)) r.actualizadas++;
      else r.sinCambios++;
    } else {
      crearProteccion_(d.rango, d.tipo, d.detalle, d.editores, r);
      r.creadas++;
    }
  });
  Object.values(existentes).forEach(p => quitarProteccion_(p, r));
}

function protegerHoja_(hoja, editores, desprotegidos, r) {
  const actuales = hoja.getProtections(SpreadsheetApp.ProtectionType.SHEET);
  const p = actuales.length ? actuales[0] : hoja.protect();
  if (actuales.length) r.actualizadas++;
  else r.creadas++;
  p.setDescription(CONFIG.PREFIJO + 'hoja|' + hoja.getName());
  p.setUnprotectedRanges(desprotegidos);
  ajustarEditores_(p, editores, r);
}

function crearProteccion_(rango, tipo, detalle, editores, r) {
  const p = rango.protect().setDescription(CONFIG.PREFIJO + tipo + '|' + detalle);
  ajustarEditores_(p, editores, r || nuevoResumen_());
  return p;
}

function quitarProteccion_(p, r) {
  try {
    p.remove();
    r.eliminadas++;
  } catch (err) {
    r.fallos.push(`No se pudo quitar la protección «${p.getDescription()}» (${p.getRange().getA1Notation()}): ${err.message}`);
  }
}

/** Deja exactamente esos editores (más la cuenta que ejecuta y el dueño). Devuelve true si cambió algo. */
function ajustarEditores_(p, deseados, r) {
  const yo = correoEfectivo_();
  const dueno = correoDueno_();
  const objetivo = new Set([...deseados, yo].filter(Boolean).map(c => c.toLowerCase()));
  if (p.isWarningOnly()) p.setWarningOnly(false);
  const actuales = p.getEditors().map(x => x.getEmail().toLowerCase());
  const sobran = actuales.filter(c => !objetivo.has(c) && c !== dueno);
  const faltan = [...objetivo].filter(c => !actuales.includes(c));
  if (!sobran.length && !faltan.length && !p.canDomainEdit()) return false;

  // Patrón recomendado por Google: asegurar la cuenta propia antes de quitar a los demás.
  p.addEditor(yo);
  if (sobran.length) p.removeEditors(sobran);
  if (p.canDomainEdit()) p.setDomainEdit(false);
  const agregar = faltan.filter(c => c !== yo);
  if (agregar.length) {
    try {
      p.addEditors(agregar);
    } catch (err) {
      agregar.forEach(c => {
        try { p.addEditor(c); } catch (err2) { r.fallos.push(`${c}: ${err2.message}`); }
      });
    }
  }
  return true;
}

function controlarTiempo_() {
  if (Date.now() - INICIO > CONFIG.LIMITE_MS) throw new Error('TIEMPO');
}

function nuevoResumen_() {
  return { creadas: 0, actualizadas: 0, sinCambios: 0, eliminadas: 0, fallos: [], completo: true };
}

function resumen_(r, u) {
  const problemas = [];
  if (!r.completo) problemas.push('Se alcanzó el límite de tiempo. Ejecute «Aplicar permisos ahora» de nuevo para terminar.');
  if (u && u.errores.length) problemas.push('Hoja Usuarios:\n' + u.errores.map(e => '  • ' + e).join('\n'));
  if (u && u.externos.length) {
    problemas.push(`Correos fuera de @${CONFIG.DOMINIO} (el script no puede identificarlos; sus ediciones se revierten):\n` +
      u.externos.map(e => '  • ' + e).join('\n'));
  }
  const avisos = [...new Set(r.fallos)];  // el mismo aviso puede repetirse en cada hoja de programa
  if (avisos.length) {
    problemas.push('Avisos:\n' + avisos.slice(0, 20).map(e => '  • ' + e).join('\n') +
      (avisos.length > 20 ? `\n  … y ${avisos.length - 20} más` : ''));
  }
  const conteo = `Protecciones creadas: ${r.creadas} · actualizadas: ${r.actualizadas} · ` +
    `sin cambios: ${r.sinCambios} · eliminadas: ${r.eliminadas}`;
  if (!problemas.length) return '✓ Sin problemas.\n\n' + conteo;
  return `⚠️ HAY ${problemas.length} PUNTO(S) QUE REVISAR (pueden impedir que alguien edite):\n\n` +
    problemas.join('\n\n') + '\n\n' + conteo + '\n\nPara revisar a una persona: menú CASA → Diagnosticar un correo…';
}

// ---------------------------------------------------------------------------
// Usuarios y roles
// ---------------------------------------------------------------------------

function leerUsuarios_(ss) {
  const hoja = ss.getSheetByName(CONFIG.HOJA_USUARIOS);
  if (!hoja) throw new Error(`No existe la hoja «${CONFIG.HOJA_USUARIOS}»`);
  const u = { admins: new Set(), casa: new Set(), programas: {}, todos: [], externos: [], errores: [] };
  const hojas = {};
  hojasPrograma_(ss).forEach(h => { hojas[claveTexto_(h.getName())] = h.getName(); });
  const dueno = correoDueno_();
  if (dueno) u.admins.add(dueno);

  const ultima = hoja.getLastRow();
  const filas = ultima < 2 ? [] : hoja.getRange(2, 1, ultima - 1, 3).getValues();
  filas.forEach((f, i) => {
    const correo = String(f[0]).trim().toLowerCase();
    const rol = String(f[1]).trim().toUpperCase();
    const programas = String(f[2]).split(',').map(s => s.trim()).filter(Boolean);
    const n = i + 2;
    if (!correo && !rol) return;
    if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(correo)) return u.errores.push(`Fila ${n}: correo inválido «${f[0]}»`);
    if (!ROLES.includes(rol)) return u.errores.push(`Fila ${n}: rol inválido «${f[1]}» (use ${ROLES.join(', ')})`);
    if (!correo.endsWith('@' + CONFIG.DOMINIO) && !u.externos.includes(correo)) u.externos.push(correo);
    if (!u.todos.includes(correo)) u.todos.push(correo);

    if (rol === 'ADMIN') u.admins.add(correo);
    else if (rol === 'CASA') u.casa.add(correo);
    else {
      if (!programas.length) return u.errores.push(`Fila ${n}: ${rol} requiere programa`);
      programas.forEach(escrito => {
        const p = hojas[claveTexto_(escrito)];
        if (!p) {
          u.errores.push(`Fila ${n}: el programa «${escrito}» no coincide con ninguna hoja ` +
            `(hojas: ${Object.values(hojas).join(', ')})`);
          return;
        }
        u.programas[p] = u.programas[p] || { gestores: new Set(), profesores: new Set() };
        u.programas[p][rol === 'GESTOR' ? 'gestores' : 'profesores'].add(correo);
      });
    }
  });
  return u;
}

function rolesDe_(u, correo, programa) {
  const g = u.programas[programa];
  return {
    admin: u.admins.has(correo),
    casa: u.casa.has(correo),
    gestor: !!g && g.gestores.has(correo),
    profesor: !!g && g.profesores.has(correo),
  };
}

function editoresFila_(u, programa, dueno) {
  const g = u.programas[programa] || { gestores: new Set(), profesores: new Set() };
  const lista = [...u.admins, ...u.casa]; // CASA sigue limitada a sus columnas por la protección de cada bloque
  if (CONFIG.GESTOR_EDITA_FILAS_DE_PROFESORES) lista.push(...g.gestores);
  if (g.profesores.has(dueno) || g.gestores.has(dueno)) lista.push(dueno);
  return lista;
}

/** tipo = «PGC» con '-' donde el rol no edita (p. ej. «PG-», «--C»). */
function rolPuedeEditar_(tipo, roles) {
  return (tipo[0] === 'P' && roles.profesor) || (tipo[1] === 'G' && roles.gestor) || (tipo[2] === 'C' && roles.casa);
}

function esDelPrograma_(tipo) {
  return !!tipo && tipo !== 'sistema' && tipo !== 'fuera' && (tipo[0] === 'P' || tipo[1] === 'G');
}

/** Lee la hoja «Permisos»: {columnas: {nombre: {tipo, identifica}}, errores: []}. */
function leerPermisos_(ss) {
  const resultado = { columnas: {}, errores: [] };
  const hoja = ss.getSheetByName(CONFIG.HOJA_PERMISOS);
  if (!hoja) {
    resultado.sinHoja = true;
    resultado.errores.push(`No existe la hoja «${CONFIG.HOJA_PERMISOS}» (libro generado con una versión anterior de la ` +
      'app): se usan los colores de los encabezados (azul = programa, verde = CASA, amarillo = ambos, gris = solo ADMIN).');
    return resultado;
  }
  const ultima = hoja.getLastRow();
  if (ultima < 2) return resultado;
  const si = v => ['SI', 'SÍ', 'X', 'TRUE', 'VERDADERO'].includes(String(v).trim().toUpperCase());
  hoja.getRange(2, 1, ultima - 1, 5).getValues().forEach((f, i) => {
    const nombre = normalizar_(f[0]);
    if (!nombre) return;
    if (resultado.columnas[nombre]) {
      resultado.errores.push(`Permisos, fila ${i + 2}: «${nombre}» está repetida; se usa la primera.`);
      return;
    }
    const tipo = (si(f[1]) ? 'P' : '-') + (si(f[2]) ? 'G' : '-') + (si(f[3]) ? 'C' : '-');
    resultado.columnas[nombre] = { tipo, identifica: si(f[4]) };
  });
  return resultado;
}

function exigirAdmin_(ss) {
  const u = leerUsuarios_(ss);
  const correo = Session.getActiveUser().getEmail().toLowerCase();
  if (!u.admins.has(correo)) {
    SpreadsheetApp.getUi().alert(`Esta opción es solo para ADMIN. Su cuenta (${correo || 'no identificada'}) ` +
      'no figura como ADMIN en la hoja Usuarios ni es dueña del archivo.');
    return null;
  }
  return u;
}

// ---------------------------------------------------------------------------
// Estructura de las hojas
// ---------------------------------------------------------------------------

function esHojaPrograma_(hoja) {
  const nombre = hoja.getName();
  if (HOJAS_ESPECIALES.includes(nombre) || nombre.startsWith('_')) return false;
  if (hoja.getMaxRows() < CONFIG.FILA_ENCABEZADO) return false;
  return normalizar_(hoja.getRange(CONFIG.FILA_ENCABEZADO, 1).getValue()) === CONFIG.COL_PROGRAMA;
}

function hojasPrograma_(ss) {
  return ss.getSheets().filter(esHojaPrograma_);
}

/**
 * Lee la fila de encabezados. tipo de cada columna: 'sistema', o «PGC» según la hoja Permisos
 * ('---' = solo ADMIN, también para columnas que no aparecen en Permisos).
 */
function columnas_(hoja, permisos) {
  const total = hoja.getLastColumn();
  const encabezado = hoja.getRange(CONFIG.FILA_ENCABEZADO, 1, 1, total);
  const nombres = encabezado.getValues()[0].map(normalizar_);
  const fondos = permisos.sinHoja ? encabezado.getBackgrounds()[0].map(c => String(c).toLowerCase()) : [];
  let ancho = 0;
  nombres.forEach((nombre, i) => { if (nombre) ancho = i + 1; });
  const sistema = [CONFIG.COL_PROGRAMA, CONFIG.COL_REGISTRADO, CONFIG.COL_FECHA];
  const desconocidas = [];
  const claves = [];
  const tipos = nombres.slice(0, ancho).map((nombre, i) => {
    if (!nombre) return null;
    if (sistema.includes(nombre)) return 'sistema';
    if (permisos.sinHoja) return COLORES_PERMISO[fondos[i]] || SIN_EDITORES;
    const permiso = permisos.columnas[nombre];
    if (!permiso) {
      desconocidas.push(nombre);
      return SIN_EDITORES;
    }
    if (permiso.identifica) claves.push(i + 1);
    return permiso.tipo;
  });
  const indice = {};
  nombres.slice(0, ancho).forEach((nombre, i) => { if (nombre) indice[nombre] = i + 1; });
  [CONFIG.COL_REGISTRADO, CONFIG.COL_FECHA].forEach(c => {
    if (!indice[c]) throw new Error(`La hoja «${hoja.getName()}» no tiene la columna «${c}»`);
  });
  // Sin columnas marcadas como «Identifica el curso», cualquier columna del programa identifica la fila.
  const identifican = claves.length ? claves : tipos.map((t, i) => (esDelPrograma_(t) ? i + 1 : 0)).filter(Boolean);
  return { nombres, tipos, indice, ancho, identifican, desconocidas };
}

/** Bloques contiguos [colIni, colFin] (base 1) de un tipo de columna. */
function bloques_(tipos, tipo) {
  const bloques = [];
  tipos.forEach((t, i) => {
    const c = i + 1;
    if (t !== tipo) return;
    const ultimo = bloques[bloques.length - 1];
    if (ultimo && ultimo[1] === c - 1) ultimo[1] = c;
    else bloques.push([c, c]);
  });
  return bloques;
}

function rangoFila_(hoja, cols, fila) {
  return hoja.getRange(fila, 1, 1, cols.ancho);
}

function tipoProteccion_(p) {
  const d = p.getDescription() || '';
  return d.startsWith(CONFIG.PREFIJO) ? d.slice(CONFIG.PREFIJO.length).split('|')[0] : null;
}

function restaurarFormulaPrograma_(hoja, cols, r1, r2) {
  const col = cols.indice[CONFIG.COL_PROGRAMA];
  if (!col) return;
  const claves = cols.identifican.map(letraColumna_);
  const ini = Math.max(r1, CONFIG.PRIMERA_FILA);
  if (r2 < ini) return;
  const programa = hoja.getName().replace(/"/g, '""');
  // El separador de argumentos depende de la configuración regional del libro («;» en España o Colombia):
  // se escribe con «,» y, si Sheets no puede leerla, se reescribe con «;».
  for (const sep of [',', ';']) {
    const formulas = [];
    for (let r = ini; r <= r2; r++) {
      formulas.push([`=IF(${claves.map(l => l + r).join('&')}=""${sep}""${sep}"${programa}")`]);
    }
    hoja.getRange(ini, col, formulas.length, 1).setFormulas(formulas);
    SpreadsheetApp.flush();
    if (!String(hoja.getRange(ini, col).getDisplayValue()).startsWith('#')) return;
  }
}

/** Escribe la fórmula del Consolidado. Devuelve null si quedó funcionando, o el problema para mostrar. */
function escribirConsolidado_(ss, programas) {
  const hoja = ss.getSheetByName(CONFIG.HOJA_CONSOLIDADO);
  if (!hoja) throw new Error(`No existe la hoja «${CONFIG.HOJA_CONSOLIDADO}»`);
  const cols = columnas_(programas[0], leerPermisos_(ss));
  const fila = CONFIG.PRIMERA_FILA;
  const ultima = letraColumna_(cols.ancho);
  const claves = cols.identifican.map(letraColumna_);
  const ref = h => `'${h.getName().replace(/'/g, "''")}'`;
  const datos = programas.map(h => `${ref(h)}!A${fila}:${ultima}`);
  const condiciones = programas.map(h => `(${claves.map(l => `${ref(h)}!${l}${fila}:${l}`).join('&')})<>""`);
  // Espacio suficiente para que quepan todas las filas de todos los programas.
  const necesarias = fila - 1 + programas.reduce((s, h) => s + h.getMaxRows() - fila + 1, 0);
  if (hoja.getMaxRows() < necesarias) hoja.insertRowsAfter(hoja.getMaxRows(), necesarias - hoja.getMaxRows());
  hoja.getRange(fila, 1, hoja.getMaxRows() - fila + 1, cols.ancho).clearContent();
  hoja.getRange(2, 1).setValue('Se actualiza automáticamente con lo registrado en cada hoja de programa (solo lectura). ' +
    'Para filtrar, use Datos → Vistas de filtro.');
  const colFecha = cols.indice[CONFIG.COL_FECHA];
  hoja.getRange(fila, colFecha, hoja.getMaxRows() - fila + 1, 1).setNumberFormat('dd/mm/yyyy hh:mm');

  // La sintaxis que acepta Sheets depende de la configuración regional del libro (con coma decimal,
  // como España o Colombia, los argumentos se separan con «;») y de si la cuenta tiene VSTACK.
  // Se prueba cada variante y se deja la primera que Sheets calcula sin error.
  const apilarVstack = sep => partes => `VSTACK(${partes.join(sep)})`;
  const apilarLlaves = partes => `{${partes.join(';')}}`;  // «;» apila filas en todas las configuraciones
  const variantes = [
    { sep: ',', apilar: apilarVstack(',') },
    { sep: ';', apilar: apilarVstack(';') },
    { sep: ',', apilar: apilarLlaves },
    { sep: ';', apilar: apilarLlaves },
  ];
  const celda = hoja.getRange(fila, 1);
  let error = '';
  for (const v of variantes) {
    celda.setFormula(`=ARRAYFORMULA(IFNA(FILTER(${v.apilar(datos)}${v.sep}${v.apilar(condiciones)})${v.sep}"Sin registros"))`);
    SpreadsheetApp.flush();
    const valor = String(celda.getDisplayValue());
    if (!valor.startsWith('#')) return null;
    error = valor;
  }
  return error === '#REF!'
    ? 'El Consolidado no puede mostrar los datos porque hay algo escrito debajo de A4 en esa hoja: borre esas celdas ' +
      'y use Aplicar permisos ahora.'
    : `La fórmula del Consolidado da error (${error}) con todas las variantes probadas. Avise a quien mantiene la app.`;
}

function hojaAuditoria_(ss) {
  let hoja = ss.getSheetByName(CONFIG.HOJA_AUDITORIA);
  if (!hoja) {
    hoja = ss.insertSheet(CONFIG.HOJA_AUDITORIA);
    hoja.appendRow(['Fecha', 'Usuario', 'Hoja', 'Rango', 'Motivo', 'Revertido', 'Valor nuevo', 'Valor anterior']);
    hoja.getRange(1, 1, 1, 8).setFontWeight('bold').setBackground('#d9d9d9');
    hoja.setFrozenRows(1);
    hoja.hideSheet();
    // Protegida desde el inicio; la próxima sincronización agrega a los ADMIN.
    const p = hoja.protect().setDescription(CONFIG.PREFIJO + 'hoja|' + CONFIG.HOJA_AUDITORIA);
    ajustarEditores_(p, [], nuevoResumen_());
  }
  return hoja;
}

function registrarAuditoria_(ss, correo, hoja, rango, motivo, revertido, nuevo, anterior) {
  hojaAuditoria_(ss).appendRow([new Date(), correo || '(no identificado)', hoja, rango, motivo, revertido,
    nuevo === undefined ? '' : String(nuevo), anterior === undefined ? '' : String(anterior)]);
}

// ---------------------------------------------------------------------------
// Utilidades
// ---------------------------------------------------------------------------

function correoDe_(e) {
  return e && e.user ? String(e.user.getEmail() || '').trim().toLowerCase() : '';
}

function correoEfectivo_() {
  return Session.getEffectiveUser().getEmail().toLowerCase();
}

function correoDueno_() {
  const dueno = SpreadsheetApp.getActiveSpreadsheet().getOwner(); // null en unidades compartidas
  return dueno ? dueno.getEmail().toLowerCase() : '';
}

function normalizar_(v) {
  return String(v).replace(/\s+/g, ' ').trim();
}

/** Para comparar nombres sin importar mayúsculas, espacios ni tildes: «  Afín » → «afin». */
function claveTexto_(v) {
  return normalizar_(v).toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
}

function letraColumna_(n) {
  let s = '';
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}
