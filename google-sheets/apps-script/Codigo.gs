/**
 * Cronograma 2027 Bloque Neón — permisos por rol y propiedad de filas en Google Sheets.
 *
 * Roles (hoja «Usuarios»):
 *   ADMIN     todo el libro; administra permisos.
 *   CASA      columnas verdes de todas las hojas de programa.
 *   GESTOR    columnas azules de su(s) programa(s), incluidas las filas de sus profesores.
 *   PROFESOR  columnas azules de su(s) programa(s); cada fila que registra queda solo para quien la registró.
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
 * @uniandes.edu.co, idealmente la cuenta dueña del archivo. Ver google-sheets/README.md.
 */

const CONFIG = {
  DOMINIO: 'uniandes.edu.co',
  HOJA_USUARIOS: 'Usuarios',
  HOJA_CONSOLIDADO: 'Consolidado',
  HOJA_INSTRUCCIONES: 'Instrucciones',
  HOJA_AUDITORIA: '_Auditoria',
  FILA_ENCABEZADO: 3,
  PRIMERA_FILA: 4,
  COLOR_AZUL: '#00b0f0',
  COL_PROGRAMA: 'Programa',
  COL_REGISTRADO: 'Registrado por',
  COL_FECHA: 'Fecha de registro',
  COLUMNAS_CLAVE: ['Código de curso', 'CRN', 'Curso'],
  PREFIJO: '[CASA] ',
  GESTOR_EDITA_FILAS_DE_PROFESORES: true,
  ALERTAS_POR_CORREO: true,
  HORA_RESTAURACION: 2,
  LIMITE_MS: 5 * 60 * 1000, // Apps Script corta a los 6 minutos
};

const ROLES = ['ADMIN', 'CASA', 'GESTOR', 'PROFESOR'];
const HOJAS_ESPECIALES = [CONFIG.HOJA_USUARIOS, CONFIG.HOJA_CONSOLIDADO, CONFIG.HOJA_INSTRUCCIONES, CONFIG.HOJA_AUDITORIA];
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
  escribirConsolidado_(ss, programas);
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
  ui.alert(`Acceso otorgado a ${nuevos.length - fallos.length} correo(s).` +
    (fallos.length ? `\n\nNo se pudo con:\n${fallos.join('\n')}` : '') +
    '\n\nGoogle puede enviar una notificación al correo; además comparta el enlace por Teams o Outlook.');
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
    if (esAdmin && nombre === CONFIG.HOJA_USUARIOS) {
      ss.toast('Cuando termine de editar «Usuarios», use CASA → Aplicar permisos ahora.', 'CASA', 8);
    } else if (!esAdmin && HOJAS_ESPECIALES.includes(nombre)) {
      incidente_(e, u, correo, null, 'Edición en hoja restringida');
    }
    return;
  }

  const cols = columnas_(hoja);
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
  if (r2 >= CONFIG.PRIMERA_FILA && tipos.has('azul')) {
    reclamarOLiberarFilas_(hoja, cols, Math.max(r1, CONFIG.PRIMERA_FILA), r2, correo, esAdmin, u);
  }
}

/** Devuelve por qué la edición no está permitida para ese correo, o null si lo está. */
function motivoRechazo_(u, correo, hoja, cols, r1, r2, tipos) {
  const roles = rolesDe_(u, correo, hoja.getName());
  if (!correo) return 'Usuario no identificado (cuenta fuera de @' + CONFIG.DOMINIO + ')';
  if (r1 < CONFIG.PRIMERA_FILA) return 'Edición de títulos o encabezados';
  if (tipos.has('sistema') || tipos.has('fuera')) return 'Edición de columnas del sistema o fuera del área de registro';
  if (tipos.has('verde') && !roles.casa) return 'Columna CASA sin rol CASA';
  if (tipos.has('azul')) {
    if (!roles.gestor && !roles.profesor) return 'Columna de programa sin rol en ' + hoja.getName();
    if (!(roles.gestor && CONFIG.GESTOR_EDITA_FILAS_DE_PROFESORES)) {
      const duenos = hoja.getRange(r1, cols.indice[CONFIG.COL_REGISTRADO], r2 - r1 + 1, 1).getValues();
      const ajena = duenos.some(d => {
        const dueno = String(d[0]).trim().toLowerCase();
        return dueno && dueno !== correo;
      });
      if (ajena) return 'Fila registrada por otra persona';
    }
  }
  return null;
}

/**
 * Fila libre con datos azules → queda a nombre de quien la editó y se protege.
 * Fila registrada sin datos azules → se libera.
 */
function reclamarOLiberarFilas_(hoja, cols, r1, r2, correo, esAdmin, u) {
  const n = r2 - r1 + 1;
  const valores = hoja.getRange(r1, 1, n, cols.ancho).getValues();
  const colReg = cols.indice[CONFIG.COL_REGISTRADO];
  const colFecha = cols.indice[CONFIG.COL_FECHA];
  const azules = cols.tipos.map((t, i) => (t === 'azul' ? i : -1)).filter(i => i >= 0);

  valores.forEach((fila, i) => {
    const r = r1 + i;
    const dueno = String(fila[colReg - 1]).trim().toLowerCase();
    const tieneDatos = azules.some(c => fila[c] !== '' && fila[c] !== null);
    if (!dueno && tieneDatos && !esAdmin && correo) {
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
  const nombres = new Set(programas.map(h => h.getName()));
  Object.keys(u.programas).filter(p => !nombres.has(p))
    .forEach(p => u.errores.push(`Programa «${p}» de la hoja Usuarios no existe como hoja`));

  try {
    HOJAS_ESPECIALES.forEach(nombre => {
      const hoja = ss.getSheetByName(nombre);
      if (hoja) protegerHoja_(hoja, [...u.admins], [], r);
    });
    programas.forEach(hoja => sincronizarHoja_(ss, hoja, u, r));
  } catch (err) {
    if (err.message !== 'TIEMPO') throw err;
    r.completo = false;
  }
  return r;
}

function sincronizarHoja_(ss, hoja, u, r) {
  const nombre = hoja.getName();
  const cols = columnas_(hoja);
  const n = hoja.getMaxRows() - CONFIG.PRIMERA_FILA + 1;
  const g = u.programas[nombre] || { gestores: new Set(), profesores: new Set() };
  const admins = [...u.admins];
  const deseadas = [];
  const desprotegidos = [];

  bloques_(cols.tipos, 'azul').forEach(([a, b]) => {
    const rango = hoja.getRange(CONFIG.PRIMERA_FILA, a, n, b - a + 1);
    desprotegidos.push(rango);
    deseadas.push({ tipo: 'azul', detalle: nombre, rango, editores: [...admins, ...g.gestores, ...g.profesores] });
  });
  bloques_(cols.tipos, 'verde').forEach(([a, b]) => {
    const rango = hoja.getRange(CONFIG.PRIMERA_FILA, a, n, b - a + 1);
    desprotegidos.push(rango);
    deseadas.push({ tipo: 'verde', detalle: nombre, rango, editores: [...admins, ...u.casa] });
  });
  const registros = hoja.getRange(CONFIG.PRIMERA_FILA, cols.indice[CONFIG.COL_REGISTRADO], n, 1).getValues();
  registros.forEach((v, i) => {
    const dueno = String(v[0]).trim().toLowerCase();
    if (!dueno) return;
    const fila = CONFIG.PRIMERA_FILA + i;
    deseadas.push({ tipo: 'fila', detalle: dueno, rango: rangoFila_(hoja, cols, fila), editores: editoresFila_(u, nombre, dueno) });
  });

  // Toda la hoja queda solo para ADMIN, salvo las zonas azul y verde, que tienen su propia protección.
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
  const lineas = [
    `Protecciones creadas: ${r.creadas} · actualizadas: ${r.actualizadas} · sin cambios: ${r.sinCambios} · eliminadas: ${r.eliminadas}`,
  ];
  if (!r.completo) lineas.push('\n⚠️ Se alcanzó el límite de tiempo. Ejecute «Aplicar permisos ahora» de nuevo para terminar.');
  if (u && u.errores.length) lineas.push('\nRevise la hoja Usuarios:\n' + u.errores.join('\n'));
  if (u && u.externos.length) {
    lineas.push(`\nCorreos fuera de @${CONFIG.DOMINIO} (el script no puede identificarlos al editar, ` +
      `así que sus ediciones se revertirán):\n` + u.externos.join('\n'));
  }
  if (r.fallos.length) lineas.push('\nAvisos:\n' + r.fallos.slice(0, 20).join('\n'));
  return lineas.join('\n');
}

// ---------------------------------------------------------------------------
// Usuarios y roles
// ---------------------------------------------------------------------------

function leerUsuarios_(ss) {
  const hoja = ss.getSheetByName(CONFIG.HOJA_USUARIOS);
  if (!hoja) throw new Error(`No existe la hoja «${CONFIG.HOJA_USUARIOS}»`);
  const u = { admins: new Set(), casa: new Set(), programas: {}, todos: [], externos: [], errores: [] };
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
      programas.forEach(p => {
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
  const lista = [...u.admins, ...u.casa]; // CASA sigue bloqueada en azul por la protección de la zona azul
  if (CONFIG.GESTOR_EDITA_FILAS_DE_PROFESORES) lista.push(...g.gestores);
  if (g.profesores.has(dueno) || g.gestores.has(dueno)) lista.push(dueno);
  return lista;
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

/** Lee la fila de encabezados: tipo de cada columna según su color (azul / verde) o si es del sistema. */
function columnas_(hoja) {
  const total = hoja.getLastColumn();
  const rango = hoja.getRange(CONFIG.FILA_ENCABEZADO, 1, 1, total);
  const nombres = rango.getValues()[0].map(normalizar_);
  const fondos = rango.getBackgrounds()[0].map(c => String(c).toLowerCase());
  let ancho = 0;
  nombres.forEach((nombre, i) => { if (nombre) ancho = i + 1; });
  const sistema = [CONFIG.COL_PROGRAMA, CONFIG.COL_REGISTRADO, CONFIG.COL_FECHA];
  const tipos = nombres.slice(0, ancho).map((nombre, i) => {
    if (!nombre) return null;
    if (sistema.includes(nombre)) return 'sistema';
    return fondos[i] === CONFIG.COLOR_AZUL ? 'azul' : 'verde';
  });
  const indice = {};
  nombres.slice(0, ancho).forEach((nombre, i) => { if (nombre) indice[nombre] = i + 1; });
  [CONFIG.COL_REGISTRADO, CONFIG.COL_FECHA].forEach(c => {
    if (!indice[c]) throw new Error(`La hoja «${hoja.getName()}» no tiene la columna «${c}»`);
  });
  return { nombres, tipos, indice, ancho };
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
  const claves = CONFIG.COLUMNAS_CLAVE.map(n => cols.indice[n]).filter(Boolean).map(letraColumna_);
  const ini = Math.max(r1, CONFIG.PRIMERA_FILA);
  if (r2 < ini) return;
  const formulas = [];
  for (let r = ini; r <= r2; r++) {
    formulas.push([`=IF(${claves.map(l => l + r).join('&')}="","","${hoja.getName()}")`]);
  }
  hoja.getRange(ini, col, formulas.length, 1).setFormulas(formulas);
}

function escribirConsolidado_(ss, programas) {
  const hoja = ss.getSheetByName(CONFIG.HOJA_CONSOLIDADO);
  if (!hoja) throw new Error(`No existe la hoja «${CONFIG.HOJA_CONSOLIDADO}»`);
  const cols = columnas_(programas[0]);
  const fila = CONFIG.PRIMERA_FILA;
  const ultima = letraColumna_(cols.ancho);
  const claves = CONFIG.COLUMNAS_CLAVE.map(n => cols.indice[n]).filter(Boolean).map(letraColumna_);
  const ref = h => `'${h.getName().replace(/'/g, "''")}'`;
  const datos = programas.map(h => `${ref(h)}!A${fila}:${ultima}`);
  const condiciones = programas.map(h => `(${claves.map(l => `${ref(h)}!${l}${fila}:${l}`).join('&')})<>""`);
  // Espacio suficiente para que quepan todas las filas de todos los programas.
  const necesarias = fila - 1 + programas.reduce((s, h) => s + h.getMaxRows() - fila + 1, 0);
  if (hoja.getMaxRows() < necesarias) hoja.insertRowsAfter(hoja.getMaxRows(), necesarias - hoja.getMaxRows());
  hoja.getRange(fila, 1, hoja.getMaxRows() - fila + 1, cols.ancho).clearContent();
  hoja.getRange(fila, 1).setFormula(
    `=ARRAYFORMULA(IFNA(FILTER(VSTACK(${datos.join(',')}),VSTACK(${condiciones.join(',')})),"Sin registros"))`);
  hoja.getRange(2, 1).setValue('Se actualiza automáticamente con lo registrado en cada hoja de programa (solo lectura). ' +
    'Para filtrar, use Datos → Vistas de filtro.');
  const colFecha = cols.indice[CONFIG.COL_FECHA];
  hoja.getRange(fila, colFecha, hoja.getMaxRows() - fila + 1, 1).setNumberFormat('dd/mm/yyyy hh:mm');
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

function letraColumna_(n) {
  let s = '';
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}
