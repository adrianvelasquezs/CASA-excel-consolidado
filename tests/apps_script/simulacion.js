// Simulación mínima de Google Sheets para probar apps_script/Codigo.gs sin una cuenta de Google.
// Uso: node simulacion.js Codigo.gs estructura.json  (la prueba tests/test_apps_script.py genera el JSON)
const fs = require('fs');
const vm = require('vm');
const [, , archivoGs, archivoEstructura] = process.argv;
const est = JSON.parse(fs.readFileSync(archivoEstructura, 'utf8'));

const DUENO = 'admin.casa@uniandes.edu.co';
let usuarioActual = DUENO;
const correos = [];
const user = e => ({ getEmail: () => e });
const letra = n => { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };

class Proteccion {
  constructor(hoja, tipo, rango) {
    this.hoja = hoja; this.tipo = tipo; this.rango = rango; this.desc = '';
    this.editores = new Set([usuarioActual]); this.dominio = true; this.aviso = false; this.desprot = [];
  }
  setDescription(d) { this.desc = d; return this; }
  getDescription() { return this.desc; }
  addEditor(e) { this.editores.add(typeof e === 'string' ? e : e.getEmail()); return this; }
  addEditors(l) { l.forEach(e => this.addEditor(e)); return this; }
  removeEditors(l) { l.forEach(e => { const c = typeof e === 'string' ? e : e.getEmail(); if (c !== DUENO) this.editores.delete(c); }); return this; }
  getEditors() { return [...this.editores, DUENO].filter((v, i, a) => a.indexOf(v) === i).map(user); }
  canDomainEdit() { return this.dominio; }
  setDomainEdit(v) { this.dominio = v; return this; }
  isWarningOnly() { return this.aviso; }
  setWarningOnly(v) { this.aviso = v; return this; }
  getRange() { return this.rango; }
  setUnprotectedRanges(l) { this.desprot = l; return this; }
  remove() { this.hoja.protecciones = this.hoja.protecciones.filter(p => p !== this); }
}

class Rango {
  constructor(hoja, r, c, nr, nc) { Object.assign(this, { hoja, r, c, nr: nr || 1, nc: nc || 1 }); }
  celdas(fn) { for (let i = 0; i < this.nr; i++) for (let j = 0; j < this.nc; j++) fn(this.r + i, this.c + j, i, j); }
  getValues() { const o = []; for (let i = 0; i < this.nr; i++) { o.push([]); for (let j = 0; j < this.nc; j++) o[i].push(this.hoja.valor(this.r + i, this.c + j)); } return o; }
  getValue() { return this.hoja.valor(this.r, this.c); }
  // Imita cómo Sheets calcula una fórmula según la configuración regional y las funciones disponibles.
  getDisplayValue() {
    const f = this.hoja.formulas[`${this.r},${this.c}`];
    if (!f) return String(this.hoja.valor(this.r, this.c));
    const sinTextos = f.replace(/"[^"]*"/g, '').replace(/'[^']*'/g, '');
    if (process.env.LOCALE_PUNTO_Y_COMA && sinTextos.includes(',')) return '#ERROR!';
    if (!process.env.LOCALE_PUNTO_Y_COMA && /[A-Z]+\([^{}]*;/.test(sinTextos)) return '#ERROR!';
    if (process.env.SIN_VSTACK && f.includes('VSTACK(')) return '#NAME?';
    return 'Sin registros';
  }
  setValue(v) { this.celdas((r, c) => this.hoja.set(r, c, v)); return this; }
  setFormula(f) { this.hoja.formulas[`${this.r},${this.c}`] = f; return this; }
  setFormulas(m) { this.celdas((r, c, i, j) => { this.hoja.formulas[`${r},${c}`] = m[i][j]; }); return this; }
  clearContent() { this.celdas((r, c) => { this.hoja.set(r, c, ''); delete this.hoja.formulas[`${r},${c}`]; }); return this; }
  getBackgrounds() { const o = [[]]; for (let j = 0; j < this.nc; j++) o[0].push(this.hoja.fondos[this.c + j] || '#ffffff'); return o; }
  getSheet() { return this.hoja; }
  getRow() { return this.r; } getLastRow() { return this.r + this.nr - 1; }
  getColumn() { return this.c; } getLastColumn() { return this.c + this.nc - 1; }
  getNumRows() { return this.nr; } getNumColumns() { return this.nc; }
  getA1Notation() { const a = letra(this.c) + this.r, b = letra(this.c + this.nc - 1) + (this.r + this.nr - 1); return this.nr === 1 && this.nc === 1 ? a : a + ':' + b; }
  contiene(r, c) { return r >= this.r && r < this.r + this.nr && c >= this.c && c < this.c + this.nc; }
  protect() { const p = new Proteccion(this.hoja, 'RANGE', this); this.hoja.protecciones.push(p); return p; }
  setNumberFormat() { return this; } setFontWeight() { return this; } setBackground() { return this; }
  setDataValidation(regla) { this.hoja.validaciones.push({ rango: this.getA1Notation(), regla }); return this; }
}

class Hoja {
  get validaciones() { return this._validaciones || (this._validaciones = []); }
  constructor(nombre, maxFilas, maxCols) { Object.assign(this, { nombre, maxFilas, maxCols, datos: {}, fondos: {}, formulas: {}, protecciones: [], oculta: false }); }
  valor(r, c) { const v = this.datos[`${r},${c}`]; return v === undefined ? '' : v; }
  set(r, c, v) { this.datos[`${r},${c}`] = v; }
  getName() { return this.nombre; }
  setName(n) { this.nombre = n; return this; }
  getIndex() { return hojas.indexOf(this) + 1; }
  // Copia contenido, formatos y fórmulas; como en Google, las protecciones no se copian.
  copyTo() {
    const h = new Hoja('Copia de ' + this.nombre, this.maxFilas, this.maxCols);
    Object.assign(h.datos, this.datos); Object.assign(h.fondos, this.fondos); Object.assign(h.formulas, this.formulas);
    hojas.push(h);
    return h;
  }
  getRange(r, c, nr, nc) { return new Rango(this, r, c, nr, nc); }
  getMaxRows() { return this.maxFilas; }
  getLastRow() { let m = 0; Object.keys(this.datos).forEach(k => { const [r] = k.split(',').map(Number); if (this.datos[k] !== '') m = Math.max(m, r); }); return m; }
  getLastColumn() { let m = 0; Object.keys(this.datos).forEach(k => { const [, c] = k.split(',').map(Number); if (this.datos[k] !== '') m = Math.max(m, c); }); return m; }
  getProtections(t) { return this.protecciones.filter(p => p.tipo === t); }
  protect() { const p = new Proteccion(this, 'SHEET', new Rango(this, 1, 1, this.maxFilas, this.maxCols)); this.protecciones.push(p); return p; }
  appendRow(v) { const r = this.getLastRow() + 1; v.forEach((x, j) => this.set(r, j + 1, x)); }
  insertRowsAfter(_, n) { this.maxFilas += n; }
  setFrozenRows() {} hideSheet() { this.oculta = true; } showSheet() { this.oculta = false; }
  // Regla de Google asumida: hay que ser editor de TODAS las protecciones que cubren la celda.
  puedeEditar(correo, r, c) {
    if (correo === DUENO) return true;
    return this.protecciones.every(p => {
      if (p.tipo === 'RANGE' && !p.rango.contiene(r, c)) return true;
      if (p.tipo === 'SHEET' && p.desprot.some(d => d.contiene(r, c))) return true;
      return p.editores.has(correo);
    });
  }
}

const hojas = [];
Object.entries(est).forEach(([nombre, info]) => {
  const h = new Hoja(nombre, info.max_row, 40);
  if (info.header.some(([v]) => v === 'Programa')) {
    info.header.forEach(([v, color], j) => { if (v) { h.set(3, j + 1, v); h.fondos[j + 1] = color ? '#' + color.slice(2).toLowerCase() : '#ffffff'; } });
  }
  if (nombre === 'Usuarios') info.usuarios_hdr.forEach((v, j) => h.set(1, j + 1, v));
  if (info.filas) info.filas.forEach((fila, i) => fila.forEach((v, j) => { if (v !== '') h.set(i + 1, j + 1, v); }));
  hojas.push(h);
});
const colaboradores = new Set([DUENO]);  // personas con acceso de edición al archivo
const ss = {
  getSheets: () => hojas,
  getSheetByName: n => hojas.find(h => h.nombre === n) || null,
  insertSheet: n => { const h = new Hoja(n, 1000, 26); hojas.push(h); return h; },
  setActiveSheet: h => { ss.activa = h; return h; },
  deleteSheet: h => { hojas.splice(hojas.indexOf(h), 1); },
  moveActiveSheet: pos => { hojas.splice(hojas.indexOf(ss.activa), 1); hojas.splice(pos - 1, 0, ss.activa); },
  getOwner: () => user(DUENO),
  getEditors: () => [...colaboradores].map(user),
  getViewers: () => [],
  addEditor: c => colaboradores.add(c), toast: () => {}, getUrl: () => 'https://docs.google.com/spreadsheets/d/x',
  getId: () => 'x',
};
const alertas = [];
let respuestaPrompt = '';  // lo que «escribe» el usuario en los cuadros de diálogo
const respuestasAlerta = [];  // botones que «pulsa» el usuario en los avisos (por defecto, Sí)
const props = {};
const ctx = {
  console: { log() {}, warn: m => console.log('   [warn]', m), error: m => console.log('   [error]', m) },
  SpreadsheetApp: {
    getActiveSpreadsheet: () => ss,
    ProtectionType: { RANGE: 'RANGE', SHEET: 'SHEET' },
    flush() {},
    newDataValidation: () => {
      const regla = { lista: null };
      const b = { requireValueInList: l => { regla.lista = l; return b; }, setAllowInvalid: () => b, setHelpText: () => b, build: () => regla };
      return b;
    },
    getUi: () => ({ prompt: () => ({ getSelectedButton: () => 'OK', getResponseText: () => respuestaPrompt }), alert: (...a) => { alertas.push(a.filter(x => typeof x === 'string').join(' | ')); return respuestasAlerta.length ? respuestasAlerta.shift() : 'YES'; }, ButtonSet: { OK: 1, YES_NO: 2, OK_CANCEL: 3, YES_NO_CANCEL: 4 }, Button: { YES: 'YES', OK: 'OK', NO: 'NO', CANCEL: 'CANCEL' },
      createMenu: () => ({ addItem() { return this; }, addSeparator() { return this; }, addToUi() {} }) }),
  },
  Session: { getActiveUser: () => user(usuarioActual), getEffectiveUser: () => user(DUENO) },
  LockService: { getDocumentLock: () => ({ tryLock: () => true, releaseLock() {} }) },
  PropertiesService: { getDocumentProperties: () => ({ getProperty: k => props[k] || null, setProperty: (k, v) => { props[k] = v; } }) },
  ScriptApp: { getProjectTriggers: () => [], deleteTrigger() {}, newTrigger: () => { const b = { forSpreadsheet: () => b, onEdit: () => b, timeBased: () => b, everyDays: () => b, atHour: () => b, create: () => b }; return b; } },
  MailApp: { sendEmail: m => correos.push(m) },
  Date, Math, String, Object, Set, Array, Error,
};
vm.createContext(ctx);
vm.runInContext(fs.readFileSync(archivoGs, 'utf8'), ctx);

// ---------------------------------------------------------------- utilidades de prueba
let fallos = 0;
const ok = (cond, msg) => { console.log((cond ? '  ✔ ' : '  ✘ ') + msg); if (!cond) fallos++; };
const H = n => ss.getSheetByName(n);
const col = (hoja, nombre) => { for (let c = 1; c < 40; c++) if (H(hoja).valor(3, c) === nombre) return c; };
function editar(correo, hoja, r, c, valor, { saltarProteccion = false } = {}) {
  const h = H(hoja);
  if (!saltarProteccion && !h.puedeEditar(correo, r, c)) return 'BLOQUEADO';
  const anterior = h.valor(r, c);
  h.set(r, c, valor);
  const e = { source: ss, range: h.getRange(r, c), user: user(correo.endsWith('@uniandes.edu.co') ? correo : ''), value: valor, oldValue: anterior === '' ? undefined : anterior };
  ctx.__e = e;
  vm.runInContext('alEditar(__e)', ctx);
  return 'EDITADO';
}

// ---------------------------------------------------------------- datos
const U = H('Usuarios');
[
  ['coordinadora@uniandes.edu.co', 'ADMIN', ''],
  ['analista.casa@uniandes.edu.co', 'CASA', ''],
  ['gestor.afin@uniandes.edu.co', 'GESTOR', 'AFIN'],
  ['prof.a@uniandes.edu.co', 'PROFESOR', 'AFIN'],
  ['prof.b@uniandes.edu.co', 'PROFESOR', 'AFIN, MF'],
  ['prof.mf@uniandes.edu.co', 'PROFESOR', 'MF'],
  ['externo@gmail.com', 'PROFESOR', 'MF'],
  ['mal@uniandes.edu.co', 'JEFE', ''],
  ['gestora.min@uniandes.edu.co', 'GESTOR', ' afín '],      // minúsculas, tilde y espacios
  ['gestora.mal@uniandes.edu.co', 'GESTOR', 'AFIN Finanzas'], // no coincide con ninguna hoja
].forEach((f, i) => f.forEach((v, j) => U.set(i + 2, j + 1, v)));

const B = col('AFIN', 'Código de curso'), D = col('AFIN', 'Curso'), L = col('AFIN', 'Número de Competencia a evaluar');
const REG = col('AFIN', 'Registrado por'), PROG = col('AFIN', 'Programa');
const REPORTE = col('AFIN', '¿Se entregó reporte?'), RAZONES = col('AFIN', 'Razones de porqué no se realizó medición');
const ULTIMA = letra(est.AFIN.header.filter(([v]) => v).length);

if (process.env.SIN_PERMISOS) {
  // Libro generado con una versión anterior (sin hoja Permisos): los permisos salen de los colores.
  console.log('\nLibro sin hoja Permisos (respaldo por colores)');
  usuarioActual = DUENO;
  vm.runInContext('configurarLibro()', ctx);
  ok(alertas[alertas.length - 1].includes('se usan los colores de los encabezados'), 'Avisa que usa los colores de los encabezados');
  ok(editar('gestor.afin@uniandes.edu.co', 'AFIN', 10, B, 'X') === 'EDITADO' && H('AFIN').valor(10, B) === 'X', 'Gestor edita columna azul');
  ok(editar('prof.a@uniandes.edu.co', 'AFIN', 11, B, 'Y') === 'EDITADO' && H('AFIN').valor(11, B) === 'Y', 'Profesor edita columna azul');
  ok(editar('analista.casa@uniandes.edu.co', 'AFIN', 10, L, 'C1') === 'EDITADO' && H('AFIN').valor(10, L) === 'C1', 'CASA edita columna verde');
  ok(editar('prof.a@uniandes.edu.co', 'AFIN', 11, L, 'x') === 'BLOQUEADO', 'Profesor no edita columna verde');
  ok(editar('analista.casa@uniandes.edu.co', 'AFIN', 12, B, 'x') === 'BLOQUEADO', 'CASA no edita columna azul');
  console.log(fallos ? `\n${fallos} prueba(s) fallaron` : '\nTodas las pruebas pasaron');
  process.exit(fallos ? 1 : 0);
}

console.log('\n1) Configurar libro (dueño)');
usuarioActual = DUENO;
vm.runInContext('configurarLibro()', ctx);
const pAFIN = H('AFIN').protecciones;
ok(pAFIN.filter(p => p.tipo === 'SHEET').length === 1, 'AFIN: 1 protección de hoja');
ok(pAFIN.filter(p => p.desc.includes('colPG-|')).length === 3, 'AFIN: 3 bloques de programa (B:K, M, W) → ' + pAFIN.filter(p => p.desc.includes('colPG-|')).map(p => p.rango.getA1Notation()).join(' '));
ok(pAFIN.filter(p => p.desc.includes('col--C|')).length === 3, 'AFIN: 3 bloques de CASA → ' + pAFIN.filter(p => p.desc.includes('col--C|')).map(p => p.rango.getA1Notation()).join(' '));
ok(H('Permisos').protecciones.length === 1, 'Hoja Permisos protegida (solo ADMIN)');
ok(H('Usuarios').protecciones.length === 1 && H('Consolidado').protecciones.length === 1, 'Usuarios y Consolidado protegidas');
ok(editar('prof.a@uniandes.edu.co', 'Permisos', 2, 4, 'SI') === 'BLOQUEADO', 'Profesor no edita la hoja Permisos');
const formula = H('Consolidado').formulas['4,1'] || '';
ok(formula.startsWith('=ARRAYFORMULA(IFNA(FILTER(') && formula.includes(`'Pregrado-CEMS'!A4:${ULTIMA}`), 'Fórmula del Consolidado (' + formula.length + ' caracteres)');
console.log('   ' + formula.slice(0, 160) + '…');
const separador = process.env.LOCALE_PUNTO_Y_COMA ? ';' : ',';
ok(formula.endsWith(`)${separador}"Sin registros"))`) && formula.includes(process.env.SIN_VSTACK ? 'FILTER({' : 'FILTER(VSTACK('),
   `Consolidado: usa la variante que Sheets acepta (separador «${separador}», ${process.env.SIN_VSTACK ? 'llaves' : 'VSTACK'})`);
console.log('   alerta: ' + alertas[alertas.length - 1].replace(/\n/g, '\n           '));
ok(alertas[alertas.length - 1].includes('externo@gmail.com') && alertas[alertas.length - 1].includes('JEFE'), 'Avisa correo externo y rol inválido');

console.log('\n2) Permisos en la interfaz (protecciones)');
ok(editar('prof.a@uniandes.edu.co', 'AFIN', 3, D, 'X') === 'BLOQUEADO', 'Profesor no edita encabezados');
ok(editar('prof.a@uniandes.edu.co', 'AFIN', 10, PROG, 'X') === 'BLOQUEADO', 'Profesor no edita columna Programa');
ok(editar('prof.a@uniandes.edu.co', 'AFIN', 10, L, 'X') === 'BLOQUEADO', 'Profesor no edita columna verde');
ok(editar('analista.casa@uniandes.edu.co', 'AFIN', 10, B, 'X') === 'BLOQUEADO', 'CASA no edita columna azul');
ok(editar('prof.mf@uniandes.edu.co', 'AFIN', 10, B, 'X') === 'BLOQUEADO', 'Profesor de MF no edita AFIN');
ok(editar('prof.a@uniandes.edu.co', 'MF', 10, B, 'X') === 'BLOQUEADO', 'Profesor de AFIN no edita MF');
ok(editar('prof.a@uniandes.edu.co', 'Usuarios', 20, 1, 'yo@uniandes.edu.co') === 'BLOQUEADO', 'Profesor no edita Usuarios');

console.log('\n3) Propiedad de filas');
ok(editar('prof.a@uniandes.edu.co', 'AFIN', 5, B, 'AFIN-1001') === 'EDITADO', 'Profesor A registra un curso en la fila 5');
ok(H('AFIN').valor(5, REG) === 'prof.a@uniandes.edu.co', 'Fila 5 queda a nombre de A');
const pf = H('AFIN').protecciones.find(p => p.desc.startsWith('[CASA] fila|'));
ok(pf && pf.rango.getA1Notation() === `A5:${ULTIMA}5`, `Fila 5 protegida (A5:${ULTIMA}5)`);
ok(editar('prof.a@uniandes.edu.co', 'AFIN', 5, D, 'Finanzas') === 'EDITADO', 'A sigue editando su fila');
ok(editar('prof.b@uniandes.edu.co', 'AFIN', 5, D, 'Otro') === 'BLOQUEADO', 'Profesor B no edita la fila de A');
ok(editar('gestor.afin@uniandes.edu.co', 'AFIN', 5, D, 'Finanzas I') === 'EDITADO', 'Gestor de AFIN sí edita la fila de A');
ok(editar('analista.casa@uniandes.edu.co', 'AFIN', 5, L, 'C1') === 'EDITADO', 'CASA edita la columna verde de la fila de A');
ok(editar('analista.casa@uniandes.edu.co', 'AFIN', 5, B, 'X') === 'BLOQUEADO', 'CASA sigue sin editar azul en la fila de A');
ok(REPORTE && RAZONES, 'Las columnas nuevas existen en la hoja');
ok(editar('analista.casa@uniandes.edu.co', 'AFIN', 5, REPORTE, 'SI') === 'EDITADO', 'CASA llena «¿Se entregó reporte?» en la fila de A');
ok(editar('analista.casa@uniandes.edu.co', 'AFIN', 9, RAZONES, 'Sin estudiantes') === 'EDITADO', 'CASA llena razones de no medición');
ok(editar('prof.a@uniandes.edu.co', 'AFIN', 5, REPORTE, 'NO') === 'BLOQUEADO', 'Profesor no edita «¿Se entregó reporte?»');
ok(editar('gestor.afin@uniandes.edu.co', 'AFIN', 5, RAZONES, 'x') === 'BLOQUEADO', 'Gestor no edita las razones (CASA)');
ok(editar('prof.a@uniandes.edu.co', 'AFIN', 7, REPORTE, 'NO', { saltarProteccion: true }) && H('AFIN').valor(7, REPORTE) === '', 'Si un profesor se salta la protección, se revierte');
ok(editar('prof.b@uniandes.edu.co', 'AFIN', 6, B, 'AFIN-2002') === 'EDITADO' && H('AFIN').valor(6, REG) === 'prof.b@uniandes.edu.co', 'B registra la fila 6');
ok(editar('prof.b@uniandes.edu.co', 'MF', 4, B, 'MF-1') === 'EDITADO', 'B (AFIN y MF) registra en MF');

console.log('\n4) Liberar fila');
editar('prof.b@uniandes.edu.co', 'AFIN', 6, B, '');
ok(H('AFIN').valor(6, REG) === '' && !H('AFIN').protecciones.some(p => p.desc.startsWith('[CASA] fila|') && p.rango.r === 6), 'Al vaciar su fila, B la libera');
ok(editar('prof.a@uniandes.edu.co', 'AFIN', 6, B, 'AFIN-3003') === 'EDITADO' && H('AFIN').valor(6, REG) === 'prof.a@uniandes.edu.co', 'Otra persona puede tomarla');

console.log('\n5) Red de seguridad (alguien quitó protecciones)');
const antes = correos.length;
ok(editar('prof.b@uniandes.edu.co', 'AFIN', 5, D, 'SABOTAJE', { saltarProteccion: true }) === 'EDITADO', 'B logra editar la fila de A');
ok(H('AFIN').valor(5, D) === 'Finanzas I', 'La edición se revierte al valor anterior');
ok(editar('prof.mf@uniandes.edu.co', 'AFIN', 20, B, 'INTRUSO', { saltarProteccion: true }) && H('AFIN').valor(20, B) === '' && H('AFIN').valor(20, REG) === '', 'Profesor de otro programa: se revierte y no reclama la fila');
editar('prof.a@uniandes.edu.co', 'AFIN', 8, PROG, 'X', { saltarProteccion: true });
ok((H('AFIN').formulas['8,' + PROG] || '').startsWith('=IF(B8&C8&D8'), 'Columna Programa: se restaura la fórmula');
editar('prof.a@uniandes.edu.co', 'Consolidado', 10, 3, 'X', { saltarProteccion: true });
ok(H('Consolidado').valor(10, 3) === '' && (H('Consolidado').formulas['4,1'] || '').startsWith('=ARRAYFORMULA'), 'Consolidado: se reescribe la fórmula');
const aud = H('_Auditoria');
ok(aud && aud.getLastRow() === 6 && aud.oculta, 'Auditoría con 5 incidentes (y oculta)');
ok(correos.length - antes === 4 && correos[antes].to.includes('coordinadora@uniandes.edu.co'), 'Alerta por correo a los ADMIN');
ok(editar('externo@gmail.com', 'MF', 9, B, 'X', { saltarProteccion: true }) && H('MF').valor(9, B) === '', 'Cuenta externa sin correo identificable: se revierte');

console.log('\n6) Sincronización idempotente y restauración');
H('AFIN').protecciones.find(p => p.desc.includes('colPG-|')).remove();
const extra = H('AFIN').getRange(30, 2, 5, 3).protect().setDescription('mía');
vm.runInContext('restaurarPermisos()', ctx);
ok(H('AFIN').protecciones.filter(p => p.desc.includes('colPG-|')).length === 3, 'Protección de programa borrada se restaura');
ok(!H('AFIN').protecciones.includes(extra), 'Protección creada a mano se elimina');
const n1 = hojas.reduce((s, h) => s + h.protecciones.length, 0);
usuarioActual = DUENO;
vm.runInContext('aplicarPermisos()', ctx);
const n2 = hojas.reduce((s, h) => s + h.protecciones.length, 0);
ok(n1 === n2, `Segunda sincronización no duplica (${n2} protecciones)`);
ok(/creadas: 0 /.test(alertas[alertas.length - 1]), 'Segunda sincronización: 0 creadas');

console.log('\n7) Menú solo para ADMIN');
usuarioActual = 'prof.a@uniandes.edu.co';
vm.runInContext('aplicarPermisos()', ctx);
ok(alertas[alertas.length - 1].includes('solo para ADMIN'), 'Profesor no puede ejecutar el menú');

console.log('\n8) Cambiar permisos en vivo (hoja Permisos)');
usuarioActual = DUENO;
const PER = H('Permisos');
const filaPermiso = nombre => { for (let r = 2; r <= PER.getLastRow(); r++) if (PER.valor(r, 1) === nombre) return r; return 0; };
const CRN = col('AFIN', 'CRN'), COMP = col('AFIN', 'Competencia');
ok(editar('analista.casa@uniandes.edu.co', 'AFIN', 40, D, 'x') === 'BLOQUEADO', 'Antes: CASA no edita «Curso»');
PER.set(filaPermiso('Curso'), 4, 'SI');                        // Curso: programa y CASA
PER.set(filaPermiso('CRN'), 2, 'NO');                          // CRN: solo gestores
PER.set(filaPermiso('Competencia'), 1, 'Competencia (otra)');  // ya no coincide con la hoja
vm.runInContext('aplicarPermisos()', ctx);
const resumenPermisos = alertas[alertas.length - 1];
ok(editar('analista.casa@uniandes.edu.co', 'AFIN', 40, D, 'Curso CASA') === 'EDITADO', 'CASA ahora edita «Curso» (columna mixta)');
ok(H('AFIN').valor(40, REG) === '', '…sin quedarse con la fila (CASA no reclama filas)');
ok(editar('analista.casa@uniandes.edu.co', 'AFIN', 5, D, 'y') === 'EDITADO', '…también en la fila registrada por A');
ok(editar('prof.a@uniandes.edu.co', 'AFIN', 5, D, 'z') === 'EDITADO', 'El dueño sigue editando «Curso»');
ok(editar('prof.b@uniandes.edu.co', 'AFIN', 5, D, 'w') === 'BLOQUEADO', 'Otro profesor sigue bloqueado en la fila de A');
ok(editar('prof.a@uniandes.edu.co', 'AFIN', 5, CRN, '1') === 'BLOQUEADO', 'Profesor ya no edita «CRN» (solo gestores)');
ok(editar('gestor.afin@uniandes.edu.co', 'AFIN', 5, CRN, '1') === 'EDITADO', 'Gestor edita «CRN»');
ok(editar('prof.a@uniandes.edu.co', 'AFIN', 5, COMP, 'x') === 'BLOQUEADO' &&
   editar('gestor.afin@uniandes.edu.co', 'AFIN', 5, COMP, 'x') === 'BLOQUEADO', 'Columna ausente en Permisos: solo ADMIN');
ok(resumenPermisos.includes('no está en la hoja Permisos'), 'Avisa de la columna que no está en Permisos');
ok(editar('prof.a@uniandes.edu.co', 'AFIN', 5, CRN, '9', { saltarProteccion: true }) && H('AFIN').valor(5, CRN) === '1',
   'Si un profesor se salta la protección de «CRN», se revierte');

console.log('\n9) Problemas frecuentes de configuración');
usuarioActual = DUENO;
vm.runInContext('aplicarPermisos()', ctx);
const res9 = alertas[alertas.length - 1];
ok(res9.startsWith('Permisos aplicados | ⚠️ HAY') && res9.includes('«AFIN Finanzas» no coincide con ninguna hoja'),
   'Programa que no coincide: aviso destacado al inicio del resumen');
ok(editar('gestora.min@uniandes.edu.co', 'AFIN', 60, B, 'AFIN-9') === 'EDITADO' && H('AFIN').valor(60, B) === 'AFIN-9',
   'Programa escrito « afín »: la gestora edita AFIN igual');
ok(editar('gestora.mal@uniandes.edu.co', 'AFIN', 61, B, 'x') === 'BLOQUEADO', 'Programa que no coincide: sigue bloqueada (y se avisa)');
editar('gestor.principal@uniandes.edu.co', 'AFIN', 62, B, 'z', { saltarProteccion: true });
ok(H('AFIN').valor(62, B) === '' && String(H('_Auditoria').valor(H('_Auditoria').getLastRow(), 5)).includes('no está en la hoja Usuarios'),
   'Correo distinto al de Usuarios (alias): se revierte y la auditoría lo explica');
const diag = c => vm.runInContext(`diagnostico_(SpreadsheetApp.getActiveSpreadsheet(), leerUsuarios_(SpreadsheetApp.getActiveSpreadsheet()), '${c}')`, ctx);
const d1 = diag('gestor.afin@uniandes.edu.co');
ok(d1.includes('No tiene acceso al archivo'), 'Diagnóstico: detecta que el archivo no se ha compartido');
ok(d1.includes('GESTOR de AFIN') && d1.includes('AFIN: puede editar'), 'Diagnóstico: muestra rol y columnas editables');
ok(diag('nadie@uniandes.edu.co').includes('No aparece en la hoja Usuarios'), 'Diagnóstico: correo que no está en Usuarios');
const pr = H('AFIN').protecciones.find(p => p.desc.includes('colPG-|'));
pr.editores.delete('gestor.afin@uniandes.edu.co');
ok(diag('gestor.afin@uniandes.edu.co').includes('aún no figura en 1 protección'), 'Diagnóstico: detecta protección donde falta');
vm.runInContext('compartirConUsuarios()', ctx);
ok(alertas[alertas.length - 1].includes('Permisos aplicados de nuevo'), 'Compartir vuelve a aplicar los permisos');
const d2 = diag('gestor.afin@uniandes.edu.co');
ok(d2.includes('Acceso al archivo: Editor') && !d2.includes('PROBLEMAS'), 'Diagnóstico tras compartir: sin problemas');

console.log('\n10) Agregar programa con el libro en uso');
usuarioActual = DUENO;
H('AFIN').set(70, B, 'dato de AFIN');
respuestaPrompt = '  NUEVO-PROG ';
vm.runInContext('agregarPrograma()', ctx);
const NP = H('NUEVO-PROG');
ok(NP && NP.getIndex() === H('Pregrado-CEMS').getIndex() + 1, 'Crea la hoja junto a los demás programas');
const encabezados = h => Array.from({ length: 45 }, (_, i) => h.valor(3, i + 1));
ok(JSON.stringify(encabezados(NP)) === JSON.stringify(encabezados(H('AFIN'))), 'Tiene las mismas columnas que los demás programas');
ok(NP.valor(70, B) === '' && NP.valor(5, REG) === '' && NP.valor(5, D) === '', 'Queda vacía (no copia datos ni dueños de filas)');
ok(NP.formulas[`4,${PROG}`] === `=IF(B4&C4&D4=""${separador}""${separador}"NUEVO-PROG")`,
   `La columna Programa usa el nombre nuevo (separador «${separador}»)`);
ok(H('Consolidado').formulas['4,1'].includes("'NUEVO-PROG'!A4"), 'Queda incluida en el Consolidado');
ok(NP.protecciones.some(p => p.desc.includes('colPG-|')) && NP.protecciones.some(p => p.tipo === 'SHEET'), 'Sus columnas quedan protegidas');
const lista = H('Usuarios').validaciones.slice(-1)[0];
ok(lista && lista.regla.lista.includes('NUEVO-PROG'), 'La lista de programas de Usuarios incluye el nuevo');
ok(alertas[alertas.length - 1].startsWith('Programa agregado'), 'Confirma la creación con los siguientes pasos');
U.set(20, 1, 'gestor.nuevo@uniandes.edu.co'); U.set(20, 2, 'GESTOR'); U.set(20, 3, 'nuevo-prog');
vm.runInContext('aplicarPermisos()', ctx);
ok(editar('gestor.nuevo@uniandes.edu.co', 'NUEVO-PROG', 10, B, 'NP-1') === 'EDITADO' && NP.valor(10, B) === 'NP-1',
   'Su gestor (escrito en minúsculas en Usuarios) edita la hoja nueva');
ok(editar('gestor.afin@uniandes.edu.co', 'NUEVO-PROG', 11, B, 'x') === 'BLOQUEADO', 'El gestor de otro programa no la edita');
for (const [nombre, esperado] of [['nuevo-prog', 'Ya existe'], ['Usuarios', 'reservado'], ['a/b', 'no sirve'], ['', 'Escriba']]) {
  const antes = hojas.length;
  respuestaPrompt = nombre;
  vm.runInContext('agregarPrograma()', ctx);
  ok(hojas.length === antes && alertas[alertas.length - 1].includes(esperado), `Rechaza el nombre «${nombre}» (${esperado})`);
}

console.log('\n11) Quitar programa');
usuarioActual = DUENO;
const consolidado = () => H('Consolidado').formulas['4,1'];
H('MM').set(5, B, 'MM-1');
respuestaPrompt = 'mm'; respuestasAlerta.push('YES');  // archivar
vm.runInContext('quitarPrograma()', ctx);
const archMM = hojas.find(h => h.nombre === '_Archivado MM');
ok(alertas.some(a => a.startsWith('Quitar «MM»') && a.includes('1 curso(s) registrado(s)')), 'Muestra cuántos cursos tiene registrados');
ok(archMM && archMM.oculta && archMM.valor(5, B) === 'MM-1', 'Archivar: la hoja se oculta y conserva los datos');
ok(!consolidado().includes("'MM'!") && !consolidado().includes('_Archivado'), 'Archivar: sale del Consolidado');
ok(archMM.protecciones.length === 1 && archMM.protecciones[0].desprot.length === 0, 'Archivar: queda una sola protección, de hoja completa');
ok(!archMM.puedeEditar('gestor.afin@uniandes.edu.co', 5, B) && !archMM.puedeEditar('analista.casa@uniandes.edu.co', 5, L),
   'Archivar: nadie fuera de ADMIN puede editarla');
ok(!H('Usuarios').validaciones.slice(-1)[0].regla.lista.includes('MM'), 'Archivar: sale de la lista de programas de Usuarios');

archMM.protecciones[0].desprot = [archMM.getRange(4, B, 10, 1)];  // alguien le quita protección
vm.runInContext('restaurarPermisos()', ctx);
ok(archMM.protecciones[0].desprot.length === 0, 'La restauración diaria vuelve a proteger la hoja archivada');

respuestaPrompt = 'MSCM'; respuestasAlerta.push('NO', 'NO');  // eliminar, pero no confirma
vm.runInContext('quitarPrograma()', ctx);
ok(H('MSCM'), 'Eliminar sin confirmar: no borra nada');
respuestaPrompt = 'MSCM'; respuestasAlerta.push('NO', 'YES');  // eliminar y confirmar
vm.runInContext('quitarPrograma()', ctx);
ok(!H('MSCM') && !consolidado().includes("'MSCM'!"), 'Eliminar: borra la hoja y la saca del Consolidado');
respuestaPrompt = 'MF'; respuestasAlerta.push('CANCEL');
vm.runInContext('quitarPrograma()', ctx);
ok(H('MF') && consolidado().includes("'MF'!"), 'Cancelar: no cambia nada');
respuestaPrompt = 'NOEXISTE';
vm.runInContext('quitarPrograma()', ctx);
ok(alertas[alertas.length - 1].startsWith('No se encontró el programa'), 'Programa inexistente: avisa y muestra los programas');

respuestaPrompt = 'MF'; respuestasAlerta.push('YES');
vm.runInContext('quitarPrograma()', ctx);
ok(alertas[alertas.length - 1].includes('prof.mf@uniandes.edu.co'), 'Avisa qué personas de Usuarios tenían el programa');
const archMF = hojas.find(h => h.nombre === '_Archivado MF');
archMF.setName('MF'); archMF.showSheet();  // restaurar: quitar el prefijo y mostrar
vm.runInContext('aplicarPermisos()', ctx);
ok(consolidado().includes("'MF'!") && editar('prof.mf@uniandes.edu.co', 'MF', 40, B, 'MF-9') === 'EDITADO' && H('MF').valor(40, B) === 'MF-9',
   'Restaurar: al quitar el prefijo y aplicar permisos el programa vuelve completo');

H('Consolidado').formulas['4,1'] = "=ARRAYFORMULA(IFNA(FILTER(VSTACK('AFIN'!A4:AK),'roto'),\"x\"))";
usuarioActual = DUENO;
vm.runInContext('aplicarPermisos()', ctx);
const reparada = H('Consolidado').formulas['4,1'];
ok(reparada.includes("'Pregrado-CEMS'!A4") && reparada.endsWith('"Sin registros"))'),
   'Aplicar permisos ahora reescribe un Consolidado dañado');

console.log(fallos ? `\n${fallos} prueba(s) fallaron` : '\nTodas las pruebas pasaron');
process.exit(fallos ? 1 : 0);
