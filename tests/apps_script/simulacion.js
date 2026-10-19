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
}

class Hoja {
  constructor(nombre, maxFilas, maxCols) { Object.assign(this, { nombre, maxFilas, maxCols, datos: {}, fondos: {}, formulas: {}, protecciones: [], oculta: false }); }
  valor(r, c) { const v = this.datos[`${r},${c}`]; return v === undefined ? '' : v; }
  set(r, c, v) { this.datos[`${r},${c}`] = v; }
  getName() { return this.nombre; }
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
  hojas.push(h);
});
const ss = {
  getSheets: () => hojas,
  getSheetByName: n => hojas.find(h => h.nombre === n) || null,
  insertSheet: n => { const h = new Hoja(n, 1000, 26); hojas.push(h); return h; },
  getOwner: () => user(DUENO),
  getEditors: () => [user(DUENO)],
  addEditor: () => {}, toast: () => {}, getUrl: () => 'https://docs.google.com/spreadsheets/d/x',
  getId: () => 'x',
};
const alertas = [];
const props = {};
const ctx = {
  console: { log() {}, warn: m => console.log('   [warn]', m), error: m => console.log('   [error]', m) },
  SpreadsheetApp: {
    getActiveSpreadsheet: () => ss,
    ProtectionType: { RANGE: 'RANGE', SHEET: 'SHEET' },
    getUi: () => ({ alert: (...a) => { alertas.push(a.filter(x => typeof x === 'string').join(' | ')); return 'YES'; }, ButtonSet: { OK: 1, YES_NO: 2 }, Button: { YES: 'YES' },
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
].forEach((f, i) => f.forEach((v, j) => U.set(i + 2, j + 1, v)));

const B = col('AFIN', 'Código de curso'), D = col('AFIN', 'Curso'), L = col('AFIN', 'Número de Competencia a evaluar');
const REG = col('AFIN', 'Registrado por'), PROG = col('AFIN', 'Programa');

console.log('\n1) Configurar libro (dueño)');
usuarioActual = DUENO;
vm.runInContext('configurarLibro()', ctx);
const pAFIN = H('AFIN').protecciones;
ok(pAFIN.filter(p => p.tipo === 'SHEET').length === 1, 'AFIN: 1 protección de hoja');
ok(pAFIN.filter(p => p.desc.includes('azul|')).length === 3, 'AFIN: 3 bloques azules (B:K, M, W) → ' + pAFIN.filter(p => p.desc.includes('azul|')).map(p => p.rango.getA1Notation()).join(' '));
ok(pAFIN.filter(p => p.desc.includes('verde|')).length === 3, 'AFIN: 3 bloques verdes → ' + pAFIN.filter(p => p.desc.includes('verde|')).map(p => p.rango.getA1Notation()).join(' '));
ok(H('Usuarios').protecciones.length === 1 && H('Consolidado').protecciones.length === 1, 'Usuarios y Consolidado protegidas');
const formula = H('Consolidado').formulas['4,1'] || '';
ok(formula.startsWith('=ARRAYFORMULA(IFNA(FILTER(VSTACK(') && formula.includes("'Pregrado-CEMS'!A4:AH"), 'Fórmula del Consolidado (' + formula.length + ' caracteres)');
console.log('   ' + formula.slice(0, 160) + '…');
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
ok(pf && pf.rango.getA1Notation() === 'A5:AH5', 'Fila 5 protegida (A5:AH5)');
ok(editar('prof.a@uniandes.edu.co', 'AFIN', 5, D, 'Finanzas') === 'EDITADO', 'A sigue editando su fila');
ok(editar('prof.b@uniandes.edu.co', 'AFIN', 5, D, 'Otro') === 'BLOQUEADO', 'Profesor B no edita la fila de A');
ok(editar('gestor.afin@uniandes.edu.co', 'AFIN', 5, D, 'Finanzas I') === 'EDITADO', 'Gestor de AFIN sí edita la fila de A');
ok(editar('analista.casa@uniandes.edu.co', 'AFIN', 5, L, 'C1') === 'EDITADO', 'CASA edita la columna verde de la fila de A');
ok(editar('analista.casa@uniandes.edu.co', 'AFIN', 5, B, 'X') === 'BLOQUEADO', 'CASA sigue sin editar azul en la fila de A');
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
ok(aud && aud.getLastRow() === 5 && aud.oculta, 'Auditoría con 4 incidentes (y oculta)');
ok(correos.length - antes === 4 && correos[antes].to.includes('coordinadora@uniandes.edu.co'), 'Alerta por correo a los ADMIN');
ok(editar('externo@gmail.com', 'MF', 9, B, 'X', { saltarProteccion: true }) && H('MF').valor(9, B) === '', 'Cuenta externa sin correo identificable: se revierte');

console.log('\n6) Sincronización idempotente y restauración');
H('AFIN').protecciones.find(p => p.desc.includes('azul|')).remove();
const extra = H('AFIN').getRange(30, 2, 5, 3).protect().setDescription('mía');
vm.runInContext('restaurarPermisos()', ctx);
ok(H('AFIN').protecciones.filter(p => p.desc.includes('azul|')).length === 3, 'Protección azul borrada se restaura');
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

console.log(fallos ? `\n${fallos} prueba(s) fallaron` : '\nTodas las pruebas pasaron');
process.exit(fallos ? 1 : 0);
