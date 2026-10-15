/**
 * Musculosa Los Guerreros - Club Manuel Belgrano - Playoff 2026
 * Backend del relevamiento de pedidos (Google Apps Script, vinculado a la planilla).
 * Ver README.md para implementarlo como Aplicacion web.
 */

// ===== Configuracion =====
var PRECIO = 38000;                 // precio base en pesos
var DESCUENTO = 10000;              // descuento del sale en pesos
var FECHA_SALE = '2026-10-15';      // dia del sale (hora Argentina, 00:00 a 23:59:59)
var ALIAS = 'dami.lizzano';         // alias para transferir
var TZ = 'America/Argentina/Buenos_Aires';
var TZ_OFFSET = '-03:00';           // Argentina no tiene horario de verano
var SHEET_ID = ''; // opcional: solo si el script NO está pegado dentro del Sheet
var CARPETA = 'Comprobantes Musculosa 2026';
var MAX_BYTES = 5 * 1024 * 1024;    // 5 MB por comprobante

var DIVISIONES = ['M15', 'M16', 'M17', 'M19', 'PS'];
var TALLES = ['XS', 'S', 'M', 'L', 'XL', 'XXL', 'XXXL'];
var MIMES = ['image/jpeg', 'image/png', 'image/webp', 'image/heic', 'image/heif', 'application/pdf'];
var ESTADOS = ['Pendiente de pago', 'Comprobante recibido', 'Pagado', 'Cancelado'];

var HEADERS = ['ID', 'Fecha/Hora', 'Nombre y apellido', 'Apodo', 'División', 'Talle', 'WhatsApp',
  'Precio base', 'Descuento', 'Total a pagar', 'Sale aplicado', 'Estado', 'Comprobante',
  'Fecha comprobante', 'Clave'];
// Indices de columna (1-based)
var C = { ID: 1, FECHA: 2, DESC: 9, TOTAL: 10, SALE: 11, ESTADO: 12, COMP: 13, FCOMP: 14, CLAVE: 15 };

// ===== Logica del sale (pura, testeable) =====
function saleInicio_() { return Date.parse(FECHA_SALE + 'T00:00:00' + TZ_OFFSET); }
function saleFin_() { return Date.parse(FECHA_SALE + 'T23:59:59.999' + TZ_OFFSET); }
function enSale_(ms) { return ms >= saleInicio_() && ms <= saleFin_(); }

/** El descuento vale si el PEDIDO se hizo el dia del sale y el COMPROBANTE llego antes de las 23:59 de ese dia. */
function aplicaSale_(msPedido, msComprobante) {
  return enSale_(msPedido) && msComprobante != null && enSale_(msComprobante);
}

// ===== Endpoints =====
function doGet(e) {
  var p = (e && e.parameter) || {};
  if (p.id) return json_(consultar_(String(p.id), String(p.k || '')));
  var now = Date.now();
  return json_({ ok: true, serverTime: now, saleActivo: enSale_(now), saleFin: saleFin_(),
    precio: PRECIO, descuento: DESCUENTO, alias: ALIAS });
}

function doPost(e) {
  try {
    var body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (body.action === 'crear') return json_(crear_(body));
    if (body.action === 'comprobante') return json_(comprobante_(body));
    return json_({ ok: false, error: 'Acción inválida.' });
  } catch (err) {
    console.error(err);
    return json_({ ok: false, error: 'Error del servidor. Probá de nuevo en un rato.' });
  }
}

// ===== Acciones =====
function crear_(b) {
  // Honeypot: un humano no completa este campo. Respondemos "ok" falso para no darle pistas al bot.
  if (b.web) return { ok: true, id: 'G-0000', clave: 'x', total: PRECIO, saleProvisorio: false };

  var nombre = texto_(b.nombre, 80), apodo = texto_(b.apodo, 30);
  var division = String(b.division || ''), talle = String(b.talle || '');
  var wa = String(b.whatsapp || '').replace(/\D/g, '');
  if (nombre.length < 3) return err_('Completá tu nombre y apellido.');
  if (!apodo) return err_('Completá el nombre/apodo para la musculosa.');
  if (DIVISIONES.indexOf(division) < 0) return err_('Elegí una división válida.');
  if (TALLES.indexOf(talle) < 0) return err_('Elegí un talle válido.');
  if (wa.length < 10 || wa.length > 13) return err_('El WhatsApp no parece válido.');

  // Idempotencia: si el cliente reintenta el mismo envio, devolvemos el mismo pedido.
  var nonce = texto_(b.nonce, 64), cache = CacheService.getScriptCache();
  if (nonce) {
    var prev = cache.get('n:' + nonce);
    if (prev) return JSON.parse(prev);
  }

  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = hoja_();
    var now = new Date();
    var sale = enSale_(now.getTime());
    var desc = sale ? DESCUENTO : 0;
    var props = PropertiesService.getScriptProperties();
    var n = Number(props.getProperty('ULTIMO_ID') || sh.getLastRow() - 1) + 1; // contador: no se repite aunque borren filas
    props.setProperty('ULTIMO_ID', String(n));
    var id = 'G-' + ('000' + n).slice(-4);
    var clave = Utilities.getUuid().replace(/-/g, '').slice(0, 10);
    sh.appendRow([id, now, safe_(nombre), safe_(apodo), division, talle, "'" + wa,
      PRECIO, desc, PRECIO - desc, sale ? 'Provisorio (falta comprobante)' : 'No',
      'Pendiente de pago', '', '', clave]);
    SpreadsheetApp.flush();
    var res = { ok: true, id: id, clave: clave, fecha: now.getTime(), precio: PRECIO,
      descuento: desc, total: PRECIO - desc, saleProvisorio: sale, saleFin: saleFin_(), alias: ALIAS };
    if (nonce) cache.put('n:' + nonce, JSON.stringify(res), 21600);
    return res;
  } finally {
    lock.releaseLock();
  }
}

function comprobante_(b) {
  var id = String(b.id || ''), clave = String(b.clave || '');
  var mime = String(b.mime || '').toLowerCase();
  if (MIMES.indexOf(mime) < 0) return err_('Formato no permitido. Subí una imagen o un PDF.');
  var bytes;
  try { bytes = Utilities.base64Decode(String(b.data || '')); } catch (x) { return err_('Archivo dañado.'); }
  if (!bytes.length) return err_('El archivo está vacío.');
  if (bytes.length > MAX_BYTES) return err_('El archivo pesa más de 5 MB.');

  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = hoja_();
    var fila = buscarFila_(sh, id);
    if (!fila) return err_('No encontramos el pedido ' + id + '.');
    var row = sh.getRange(fila, 1, 1, HEADERS.length).getValues()[0];
    if (String(row[C.CLAVE - 1]) !== clave) return err_('El link del pedido no es válido.');
    if (row[C.ESTADO - 1] === 'Pagado' || row[C.ESTADO - 1] === 'Cancelado') {
      return err_('Este pedido ya está cerrado (' + row[C.ESTADO - 1] + ').');
    }

    var now = new Date();
    var ext = mime === 'application/pdf' ? 'pdf' : mime.split('/')[1].replace('jpeg', 'jpg');
    var file = carpeta_().createFile(Utilities.newBlob(bytes, mime, id + '_' + now.getTime() + '.' + ext));
    // El archivo queda privado (solo el dueño de la planilla lo ve).

    var msPedido = new Date(row[C.FECHA - 1]).getTime();
    var sale = aplicaSale_(msPedido, now.getTime());
    var desc = sale ? DESCUENTO : 0;
    var saleTxt = sale ? 'Sí' : (enSale_(msPedido) ? 'No (comprobante fuera de término)' : 'No');
    sh.getRange(fila, C.DESC, 1, 6).setValues([[desc, PRECIO - desc, saleTxt, 'Comprobante recibido',
      file.getUrl(), now]]);
    return { ok: true, id: id, total: PRECIO - desc, descuento: desc, saleAplicado: sale };
  } finally {
    lock.releaseLock();
  }
}

function consultar_(id, clave) {
  var sh = hoja_();
  var fila = buscarFila_(sh, id);
  if (!fila) return err_('No encontramos el pedido.');
  var r = sh.getRange(fila, 1, 1, HEADERS.length).getValues()[0];
  if (String(r[C.CLAVE - 1]) !== clave) return err_('El link del pedido no es válido.');
  return { ok: true, id: id, clave: clave, apodo: String(r[3]).replace(/^'/, ''), division: r[4], talle: r[5],
    nombre: String(r[2]).replace(/^'/, ''), total: r[C.TOTAL - 1], descuento: r[C.DESC - 1],
    saleProvisorio: Number(r[C.DESC - 1]) > 0, estado: r[C.ESTADO - 1], alias: ALIAS, saleFin: saleFin_() };
}

// ===== Helpers =====
function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function err_(msg) { return { ok: false, error: msg }; }

function texto_(v, max) {
  return String(v == null ? '' : v).replace(/[\u0000-\u001F\u007F]/g, ' ').replace(/\s+/g, ' ').trim().slice(0, max);
}
/** Evita inyeccion de formulas en la planilla. */
function safe_(s) { return /^[=+\-@\t\r]/.test(s) ? "'" + s : s; }

function buscarFila_(sh, id) {
  if (!/^G-\d{4,}$/.test(id) || sh.getLastRow() < 2) return 0;
  var cell = sh.getRange(2, C.ID, sh.getLastRow() - 1, 1).createTextFinder(id).matchEntireCell(true).findNext();
  return cell ? cell.getRow() : 0;
}

function libro_() {
  return SpreadsheetApp.getActive() || SpreadsheetApp.openById(SHEET_ID);
}

function hoja_() {
  var ss = libro_();
  var sh = ss.getSheetByName('Pedidos');
  if (sh && sh.getLastRow() > 0) return sh;
  sh = sh || ss.insertSheet('Pedidos');
  sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS])
    .setFontWeight('bold').setBackground('#0B1F4B').setFontColor('#FFFFFF');
  sh.setFrozenRows(1);
  sh.getRange('B:B').setNumberFormat('dd/MM/yyyy HH:mm:ss');
  sh.getRange('N:N').setNumberFormat('dd/MM/yyyy HH:mm:ss');
  sh.getRange('H:J').setNumberFormat('$#,##0');
  sh.getRange('L2:L').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(ESTADOS, true).build());
  var est = sh.getRange('L2:L');
  var regla = function (txt, color) {
    return SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(txt).setBackground(color).setRanges([est]).build();
  };
  sh.setConditionalFormatRules([regla('Pendiente de pago', '#FDE2E1'), regla('Comprobante recibido', '#FFF4CC'),
    regla('Pagado', '#D9F2DC'), regla('Cancelado', '#E0E0E0')]);
  sh.hideColumns(C.CLAVE);
  resumen_(ss);
  return sh;
}

function resumen_(ss) {
  if (ss.getSheetByName('Resumen')) return;
  var r = ss.insertSheet('Resumen');
  var rows = [['Talle', 'Pedidos', 'Con comprobante', 'Pagados']];
  TALLES.forEach(function (t, i) {
    var a = 'A' + (i + 2);
    rows.push([t, '=COUNTIFS(Pedidos!F:F,' + a + ',Pedidos!L:L,"<>Cancelado")',
      '=COUNTIFS(Pedidos!F:F,' + a + ',Pedidos!L:L,"Comprobante recibido")',
      '=COUNTIFS(Pedidos!F:F,' + a + ',Pedidos!L:L,"Pagado")']);
  });
  rows.push(['Total', '=SUM(B2:B8)', '=SUM(C2:C8)', '=SUM(D2:D8)']);
  rows.push(['', '', '', '']);
  rows.push(['División', 'Pedidos', 'Con comprobante', 'Pagados']);
  DIVISIONES.forEach(function (d, i) {
    var a = 'A' + (i + 12);
    rows.push([d, '=COUNTIFS(Pedidos!E:E,' + a + ',Pedidos!L:L,"<>Cancelado")',
      '=COUNTIFS(Pedidos!E:E,' + a + ',Pedidos!L:L,"Comprobante recibido")',
      '=COUNTIFS(Pedidos!E:E,' + a + ',Pedidos!L:L,"Pagado")']);
  });
  rows.push(['', '', '', '']);
  rows.push(['Total cobrado (Pagado)', '=SUMIFS(Pedidos!J:J,Pedidos!L:L,"Pagado")', '', '']);
  rows.push(['A cobrar (con comprobante)', '=SUMIFS(Pedidos!J:J,Pedidos!L:L,"Comprobante recibido")', '', '']);
  r.getRange(1, 1, rows.length, 4).setValues(rows);
  [1, 11].forEach(function (h) {
    r.getRange(h, 1, 1, 4).setFontWeight('bold').setBackground('#C8102E').setFontColor('#FFFFFF');
  });
  r.getRange('A9:D9').setFontWeight('bold');
  r.getRange('B18:B19').setNumberFormat('$#,##0');
  r.autoResizeColumn(1);
}

function carpeta_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('CARPETA_ID');
  if (id) { try { return DriveApp.getFolderById(id); } catch (x) { /* borrada: se recrea */ } }
  var f = DriveApp.createFolder(CARPETA);
  props.setProperty('CARPETA_ID', f.getId());
  return f;
}

/** Ejecutalo desde el editor (Ejecutar > testSale) para chequear la regla del sale. Mira el registro. */
function testSale() {
  var t = function (s) { return Date.parse(s); };
  var casos = [
    ['pedido y comprobante el 15', t('2026-10-15T10:00:00-03:00'), t('2026-10-15T23:00:00-03:00'), true],
    ['pedido 23:59:30, comprobante 23:59:50', t('2026-10-15T23:59:30-03:00'), t('2026-10-15T23:59:50-03:00'), true],
    ['pedido el 15, comprobante el 16 00:00:01', t('2026-10-15T20:00:00-03:00'), t('2026-10-16T00:00:01-03:00'), false],
    ['pedido el 14', t('2026-10-14T23:59:00-03:00'), t('2026-10-15T09:00:00-03:00'), false],
    ['pedido el 16', t('2026-10-16T00:00:00-03:00'), t('2026-10-16T00:01:00-03:00'), false],
    ['pedido 15 00:00 justo', t('2026-10-15T00:00:00-03:00'), t('2026-10-15T00:05:00-03:00'), true],
    ['sin comprobante', t('2026-10-15T10:00:00-03:00'), null, false],
    ['UTC 02:59 del 16 = 23:59 AR del 15', t('2026-10-16T02:59:00Z'), t('2026-10-16T02:59:30Z'), true],
    ['UTC 03:00 del 16 = 00:00 AR del 16', t('2026-10-15T12:00:00Z'), t('2026-10-16T03:00:00Z'), false]
  ];
  casos.forEach(function (c) {
    var r = aplicaSale_(c[1], c[2]);
    if (r !== c[3]) throw new Error('FALLA: ' + c[0]);
  });
  console.log('testSale OK (' + casos.length + ' casos)');
}
