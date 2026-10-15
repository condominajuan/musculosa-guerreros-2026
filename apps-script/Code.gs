/**
 * Musculosa Los Guerreros - Club Manuel Belgrano - Playoff 2026
 * Backend del relevamiento de pedidos (Google Apps Script, vinculado a la planilla).
 * Ver README.md. Despues de pegar este codigo: Ejecutar > configurarPlanilla (una sola vez).
 */

// ===== Configuracion =====
var PRECIO = 38000;
var ALIAS = 'playoffs.2026';
var CARPETA = 'Comprobantes Musculosa 2026';
var MAX_BYTES = 5 * 1024 * 1024;    // 5 MB por comprobante
var SHEET_ID = '';                  // opcional: solo si el script NO esta pegado dentro del Sheet
var HOJA_WEB = 'Pedidos Web';       // pedidos que entran por la landing
var HOJA_FORMS = 'Pedidos Forms';   // pedidos del sale hecho por Google Forms

var DIVISIONES = ['M15', 'M16', 'M17', 'M19', 'PS'];
var TALLES = ['XS', 'S', 'M', 'L', 'XL', 'XXL', 'XXXL'];
var MIMES = ['image/jpeg', 'image/png', 'image/webp', 'image/heic', 'image/heif', 'application/pdf'];
var ESTADOS = ['Pendiente de pago', 'Comprobante recibido', 'Pago confirmado', 'Cancelado'];

// Orden de columnas de la hoja Pedidos. El codigo busca cada columna por su titulo
// (si la planilla tiene columnas de mas, como una vieja "Clave", se ignoran).
var HEADERS = ['ID', 'Fecha/Hora', 'Nombre y apellido', 'División', 'Talle', 'WhatsApp', 'Precio base',
  'Estado', 'Pago confirmado', 'Comprobante', 'Fecha comprobante', 'Avisar', 'Aviso enviado'];

// ===== Endpoints =====
function doGet() {
  return json_({ ok: true, serverTime: Date.now(), precio: PRECIO, alias: ALIAS });
}

function doPost(e) {
  try {
    var body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (body.action === 'comprobante') return json_(pedido_(body));
    return json_({ ok: false, error: 'Acción inválida.' });
  } catch (err) {
    console.error(err);
    return json_({ ok: false, error: 'Error del servidor. Probá de nuevo en un rato.' });
  }
}

// ===== Accion =====
/**
 * El pedido se anota recien cuando llega el comprobante: datos + archivo en un solo POST.
 * Idempotente por nonce: un reintento devuelve el mismo pedido sin duplicar fila ni archivo.
 */
function pedido_(b) {
  // Honeypot: un humano no completa este campo. Respondemos "ok" falso para no darle pistas al bot.
  if (b.web) return { ok: true, id: 'G-0000', total: PRECIO };

  var nombre = texto_(b.nombre, 80);
  var division = String(b.division || ''), talle = String(b.talle || '');
  var wa = String(b.whatsapp || '').replace(/\D/g, '');
  var nonce = texto_(b.nonce, 64);
  if (nombre.length < 3) return err_('Completá tu nombre y apellido.');
  if (DIVISIONES.indexOf(division) < 0) return err_('Elegí una división válida.');
  if (TALLES.indexOf(talle) < 0) return err_('Elegí un talle válido.');
  if (wa.length < 10 || wa.length > 13) return err_('El WhatsApp no parece válido.');
  if (nonce.length < 8) return err_('Pedido inválido. Recargá la página y probá de nuevo.');

  var mime = String(b.mime || '').toLowerCase();
  if (MIMES.indexOf(mime) < 0) return err_('Formato no permitido. Subí una imagen o un PDF.');
  var bytes;
  try { bytes = Utilities.base64Decode(String(b.data || '')); } catch (x) { return err_('Archivo dañado.'); }
  if (!bytes.length) return err_('El archivo está vacío.');
  if (bytes.length > MAX_BYTES) return err_('El archivo pesa más de 5 MB.');

  var cache = CacheService.getScriptCache(), key = 'n:' + nonce;
  var prev = cache.get(key);
  if (prev) return JSON.parse(prev);

  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    prev = cache.get(key);                 // otro reintento pudo terminar mientras esperabamos el lock
    if (prev) return JSON.parse(prev);

    var sh = hoja_(), col = cols_(sh);
    var now = new Date();
    var props = PropertiesService.getScriptProperties();
    var n = Number(props.getProperty('ULTIMO_ID') || ultimaFila_(sh, col) - 1) + 1; // no se repite aunque borren filas
    var id = 'G-' + ('000' + n).slice(-4);

    // 1) Archivo primero: si Drive falla, no queda fila ni ID consumido.
    var ext = mime === 'application/pdf' ? 'pdf' : mime.split('/')[1].replace('jpeg', 'jpg');
    var file = carpeta_().createFile(Utilities.newBlob(bytes, mime, id + '_' + now.getTime() + '.' + ext));
    props.setProperty('ULTIMO_ID', String(n));

    // 2) Fila completa de una sola vez.
    var fila = ultimaFila_(sh, col) + 1;
    var row = [];
    for (var i = 0; i < sh.getLastColumn(); i++) row.push('');
    var set = function (h, v) { row[col[h] - 1] = v; };
    set('ID', id); set('Fecha/Hora', now); set('Nombre y apellido', safe_(nombre));
    set('División', division); set('Talle', talle); set('WhatsApp', "'" + wa);
    set('Precio base', PRECIO); set('Estado', 'Comprobante recibido'); set('Pago confirmado', false);
    set('Comprobante', file.getUrl()); set('Fecha comprobante', now);
    set('Aviso enviado', false);
    sh.getRange(fila, 1, 1, row.length).setValues([row]);
    sh.getRange(fila, col['Pago confirmado']).insertCheckboxes();
    sh.getRange(fila, col['Aviso enviado']).insertCheckboxes();
    SpreadsheetApp.flush();

    var res = { ok: true, id: id, total: PRECIO };
    cache.put(key, JSON.stringify(res), 21600);
    return res;
  } finally {
    lock.releaseLock();
  }
}

// ===== Checkbox "Pago confirmado" (trigger simple, no pide permisos) =====
function onEdit(e) {
  if (!e || !e.range) return;
  var sh = e.range.getSheet();
  if (sh.getName() !== HOJA_WEB || e.range.getRow() < 2) return;
  var col = cols_(sh);
  if (e.range.getColumn() !== col['Pago confirmado']) return;
  var r = e.range.getRow();
  var estadoCel = sh.getRange(r, col['Estado']);
  if (estadoCel.getValue() === 'Cancelado') return;
  var pago = e.range.getValue() === true;
  estadoCel.setValue(pago ? 'Pago confirmado'
    : (sh.getRange(r, col['Comprobante']).getValue() ? 'Comprobante recibido' : 'Pendiente de pago'));
  // Al confirmar el pago aparece el link para avisar por WhatsApp; al destildar se borra.
  sh.getRange(r, col['Avisar']).setFormula(pago ? avisoFormula_(sh, col, r) : '');
}

// ===== Aviso por WhatsApp (link wa.me con el mensaje armado; se envia con un clic) =====
function mensaje_(nombre, id, division, talle) {
  var primero = String(nombre).trim().split(' ')[0];
  return 'Hola ' + primero + '! Te confirmamos tu pedido ' + id + ': musculosa Los Guerreros, división ' + division +
    ', talle ' + talle + '. Pago confirmado. Cuando tengamos novedades de la entrega te avisamos por acá. ' +
    'Gracias por sumarte! Club Manuel Belgrano';
}

/** Normaliza un celular argentino a 549 + area + numero (10 digitos, sin 0 ni 15). Devuelve '' si no se puede. */
function telAR_(raw) {
  var d = String(raw).replace(/\D/g, '').replace(/^00/, '');
  if (d.indexOf('54') === 0) d = d.slice(2);
  if (d.charAt(0) === '9' && d.length > 10) d = d.slice(1);
  d = d.replace(/^0/, '');
  if (d.length === 12) {                       // area con 0 y 15: 011 15 xxxx xxxx ya sin 0 -> 11 15 ...
    d = d.replace(/^(\d{2,4})15(\d{6,8})$/, '$1$2');
  } else if (d.length === 11) {
    for (var a = 2; a <= 4; a++) if (d.substr(a, 2) === '15' && d.length - 2 === 10) { d = d.slice(0, a) + d.slice(a + 2); break; }
  }
  return d.length === 10 && d.indexOf('15') !== 0 ? '549' + d : '';   // ningun codigo de area empieza con 15
}

function avisoFormula_(sh, col, r) {
  var v = function (h) { return sh.getRange(r, col[h]).getValue(); };
  var tel = telAR_(v('WhatsApp'));
  if (!tel) return '="Revisar número"';
  var url = 'https://wa.me/' + tel + '?text=' + encodeURIComponent(mensaje_(v('Nombre y apellido'), v('ID'), v('División'), v('Talle')));
  return '=HYPERLINK("' + url + '","Avisar por WhatsApp")';
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

/** Mapa titulo -> numero de columna (1-based), leido de la fila 1. */
function cols_(sh) {
  var h = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0], m = {};
  h.forEach(function (t, i) { m[String(t).trim()] = i + 1; });
  HEADERS.forEach(function (t) { if (!m[t]) throw new Error('Falta la columna "' + t + '" en Pedidos. Ejecutá configurarPlanilla.'); });
  return m;
}

/** Ultima fila con un ID. No usa getLastRow(): las casillas y formatos de filas vacias lo inflan. */
function ultimaFila_(sh, col) {
  var ids = sh.getRange(1, col['ID'], sh.getMaxRows(), 1).getValues();
  for (var i = ids.length - 1; i >= 1; i--) if (ids[i][0] !== '') return i + 1;
  return 1;
}

function libro_() {
  return SpreadsheetApp.getActive() || SpreadsheetApp.openById(SHEET_ID);
}

function hoja_() {
  var sh = libro_().getSheetByName(HOJA_WEB);
  if (!sh) { configurarPlanilla(); sh = libro_().getSheetByName(HOJA_WEB); }
  return sh;
}

function carpeta_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('CARPETA_ID');
  if (id) { try { return DriveApp.getFolderById(id); } catch (x) { /* borrada: se recrea */ } }
  var f = DriveApp.createFolder(CARPETA);
  props.setProperty('CARPETA_ID', f.getId());
  return f;
}

// ===== Configuracion de la planilla (ejecutar una vez desde el editor; se puede repetir sin romper nada) =====
function configurarPlanilla() {
  var ss = libro_();
  var sh = ss.getSheetByName(HOJA_WEB) || ss.insertSheet(HOJA_WEB);

  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  } else {
    // Planilla existente: agrega la columna "Pago confirmado" justo despues de "Estado" si falta.
    var h = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
    if (h.indexOf('Pago confirmado') < 0) {
      var iE = h.indexOf('Estado');
      if (iE < 0) throw new Error('No encuentro la columna Estado en ' + HOJA_WEB + '.');
      sh.insertColumnAfter(iE + 1);
      sh.getRange(1, iE + 2).setValue('Pago confirmado');
    }
  }
  // Columnas nuevas (planilla existente): se agregan en orden, cada una despues de la anterior.
  [['Avisar', 'Fecha comprobante'], ['Aviso enviado', 'Avisar']].forEach(function (par) {
    var hh = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(function (x) { return String(x).trim(); });
    if (hh.indexOf(par[0]) < 0) {
      var ip = hh.indexOf(par[1]);
      if (ip < 0) throw new Error('No encuentro la columna ' + par[1] + ' en ' + HOJA_WEB + '.');
      sh.insertColumnAfter(ip + 1);
      sh.getRange(1, ip + 2).setValue(par[0]);
    }
  });
  var col = cols_(sh);
  var maxF = Math.max(sh.getMaxRows(), 2);
  var nF = maxF - 1;

  sh.getRange(1, 1, 1, sh.getLastColumn()).setFontWeight('bold').setBackground('#0B1F4B').setFontColor('#FFFFFF');
  sh.setFrozenRows(1);
  sh.getRange(2, col['Fecha/Hora'], nF, 1).setNumberFormat('dd/MM/yyyy HH:mm:ss');
  sh.getRange(2, col['Fecha comprobante'], nF, 1).setNumberFormat('dd/MM/yyyy HH:mm:ss');
  sh.getRange(2, col['Precio base'], nF, 1).setNumberFormat('$#,##0');
  sh.getRange(2, col['WhatsApp'], nF, 1).setNumberFormat('@');

  // Casillas SOLO en filas con pedido (las filas vacias quedan limpias, asi no inflan getLastRow).
  var datos = ultimaFila_(sh, col) - 1;
  if (datos + 2 <= maxF) {
    ['Pago confirmado', 'Aviso enviado'].forEach(function (h) {
      var vacias = sh.getRange(datos + 2, col[h], maxF - datos - 1, 1);
      vacias.clearDataValidations(); vacias.clearContent();
    });
  }
  if (datos > 0) {
    sh.getRange(2, col['Pago confirmado'], datos, 1).insertCheckboxes();
    sh.getRange(2, col['Aviso enviado'], datos, 1).insertCheckboxes();
    var est = sh.getRange(2, col['Estado'], datos, 1).getValues();
    sh.getRange(2, col['Pago confirmado'], datos, 1).setValues(est.map(function (r) {
      return [r[0] === 'Pago confirmado' || r[0] === 'Pagado'];
    }));
    sh.getRange(2, col['Estado'], datos, 1).setValues(est.map(function (r) {
      return [r[0] === 'Pagado' ? 'Pago confirmado' : r[0]];
    }));
  }

  // Link de aviso para los pedidos que ya estan confirmados.
  for (var r = 2; r <= datos + 1; r++) {
    if (sh.getRange(r, col['Pago confirmado']).getValue() === true) {
      sh.getRange(r, col['Avisar']).setFormula(avisoFormula_(sh, col, r));
    }
  }

  var rEst = sh.getRange(2, col['Estado'], nF, 1);
  rEst.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(ESTADOS, true).build());
  var regla = function (txt, color) {
    return SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(txt).setBackground(color).setRanges([rEst]).build();
  };
  // Pago confirmado: se pinta TODA la fila (va primero para que mande sobre el color de la celda Estado).
  var filaPaga = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$' + letra_(col['Estado']) + '2="Pago confirmado"').setBackground('#D9F2DC')
    .setRanges([sh.getRange(2, 1, nF, sh.getLastColumn())]).build();
  sh.setConditionalFormatRules([filaPaga, regla('Pendiente de pago', '#FDE2E1'), regla('Comprobante recibido', '#FFF4CC'),
    regla('Cancelado', '#E0E0E0')]);
  if (!sh.getFilter()) sh.getRange(1, 1, maxF, sh.getLastColumn()).createFilter();

  armarTodos_(ss, sh, col);
  armarDashboard_(ss);
  var viejo = ss.getSheetByName('Resumen');
  if (viejo) ss.deleteSheet(viejo);
  SpreadsheetApp.flush();
}

function letra_(n) { return String.fromCharCode(64 + n); }

/** Hoja "Todos": une los pedidos del Forms (sale) con los de la web, en una sola tabla con filtros. */
function armarTodos_(ss, ped, col) {
  var t = ss.getSheetByName('Todos') || ss.insertSheet('Todos');
  t.clear();
  var P = function (h) { return "'" + HOJA_WEB + "'!$" + letra_(col[h]) + "$2:$" + letra_(col[h]); };
  var F = "'" + HOJA_FORMS + "'!";
  t.getRange(1, 1, 1, 7).setValues([['Origen', 'Fecha', 'Nombre', 'División', 'Talle', 'WhatsApp', 'Estado']])
    .setFontWeight('bold').setBackground('#0B1F4B').setFontColor('#FFFFFF');
  // Forms: la columna E (check) = pago confirmado. Web: se toma el Estado de Pedidos.
  var forms = 'IFERROR(FILTER({ARRAYFORMULA(IF(' + F + '$A$2:$A<>"","Forms (sale)","")),' + F + '$A$2:$A,' +
    'ARRAYFORMULA(TRIM(' + F + '$C$2:$C)),ARRAYFORMULA(TRIM(' + F + '$B$2:$B)),ARRAYFORMULA(TRIM(' + F + '$D$2:$D)),ARRAYFORMULA(IF(' + F + '$A$2:$A<>"","","")),' +
    'ARRAYFORMULA(IF(TRIM(' + F + '$E$2:$E)="✔","Pago confirmado","Pendiente de pago"))},' + F + '$A$2:$A<>""),{"","","","","","",""})';
  var web = 'IFERROR(FILTER({ARRAYFORMULA(IF(' + P('ID') + '<>"","Web","")),' + P('Fecha/Hora') + ',' + P('Nombre y apellido') + ',' +
    P('División') + ',' + P('Talle') + ',' + P('WhatsApp') + ',' + P('Estado') + '},' + P('ID') + '<>""),{"","","","","","",""})';
  t.getRange('A2').setFormula('=LET(a,{' + forms + ';' + web + '},FILTER(a,INDEX(a,,3)<>""))');
  t.getRange('B2:B').setNumberFormat('dd/MM/yyyy HH:mm:ss');
  t.setFrozenRows(1);
  if (t.getFilter()) t.getFilter().remove();
  t.getRange('A1:G1000').createFilter();
  t.autoResizeColumns(1, 7);
}

/** Hoja "Dashboard": cuantas musculosas por division x talle (lo que hay que pedirle al proveedor) y estado de pagos. */
function armarDashboard_(ss) {
  var d = ss.getSheetByName('Dashboard') || ss.insertSheet('Dashboard', 0);
  d.clear();
  d.setHiddenGridlines(true);
  var head = function (r, c, n) {
    d.getRange(r, c, 1, n).setFontWeight('bold').setBackground('#0B1F4B').setFontColor('#FFFFFF').setHorizontalAlignment('center');
  };
  var T = 'Todos!';
  var nT = TALLES.length, nD = DIVISIONES.length;

  d.getRange('A1').setValue('Musculosa Los Guerreros 2026').setFontSize(16).setFontWeight('bold');

  // KPIs
  d.getRange('A3:E3').setValues([['Reservadas', 'Pago confirmado', 'Comprobante por revisar', 'Pendientes de pago', 'Cancelados']]);
  head(3, 1, 5);
  d.getRange('A4:E4').setFormulas([['=COUNTA(' + T + 'C2:C)-COUNTIFS(' + T + 'G2:G,"Cancelado")',
    '=COUNTIFS(' + T + 'G:G,"Pago confirmado")', '=COUNTIFS(' + T + 'G:G,"Comprobante recibido")',
    '=COUNTIFS(' + T + 'G:G,"Pendiente de pago")', '=COUNTIFS(' + T + 'G:G,"Cancelado")']])
    .setFontSize(18).setFontWeight('bold').setHorizontalAlignment('center');

  // Matriz 1: todas las reservas (sin canceladas)
  var matriz = function (fila, titulo, estadoCrit) {
    d.getRange(fila, 1).setValue(titulo).setFontWeight('bold').setFontSize(12);
    var enc = ['División'].concat(TALLES, ['Total']);
    d.getRange(fila + 1, 1, 1, enc.length).setValues([enc]);
    head(fila + 1, 1, enc.length);
    var rows = [], etiquetas = [];
    DIVISIONES.forEach(function (div, i) {
      var r = fila + 2 + i, f = [];
      etiquetas.push([div]);
      for (var c = 0; c < nT; c++) {
        f.push('=COUNTIFS(' + T + '$D$2:$D,$A' + r + ',' + T + '$E$2:$E,' + letra_(c + 2) + '$' + (fila + 1) + ',' + estadoCrit + ')');
      }
      f.push('=SUM(B' + r + ':' + letra_(nT + 1) + r + ')');
      rows.push(f);
    });
    var rt = fila + 2 + nD, tot = [];
    etiquetas.push(['Total']);
    for (var c = 0; c <= nT; c++) tot.push('=SUM(' + letra_(c + 2) + (fila + 2) + ':' + letra_(c + 2) + (rt - 1) + ')');
    rows.push(tot);
    // Las etiquetas van con setValues: en setFormulas, 'PS' o 'M15' se tomaban como formula (=PS, =M15).
    d.getRange(fila + 2, 1, etiquetas.length, 1).setValues(etiquetas);
    d.getRange(fila + 2, 2, rows.length, nT + 1).setFormulas(rows).setHorizontalAlignment('center');
    d.getRange(rt, 1, 1, nT + 2).setFontWeight('bold').setBorder(true, null, null, null, null, null);
    d.getRange(fila + 2, nT + 2, nD + 1, 1).setFontWeight('bold');
    d.getRange(fila + 2, 1, nD, 1).setFontWeight('bold').setHorizontalAlignment('left');
    var rg = d.getRange(fila + 2, 2, nD, nT);
    d.setConditionalFormatRules(d.getConditionalFormatRules().concat([SpreadsheetApp.newConditionalFormatRule()
      .setGradientMinpoint('#FFFFFF').setGradientMaxpoint('#F4A3A9').setRanges([rg]).build()]));
  };
  matriz(6, 'Todas las reservas por división y talle (sin canceladas)', T + '$G$2:$G,"<>Cancelado"');
  var f2 = 6 + 2 + nD + 3;
  matriz(f2, 'Solo pagos confirmados por división y talle', T + '$G$2:$G,"Pago confirmado"');

  // Estado de pagos por division
  var f3 = f2 + 2 + nD + 3;
  d.getRange(f3, 1).setValue('Estado por división').setFontWeight('bold').setFontSize(12);
  d.getRange(f3 + 1, 1, 1, 4).setValues([['División', 'Pago confirmado', 'Comprobante por revisar', 'Pendiente de pago']]);
  head(f3 + 1, 1, 4);
  var rows3 = DIVISIONES.map(function (div, i) {
    var r = f3 + 2 + i;
    return ['Pago confirmado', 'Comprobante recibido', 'Pendiente de pago'].map(function (e) {
      return '=COUNTIFS(' + T + '$D$2:$D,$A' + r + ',' + T + '$G$2:$G,"' + e + '")';
    });
  });
  d.getRange(f3 + 2, 1, nD, 1).setValues(DIVISIONES.map(function (x) { return [x]; }));
  d.getRange(f3 + 2, 2, nD, 3).setFormulas(rows3).setHorizontalAlignment('center');
  d.getRange(f3 + 2, 1, nD, 1).setFontWeight('bold').setHorizontalAlignment('left');
  d.setColumnWidths(1, 1, 130);
  d.setColumnWidths(2, 8, 95);
}

/** Da acceso de solo lectura a la carpeta de comprobantes. Ejecutalo una vez desde el editor. Los archivos nuevos heredan el acceso. */
function compartirComprobantes() {
  var emails = ['lizzanojuanignacio@gmail.com'];
  var f = carpeta_();
  emails.forEach(function (e) { f.addViewer(e); });
  console.log('Carpeta compartida (lectura) con: ' + emails.join(', ') + ' -> ' + f.getUrl());
}
