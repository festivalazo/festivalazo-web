/**
 * Festivalazo 2027 · Receptor de inscripciones y comidas
 * Proyecto de Apps Script en la cuenta festivalaz01@gmail.com, publicado como aplicación web.
 * Una sola planilla, "Festivalazo 2027 · Inscripciones", con dos pestañas:
 *   - Inscripciones: lo que llega de festivalazo.com/inscripcion.html
 *   - Comidas: lo que llega de festivalazo.com/comida.html (campo tipo = "comida")
 * Las dos se cruzan por email: en Inscripciones aparece qué comida pidió cada persona,
 * y en Comidas aparece si esa persona ya está inscripta al festival.
 */
var ORG_EMAIL = 'festivalaz01@gmail.com';
var SHEET_NAME = 'Festivalazo 2027 · Inscripciones';
var PAGO = '\nDónde pagar:\nAlias: festivalazo.elorigen\nCuenta Naranja, a nombre de Aricelli Villoldo\n\n';

var TABS = {
  Inscripciones: [
    ['fecha', 'Fecha'], ['nombre', 'Nombre y apellido'], ['email', 'Email'], ['whatsapp', 'Teléfono'],
    ['instagram', 'Instagram'], ['desde', 'De dónde viene'], ['emergencia', 'Contacto de emergencia'],
    ['entrada', 'Entrada'], ['precio', 'Precio'], ['pago', 'Cómo paga'], ['dupla', 'Compa de dupla'],
    ['remera', 'Remera'], ['residente', 'Amigo/a de residente'], ['clases', 'Clases regulares'],
    ['acroyoga', 'Experiencia acroyoga'], ['acroduo', 'Experiencia acroduo'], ['verticales', 'Paradas de manos'],
    ['intereses', 'Qué le interesa'], ['cuidado', 'Cuidando (1-10)'], ['otra', 'Otra práctica'],
    ['variete', 'Propuesta varieté'], ['salud', 'Salud / lesiones'], ['acuerdo', 'Acuerdo de cuidado'],
    ['politica', 'Acepta política de cancelación'], ['mensaje', 'Mensaje'],
    ['pagado', 'Pago recibido'], ['notas', 'Notas'], ['comida', 'Comida pedida']
  ],
  Comidas: [
    ['fecha', 'Fecha'], ['nombre', 'Nombre y apellido'], ['email', 'Email'], ['whatsapp', 'Teléfono'],
    ['opcion', 'Opción'], ['precio', 'Precio'], ['sueltas', 'Comidas sueltas'], ['dieta', 'Dieta'],
    ['sintacc', 'Sin TACC'], ['alergias', 'Alergias'], ['mensaje', 'Mensaje'],
    ['inscripto', '¿Inscripto al festival?'], ['pagado', 'Pago recibido'], ['notas', 'Notas']
  ]
};

function getBook_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('SHEET_ID');
  var ss = id ? SpreadsheetApp.openById(id) : null;
  if (!ss) {
    ss = SpreadsheetApp.create(SHEET_NAME);
    props.setProperty('SHEET_ID', ss.getId());
  }
  return ss;
}

// Devuelve la pestaña pedida, creándola o completando los títulos si hace falta.
function getTab_(ss, name) {
  var sheet = ss.getSheetByName(name);
  if (!sheet && name === 'Inscripciones') {
    sheet = ss.getSheets()[0];
    sheet.setName(name);
  }
  if (!sheet) sheet = ss.insertSheet(name);
  var cols = TABS[name];
  var head = sheet.getRange(1, 1, 1, cols.length);
  var want = cols.map(function (c) { return c[1]; });
  if (head.getValues()[0].join('|') !== want.join('|')) {
    head.setValues([want]).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// Busca la fila (1-based) de un email en la columna Email de una pestaña; 0 si no está.
function findEmail_(sheet, name, email) {
  if (!email || sheet.getLastRow() < 2) return 0;
  var col = colOf_(name, 'email');
  var vals = sheet.getRange(2, col, sheet.getLastRow() - 1, 1).getValues();
  var e = String(email).trim().toLowerCase();
  for (var i = vals.length - 1; i >= 0; i--) {
    if (String(vals[i][0]).trim().toLowerCase() === e) return i + 2;
  }
  return 0;
}

function colOf_(name, key) {
  var cols = TABS[name];
  for (var i = 0; i < cols.length; i++) if (cols[i][0] === key) return i + 1;
  return 0;
}

function money_(n) {
  return n ? '$' + Number(n).toLocaleString('es-AR') : 'a confirmar';
}

function rowFor_(name, d) {
  return TABS[name].map(function (c) { return d[c[0]] == null ? '' : d[c[0]]; });
}

function bodyFor_(name, d) {
  return TABS[name].filter(function (c) { return d[c[0]]; })
    .map(function (c) { return c[1] + ': ' + d[c[0]]; }).join('\n');
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var d = JSON.parse(e.postData.contents);
    d.fecha = new Date();
    // Evita que un texto que empiece con = + - @ se tome como fórmula
    Object.keys(d).forEach(function (k) {
      if (typeof d[k] === 'string' && /^[=+\-@]/.test(d[k])) d[k] = "'" + d[k];
    });
    var ss = getBook_();
    var ins = getTab_(ss, 'Inscripciones');
    var com = getTab_(ss, 'Comidas');

    if (d.tipo === 'comida') guardarComida_(d, ins, com);
    else guardarInscripcion_(d, ins, com);

    return ContentService.createTextOutput(JSON.stringify({ ok: true })).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: String(err) })).setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

function guardarInscripcion_(d, ins, com) {
  var r = findEmail_(com, 'Comidas', d.email);
  if (r) {
    d.comida = com.getRange(r, colOf_('Comidas', 'opcion')).getValue();
    com.getRange(r, colOf_('Comidas', 'inscripto')).setValue('Sí');
  }
  ins.appendRow(rowFor_('Inscripciones', d));

  MailApp.sendEmail({
    to: ORG_EMAIL,
    replyTo: d.email,
    subject: 'Nueva inscripción Festivalazo 2027 · ' + d.nombre + ' · ' + d.entrada,
    body: bodyFor_('Inscripciones', d)
  });

  var dosCuotas = String(d.pago || '').indexOf('Dos') === 0;
  MailApp.sendEmail({
    to: d.email,
    replyTo: ORG_EMAIL,
    name: 'Festivalazo',
    subject: 'Festivalazo 2027 · Recibimos tu inscripción',
    body: 'Hola ' + d.nombre + ',\n\n'
      + '¡Gracias por inscribirte a Festivalazo 2027, la décima edición! Nos vemos del 8 al 11 de octubre.\n\n'
      + 'Tu entrada: ' + d.entrada + ' · ' + money_(d.precio) + '\n'
      + (dosCuotas ? 'Elegiste pagar en dos cuotas: 50% ahora y 50% antes del 10 de septiembre.\n' : '')
      + PAGO
      + 'Tu lugar queda reservado cuando recibamos el pago. Respondé este mail con el comprobante.\n\n'
      + (d.comida ? '' : '¿Querés sumar la comida del finde? Anotate en https://festivalazo.com/comida.html\n\n')
      + 'Recordá que el camping se abona aparte, en efectivo, en el lugar.\n\n'
      + 'Un abrazo grande,\nEquipo Festivalazo'
  });
}

function guardarComida_(d, ins, com) {
  var r = findEmail_(ins, 'Inscripciones', d.email);
  d.inscripto = r ? 'Sí' : 'Todavía no';
  if (r) ins.getRange(r, colOf_('Inscripciones', 'comida')).setValue(d.opcion);
  com.appendRow(rowFor_('Comidas', d));

  MailApp.sendEmail({
    to: ORG_EMAIL,
    replyTo: d.email,
    subject: 'Nuevo pedido de comida Festivalazo 2027 · ' + d.nombre + ' · ' + d.opcion,
    body: bodyFor_('Comidas', d)
  });

  MailApp.sendEmail({
    to: d.email,
    replyTo: ORG_EMAIL,
    name: 'Festivalazo',
    subject: 'Festivalazo 2027 · Recibimos tu pedido de comida',
    body: 'Hola ' + d.nombre + ',\n\n'
      + '¡Gracias por sumarte a la comida de Festivalazo 2027!\n\n'
      + 'Tu pedido: ' + d.opcion + (d.sueltas ? ' (' + d.sueltas + ')' : '') + ' · ' + d.dieta
      + (d.sintacc ? ' · sin TACC' : '') + '\n'
      + 'Importe: ' + (d.precio ? money_(d.precio) : 'te lo confirmamos en otro mail') + '\n'
      + (d.precio ? PAGO + 'Respondé este mail con el comprobante.\n\n' : '\n')
      + (r ? '' : 'Todavía no encontramos tu inscripción al festival con este email. Si no te inscribiste, hacelo en https://festivalazo.com/inscripcion.html\n\n')
      + 'Un abrazo grande,\nEquipo Festivalazo'
  });
}

// Para probar desde el editor: crea la planilla y deja las dos pestañas con sus títulos.
function preparar() {
  var ss = getBook_();
  getTab_(ss, 'Inscripciones');
  getTab_(ss, 'Comidas');
  Logger.log(ss.getUrl());
}
