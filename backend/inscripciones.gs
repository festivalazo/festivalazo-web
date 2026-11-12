/**
 * Festivalazo 2027 · Receptor de inscripciones
 * Proyecto de Apps Script en la cuenta festivalaz01@gmail.com, publicado como aplicación web.
 * Cada inscripción de festivalazo.com/inscripcion.html se guarda como fila nueva en la planilla
 * "Festivalazo 2027 · Inscripciones" (se crea sola la primera vez), y llega un mail a la organización.
 * A la persona inscripta le llega un mail con los datos para pagar.
 */
var ORG_EMAIL = 'festivalaz01@gmail.com';
var SHEET_NAME = 'Festivalazo 2027 · Inscripciones';

var COLUMNS = [
  ['fecha', 'Fecha'], ['nombre', 'Nombre y apellido'], ['email', 'Email'], ['whatsapp', 'Teléfono'],
  ['instagram', 'Instagram'], ['desde', 'De dónde viene'], ['emergencia', 'Contacto de emergencia'],
  ['entrada', 'Entrada'], ['precio', 'Precio'], ['pago', 'Cómo paga'], ['dupla', 'Compa de dupla'],
  ['remera', 'Remera'], ['residente', 'Amigo/a de residente'], ['clases', 'Clases regulares'],
  ['acroyoga', 'Experiencia acroyoga'], ['acroduo', 'Experiencia acroduo'], ['verticales', 'Paradas de manos'],
  ['intereses', 'Qué le interesa'], ['cuidado', 'Cuidando (1-10)'], ['otra', 'Otra práctica'],
  ['variete', 'Propuesta varieté'], ['salud', 'Salud / lesiones'], ['acuerdo', 'Acuerdo de cuidado'],
  ['politica', 'Acepta política de cancelación'], ['mensaje', 'Mensaje'],
  ['pagado', 'Pago recibido'], ['notas', 'Notas']
];

function getSheet_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('SHEET_ID');
  var ss = id ? SpreadsheetApp.openById(id) : null;
  if (!ss) {
    ss = SpreadsheetApp.create(SHEET_NAME);
    props.setProperty('SHEET_ID', ss.getId());
  }
  var sheet = ss.getSheets()[0];
  if (sheet.getRange(1, 1).getValue() !== 'Fecha') {
    sheet.getRange(1, 1, 1, COLUMNS.length).setValues([COLUMNS.map(function (c) { return c[1]; })])
      .setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function money_(n) {
  return '$' + Number(n || 0).toLocaleString('es-AR');
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
    getSheet_().appendRow(COLUMNS.map(function (c) { return d[c[0]] == null ? '' : d[c[0]]; }));

    MailApp.sendEmail({
      to: ORG_EMAIL,
      replyTo: d.email,
      subject: 'Nueva inscripción Festivalazo 2027 · ' + d.nombre + ' · ' + d.entrada,
      body: COLUMNS.filter(function (c) { return d[c[0]]; })
        .map(function (c) { return c[1] + ': ' + d[c[0]]; }).join('\n')
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
        + '\nDónde pagar:\nAlias: festivalazo.elorigen\nCuenta Naranja, a nombre de Aricelli Villoldo\n\n'
        + 'Tu lugar queda reservado cuando recibamos el pago. Respondé este mail con el comprobante.\n\n'
        + 'Recordá que el camping se abona aparte, en efectivo, en el lugar.\n\n'
        + 'Un abrazo grande,\nEquipo Festivalazo'
    });

    return ContentService.createTextOutput(JSON.stringify({ ok: true })).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: String(err) })).setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

// Para probar desde el editor: crea la planilla y deja la fila de títulos.
function preparar() {
  Logger.log(getSheet_().getParent().getUrl());
}
