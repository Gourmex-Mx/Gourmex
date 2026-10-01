/**
 * BACKEND · ENCUESTA COMEDOR GOURMEX
 * Recibe las respuestas de la página y las guarda en esta Google Sheet.
 *
 * 1. Crea una Google Sheet nueva: "Respuestas Encuesta GOURMEX"
 * 2. Menú Extensiones → Apps Script. Borra todo y pega este código. Guarda.
 * 3. Implementar → Nueva implementación → tipo "Aplicación web"
 *      Ejecutar como: Yo
 *      Quién tiene acceso: Cualquier persona
 * 4. Copia la URL que termina en /exec y pégala en index.html (SCRIPT_URL)
 */

var HOJA = 'Respuestas';
var COLS = ['Fecha y hora de envío','No. Nómina','Turno','Fecha','Sabor','Variedad del menú','Servicio en barra','Atención del personal','Platillo sugerido'];

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(HOJA) || ss.insertSheet(HOJA);
    if (sh.getLastRow() === 0) {
      sh.appendRow(COLS);
      sh.getRange(1, 1, 1, COLS.length).setFontWeight('bold').setBackground('#113F70').setFontColor('#FFFFFF');
      sh.setFrozenRows(1);
    }
    var p = e.parameter;
    sh.appendRow([new Date(), p.nomina, p.turno, p.fecha, p.sabor, p.variedad, p.barra, p.atencion, p.platillo]);
    return ContentService.createTextOutput('ok');
  } finally {
    lock.releaseLock();
  }
}
