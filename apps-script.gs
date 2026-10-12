/**
 * Recibe las respuestas de la encuesta y las guarda en la hoja de cálculo.
 * Pegar en: Hoja de Google > Extensiones > Apps Script.
 */
const HOJA = 'Respuestas';
const ENCABEZADOS = ['Fecha','Rol','Experiencia','P1','P2','P3','P4','P5','P6','P7','P8'];

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const d = JSON.parse(e.postData.contents);
    if (!Array.isArray(d.resp) || d.resp.length !== 8) throw new Error('Datos inválidos');
    const resp = d.resp.map(Number);
    if (resp.some(v => !(v >= 1 && v <= 5))) throw new Error('Valores fuera de rango');

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(HOJA) || ss.insertSheet(HOJA);
    if (sh.getLastRow() === 0) {
      sh.appendRow(ENCABEZADOS);
      sh.getRange(1, 1, 1, ENCABEZADOS.length).setFontWeight('bold');
      sh.setFrozenRows(1);
    }
    sh.appendRow([new Date(), String(d.rol).slice(0, 80), String(d.exp).slice(0, 40)].concat(resp));
    return salida({ ok: true });
  } catch (err) {
    return salida({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// Abrir la URL en el navegador para comprobar que está activo
function doGet() { return salida({ estado: 'activo' }); }

function salida(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}