// ── Axion App Parking — Google Apps Script ──────────────────────────────────
// Pegar este código en: script.google.com → Nuevo proyecto
// Luego: Implementar → Nueva implementación → Aplicación web
//   · Ejecutar como: Yo
//   · Quién tiene acceso: Cualquier usuario
// Copiar la URL generada y pegarla en el HTML (variable SHEET_URL)
// ─────────────────────────────────────────────────────────────────────────────

const SHEET_NAME = 'Parking';

function doPost(e) {
  try {
    const data  = JSON.parse(e.postData.contents);
    guardarFila(data);
    return respuesta({ ok: true });
  } catch (err) {
    return respuesta({ ok: false, error: err.message });
  }
}

function doGet(e) {
  const p = e.parameter;
  if (p && p.tipo) {
    try {
      guardarFila(p);
      return respuesta({ ok: true });
    } catch (err) {
      return respuesta({ ok: false, error: err.message });
    }
  }
  return respuesta({ ok: true, msg: 'API Parking activa' });
}

// ── Helpers ──────────────────────────────────────────────────────────────────

function guardarFila(data) {
  const sheet = getOrCreateSheet();
  const ahora = new Date().toLocaleString('es-AR', {
    timeZone: 'America/Montevideo',
    hour12: false,
    day: '2-digit', month: '2-digit', year: 'numeric',
    hour: '2-digit', minute: '2-digit'
  });
  sheet.appendRow([
    ahora,
    data.tipo,
    data.patente,
    data.entrada,
    data.salida,
    Number(data.facturadas),
    data.tarifa,
    Number(data.monto)
  ]);
}

function getOrCreateSheet() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let sheet   = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow([
      'Timestamp', 'Tipo', 'Patente',
      'Ingreso', 'Salida', 'Horas', 'Tarifa', 'Monto'
    ]);
    // Formato encabezado
    const header = sheet.getRange(1, 1, 1, 8);
    header.setFontWeight('bold');
    header.setBackground('#C4006A');
    header.setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  }

  return sheet;
}

function respuesta(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
