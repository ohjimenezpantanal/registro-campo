// ═══════════════════════════════════════════════════════════════
// APPS SCRIPT — Hacienda Pantanal
// Versión con alertas por email
// ═══════════════════════════════════════════════════════════════
// INSTRUCCIONES SI SE BORRA:
// 1. Extensiones → Apps Script
// 2. Borre todo (Cmd+A → Delete)
// 3. Pegue este código completo
// 4. Guarde (Cmd+S)
// 5. Implementar → Gestionar implementaciones → lápiz ✏️
// 6. Nueva versión → Implementar
// ═══════════════════════════════════════════════════════════════
 
const EMAIL_ADMIN = 'ohjimenez.pantanal@gmail.com';
 
function doPost(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const data = JSON.parse(e.postData.contents);
 
    if (data.tipo === 'test') return respuesta(true, 'Conexion OK');
 
    let sheetName;
    if (data.origen === 'trabajador') {
      sheetName = '⏳ Pendientes';
    } else {
      switch(data.tipo) {
        case 'actividades': sheetName = 'Actividades'; break;
        case 'insumos':     sheetName = 'Insumos';     break;
        case 'otros':       sheetName = 'OtrosPagos';  break;
        case 'inventario':  sheetName = 'Inventario';  break;
        default: return respuesta(false, 'Tipo desconocido');
      }
    }
 
    const ws = ss.getSheetByName(sheetName);
    if (!ws) return respuesta(false, 'Pestana no encontrada: ' + sheetName);
 
    const colA = ws.getRange('A1:A' + Math.max(ws.getLastRow() + 1, 6)).getValues();
    let nextRow = 6;
    for (let i = 5; i < colA.length; i++) {
      if (colA[i][0] === '' || colA[i][0] === null) { nextRow = i + 1; break; }
    }
 
    const fila = data.fila;
    for (let i = 0; i < fila.length; i++) {
      let val = fila[i];
      if (typeof val === 'string' && val.match(/^[\d.,]+$/)) {
        val = parseFloat(val.replace(',', '.')) || val;
      }
      ws.getRange(nextRow, i + 1).setValue(val);
    }
 
    if (sheetName === '⏳ Pendientes') {
      ws.getRange(nextRow, 10).setValue('⏳ Pendiente');
    }
 
    if (data.tipo === 'inventario') {
      const colA2 = ws.getRange(2, 1, nextRow, 1).getValues();
      let maxNum = 0;
      colA2.forEach(r => { if (typeof r[0] === 'number' && r[0] > maxNum) maxNum = r[0]; });
      ws.getRange(nextRow, 1).setValue(maxNum + 1);
    }
 
    return respuesta(true, 'Guardado en ' + sheetName + ', fila ' + nextRow);
 
  } catch (err) {
    try {
      MailApp.sendEmail({
        to: EMAIL_ADMIN,
        subject: '⚠️ Error en Registro de Campo — Hacienda Pantanal',
        body: 'Error al guardar un registro:\n\n' + err.message +
              '\n\nRevise el Apps Script en Google Sheets.'
      });
    } catch(mailErr) {}
    return respuesta(false, 'Error: ' + err.message);
  }
}
 
function verificacionDiaria() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const wa = ss.getSheetByName('Actividades');
    const wp = ss.getSheetByName('⏳ Pendientes');
    const totalAct  = wa ? wa.getLastRow() - 5 : 0;
    const totalPend = wp ? wp.getLastRow() - 5 : 0;
    const fecha = new Date().toLocaleDateString('es-EC');
    const hora  = new Date().toLocaleTimeString('es-EC');
 
    MailApp.sendEmail({
      to: EMAIL_ADMIN,
      subject: '✅ Sistema OK — Hacienda Pantanal ' + fecha,
      body: 'Buenos días, Ing. Oscar.\n\n' +
            'El sistema de registro está funcionando correctamente.\n\n' +
            '📊 Estado actual:\n' +
            '• Actividades registradas: ' + totalAct + '\n' +
            '• Pendientes por aprobar: ' + totalPend + '\n\n' +
            'Verificado el ' + fecha + ' a las ' + hora + '.\n\n' +
            '— Sistema Hacienda Pantanal'
    });
  } catch(err) {
    MailApp.sendEmail({
      to: EMAIL_ADMIN,
      subject: '❌ ALERTA — Sistema Hacienda Pantanal no responde',
      body: 'No se pudo verificar el sistema.\n\nError: ' + err.message +
            '\n\nRevise el Apps Script inmediatamente.'
    });
  }
}
 
function respuesta(ok, msg) {
  return ContentService
    .createTextOutput(JSON.stringify({ ok: ok, msg: msg }))
    .setMimeType(ContentService.MimeType.JSON);
}
 
