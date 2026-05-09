// ═══════════════════════════════════════════════════════════════
// APPS SCRIPT — Hacienda Pantanal
// Versión completa con Cuaderno de Secretaria
// ═══════════════════════════════════════════════════════════════
// INSTRUCCIONES SI SE BORRA:
// 1. Extensiones → Apps Script
// 2. Borre todo (Cmd+A → Delete)
// 3. Pegue este código completo
// 4. Guarde (Cmd+S)
// 5. Ejecutar onOpen → aceptar permisos → recargar Sheets
// 6. Para el doPost: Implementar → Gestionar implementaciones
//    → lápiz ✏️ → Nueva versión → Implementar
// ═══════════════════════════════════════════════════════════════

const EMAIL_ADMIN = 'ohjimenez.pantanal@gmail.com';

// ── MENÚ PRINCIPAL ────────────────────────────────────────────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🌴 Pantanal Agro')
    .addItem('📊 Generar Reporte Ejecutivo',  'generarReporteManual')
    .addItem('📄 Procesar Facturas Siigo',    'procesarFacturasManual')
    .addItem('✅ Verificar Sistema',           'verificacionDiaria')
    .addSeparator()
    .addItem('📋 Crear / Recrear Cuaderno',   'crearCuaderno')
    .addItem('👤 Agregar empleado nuevo',     'agregarEmpleado')
    .addItem('⚙️ Agregar actividad nueva',    'agregarActividad')
    .addSeparator()
    .addItem('🔄 Ejecutar conciliación',      'ejecutarConciliacion')
    .addToUi();
}


// ═══════════════════════════════════════════════════════════════
// FORMULARIO — doPost
// ═══════════════════════════════════════════════════════════════

function doPost(e) {
  try {
    const ss   = SpreadsheetApp.getActiveSpreadsheet();
    const data = JSON.parse(e.postData.contents);

    if (data.tipo === 'test') return respuesta(true, 'Conexion OK');

    let sheetName;
    if (data.origen === 'trabajador') {
      sheetName = '⏳ Pendientes';
    } else if (data.origen === 'secretaria') {
      sheetName = '📋 Cuaderno';
    } else {
      switch (data.tipo) {
  case 'actividades': sheetName = 'Actividades'; break;
  case 'insumos':     sheetName = 'Insumos';     break;
  case 'otros':       sheetName = 'OtrosPagos';  break;
  case 'inventario':  sheetName = 'Inventario';  break;
  case 'cuaderno':    sheetName = '📋 Cuaderno'; break;
  default: return respuesta(false, 'Tipo desconocido');
}
}

    const ws = ss.getSheetByName(sheetName);
    if (!ws) return respuesta(false, 'Pestana no encontrada: ' + sheetName);

    // Buscar siguiente fila vacía
    const colBuscar = sheetName === '📋 Cuaderno' ? 'B' : 'A';
    const colVals   = ws.getRange(colBuscar + '1:' + colBuscar + Math.max(ws.getLastRow() + 1, 6)).getValues();
    let nextRow = 6;
    for (let i = 5; i < colVals.length; i++) {
      if (colVals[i][0] === '' || colVals[i][0] === null) { nextRow = i + 1; break; }
    }

    const fila = data.fila;
Logger.log('origen: ' + data.origen + ' | fila: ' + JSON.stringify(fila));
    if (sheetName === '📋 Cuaderno') {
      // B=fecha, C=finca, D=empleado, E=actividad, F=cantidad, G=precio, I=notas
      const cols = [2, 3, 4, 5, 6, 7, 9];
      for (let i = 0; i < fila.length && i < cols.length; i++) {
        let val = fila[i];
        if (typeof val === 'string' && val.match(/^[\d.,]+$/)) {
          val = parseFloat(val.replace(',', '.')) || val;
        }
        ws.getRange(nextRow, cols[i]).setValue(val);
      }
    } else {
      // Resto de hojas: escribir desde col A
      for (let i = 0; i < fila.length; i++) {
        let val = fila[i];
        if (typeof val === 'string' && val.match(/^[\d.,]+$/)) {
          val = parseFloat(val.replace(',', '.')) || val;
        }
        ws.getRange(nextRow, i + 1).setValue(val);
      }
    }

    if (sheetName === '⏳ Pendientes') {
      ws.getRange(nextRow, 10).setValue('⏳ Pendiente');
    }

    if (data.tipo === 'inventario') {
      const colA2 = ws.getRange(2, 1, nextRow, 1).getValues();
      let maxNum  = 0;
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
    } catch (mailErr) {}
    return respuesta(false, 'Error: ' + err.message);
  }
}

function respuesta(ok, msg) {
  return ContentService
    .createTextOutput(JSON.stringify({ ok: ok, msg: msg }))
    .setMimeType(ContentService.MimeType.JSON);
}


// ═══════════════════════════════════════════════════════════════
// VERIFICACIÓN DIARIA
// ═══════════════════════════════════════════════════════════════

function verificacionDiaria() {
  try {
    const ss        = SpreadsheetApp.getActiveSpreadsheet();
    const wa        = ss.getSheetByName('Actividades');
    const wp        = ss.getSheetByName('⏳ Pendientes');
    const totalAct  = wa ? wa.getLastRow() - 5 : 0;
    const totalPend = wp ? wp.getLastRow() - 5 : 0;
    const fecha     = new Date().toLocaleDateString('es-EC');
    const hora      = new Date().toLocaleTimeString('es-EC');

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
  } catch (err) {
    MailApp.sendEmail({
      to: EMAIL_ADMIN,
      subject: '❌ ALERTA — Sistema Hacienda Pantanal no responde',
      body: 'No se pudo verificar el sistema.\n\nError: ' + err.message +
            '\n\nRevise el Apps Script inmediatamente.'
    });
  }
}


// ═══════════════════════════════════════════════════════════════
// GENERADOR DE REPORTES EJECUTIVOS
// ═══════════════════════════════════════════════════════════════

const EMAIL_REPORTE = 'ohjimenez.pantanal@gmail.com';
const MESES = ['Enero','Febrero','Marzo','Abril','Mayo','Junio',
               'Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'];

function generarReporteManual() {
  const ui = SpreadsheetApp.getUi();

  const desde = ui.prompt(
    'Reporte Ejecutivo — Pantanal Agro',
    'Ingrese el MES DE INICIO (1-12):',
    ui.ButtonSet.OK_CANCEL
  );
  if (desde.getSelectedButton() !== ui.Button.OK) return;

  const hasta = ui.prompt(
    'Reporte Ejecutivo — Pantanal Agro',
    'Ingrese el MES DE FIN (1-12):',
    ui.ButtonSet.OK_CANCEL
  );
  if (hasta.getSelectedButton() !== ui.Button.OK) return;

  const mesInicio = parseInt(desde.getResponseText());
  const mesFin    = parseInt(hasta.getResponseText());

  if (isNaN(mesInicio) || isNaN(mesFin) || mesInicio < 1 || mesFin > 12 || mesInicio > mesFin) {
    ui.alert('Período inválido. Ingrese números del 1 al 12.');
    return;
  }

  const periodo = MESES[mesInicio - 1] + ' - ' + MESES[mesFin - 1] + ' 2026';
  ui.alert('Generando reporte para: ' + periodo + '\n\nRecibirá el PDF en su email en unos segundos.');
  generarYEnviarReporte(mesInicio, mesFin, periodo);
}

function reporteAutomaticoMensual() {
  const hoy     = new Date();
  const mesAct  = hoy.getMonth() + 1;
  const mesAnt  = mesAct === 1 ? 12 : mesAct - 1;
  const periodo = MESES[mesAnt - 1] + ' 2026';
  generarYEnviarReporte(mesAnt, mesAnt, periodo);
}

function generarYEnviarReporte(mesInicio, mesFin, periodo) {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const wd  = ss.getSheetByName('📈 Dashboard');
  const wv  = ss.getSheetByName('💵 Ventas');
  const we  = ss.getSheetByName('📊 EBITDA');

  const tmTotal  = wd.getRange('A97').getValue() || 0;
  const ctTotal  = wd.getRange('C97').getValue() || 0;
  const ctTm     = wd.getRange('G97').getValue() || 0;
  const ingTotal = wv ? wv.getRange('F106').getValue() || 0 : 0;
  const utilidad = ingTotal - ctTotal;
  const margen   = ingTotal > 0 ? (utilidad / ingTotal * 100) : 0;
  const ebitda   = we ? we.getRange('B35').getValue() || 0 : utilidad;

  const ventas_mes = [];
  for (let mo = mesInicio; mo <= mesFin; mo++) {
    const r = 110 + mo;
    if (wv) {
      ventas_mes.push([
        MESES[mo - 1],
        wv.getRange(r, 2).getValue() || 0,
        wv.getRange(r, 3).getValue() || 0,
        wv.getRange(r, 4).getValue() || 0,
        wv.getRange(r, 5).getValue() || 0,
        wv.getRange(r, 6).getValue() || 0
      ]);
    }
  }

  const saldoPrest = we ? we.getRange('B22').getValue() || 37000 : 37000;
  const cuotaPrest = we ? we.getRange('B23').getValue() || 4000  : 4000;

  const html    = construirHtmlReporte({ periodo, tmTotal, ctTotal, ctTm, ingTotal,
                                         utilidad, margen, ebitda, ventas_mes,
                                         saldoPrest, cuotaPrest, mesInicio, mesFin });
  const blob    = Utilities.newBlob(html, 'text/html', 'reporte.html');
  const pdfBlob = blob.getAs('application/pdf');
  const fileName = 'Pantanal_Agro_Reporte_' + periodo.replace(/ /g, '_') + '.pdf';
  pdfBlob.setName(fileName);

  const folder  = obtenerCarpetaReportes();
  const file    = folder.createFile(pdfBlob);
  const fileUrl = file.getUrl();

  GmailApp.sendEmail(
    EMAIL_REPORTE,
    '📊 Reporte Ejecutivo — ' + periodo + ' | Pantanal Agro',
    'Estimado Ing. Oscar,\n\n' +
    'Adjunto el reporte ejecutivo del período ' + periodo + '.\n\n' +
    'RESUMEN:\n' +
    '• TM producidas: ' + tmTotal.toFixed(2) + ' TM\n' +
    '• Ingresos: $' + ingTotal.toLocaleString() + '\n' +
    '• Utilidad neta: $' + utilidad.toFixed(0) + '\n' +
    '• Margen: ' + margen.toFixed(1) + '%\n' +
    '• EBITDA: $' + ebitda.toFixed(0) + '\n\n' +
    'El reporte también fue guardado en Drive:\n' + fileUrl + '\n\n' +
    '— Sistema Pantanal Agro',
    { attachments: [pdfBlob] }
  );
}

function construirHtmlReporte(d) {
  const verde = '#1B4332'; const verdeM = '#2D6A4F'; const verdeCl = '#40916C';
  const verdePal = '#B7E4C7'; const naranja = '#F59E0B'; const azul = '#1A56DB';

  let filasVentas = '';
  let totalTm = 0, totalIng = 0, totalCt = 0, totalUt = 0;
  d.ventas_mes.forEach(([mes, tm, prec, ing, ct, ut]) => {
    totalTm += tm; totalIng += ing; totalCt += ct; totalUt += ut;
    filasVentas += `
      <tr>
        <td style="font-weight:600">${mes}</td>
        <td>${tm.toFixed(2)}</td>
        <td>$${prec.toFixed(0)}</td>
        <td>$${ing.toLocaleString('es-EC', { minimumFractionDigits: 0 })}</td>
        <td>$${ct.toFixed(0)}</td>
        <td style="color:${ut > 0 ? '#065F46' : '#991B1B'};font-weight:700">$${ut.toFixed(0)}</td>
      </tr>`;
  });
  filasVentas += `
    <tr style="background:${verdePal};font-weight:700">
      <td>TOTAL</td><td>${totalTm.toFixed(2)}</td><td>—</td>
      <td>$${totalIng.toLocaleString('es-EC', { minimumFractionDigits: 0 })}</td>
      <td>$${totalCt.toFixed(0)}</td>
      <td style="color:#065F46">$${totalUt.toFixed(0)}</td>
    </tr>`;

  return `<!DOCTYPE html><html><head><meta charset="UTF-8">
  <style>
    body{font-family:Arial,sans-serif;font-size:11px;color:#1A1A1A;margin:0;padding:0}
    .header{background:${verde};color:white;padding:16px 20px}
    .header-top{display:flex;justify-content:space-between;align-items:center}
    .empresa{font-size:18px;font-weight:700}
    .sub{font-size:10px;opacity:0.8;margin-top:3px}
    .naranja-bar{background:${naranja};height:4px}
    .kpis{display:flex;gap:10px;padding:12px 20px;background:#F8F9FA}
    .kpi{flex:1;background:white;border-radius:8px;padding:10px;text-align:center;border-top:4px solid ${verdeCl}}
    .kpi-val{font-size:18px;font-weight:700;color:${verde}}
    .kpi-lbl{font-size:9px;color:#666;text-transform:uppercase;margin-top:3px}
    .section{padding:10px 20px}
    .section-title{background:${verde};color:white;padding:5px 10px;border-radius:4px;font-weight:700;font-size:11px;margin-bottom:6px}
    table{width:100%;border-collapse:collapse;font-size:10px}
    th{background:${verdeM};color:white;padding:5px;text-align:center}
    td{padding:4px 6px;text-align:center;border-bottom:1px solid #eee}
    tr:nth-child(even){background:#F8F9FA}
    .er-row{display:flex;justify-content:space-between;padding:4px 8px;border-bottom:1px solid #eee}
    .er-bold{font-weight:700;background:${verdePal}}
    .two-col{display:grid;grid-template-columns:1fr 1fr;gap:12px}
    .footer{background:${verde};color:white;padding:8px 20px;font-size:9px;display:flex;justify-content:space-between;margin-top:10px}
  </style></head><body>
  <div class="header">
    <div class="header-top">
      <div><div class="empresa">LLH - Hacienda Pantanal</div>
      <div class="sub">Palma Africana | Quininde, Esmeraldas, Ecuador</div></div>
      <div style="text-align:right">
        <div style="font-size:13px;font-weight:700">REPORTE EJECUTIVO</div>
        <div style="font-size:10px;opacity:0.8">${d.periodo}</div>
      </div>
    </div>
  </div>
  <div class="naranja-bar"></div>
  <div class="kpis">
    <div class="kpi"><div class="kpi-val">${d.tmTotal.toFixed(1)} TM</div><div class="kpi-lbl">TM Producidas</div></div>
    <div class="kpi" style="border-top-color:${azul}"><div class="kpi-val">$${Math.round(d.ingTotal).toLocaleString()}</div><div class="kpi-lbl">Ingresos Brutos</div></div>
    <div class="kpi" style="border-top-color:${verdeM}"><div class="kpi-val">$${Math.round(d.ebitda).toLocaleString()}</div><div class="kpi-lbl">EBITDA</div></div>
    <div class="kpi" style="border-top-color:${naranja}"><div class="kpi-val">$${d.utilidad.toFixed(0)}</div><div class="kpi-lbl">Utilidad Neta (${d.margen.toFixed(1)}%)</div></div>
    <div class="kpi" style="border-top-color:#6B21A8"><div class="kpi-val">$${d.ctTm.toFixed(2)}</div><div class="kpi-lbl">CT / TM</div></div>
  </div>
  <div class="section">
    <div class="section-title">VENTAS Y PRODUCCION MENSUAL</div>
    <table>
      <tr><th>Mes</th><th>TM Vendidas</th><th>Precio $/TM</th><th>Ingreso ($)</th><th>Costo ($)</th><th>Margen ($)</th></tr>
      ${filasVentas}
    </table>
  </div>
  <div class="two-col" style="padding:0 20px">
    <div>
      <div class="section-title">ESTADO DE RESULTADOS</div>
      <div class="er-row"><span>(+) Ingresos</span><span>$${Math.round(d.ingTotal).toLocaleString()}</span></div>
      <div class="er-row"><span>(-) Costos operativos</span><span>-$${Math.round(d.ctTotal).toLocaleString()}</span></div>
      <div class="er-row er-bold"><span>= EBITDA</span><span>$${Math.round(d.ebitda).toLocaleString()}</span></div>
      <div class="er-row"><span>(-) Intereses préstamo</span><span>-$${Math.round(d.saldoPrest * 0.107 / 12 * (d.mesFin - d.mesInicio + 1)).toLocaleString()}</span></div>
      <div class="er-row er-bold" style="color:${verde}"><span>= Utilidad Neta</span><span>$${Math.round(d.utilidad).toLocaleString()}</span></div>
    </div>
    <div>
      <div class="section-title">CRÉDITO BANCARIO</div>
      <div class="er-row"><span>Banco</span><span style="font-weight:700">Produbanco</span></div>
      <div class="er-row"><span>Saldo actual</span><span style="font-weight:700">$${d.saldoPrest.toLocaleString()}</span></div>
      <div class="er-row"><span>Cuota mensual</span><span style="font-weight:700">$${d.cuotaPrest.toLocaleString()}</span></div>
      <div class="er-row"><span>Tasa interés</span><span style="font-weight:700">10.7% anual</span></div>
      <div class="er-row"><span>Vencimiento</span><span style="font-weight:700">Febrero 2027</span></div>
      <div class="er-row er-bold" style="color:${azul}"><span>Ratio D/EBITDA</span><span>~1.8x ✅</span></div>
    </div>
  </div>
  <div class="section" style="margin-top:8px">
    <div class="section-title">PROYECCIÓN 2026 — Si mantiene el ritmo actual</div>
    <div style="display:flex;gap:8px">
      <div style="flex:1;background:#F8F9FA;border-radius:6px;padding:8px;text-align:center">
        <div style="font-size:16px;font-weight:700;color:${verde}">${(d.tmTotal * 3).toFixed(0)} TM</div>
        <div style="font-size:9px;color:#666">TM proyectadas 2026</div>
      </div>
      <div style="flex:1;background:#F8F9FA;border-radius:6px;padding:8px;text-align:center">
        <div style="font-size:16px;font-weight:700;color:${azul}">$${Math.round(d.ingTotal * 3).toLocaleString()}</div>
        <div style="font-size:9px;color:#666">Ingreso proyectado</div>
      </div>
      <div style="flex:1;background:#F8F9FA;border-radius:6px;padding:8px;text-align:center">
        <div style="font-size:16px;font-weight:700;color:${verdeM}">$${Math.round(d.utilidad * 3).toLocaleString()}</div>
        <div style="font-size:9px;color:#666">Utilidad proyectada</div>
      </div>
      <div style="flex:1;background:#F8F9FA;border-radius:6px;padding:8px;text-align:center">
        <div style="font-size:16px;font-weight:700;color:#6B21A8">5,166 TM</div>
        <div style="font-size:9px;color:#666">Meta anual (30 TM/ha)</div>
      </div>
    </div>
  </div>
  <div class="footer">
    <span>LLH - Hacienda Pantanal | Oscar Jimenez</span>
    <span>Información confidencial — Solo uso interno y bancario</span>
    <span>${d.periodo}</span>
  </div>
  </body></html>`;
}

function obtenerCarpetaReportes() {
  const nombre   = 'Pantanal Agro — Reportes';
  const carpetas = DriveApp.getFoldersByName(nombre);
  if (carpetas.hasNext()) return carpetas.next();
  return DriveApp.createFolder(nombre);
}


// ═══════════════════════════════════════════════════════════════
// PROCESADOR DE FACTURAS SIIGO
// ═══════════════════════════════════════════════════════════════

const CARPETA_PENDIENTES  = 'Facturas Siigo - Pendientes';
const CARPETA_PROCESADAS  = 'Facturas Siigo - Procesadas';
const CARPETA_CON_ERRORES = 'Facturas Siigo - Con Errores';

function crearCarpetas() {
  [CARPETA_PENDIENTES, CARPETA_PROCESADAS, CARPETA_CON_ERRORES].forEach(nombre => {
    if (!DriveApp.getFoldersByName(nombre).hasNext()) DriveApp.createFolder(nombre);
  });
  SpreadsheetApp.getUi().alert(
    '✅ Carpetas creadas en Google Drive:\n\n' +
    '📁 ' + CARPETA_PENDIENTES + '\n' +
    '📁 ' + CARPETA_PROCESADAS + '\n' +
    '📁 ' + CARPETA_CON_ERRORES
  );
}

function instalarActivadorFacturas() {
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === 'procesarFacturasPendientes')
      ScriptApp.deleteTrigger(trigger);
  });
  ScriptApp.newTrigger('procesarFacturasPendientes').timeBased().everyMinutes(15).create();
  SpreadsheetApp.getUi().alert('✅ Activador instalado. Revisará la carpeta cada 15 minutos.');
}

function procesarFacturasPendientes() {
  const carpetaPend = obtenerCarpeta(CARPETA_PENDIENTES);
  if (!carpetaPend) return;
  const archivos = carpetaPend.getFilesByType(MimeType.PDF);
  let procesados = 0; let errores = 0;
  while (archivos.hasNext()) {
    const archivo = archivos.next();
    try {
      const resultado = procesarFacturaSiigo(archivo);
      if (resultado.exito) { moverArchivo(archivo, CARPETA_PROCESADAS); procesados++; }
      else                  { moverArchivo(archivo, CARPETA_CON_ERRORES); errores++; }
    } catch (e) { moverArchivo(archivo, CARPETA_CON_ERRORES); errores++; }
  }
  if (procesados > 0 || errores > 0) enviarResumenProcesamiento(procesados, errores);
}

function procesarFacturaSiigo(archivo) {
  const texto = extraerTextoPDF(archivo);
  if (!texto || texto.length < 50) return { exito: false, error: 'No se pudo leer el PDF' };
  const fecha = extraerFecha(texto);
  if (!fecha) return { exito: false, error: 'No se encontró fecha de emisión' };
  const lineas = extraerLineasProducto(texto);
  if (!lineas || lineas.length === 0) return { exito: false, error: 'No se encontraron líneas de producto' };
  const numFactura = extraerNumeroFactura(texto);
  if (numFactura && facturaYaProcesada(numFactura)) return { exito: false, error: 'Factura ya procesada' };
  const registrosCargados = cargarEnVentas(fecha, lineas, numFactura, archivo.getName());
  return { exito: true, registros: registrosCargados };
}

function extraerTextoPDF(archivo) {
  try {
    const blob     = archivo.getBlob();
    const tempFile = DriveApp.getRootFolder().createFile(blob);
    const copy     = Drive.Files.copy(
      { title: 'temp_doc_' + Date.now(), mimeType: MimeType.GOOGLE_DOCS },
      tempFile.getId(), { convert: true }
    );
    const doc   = DocumentApp.openById(copy.id);
    const texto = doc.getBody().getText();
    tempFile.setTrashed(true);
    DriveApp.getFileById(copy.id).setTrashed(true);
    return texto;
  } catch (e) { Logger.log('Error extrayendo texto: ' + e.message); return null; }
}

function extraerFecha(texto) {
  const patrones = [
    /Fecha\s+(?:de\s+)?Emisi[oó]n[:\s]+(\d{2}\/\d{2}\/\d{4})/i,
    /Fecha\s+Emision[:\s]+(\d{2}\/\d{2}\/\d{4})/i,
    /(\d{2}\/\d{2}\/\d{4})/,
  ];
  for (const patron of patrones) {
    const match = texto.match(patron);
    if (match) {
      const p = match[1].split('/');
      return new Date(parseInt(p[2]), parseInt(p[1]) - 1, parseInt(p[0]));
    }
  }
  return null;
}

function extraerNumeroFactura(texto) {
  const match = texto.match(/(?:FACTURA|Factura)\s+No\.\s*([\d\-]+)/i);
  return match ? match[1].trim() : null;
}

function extraerLineasProducto(texto) {
  const lineas = [];
  const mapaFincas = {
    'andino': 'Andino', 'corrales': 'Los Corrales', 'los corrales': 'Los Corrales',
    'chipo': 'Chipo', 'marujita': 'Marujita', 'la marujita': 'Marujita',
    'castañeda': 'Castañeda', 'castaneda': 'Castañeda',
  };
  const filas = texto.split('\n');
  const productosOrdenados = [];
  for (let i = 0; i < filas.length; i++) {
    const fila = filas[i].trim();
    if (!fila || !/finca/i.test(fila)) continue;
    let nombreFinca = null;
    const filaLower = fila.toLowerCase();
    for (const [clave, valor] of Object.entries(mapaFincas)) {
      if (filaLower.includes(clave)) { nombreFinca = valor; break; }
    }
    if (!nombreFinca) continue;
    const matchCant = fila.match(/^(?:\d{2}\s+)?(\d{1,3}\.\d{1,3})\s+Fruta/i);
    let cantidad = null;
    if (matchCant) {
      cantidad = parseFloat(matchCant[1]);
    } else {
      const nums = fila.match(/\b(\d{1,3}\.\d{2,3})\b/g);
      if (nums) { for (const n of nums) { const v = parseFloat(n); if (v >= 0.5 && v <= 500) { cantidad = v; break; } } }
    }
    if (cantidad) productosOrdenados.push({ finca: nombreFinca, tm: cantidad });
  }
  const precios = [];
  for (const fila of filas) {
    const f = fila.trim();
    if (/\d{2,3}(?:\.\d{2})?\s+\$0\.00/.test(f)) {
      const matches = f.matchAll(/(\d{2,3}(?:\.\d{2})?)\s+\$0\.00/g);
      for (const m of matches) { const p = parseFloat(m[1]); if (p >= 100 && p <= 400) precios.push(p); }
    }
  }
  if (precios.length === 0) {
    for (const fila of filas) {
      const matchP = fila.trim().match(/^(\d{2,3}(?:\.\d{2})?)\s+\$0\.00/);
      if (matchP) { const p = parseFloat(matchP[1]); if (p >= 100 && p <= 400) precios.push(p); }
    }
  }
  for (let i = 0; i < productosOrdenados.length; i++) {
    const precio = precios[i] || precios[0];
    if (precio) lineas.push({ finca: productosOrdenados[i].finca, tm: productosOrdenados[i].tm, precio });
  }
  return lineas;
}

function facturaYaProcesada(numFactura) { return false; }

function cargarEnVentas(fecha, lineas, numFactura, nombreArchivo) {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const wv   = ss.getSheetByName('💵 Ventas');
  if (!wv) throw new Error('No se encontró la pestaña Ventas');
  const colA = wv.getRange('A1:A105').getValues();
  let nextRow = 6;
  for (let i = 5; i < colA.length; i++) {
    if (!colA[i][0] && !wv.getRange(i + 1, 2).getValue()) { nextRow = i + 1; break; }
  }
  let registros = 0;
  for (const linea of lineas) {
    if (nextRow > 105) break;
    wv.getRange(nextRow, 1).setValue(calcularSemana(fecha));
    wv.getRange(nextRow, 2).setValue(fecha).setNumberFormat('DD/MM/YYYY');
    wv.getRange(nextRow, 3).setValue(linea.finca);
    wv.getRange(nextRow, 4).setValue(linea.tm).setNumberFormat('#,##0.000');
    wv.getRange(nextRow, 5).setValue(linea.precio).setNumberFormat('$#,##0.00');
    nextRow++; registros++;
  }
  return registros;
}

function calcularSemana(fecha) {
  const inicio = new Date(fecha.getFullYear(), 0, 1);
  return Math.ceil(((fecha - inicio) / 86400000 + inicio.getDay() + 1) / 7);
}

function moverArchivo(archivo, nombreCarpeta) {
  try {
    const dest = obtenerCarpeta(nombreCarpeta);
    if (!dest) return;
    dest.addFile(archivo);
    const orig = obtenerCarpeta(CARPETA_PENDIENTES);
    if (orig) orig.removeFile(archivo);
  } catch (e) { Logger.log('Error moviendo archivo: ' + e.message); }
}

function obtenerCarpeta(nombre) {
  const c = DriveApp.getFoldersByName(nombre);
  return c.hasNext() ? c.next() : null;
}

function enviarResumenProcesamiento(procesados, errores) {
  const fecha = new Date().toLocaleDateString('es-EC');
  MailApp.sendEmail({
    to: EMAIL_ADMIN,
    subject: `📄 Facturas procesadas — ${fecha} | Pantanal Agro`,
    body: `Estimado Ing. Oscar,\n\n` +
          `✅ Procesadas correctamente: ${procesados}\n` +
          `❌ Con errores: ${errores}\n\n` +
          `— Sistema LLH Pantanal Agro`
  });
}

function procesarFacturasManual() {
  const ui   = SpreadsheetApp.getUi();
  const resp = ui.alert(
    'Procesar Facturas Siigo',
    'Se procesarán todos los PDFs en "Facturas Siigo - Pendientes".\n¿Continuar?',
    ui.ButtonSet.YES_NO
  );
  if (resp !== ui.Button.YES) return;
  procesarFacturasPendientes();
  ui.alert('✅ Procesamiento completado. Revise su email para el resumen.');
}


// ═══════════════════════════════════════════════════════════════
// CUADERNO DE SECRETARIA
// ═══════════════════════════════════════════════════════════════

const FINCAS_LISTA = [
  "Castañeda", "Marujita", "Andino", "Los Corrales", "Chipo"
];

const EMPLEADOS_LISTA = [
  "Aldahir Bravo", "Cesar Villalva", "Elio Quijije", "Felix Arellano",
  "Horacio Rivera", "Italo Barragan", "Jessica Rivera", "Klever Zambrano",
  "Manuel Morales", "Merlin Valencia", "Milton Barragan", "Oliver Cedeño",
  "Vicente Narvaez", "Ignacio Barragan", "Diviel Jimenez", "Antonio Suarez",
  "Valentin Loor", "Francisco Mero", "Ruben Luna", "Edison Mendoza",
  "Darwin Sesme", "Otro (temporal)"
];

const ACTIVIDADES_LISTA = [
  "Aseroría Técnica Fincas", "Chapia", "Coronas Químicas", "Corona Y Desvetillada",
  "Cosecha", "Cuidado De Animales De Trabajo", "Cuidado Ganado", "Despejando Plantas",
  "Poda", "Poda Y Desvetillada", "Polinización", "Regada De Cal",
  "Transporte De Fruta De Palma", "Fertilización", "Alambradas", "Actividades varias",
  "Fumigación caminos", "Transporte de fertilizantes", "Transporte de agua",
  "Chapia en esteros", "Limpieza rastrojos", "Otro (temporal)"
];

const FILA_INICIO_DATOS = 5;
const FILAS_DATOS        = 200;

function crearCuaderno() {
  const ss          = SpreadsheetApp.getActiveSpreadsheet();
  const NOMBRE_HOJA = "📋 Cuaderno";

  let hoja = ss.getSheetByName(NOMBRE_HOJA);
  if (hoja) {
    const ui   = SpreadsheetApp.getUi();
    const resp = ui.alert(
      "Hoja existente",
      "La pestaña '" + NOMBRE_HOJA + "' ya existe. ¿Recrear desde cero?",
      ui.ButtonSet.YES_NO
    );
    if (resp !== ui.Button.YES) return;
    ss.deleteSheet(hoja);
  }
  hoja = ss.insertSheet(NOMBRE_HOJA);

  const pendientes = ss.getSheetByName("⏳ Pendientes");
  if (pendientes) { ss.setActiveSheet(hoja); ss.moveActiveSheet(pendientes.getIndex()); }

  const rTitulo = hoja.getRange("A1:I1");
  rTitulo.merge();
  rTitulo.setValue("📋 CUADERNO DE LA SECRETARIA — Registro de actividades");
  rTitulo.setBackground("#1a6b3c").setFontColor("#ffffff").setFontWeight("bold")
         .setFontSize(13).setVerticalAlignment("middle");
  hoja.setRowHeight(1, 36);

  const rInstruccion = hoja.getRange("A2:I2");
  rInstruccion.merge();
  rInstruccion.setValue(
    "💡 Complete los campos en amarillo. Semana y Total se calculan solos. " +
    "Menú 🌴 Pantanal Agro → Agregar para ampliar listas."
  );
  rInstruccion.setBackground("#fff8e1").setFontColor("#5d4037").setFontSize(11).setWrap(true);
  hoja.setRowHeight(2, 32);

  hoja.getRange("A3:I3").setBackground("#f5f5f5");
  hoja.setRowHeight(3, 8);

  const COLS = [
    { titulo: "Semana\n(auto)", color: "#e8f5e9", ancho: 72  },
    { titulo: "Fecha *",        color: "#fff9c4", ancho: 110 },
    { titulo: "Finca *",        color: "#e8eaf6", ancho: 115 },
    { titulo: "Empleado *",     color: "#e8eaf6", ancho: 210 },
    { titulo: "Actividad *",    color: "#e8eaf6", ancho: 165 },
    { titulo: "Cantidad *",     color: "#fff9c4", ancho: 85  },
    { titulo: "Precio Unit. $", color: "#fff9c4", ancho: 100 },
    { titulo: "Total $\n(auto)",color: "#e8f5e9", ancho: 95  },
    { titulo: "Notas",          color: "#fafafa", ancho: 220 },
  ];

  COLS.forEach((col, i) => {
    const c = hoja.getRange(4, i + 1);
    c.setValue(col.titulo).setBackground(col.color).setFontWeight("bold")
     .setFontSize(11).setVerticalAlignment("middle").setHorizontalAlignment("center").setWrap(true);
    hoja.setColumnWidth(i + 1, col.ancho);
  });
  hoja.setRowHeight(4, 42);
  hoja.getRange("A4:I4").setBorder(null, null, true, null, null, null, "#1a6b3c", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  const FI = FILA_INICIO_DATOS;
  const FD = FILAS_DATOS;

  for (let f = FI; f < FI + FD; f++) {
    hoja.getRange(f, 1).setFormula(`=IF(B${f}="";"";INT((B${f}-DATE(YEAR(B${f});1;1)+WEEKDAY(DATE(YEAR(B${f});1;1);2))/7)+1)`);
    hoja.getRange(f, 8).setFormula(`=IF(OR(F${f}="";G${f}="");"";F${f}*G${f})`);
    hoja.getRange(f, 10).setFormula(
      `=IF(B${f}="";"";TEXT(B${f};"yyyy-mm-dd")&"|"&D${f}&"|"&E${f}&"|"&C${f})`
);
  }

  COLS.forEach((col, i) => {
    hoja.getRange(FI, i + 1, FD, 1).setBackground(col.color).setVerticalAlignment("middle");
  });
  hoja.getRange(FI, 1, FD, 1).setNumberFormat("0");
  hoja.getRange(FI, 2, FD, 1).setNumberFormat("yyyy-mm-dd");
  hoja.getRange(FI, 6, FD, 1).setNumberFormat("0.##");
  hoja.getRange(FI, 7, FD, 1).setNumberFormat("$ 0.00");
  hoja.getRange(FI, 8, FD, 1).setNumberFormat("$ 0.00");
  [1, 2, 6, 7, 8].forEach(col => hoja.getRange(FI, col, FD, 1).setHorizontalAlignment("center"));
  hoja.getRange(FI, 1, FD, 9).setBorder(true, true, true, true, true, true, "#c8e6c9", SpreadsheetApp.BorderStyle.SOLID);

  _aplicarDropdown(hoja, FI, FD, 3, FINCAS_LISTA,      "Seleccione una finca.");
  _aplicarDropdown(hoja, FI, FD, 4, EMPLEADOS_LISTA,   "Seleccione un empleado.");
  _aplicarDropdown(hoja, FI, FD, 5, ACTIVIDADES_LISTA, "Seleccione una actividad.");

  hoja.getRange(4, 10).setValue("🔑 Clave").setFontWeight("bold").setBackground("#ede7f6");
  hoja.setColumnWidth(10, 60);
  hoja.hideColumns(10);

  hoja.getRange(FI, 1, FD, 1).protect().setDescription("Semana — automática").setWarningOnly(true);
  hoja.getRange(FI, 8, FD, 1).protect().setDescription("Total — automático").setWarningOnly(true);

  hoja.setFrozenRows(4);
  hoja.setTabColor("#1a6b3c");

  SpreadsheetApp.getActiveSpreadsheet().toast(
    "✅ Cuaderno creado con " + FD + " filas listas.", "Cuaderno listo", 8
  );
}

function agregarEmpleado() {
  _agregarItemDropdown(4, "empleado", "👤 Nuevo empleado", "Escribe el nombre completo:");
}

function agregarActividad() {
  _agregarItemDropdown(5, "actividad", "⚙️ Nueva actividad", "Escribe el nombre de la actividad:");
}

function _aplicarDropdown(hoja, filaInicio, cantFilas, col, lista, ayuda) {
  const regla = SpreadsheetApp.newDataValidation()
    .requireValueInList(lista, true)
    .setAllowInvalid(true)
    .setHelpText(ayuda)
    .build();
  hoja.getRange(filaInicio, col, cantFilas, 1).setDataValidation(regla);
}

function _agregarItemDropdown(columna, tipo, tituloDialogo, mensajeDialogo) {
  const ui   = SpreadsheetApp.getUi();
  const resp = ui.prompt(tituloDialogo, mensajeDialogo, ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;

  const nuevoItem = resp.getResponseText().trim();
  if (!nuevoItem) { ui.alert("Nombre vacío. Cancelado."); return; }

  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName("📋 Cuaderno");
  if (!hoja) { ui.alert("No se encontró '📋 Cuaderno'. Créala primero."); return; }

  const rango       = hoja.getRange(FILA_INICIO_DATOS, columna, FILAS_DATOS, 1);
  const reglaActual = rango.getDataValidation();
  if (!reglaActual) { ui.alert("No hay dropdown. Recrea el Cuaderno."); return; }

  const listaActual = reglaActual.getCriteriaValues()[0];
  if (listaActual.some(item => item.toLowerCase() === nuevoItem.toLowerCase())) {
    ui.alert("'" + nuevoItem + "' ya está en la lista."); return;
  }

  listaActual.push(nuevoItem);
  listaActual.sort((a, b) => a.localeCompare(b, "es", { sensitivity: "base" }));

  hoja.getRange(FILA_INICIO_DATOS, columna, FILAS_DATOS, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(listaActual, true)
      .setAllowInvalid(false)
      .setHelpText("Seleccione un " + tipo + " de la lista.")
      .build()
  );

  ui.alert("✅ '" + nuevoItem + "' agregado a la lista de " + tipo + "s.");
}

function ejecutarConciliacion() {
  SpreadsheetApp.getUi().alert(
    "🔄 Motor de conciliación",
    "Próxima versión: comparará 📋 Cuaderno vs ⏳ Pendientes y generará 🔄 Conciliación.",
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function doGet(e) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const cat   = ss.getSheetByName('⚙️ Catálogos');
  const datos = cat.getDataRange().getValues();

  const fincas      = [];
  const actividades = [];
  const empleados   = [];

  for (let i = 6; i < datos.length; i++) {
    if (datos[i][0] && typeof datos[i][0] === 'string' && datos[i][0].trim())
      fincas.push(datos[i][0].trim());
    if (datos[i][3] && typeof datos[i][3] === 'string' && datos[i][3].trim())
      actividades.push(datos[i][3].trim());
    if (datos[i][6] && typeof datos[i][6] === 'string' && datos[i][6].trim() && isNaN(datos[i][6]))
      empleados.push(datos[i][6].trim());
  }

  const result = JSON.stringify({ fincas, actividades, empleados });
  return ContentService.createTextOutput(result).setMimeType(ContentService.MimeType.JSON);
}
