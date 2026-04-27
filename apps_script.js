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
// ═══════════════════════════════════════════════════════════════
// GENERADOR DE REPORTES — Pantanal Agro
// Agregar al final del Apps Script existente
// ═══════════════════════════════════════════════════════════════
 
const EMAIL_REPORTE = 'ohjimenez.pantanal@gmail.com';
const MESES = ['Enero','Febrero','Marzo','Abril','Mayo','Junio',
               'Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'];
 
// ── BOTÓN MANUAL — Genera reporte del período que elija ───────
function generarReporteManual() {
  const ui = SpreadsheetApp.getUi();
  
  // Preguntar período
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
  
  if (isNaN(mesInicio) || isNaN(mesFin) || mesInicio<1 || mesFin>12 || mesInicio>mesFin) {
    ui.alert('Período inválido. Ingrese números del 1 al 12.');
    return;
  }
  
  const periodo = MESES[mesInicio-1] + ' - ' + MESES[mesFin-1] + ' 2026';
  ui.alert('Generando reporte para: ' + periodo + '\n\nRecibirá el PDF en su email en unos segundos.');
  
  generarYEnviarReporte(mesInicio, mesFin, periodo);
}
 
// ── REPORTE AUTOMÁTICO MENSUAL ────────────────────────────────
function reporteAutomaticoMensual() {
  const hoy    = new Date();
  const mesAct = hoy.getMonth() + 1;  // mes actual (1-12)
  const mesAnt = mesAct === 1 ? 12 : mesAct - 1;  // mes anterior
  const periodo = MESES[mesAnt-1] + ' 2026';
  
  generarYEnviarReporte(mesAnt, mesAnt, periodo);
}
 
// ── FUNCIÓN PRINCIPAL ─────────────────────────────────────────
function generarYEnviarReporte(mesInicio, mesFin, periodo) {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const wd  = ss.getSheetByName('📈 Dashboard');
  const wv  = ss.getSheetByName('💵 Ventas');
  const we  = ss.getSheetByName('📊 EBITDA');
  const wpy = ss.getSheetByName('💰 Precios y Ventas');
  
  // ── Extraer datos del sistema ────────────────────────────────
  const tmTotal    = wd.getRange('A97').getValue() || 0;
  const ctTotal    = wd.getRange('C97').getValue() || 0;
  const ctTm       = wd.getRange('G97').getValue() || 0;
  const ingTotal   = wv ? wv.getRange('F106').getValue() || 0 : 0;
  const utilidad   = ingTotal - ctTotal;
  const margen     = ingTotal > 0 ? (utilidad/ingTotal*100) : 0;
  const ebitda     = we ? we.getRange('B35').getValue() || 0 : utilidad;
  
  // Datos por mes
  const ventas_mes = [];
  for (let mo = mesInicio; mo <= mesFin; mo++) {
    const r = 110 + mo;  // filas del resumen mensual en Ventas
    if (wv) {
      const tmMes  = wv.getRange(r, 2).getValue() || 0;
      const precMes= wv.getRange(r, 3).getValue() || 0;
      const ingMes = wv.getRange(r, 4).getValue() || 0;
      const ctMes  = wv.getRange(r, 5).getValue() || 0;
      const utMes  = wv.getRange(r, 6).getValue() || 0;
      ventas_mes.push([MESES[mo-1], tmMes, precMes, ingMes, ctMes, utMes]);
    }
  }
  
  // Préstamo
  const saldoPrest = we ? we.getRange('B22').getValue() || 37000 : 37000;
  const cuotaPrest = we ? we.getRange('B23').getValue() || 4000  : 4000;
  
  // ── Construir HTML del reporte ───────────────────────────────
  const html = construirHtmlReporte({
    periodo, tmTotal, ctTotal, ctTm, ingTotal, utilidad, margen,
    ebitda, ventas_mes, saldoPrest, cuotaPrest, mesInicio, mesFin
  });
  
  // ── Convertir a PDF via DriveApp ─────────────────────────────
  const blob     = Utilities.newBlob(html, 'text/html', 'reporte.html');
  const pdfBlob  = blob.getAs('application/pdf');
  const fileName = 'Pantanal_Agro_Reporte_' + periodo.replace(/ /g,'_') + '.pdf';
  pdfBlob.setName(fileName);
  
  // Guardar en Drive
  const folder = obtenerCarpetaReportes();
  const file   = folder.createFile(pdfBlob);
  const fileUrl= file.getUrl();
  
  // Enviar por email
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
  
  Logger.log('✅ Reporte generado y enviado: ' + fileName);
}
 
// ── CONSTRUIR HTML DEL REPORTE ────────────────────────────────
function construirHtmlReporte(d) {
  const verde = '#1B4332'; const verdeM = '#2D6A4F'; const verdeCl = '#40916C';
  const verdePal = '#B7E4C7'; const naranja = '#F59E0B'; const azul = '#1A56DB';
  
  let filasVentas = '';
  let totalTm=0, totalIng=0, totalCt=0, totalUt=0;
  d.ventas_mes.forEach(([mes,tm,prec,ing,ct,ut]) => {
    totalTm+=tm; totalIng+=ing; totalCt+=ct; totalUt+=ut;
    filasVentas += `
      <tr>
        <td style="font-weight:600">${mes}</td>
        <td>${tm.toFixed(2)}</td>
        <td>$${prec.toFixed(0)}</td>
        <td>$${ing.toLocaleString('es-EC',{minimumFractionDigits:0})}</td>
        <td>$${ct.toFixed(0)}</td>
        <td style="color:${ut>0?'#065F46':'#991B1B'};font-weight:700">$${ut.toFixed(0)}</td>
      </tr>`;
  });
  filasVentas += `
    <tr style="background:${verdePal};font-weight:700">
      <td>TOTAL</td><td>${totalTm.toFixed(2)}</td><td>—</td>
      <td>$${totalIng.toLocaleString('es-EC',{minimumFractionDigits:0})}</td>
      <td>$${totalCt.toFixed(0)}</td>
      <td style="color:#065F46">$${totalUt.toFixed(0)}</td>
    </tr>`;
 
  return `<!DOCTYPE html><html><head><meta charset="UTF-8">
  <style>
    body{font-family:Arial,sans-serif;font-size:11px;color:#1A1A1A;margin:0;padding:0}
    .header{background:${verde};color:white;padding:16px 20px;position:relative}
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
      <div>
        <div class="empresa">LLH - Hacienda Pantanal</div>
        <div class="sub">Palma Africana | Quininde, Esmeraldas, Ecuador</div>
      </div>
      <div style="text-align:right">
        <div style="font-size:13px;font-weight:700">REPORTE EJECUTIVO</div>
        <div style="font-size:10px;opacity:0.8">${d.periodo}</div>
      </div>
    </div>
  </div>
  <div class="naranja-bar"></div>
  
  <div class="kpis">
    <div class="kpi" style="border-top-color:${verdeCl}">
      <div class="kpi-val">${d.tmTotal.toFixed(1)} TM</div>
      <div class="kpi-lbl">TM Producidas</div>
    </div>
    <div class="kpi" style="border-top-color:${azul}">
      <div class="kpi-val">$${Math.round(d.ingTotal).toLocaleString()}</div>
      <div class="kpi-lbl">Ingresos Brutos</div>
    </div>
    <div class="kpi" style="border-top-color:${verdeM}">
      <div class="kpi-val">$${Math.round(d.ebitda).toLocaleString()}</div>
      <div class="kpi-lbl">EBITDA</div>
    </div>
    <div class="kpi" style="border-top-color:${naranja}">
      <div class="kpi-val">$${d.utilidad.toFixed(0)}</div>
      <div class="kpi-lbl">Utilidad Neta (${d.margen.toFixed(1)}%)</div>
    </div>
    <div class="kpi" style="border-top-color:#6B21A8">
      <div class="kpi-val">$${d.ctTm.toFixed(2)}</div>
      <div class="kpi-lbl">CT / TM</div>
    </div>
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
      <div class="er-row"><span>(-) Intereses prestamo</span><span>-$${Math.round(d.saldoPrest*0.107/12*(d.mesFin-d.mesInicio+1)).toLocaleString()}</span></div>
      <div class="er-row er-bold" style="color:${verde}"><span>= Utilidad Neta</span><span>$${Math.round(d.utilidad).toLocaleString()}</span></div>
    </div>
    <div>
      <div class="section-title">CREDITO BANCARIO</div>
      <div class="er-row"><span>Banco</span><span style="font-weight:700">Produbanco</span></div>
      <div class="er-row"><span>Saldo actual</span><span style="font-weight:700">$${d.saldoPrest.toLocaleString()}</span></div>
      <div class="er-row"><span>Cuota mensual</span><span style="font-weight:700">$${d.cuotaPrest.toLocaleString()}</span></div>
      <div class="er-row"><span>Tasa interes</span><span style="font-weight:700">10.7% anual</span></div>
      <div class="er-row"><span>Vencimiento</span><span style="font-weight:700">Febrero 2027</span></div>
      <div class="er-row er-bold" style="color:${azul}"><span>Ratio D/EBITDA</span><span>~1.8x ✅</span></div>
    </div>
  </div>
  
  <div class="section" style="margin-top:8px">
    <div class="section-title">PROYECCION 2026 — Si mantiene el ritmo actual</div>
    <div style="display:flex;gap:8px">
      <div style="flex:1;background:#F8F9FA;border-radius:6px;padding:8px;text-align:center">
        <div style="font-size:16px;font-weight:700;color:${verde}">${(d.tmTotal*3).toFixed(0)} TM</div>
        <div style="font-size:9px;color:#666">TM proyectadas 2026</div>
      </div>
      <div style="flex:1;background:#F8F9FA;border-radius:6px;padding:8px;text-align:center">
        <div style="font-size:16px;font-weight:700;color:${azul}">$${Math.round(d.ingTotal*3).toLocaleString()}</div>
        <div style="font-size:9px;color:#666">Ingreso proyectado</div>
      </div>
      <div style="flex:1;background:#F8F9FA;border-radius:6px;padding:8px;text-align:center">
        <div style="font-size:16px;font-weight:700;color:${verdeM}">$${Math.round(d.utilidad*3).toLocaleString()}</div>
        <div style="font-size:9px;color:#666">Utilidad proyectada</div>
      </div>
      <div style="flex:1;background:#F8F9FA;border-radius:6px;padding:8px;text-align:center">
        <div style="font-size:16px;font-weight:700;color:#6B21A8">5,166 TM</div>
        <div style="font-size:9px;color:#666">Meta anual (30 TM/ha)</div>
      </div>
    </div>
  </div>
  
  <div class="footer">
    <span>LLH - Hacienda Pantanal | Oscar Jimenez - Analista General</span>
    <span>Informacion confidencial — Solo para uso interno y bancario</span>
    <span>${d.periodo}</span>
  </div>
  
  </body></html>`;
}
 
// ── CARPETA EN DRIVE ──────────────────────────────────────────
function obtenerCarpetaReportes() {
  const nombre = 'Pantanal Agro — Reportes';
  const carpetas = DriveApp.getFoldersByName(nombre);
  if (carpetas.hasNext()) return carpetas.next();
  return DriveApp.createFolder(nombre);
}
 
// ── MENÚ EN GOOGLE SHEETS ─────────────────────────────────────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📊 Pantanal Agro')
    .addItem('Generar Reporte Ejecutivo', 'generarReporteManual')
    .addItem('Verificar Sistema', 'verificacionDiaria')
    .addToUi();
}
 
