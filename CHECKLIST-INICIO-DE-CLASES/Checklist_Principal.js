/**
 * SISTEMA GESTIÓN VONEX - VERSIÓN FINAL CON CORRECCIÓN DE VALORES
 */
function onEdit(e) {
  const ss = e.source;
  const sheet = ss.getActiveSheet();
  const r = e.range;
  const row = r.getRow();
  const col = r.getColumn();
  
  // Limpieza agresiva del valor y nombre de la hoja
  const val = sheet.getRange(row, col).getValue().toString().trim().toUpperCase();
  const sheetName = sheet.getName().toString().trim();

  const MIN_ROW = 9;
  const FILA_SEDES = 7;
  const EMAIL_MASTER = "marevalo@vonex.edu.pe";
  
  // Mapa de columnas según tu estructura F=6, H=8, J=10, etc.
  const SHEET_CONFIG = {
    "UNMSM-INTENSIVO": [[6,7], [8,9], [10,11], [12,13], [14,15], [16,17], [18,19], [20,21]],
    "UNMSM": [[6,7], [8,9], [10,11], [12,13], [14,15], [16,17], [18,19], [20,21], [22,23], [24,25]]
  };

  if (row < MIN_ROW) return;
  
  // Si la hoja no coincide, el script te avisará en el Excel
  if (!SHEET_CONFIG[sheetName]) {
    ss.toast("La hoja '" + sheetName + "' no está configurada en el script.", "Aviso de Configuración");
    return;
  }

  const COLUMN_PAIRS = SHEET_CONFIG[sheetName];

  for (const [colCheck, colFecha] of COLUMN_PAIRS) {
    if (col === colCheck) {
      
      if (val === "REALIZADO") {
        // 1. Poner fecha
        sheet.getRange(row, colFecha).setValue(new Date()).setNumberFormat("dd-MM-yyyy HH:mm:ss");
        
        // 2. Datos para el correo
        const sedeNombre = sheet.getRange(FILA_SEDES, colCheck).getValue() || "Sede";
        
        // Notificación visual inmediata en el Sheet
        ss.toast("Enviando Dashboard para " + sedeNombre + "...", "Procesando");

        // 3. Envío de correos envueltos en TRY/CATCH para ver errores
        try {
          enviarDashboardSede(sheet, colCheck, row, EMAIL_MASTER, sedeNombre);
          gestionarDependencias(sheet, row, colCheck, sedeNombre);
          ss.toast("✅ Dashboard enviado a " + EMAIL_MASTER, "Éxito");
        } catch (err) {
          ss.toast("❌ Error de correo: " + err.message, "Error");
        }
      }

      if (val === "NO APLICA") {
        sheet.getRange(row, colFecha).clearContent();
        const cellObs = sheet.getRange(row, 24); // Columna X
        if (cellObs.getValue() === "") cellObs.setValue("⚠️ COMPLETAR MOTIVO AQUÍ");
      }
    }
  }
}

// FUNCIONES DE APOYO (Dashboard y Dependencias)
function enviarDashboardSede(sheet, colCheck, filaEditada, emailDestino, sedeNombre) {
  // 1. Obtener datos para el asunto
  const fechaInicioCiclo = sheet.getRange("C4").getDisplayValue();
  const sheetName = sheet.getName();
  const asunto = `CHECK LIST GESTIÓN DE INICIO DE CLASES - ${sheetName} - ${fechaInicioCiclo}`;
  
  // 2. Definición de Rangos
  const rangos = [
    { n: "ANTES DEL INICIO", i: 9, f: 35, color: "#2980b9" },
    { n: "PRIMERA SEMANA", i: 38, f: 45, color: "#27ae60" },
    { n: "5TA SEMANA (DESPUÉS)", i: 48, f: 49, color: "#f39c12" }
  ];

  const actividadActual = sheet.getRange(filaEditada, 2).getValue();
  let totalHechoGlobal = 0;
  let totalTareasGlobal = 0;
  let htmlFilas = "";

  // 3. Construcción de Barras de Progreso
  rangos.forEach(r => {
    let datos = sheet.getRange(r.i, colCheck, (r.f - r.i) + 1, 1).getValues();
    let hechos = datos.filter(f => {
      let v = f[0].toString().toUpperCase().trim();
      return v === "REALIZADO" || v === "NO APLICA";
    }).length;
    let total = datos.length;
    let porc = ((hechos / total) * 100).toFixed(0);
    
    totalHechoGlobal += hechos;
    totalTareasGlobal += total;

    htmlFilas += `
      <div style="margin-bottom: 20px;">
        <div style="display: flex; justify-content: space-between; font-size: 13px; margin-bottom: 5px;">
          <span style="font-weight: bold; color: #34495e;">${r.n}</span>
          <span style="color: #7f8c8d;">${porc}% (${hechos}/${total})</span>
        </div>
        <div style="background-color: #ecf0f1; border-radius: 10px; height: 12px; width: 100%;">
          <div style="background-color: ${r.color}; width: ${porc}%; height: 12px; border-radius: 10px;"></div>
        </div>
      </div>`;
  });

  const porcGlobal = ((totalHechoGlobal / totalTareasGlobal) * 100).toFixed(0);

  // 4. Diseño del Cuerpo del Correo (HTML)
  const htmlBody = `
    <div style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; max-width: 600px; margin: auto; border: 1px solid #e0e0e0; border-radius: 8px; overflow: hidden; box-shadow: 0 4px 10px rgba(0,0,0,0.1);">
      <div style="background-color: #004691; color: white; padding: 20px; text-align: center;">
        <h2 style="margin: 0; font-size: 20px;">REPORTE DE CUMPLIMIENTO</h2>
        <p style="margin: 5px 0 0 0; opacity: 0.8; font-size: 14px;">Sede: ${sedeNombre}</p>
      </div>
      
      <div style="padding: 25px; background-color: white;">
        <p style="font-size: 15px; color: #2c3e50;">Se ha registrado un avance en la actividad:<br>
        <strong style="color: #c0392b;">"${actividadActual}"</strong></p>
        <hr style="border: 0; border-top: 1px solid #eee; margin: 20px 0;">
        
        ${htmlFilas}
        
        <div style="background-color: #f9f9f9; padding: 15px; border-radius: 8px; text-align: center; margin-top: 30px;">
          <span style="display: block; font-size: 12px; color: #7f8c8d; text-transform: uppercase; letter-spacing: 1px;">Progreso General de Sede</span>
          <span style="font-size: 32px; font-weight: bold; color: #2c3e50;">${porcGlobal}%</span>
        </div>

        <div style="text-align: center; margin-top: 30px;">
          <a href="${SpreadsheetApp.getActiveSpreadsheet().getUrl()}" 
             style="background-color: #004691; color: white; padding: 12px 25px; text-decoration: none; border-radius: 5px; font-weight: bold; font-size: 14px;">
             ABRIR CHECKLIST COMPLETO
          </a>
        </div>
      </div>
      
      <div style="background-color: #f4f7f6; color: #95a5a6; padding: 15px; text-align: center; font-size: 11px;">
        Este es un reporte automático generado por el Sistema de Gestión Vonex.<br>
        © 2026 Academia Vonex - Todos los derechos reservados.
      </div>
    </div>`;

  GmailApp.sendEmail(emailDestino, asunto, "", { htmlBody: htmlBody });
}

function gestionarDependencias(sheet, filaActual, colCheck, sede) {
  const REGLAS = { 13: [23], 19: [21], 20: [21] };
  if (REGLAS[filaActual]) {
    REGLAS[filaActual].forEach(filaDestino => {
      let email = sheet.getRange(filaDestino, 4).getValue(); // Columna D
      if (!email || !email.includes("@")) email = "marevalo@vonex.edu.pe"; 
      const actividad = sheet.getRange(filaDestino, 2).getValue();
      GmailApp.sendEmail(email, `🚀 Turno: ${actividad} (${sede})`, `Se ha habilitado la tarea: ${actividad}`);
    });
  }
}