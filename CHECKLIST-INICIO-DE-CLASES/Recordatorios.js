/**
 * SISTEMA DE VIGILANCIA VONEX - ESTÉTICA BUBBLE 3D RECUPERADA (MATRIZ)
 */

function enviarRecordatoriosDiarios() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetsNames = ["UNMSM-INTENSIVO", "UNMSM", "UNI", "UNHEVAL"]; 
  const hoy = new Date();
  const hoyCero = new Date(hoy.setHours(0,0,0,0));

  let mapaResponsables = {};

  sheetsNames.forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (!sheet) return;

    const data = sheet.getDataRange().getValues();
    const colSedesIndices = [5, 7, 9, 11, 13, 15, 17, 19]; 
    const sedesNombres = colSedesIndices.map(idx => sheet.getRange(7, idx + 1).getValue());

    for (let i = 8; i < data.length; i++) {
      const nombreEncargado = data[i][2] ? data[i][2].toString().trim() : ""; 
      const emailEncargado = data[i][3] ? data[i][3].toString().trim() : "";  
      const fechaRaw = data[i][21]; 
      const horaRaw = data[i][22];  
      const descripcion = data[i][4]; 
      const periodo = data[i][0] || "GESTIÓN";

      if (nombreEncargado && emailEncargado.includes("@") && fechaRaw instanceof Date) {
        let fechaLimite = new Date(fechaRaw.getTime());
                
        // Si horaRaw tiene datos, extraemos la hora sin importar el formato
        if (horaRaw) {
          let h = new Date(horaRaw);
          // Si es un objeto Date válido (Columna W)
          if (!isNaN(h.getTime())) {
            fechaLimite.setHours(h.getHours(), h.getMinutes(), 0);
          } 
        } else {
          // Si no hay hora, ponemos fin del día por defecto
          fechaLimite.setHours(23, 59, 0);
        }

        const fechaSolo = new Date(fechaRaw); fechaSolo.setHours(0,0,0,0);
        const diffDias = Math.floor((fechaSolo - hoyCero) / 86400000);

        if (diffDias === 7 || diffDias === 1 || diffDias === 0 || diffDias <= -1) {
          let sedesStatus = [];
          let pendienteGlobal = false;

          colSedesIndices.forEach((colIdx, index) => {
            const estado = (data[i][colIdx] || "").toString().toUpperCase().trim();
            const esPendiente = (estado !== "REALIZADO" && estado !== "NO APLICA");
            if (esPendiente) pendienteGlobal = true;
            sedesStatus.push({ nombre: sedesNombres[index], completado: !esPendiente });
          });

          if (pendienteGlobal) {
            if (!mapaResponsables[nombreEncargado]) {
              mapaResponsables[nombreEncargado] = { email: emailEncargado, tareas: [], programa: name };
            }
            mapaResponsables[nombreEncargado].tareas.push({
              actividad: descripcion,
              vence: fechaLimite,
              alerta: diffDias < 0 ? "TIEMPO VENCIDO" : (diffDias === 0 ? "VENCE HOY" : (diffDias === 1 ? "VENCE MAÑANA" : "VENCE EN 7 DÍAS")),
              color: diffDias < 0 ? "#8B0000" : (diffDias === 0 ? "#FF4757" : (diffDias === 1 ? "#FFA502" : "#2ED573")),
              sedes: sedesStatus,
              periodo: periodo
            });
          }
        }
      }
    }
  });

  for (let nombre in mapaResponsables) {
    const info = mapaResponsables[nombre];
    const fechaAsunto = Utilities.formatDate(info.tareas[0].vence, "GMT-5", "dd/MM/yyyy");
    const asunto = "RECORDATORIO CHECK LIST INICIO DE CLASES: " + info.programa + " | " + fechaAsunto;
    enviarEmailBubble3D(info.email, nombre, info.tareas, asunto, info.programa);
  }
}

function enviarEmailBubble3D(emailDestino, nombre, actividades, asunto, programa) {
  
  let cardsHtml = actividades.map(a => {
    // Generar las burbujas de sedes (Badges)
    const sedesHtml = a.sedes.map(s => {
      const bgColor = s.completado ? "#E3FCEF" : "#FEEBEB";
      const textColor = s.completado ? "#006644" : "#BF2600";
      const icon = s.completado ? "✓" : "○";
      return `<span style="display: inline-block; background: ${bgColor}; color: ${textColor}; padding: 2px 8px; border-radius: 6px; font-size: 10px; font-weight: bold; margin-right: 4px; margin-bottom: 4px; border: 1px solid ${textColor}22;">${icon} ${s.nombre}</span>`;
    }).join("");

    return `
    <div style="background: #ffffff; border-radius: 20px; padding: 20px; margin-bottom: 20px; box-shadow: 0 10px 25px rgba(0,0,0,0.07); border: 1px solid #f0f0f0;">
      <table width="100%" cellpadding="0" cellspacing="0" border="0">
        <tr>
          <td>
            <span style="background: ${a.color}; color: white; padding: 3px 10px; border-radius: 50px; font-size: 10px; font-weight: bold; text-transform: uppercase; letter-spacing: 0.5px;">${a.alerta}</span>
            <span style="color: #A4B0BE; font-size: 11px; font-weight: bold; margin-left: 10px;">${a.periodo}</span>
          </td>
          <td align="right" style="color: #2F3542; font-size: 12px; font-weight: bold;">
            ${Utilities.formatDate(a.vence, "GMT-5", "dd/MM HH:mm")}
          </td>
        </tr>
      </table>
      <div style="margin-top: 12px; color: #2D3436; font-size: 15px; font-weight: 600; line-height: 1.4; font-family: 'Segoe UI', Roboto, sans-serif;">
        ${a.actividad}
      </div>
      <div style="margin-top: 15px; border-top: 1px solid #f8f9fa; padding-top: 12px;">
        <div style="color: #A4B0BE; font-size: 9px; font-weight: bold; text-transform: uppercase; margin-bottom: 8px; letter-spacing: 1px;">Estado por Sedes:</div>
        ${sedesHtml}
      </div>
    </div>`;
  }).join("");

  const body = `
    <div style="background-color: #F4F7FA; padding: 40px 15px; font-family: 'Segoe UI', Roboto, Helvetica, Arial, sans-serif;">
      <div style="max-width: 550px; margin: auto;">
        
        <div style="background: linear-gradient(135deg, #004691 0%, #003366 100%); padding: 35px; border-radius: 30px; text-align: center; margin-bottom: 25px; box-shadow: 0 15px 30px rgba(0,70,145,0.2);">
          <h1 style="color: #FFFFFF; margin: 0; font-size: 24px; font-weight: 800;">ENCARGADO: ${nombre} </h1>
          <p style="color: #FFCC00; margin: 8px 0 0; font-size: 13px; font-weight: bold; text-transform: uppercase; letter-spacing: 1px;">${programa} - Recordatorio de Gestión &#128276;</p>
        </div>

        <div style="margin-bottom: 30px;">
          <p style="color: #57606F; font-size: 14px; text-align: center; margin-bottom: 20px;">Tienes las siguientes actividades próximas a vencer:</p>
          ${cardsHtml}
        </div>

        <div style="text-align: center; background: white; padding: 30px; border-radius: 30px; box-shadow: 0 10px 20px rgba(0,0,0,0.05);">
          <a href="${SpreadsheetApp.getActiveSpreadsheet().getUrl()}" 
             style="background: #FFCC00; color: #004691; padding: 18px 40px; border-radius: 50px; text-decoration: none; font-weight: 800; font-size: 16px; display: inline-block; box-shadow: 0 8px 15px rgba(255,204,0,0.3);">
            REVISAR MI CHECKLIST
          </a>
          <p style="color: #A4B0BE; font-size: 10px; margin-top: 20px;">© 2026 Academia Vonex | Sistema de Control Operativo</p>
        </div>
      </div>
    </div>`;

  GmailApp.sendEmail(emailDestino, asunto, "", { htmlBody: body });
}