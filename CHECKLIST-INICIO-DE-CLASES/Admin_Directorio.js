/**
 * =========================================================================
 * SISTEMA DE GESTIÓN Y RECORDATORIOS VONEX - MÓDULO ADMINISTRATIVO
 * =========================================================================
 * 
 * Funcionalidades:
 * 1. Inicialización y formateo de la hoja "ADMIN_DIRECTORIO" (Estética Vonex Cristal).
 * 2. Directorio Maestro de Sedes (LCE, SJL, IND, COM, SAN, PTP, VMT).
 * 3. Detección automática y dinámica de actividades que vencen HOY.
 * 4. Generación de mensajes y enlaces de WhatsApp individuales por Director (con formato solicitado).
 * 5. Envío manual de correos con plantilla Cristal (Glassmorphic) de alta compatibilidad.
 * 
 * NOTA DE SEGURIDAD: NO activa ningún trigger automático. Todo es de ejecución manual o bajo demanda.
 */

// CONFIGURACIÓN GLOBAL
const CONFIG_VONEX = {
  NOMBRE_HOJA_ADMIN: "ADMIN_DIRECTORIO",
  HOJA_CHECKLIST_DEFAULT: "UNMSM",
  HORA_LIMITE_DEFAULT: "6:00 p. m.",
  COLOR_PRIMARIO: "#004691",
  COLOR_SECUNDARIO: "#FFCC00",
  COLOR_FONDO_HEADER: "#0f2b5c",
  COLOR_CRISTAL_CARD: "#f8fafc"
};

/**
 * Menú personalizado en Google Sheets al abrir el archivo
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("🚀 Gestión Vonex")
    .addItem("⚙️ 1. Crear / Configurar Hoja ADMIN_DIRECTORIO", "configurarHojaAdminDirectorio")
    .addItem("📲 2. Generar Mensajes y Enlaces de WhatsApp (Hoy)", "generarEnlacesWhatsAppHoy")
    .addItem("📧 3. Enviar Correos de Recordatorio (Modo Manual)", "enviarCorreosRecordatorioManual")
    .addSeparator()
    .addItem("📊 4. Ver Resumen de Pendientes de Hoy", "mostrarResumenConsola")
    .addToUi();
}

/**
 * Crea y formatea la pestaña ADMIN_DIRECTORIO con diseño limpio y profesional
 */
function configurarHojaAdminDirectorio() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG_VONEX.NOMBRE_HOJA_ADMIN);

  if (!sheet) {
    sheet = ss.insertSheet(CONFIG_VONEX.NOMBRE_HOJA_ADMIN);
  }

  sheet.setTabColor(CONFIG_VONEX.COLOR_PRIMARIO);

  // 1. Título y Encabezado Principal
  sheet.getRange("A1:H1").merge()
    .setValue("ACADEMIA VONEX - PANEL DE CONTROL Y DIRECTORIO DE SEDES")
    .setFontFamily("Segoe UI")
    .setFontSize(14)
    .setFontWeight("bold")
    .setFontColor("#FFFFFF")
    .setBackground(CONFIG_VONEX.COLOR_PRIMARIO)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");
  sheet.setRowHeight(1, 40);

  // Subtítulo informativo
  sheet.getRange("A2:H2").merge()
    .setValue("Administración de notificaciones, enlaces de WhatsApp y recordatorios para Directores de Sede")
    .setFontFamily("Segoe UI")
    .setFontSize(10)
    .setFontColor("#64748b")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");
  sheet.setRowHeight(2, 24);

  // 2. SECCIÓN A: DIRECTORIO MAESTRO DE SEDES
  sheet.getRange("A4:H4").merge()
    .setValue("📋 DIRECTORIO MAESTRO DE DIRECTORES DE SEDE")
    .setFontFamily("Segoe UI")
    .setFontSize(11)
    .setFontWeight("bold")
    .setFontColor("#004691")
    .setBackground("#e0f2fe");

  const headersDirectorio = [
    ["Código", "Nombre Sede", "Director(a) (Nombre en WhatsApp)", "Celular WhatsApp (51...)", "Correo Electrónico", "Notificar WhatsApp", "Notificar Correo", "Observaciones"]
  ];
  sheet.getRange("A5:H5").setValues(headersDirectorio)
    .setFontFamily("Segoe UI")
    .setFontSize(10)
    .setFontWeight("bold")
    .setFontColor("#FFFFFF")
    .setBackground("#004691")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");
  sheet.setRowHeight(5, 28);

  // Datos base de las 7 sedes (Director de SJL inicializado como "Joe" según indicación)
  const sedesBase = [
    ["LCE", "La Central", "Director LCE", "51987654321", "director.lce@vonex.edu.pe", "SÍ", "SÍ", ""],
    ["SJL", "San Juan de Lurigancho", "Joe", "51987654322", "director.sjl@vonex.edu.pe", "SÍ", "SÍ", ""],
    ["IND", "Independencia", "Director IND", "51987654323", "director.ind@vonex.edu.pe", "SÍ", "SÍ", ""],
    ["COM", "Comas", "Director COM", "51987654324", "director.com@vonex.edu.pe", "SÍ", "SÍ", ""],
    ["SAN", "Santa Anita", "Director SAN", "51987654325", "director.san@vonex.edu.pe", "SÍ", "SÍ", ""],
    ["PTP", "Puente Piedra", "Director PTP", "51987654326", "director.ptp@vonex.edu.pe", "SÍ", "SÍ", ""],
    ["VMT", "Villa María del Triunfo", "Director VMT", "51987654327", "director.vmt@vonex.edu.pe", "SÍ", "SÍ", ""]
  ];

  // Si las celdas están vacías, llenar los valores base
  if (sheet.getRange("A6").getValue() === "") {
    sheet.getRange("A6:H12").setValues(sedesBase)
      .setFontFamily("Segoe UI")
      .setFontSize(10)
      .setVerticalAlignment("middle");
  }

  // Validación de datos para SÍ/NO
  const reglaSiNo = SpreadsheetApp.newDataValidation().requireValueInList(["SÍ", "NO"]).build();
  sheet.getRange("F6:G12").setDataValidation(reglaSiNo);

  // 3. SECCIÓN B: CONSOLA DE ENVÍOS DE RECORDATORIOS DE HOY
  sheet.getRange("A15:H15").merge()
    .setValue("📲 CONSOLA DE ENVÍOS DE RECORDATORIOS (ACTIVIDADES DE HOY)")
    .setFontFamily("Segoe UI")
    .setFontSize(11)
    .setFontWeight("bold")
    .setFontColor("#004691")
    .setBackground("#fef3c7");

  const headersConsola = [
    ["Sede", "Director(a)", "Tareas Vencen Hoy", "Estado Cumplimiento", "Acción WhatsApp Web / Móvil", "Acción Correo", "Último Recordatorio", "Mensaje Generado"]
  ];
  sheet.getRange("A16:H16").setValues(headersConsola)
    .setFontFamily("Segoe UI")
    .setFontSize(10)
    .setFontWeight("bold")
    .setFontColor("#FFFFFF")
    .setBackground("#1e293b")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");
  sheet.setRowHeight(16, 28);

  // Anchos de columna optimizados
  sheet.setColumnWidth(1, 80);   // Código
  sheet.setColumnWidth(2, 170);  // Sede
  sheet.setColumnWidth(3, 190);  // Director
  sheet.setColumnWidth(4, 180);  // Teléfono / Tareas
  sheet.setColumnWidth(5, 230);  // Correo / WhatsApp link
  sheet.setColumnWidth(6, 140);  // Notif / Correo
  sheet.setColumnWidth(7, 160);  // Notif / Último
  sheet.setColumnWidth(8, 250);  // Observaciones / Mensaje

  ss.toast("Hoja ADMIN_DIRECTORIO configurada con éxito.", "🚀 Vonex Control", 4);
}

/**
 * Obtiene el directorio maestro mapeado por código de sede (LCE, SJL, etc.)
 * Lee dinámicamente por nombres de encabezados de columna para máxima escalabilidad.
 */
function obtenerDirectorioSedes(sheetAdmin) {
  const allData = sheetAdmin.getDataRange().getValues();
  if (allData.length < 2) return {};

  let headerRowIndex = -1;
  let colMap = {};

  for (let r = 0; r < Math.min(allData.length, 15); r++) {
    const fila = allData[r].map(c => (c || "").toString().trim().toUpperCase());
    if (fila.some(c => c.includes("ROL")) && fila.some(c => c.includes("CORREO"))) {
      headerRowIndex = r;
      fila.forEach((h, c) => {
        if (!h) return;
        if (h.includes("ROL")) colMap.rol = c;
        if (h.includes("ABREVIATURA") || h.includes("CODIGO")) colMap.codigo = c;
        if (h.includes("SEDE") && !h.includes("ROL") && !h.includes("CODIGO") && !h.includes("ABREVIATURA")) colMap.sede = c;
        if (h.includes("NOMBRE") || h.includes("DIRECTOR") || h.includes("RESPONSABLE")) {
          if (!h.includes("SEDE")) colMap.nombre = c;
        }
        if (h.includes("CELULAR") || h.includes("TELEFONO") || (h.includes("WHATSAPP") && !h.includes("NOTIFICAR"))) colMap.celular = c;
        if (h.includes("CORREO") && !h.includes("NOTIFICAR")) colMap.correo = c;
        if (h.includes("NOTIFICAR") && h.includes("WHATSAPP")) colMap.notifWsp = c;
        if (h.includes("NOTIFICAR") && h.includes("CORREO")) colMap.notifCorreo = c;
      });
      break;
    }
  }

  if (headerRowIndex === -1) return {};

  let directorio = {};
  for (let r = headerRowIndex + 1; r < allData.length; r++) {
    const row = allData[r];
    const cod = colMap.codigo !== undefined ? (row[colMap.codigo] || "").toString().trim().toUpperCase() : "";
    if (cod && cod !== "INACTIVO") {
      directorio[cod] = {
        codigo: cod,
        nombreSede: colMap.sede !== undefined ? (row[colMap.sede] || cod).toString().trim() : cod,
        director: colMap.nombre !== undefined ? (row[colMap.nombre] || "Director(a)").toString().trim() : "Director(a)",
        celular: colMap.celular !== undefined ? (row[colMap.celular] || "").toString().replace(/[^0-9]/g, "") : "",
        correo: colMap.correo !== undefined ? (row[colMap.correo] || "").toString().trim() : "",
        notificarWsp: colMap.notifWsp !== undefined ? (row[colMap.notifWsp] || "SÍ").toString().toUpperCase() === "SÍ" : true,
        notificarCorreo: colMap.notifCorreo !== undefined ? (row[colMap.notifCorreo] || "SÍ").toString().toUpperCase() === "SÍ" : true
      };
    }
  }

  return directorio;
}

/**
 * Detecta y extrae dinámicamente las actividades que vencen HOY en la hoja de checklist
 */
function obtenerActividadesVencenHoy(nombreHojaChecklist) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(nombreHojaChecklist || CONFIG_VONEX.HOJA_CHECKLIST_DEFAULT);
  
  if (!sheet) {
    throw new Error("No se encontró la hoja de checklist: " + nombreHojaChecklist);
  }

  const data = sheet.getDataRange().getValues();
  const hoy = new Date();
  const hoyDia = hoy.getDate();
  const hoyMes = hoy.getMonth(); // 0-indexed (9 = octubre)

  // 1. Detectar columnas de Sedes en fila 7 (índice 6)
  const filaEncabezados = data[6];
  let columnasSedes = []; // [{ codigo: 'LCE', colIdx: 4 }, ...]

  for (let c = 4; c < filaEncabezados.length; c++) {
    const texto = (filaEncabezados[c] || "").toString().trim().toUpperCase();
    if (["LCE", "SJL", "IND", "COM", "SAN", "PTP", "VMT"].includes(texto)) {
      columnasSedes.push({ codigo: texto, colIdx: c });
    }
  }

  // Detectar columnas de Fecha Límite y Descripción
  let colFechaIdx = -1;
  let colHoraIdx = -1;
  let colDescIdx = 3; // Por defecto columna D (índice 3)
  let colActividadIdx = 1; // Por defecto columna B (índice 1)

  for (let c = 0; c < filaEncabezados.length; c++) {
    const header = (filaEncabezados[c] || "").toString().trim().toUpperCase();
    if (header.includes("FECHA LÍMITE") || header === "FECHA") colFechaIdx = c;
    if (header.includes("HORA LÍMITE") || header === "HORA") colHoraIdx = c;
    if (header.includes("DESCRIPCIÓN")) colDescIdx = c;
  }

  // Fallback si no tienen texto exacto en fila 7
  if (colFechaIdx === -1) colFechaIdx = 20; // Columna U
  if (colHoraIdx === -1) colHoraIdx = 21;  // Columna V

  let actividadesHoy = [];

  // Recorrer filas desde la 8 (índice 7) en adelante
  for (let r = 7; r < data.length; r++) {
    const row = data[r];
    const fechaVal = row[colFechaIdx];
    const desc = (row[colDescIdx] || "").toString().trim();
    const actividadNombre = (row[colActividadIdx] || "").toString().trim();

    if (!desc || desc.toUpperCase().startsWith("ACTIVIDADES")) continue;

    let esHoy = false;

    if (fechaVal instanceof Date) {
      if (fechaVal.getDate() === hoyDia && fechaVal.getMonth() === hoyMes) {
        esHoy = true;
      }
    } else if (typeof fechaVal === "string" && fechaVal.length > 0) {
      // Comparar cadenas tipo "22 de oct", "jue, 22 de oct"
      const textoFecha = fechaVal.toLowerCase();
      if (textoFecha.includes(hoyDia.toString()) && (textoFecha.includes("oct") || textoFecha.includes("10"))) {
        esHoy = true;
      }
    }

    if (esHoy) {
      // Estado de cada sede para esta actividad
      let estadosPorSede = {};
      columnasSedes.forEach(s => {
        const valCelda = row[s.colIdx];
        // En Sheets checkbox: true = marcado, false = desmarcado
        const completada = (valCelda === true || valCelda === "REALIZADO" || valCelda === "NO APLICA");
        estadosPorSede[s.codigo] = completada;
      });

      actividadesHoy.push({
        fila: r + 1,
        actividad: actividadNombre,
        descripcion: desc,
        hora: (row[colHoraIdx] || CONFIG_VONEX.HORA_LIMITE_DEFAULT).toString().trim(),
        estados: estadosPorSede
      });
    }
  }

  return {
    actividades: actividadesHoy,
    sedes: columnasSedes.map(s => s.codigo)
  };
}

/**
 * Genera el texto exacto solicitado para el mensaje de WhatsApp
 */
function armarMensajeWhatsApp(directorNombre, sedeCodigo, tareasPendientes, urlChecklist) {
  // Lista de emojis numerados
  const emojisNum = ["1️⃣", "2️⃣", "3️⃣", "4️⃣", "5️⃣", "6️⃣", "7️⃣", "8️⃣", "9️⃣", "🔟"];

  let lineasTareas = "";
  tareasPendientes.forEach((t, idx) => {
    const num = emojisNum[idx] || `*${idx + 1}.*`;
    let textoItem = t.descripcion.replace(/^-\s*/, "").trim();

    // Formateo visual con sangría para subtareas específicas conocidas
    if (textoItem.toLowerCase().includes("actualizar lista de salones")) {
      lineasTareas += `${num} *Actualizar lista de salones*\n   └ Incluir Nombre del tutor / Ciclo / Aula al costado del croquis\n`;
    } else if (textoItem.toLowerCase().includes("colocar un mural del croquis")) {
      lineasTareas += `${num} *Colocar un mural del croquis en la puerta de la sede*\n   └ Con los números de aulas correspondientes de cada ciclo\n`;
    } else {
      lineasTareas += `${num} *${textoItem}*\n`;
    }
  });

  const mensaje = 
`🚨 *RECORDATORIO | CHECKLIST INICIO DE CLASES* 🚨
Buen día ${directorNombre}, le recordamos que *⏰HOY a las 6:00 p. m.* vence el plazo para el cumplimiento de las siguientes actividades:
${lineasTareas}🔗 *Marque su cumplimiento directamente aquí:*
👉 ${urlChecklist}

*Gestión de Procesos | Academia Vonex*`;

  return mensaje;
}

/**
 * 📲 FUNCIÓN PRINCIPAL: Genera los enlaces de WhatsApp y llena la consola en ADMIN_DIRECTORIO
 */
function generarEnlacesWhatsAppHoy() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheetAdmin = ss.getSheetByName(CONFIG_VONEX.NOMBRE_HOJA_ADMIN);

  if (!sheetAdmin) {
    configurarHojaAdminDirectorio();
    sheetAdmin = ss.getSheetByName(CONFIG_VONEX.NOMBRE_HOJA_ADMIN);
  }

  const directorio = obtenerDirectorioSedes(sheetAdmin);
  const resultado = obtenerActividadesVencenHoy(CONFIG_VONEX.HOJA_CHECKLIST_DEFAULT);
  const actividadesHoy = resultado.actividades;
  const urlChecklist = ss.getUrl() + "#gid=" + ss.getSheetByName(CONFIG_VONEX.HOJA_CHECKLIST_DEFAULT).getSheetId();

  if (actividadesHoy.length === 0) {
    SpreadsheetApp.getUi().alert("Aviso", "No se detectaron actividades con fecha límite para el día de hoy en la hoja " + CONFIG_VONEX.HOJA_CHECKLIST_DEFAULT, SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  let filasConsola = [];
  const codigosSedes = ["LCE", "SJL", "IND", "COM", "SAN", "PTP", "VMT"];

  codigosSedes.forEach(cod => {
    const dirInfo = directorio[cod] || {
      nombreSede: cod,
      director: "Director(a)",
      celular: "",
      notificarWsp: true
    };

    // Filtrar tareas que están PENDIENTES para esta sede específica
    const pendientesSede = actividadesHoy.filter(act => act.estados[cod] === false);
    const cantPendientes = pendientesSede.length;

    let estadoTexto = "";
    let formulaWsp = "";
    let mensajeTexto = "";

    if (cantPendientes === 0) {
      estadoTexto = "✅ Al día (0 pendientes)";
      formulaWsp = "Sin pendientes";
      mensajeTexto = "Todas las actividades de hoy están completadas.";
    } else {
      estadoTexto = `⚠️ ${cantPendientes} de ${actividadesHoy.length} pendientes`;
      mensajeTexto = armarMensajeWhatsApp(dirInfo.director, cod, pendientesSede, urlChecklist);

      if (dirInfo.celular && dirInfo.celular.length >= 8) {
        const urlWsp = "https://api.whatsapp.com/send?phone=" + dirInfo.celular + "&text=" + encodeURIComponent(mensajeTexto);
        formulaWsp = `=HYPERLINK("${urlWsp}", "📲 ENVIAR WHATSAPP (${dirInfo.director})")`;
      } else {
        formulaWsp = "⚠️ Falta celular en dir.";
      }
    }

    filasConsola.push([
      cod + " - " + dirInfo.nombreSede,
      dirInfo.director,
      cantPendientes,
      estadoTexto,
      formulaWsp,
      dirInfo.notificarCorreo ? "Listo para enviar" : "Desactivado",
      "—",
      mensajeTexto
    ]);
  });

  // Escribir en la sección de consola (desde la fila 17)
  sheetAdmin.getRange("A17:H23").clearContent();
  sheetAdmin.getRange("A17:H23").setValues(filasConsola)
    .setFontFamily("Segoe UI")
    .setFontSize(10)
    .setVerticalAlignment("middle");

  // Formato visual para los botones de WhatsApp
  sheetAdmin.getRange("E17:E23")
    .setFontWeight("bold")
    .setFontColor("#004691")
    .setHorizontalAlignment("center");

  ss.toast("✅ Enlaces de WhatsApp generados con éxito en ADMIN_DIRECTORIO.", "🚀 Éxito", 5);
}

/**
 * 📧 FUNCIÓN: Envío MANUAL de correos con plantilla Cristal a los Directores
 * NOTA: Requiere confirmación del usuario para evitar envíos involuntarios.
 */
function enviarCorreosRecordatorioManual() {
  const ui = SpreadsheetApp.getUi();
  const confirmacion = ui.alert(
    "Confirmación de Envío de Correos",
    "¿Está seguro de enviar los recordatorios por correo a los directores de sede con actividades pendientes de hoy?\n\n(Solo se enviará a quienes tengan 'Notificar Correo = SÍ' en el Directorio)",
    ui.ButtonSet.YES_NO
  );

  if (confirmacion !== ui.Button.YES) {
    ui.alert("Envío cancelado", "No se ha enviado ningún correo.", ui.ButtonSet.OK);
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetAdmin = ss.getSheetByName(CONFIG_VONEX.NOMBRE_HOJA_ADMIN);
  if (!sheetAdmin) {
    ui.alert("Error", "Primero configure la hoja ADMIN_DIRECTORIO.", ui.ButtonSet.OK);
    return;
  }

  const directorio = obtenerDirectorioSedes(sheetAdmin);
  const resultado = obtenerActividadesVencenHoy(CONFIG_VONEX.HOJA_CHECKLIST_DEFAULT);
  const actividadesHoy = resultado.actividades;
  const urlChecklist = ss.getUrl() + "#gid=" + ss.getSheetByName(CONFIG_VONEX.HOJA_CHECKLIST_DEFAULT).getSheetId();

  let totalEnviados = 0;
  const codigosSedes = ["LCE", "SJL", "IND", "COM", "SAN", "PTP", "VMT"];

  codigosSedes.forEach((cod, idx) => {
    const dirInfo = directorio[cod];
    if (!dirInfo || !dirInfo.notificarCorreo || !dirInfo.correo || !dirInfo.correo.includes("@")) {
      return;
    }

    const pendientesSede = actividadesHoy.filter(act => act.estados[cod] === false);
    if (pendientesSede.length === 0) return; // Si ya cumplió todo, no se envía alerta

    const asunto = `🚨 URGENTE: ${pendientesSede.length} Actividades con Vencimiento HOY (18:00 hrs) - Sede ${dirInfo.nombreSede}`;
    const htmlBody = generarHtmlEmailCristal(dirInfo.nombreSede, dirInfo.director, CONFIG_VONEX.HOJA_CHECKLIST_DEFAULT, pendientesSede, urlChecklist);

    try {
      GmailApp.sendEmail(dirInfo.correo, asunto, "", {
        htmlBody: htmlBody,
        name: "PROCESOS"
      });
      totalEnviados++;
      // Registrar timestamp en columna G de la consola
      sheetAdmin.getRange(17 + idx, 7).setValue(new Date()).setNumberFormat("dd/MM HH:mm");
    } catch (err) {
      Logger.log("Error al enviar a " + dirInfo.correo + ": " + err.message);
    }
  });

  ui.alert("Envío Completado", `Se enviaron ${totalEnviados} correos de recordatorio exitosamente.`, ui.ButtonSet.OK);
}

/**
 * Genera el HTML institucional de ancho completo (100%) para el correo de recordatorio
 */
function generarHtmlEmailCristal(sedeNombre, directorNombre, ciclo, pendientes, urlSheet) {
  let filasTareasHtml = "";

  pendientes.forEach((t, i) => {
    const descLimpia = t.descripcion.replace(/^-\s*/, "");
    filasTareasHtml += `
      <div style="background: #ffffff; border: 1px solid #e2e8f0; border-left: 4px solid #dc2626; border-radius: 12px; padding: 14px 18px; margin-bottom: 10px;">
        <table width="100%" cellpadding="0" cellspacing="0">
          <tr>
            <td style="vertical-align: top;">
              <span style="background: #e0f2fe; color: #0369a1; font-size: 10px; font-weight: 800; padding: 3px 9px; border-radius: 6px;">ACTIVIDAD ${i + 1}</span>
              <div style="font-size: 13.5px; font-weight: 700; color: #0f172a; margin-top: 6px;">${descLimpia}</div>
            </td>
            <td align="right" style="vertical-align: top; min-width: 100px;">
              <div style="font-size: 11px; font-weight: 800; color: #dc2626; background: #fef2f2; border: 1px solid #fecaca; padding: 4px 10px; border-radius: 9999px; display: inline-block;">
                &#9200; ${t.hora}
              </div>
              <div style="font-size: 10px; color: #64748b; font-weight: 700; margin-top: 4px; text-transform: uppercase;">PENDIENTE</div>
            </td>
          </tr>
        </table>
      </div>`;
  });

  const contenidoHtml = `
    <div style="background-color:#ffffff; border:1px solid #e2e8f0; border-radius:12px; padding:18px 20px; margin-bottom:16px;">
      <p style="margin:0 0 8px 0; font-size:14.5px; font-weight:700; color:#0f172a;">
        Buen d&#237;a, ${directorNombre}. &#128075;
      </p>
      <p style="margin:0; font-size:13.5px; color:#334155; line-height:1.55;">
        Le recordamos que <strong>HOY a las 6:00 p. m.</strong> vence el plazo para el cumplimiento de las siguientes actividades operativas correspondientes al ciclo <strong>${ciclo}</strong>:
      </p>
    </div>

    <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="width:100%; margin-bottom:16px;">
      <tr>
        <td width="49%" valign="top" style="background-color:#ffffff; border:1px solid #e2e8f0; border-top:3px solid #0d3b99; border-radius:12px; padding:16px 18px;">
          <div style="font-size:11px; font-weight:800; color:#0d3b99; text-transform:uppercase; letter-spacing:0.5px; margin-bottom:6px;">Sede Asignada</div>
          <div style="font-size:16px; font-weight:800; color:#0f172a;">${sedeNombre}</div>
        </td>
        <td width="2%">&nbsp;</td>
        <td width="49%" valign="top" style="background-color:#ffffff; border:1px solid #e2e8f0; border-top:3px solid #facc15; border-radius:12px; padding:16px 18px;">
          <div style="font-size:11px; font-weight:800; color:#854d0e; text-transform:uppercase; letter-spacing:0.5px; margin-bottom:6px;">Pendientes Hoy</div>
          <div style="font-size:18px; font-weight:900; color:#dc2626;">${pendientes.length}</div>
        </td>
      </tr>
    </table>

    ${filasTareasHtml}
  `;

  return construirCorreoVonexFullWidth({
    titulo: "RECORDATORIO DE VENCIMIENTO HOY (18:00 H)",
    subtituloHtml: `SEDE: <span style="color: #fde047;">${sedeNombre}</span> &nbsp;|&nbsp; CICLO: <span style="color: #fde047;">${ciclo}</span>`,
    contenidoHtml: contenidoHtml,
    urlBoton: urlSheet,
    textoBoton: "REGISTRAR CUMPLIMIENTO"
  });
}

/**
 * Muestra un resumen rápido de las tareas de hoy en un cuadro de diálogo
 */
function mostrarResumenConsola() {
  const ui = SpreadsheetApp.getUi();
  const resultado = obtenerActividadesVencenHoy(CONFIG_VONEX.HOJA_CHECKLIST_DEFAULT);
  const actividades = resultado.actividades;

  if (actividades.length === 0) {
    ui.alert("Resumen de Hoy", "No hay actividades con vencimiento hoy.", ui.ButtonSet.OK);
    return;
  }

  let texto = `Se detectaron ${actividades.length} actividades con vencimiento HOY:\n\n`;
  actividades.forEach((a, i) => {
    texto += `${i + 1}. [Fila ${a.fila}] ${a.descripcion} (Hora: ${a.hora})\n`;
  });

  ui.alert("Actividades de Hoy (" + CONFIG_VONEX.HOJA_CHECKLIST_DEFAULT + ")", texto, ui.ButtonSet.OK);
}
