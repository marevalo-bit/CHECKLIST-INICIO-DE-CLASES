/**
 * =========================================================================
 * BACKEND DEL MÓDULO WEB INTERACTIVO - GESTIÓN Y CUMPLIMIENTO VONEX
 * =========================================================================
 * 
 * Este archivo provee los servicios web para:
 * 1. Control de acceso por roles (PROCESOS y DIRECTOR ZONAL).
 * 2. Marcación de cumplimiento en tiempo real.
 * 3. Gestión y pactación de fechas del cronograma.
 * 4. Control de automatización (Modo MANUAL por defecto vs Trigger opcional).
 * 5. Envío de recordatorios por WhatsApp y Correo (Estética renovada con badges por sede).
 * 6. Reporte ejecutivo con vista previa.
 * 7. Correo oficial de apertura a los 18 responsables.
 */

/**
 * Normaliza cadenas para comparaciones robustas e insensibles a mayúsculas y tildes.
 */
function normalizarTexto(txt) {
  return (txt || "")
    .toString()
    .trim()
    .toUpperCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "");
}

/**
 * Lee y extrae dinámicamente todo el directorio de sedes, directores, sistemas, 
 * procesos y dirección zonal directamente de la hoja "ADMIN_DIRECTORIO".
 * No se identifican columnas por números ni índices fijos: se mapean por nombres de encabezados.
 */
function leerDirectorioDesdeSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("ADMIN_DIRECTORIO");
  if (!sheet) return [];

  const allData = sheet.getDataRange().getValues();
  if (allData.length < 2) return [];

  // 1. Localizar la fila de encabezados por sus nombres de columnas (ROL, Correo Electrónico, etc.)
  let headerRowIndex = -1;
  let colMap = {};

  for (let r = 0; r < allData.length; r++) {
    const fila = allData[r].map(normalizarTexto);
    const tieneRol = fila.some(c => c === "ROL" || c.startsWith("ROL"));
    const tieneCorreo = fila.some(c => c.includes("CORREO"));

    if (tieneRol && tieneCorreo) {
      headerRowIndex = r;
      fila.forEach((header, c) => {
        if (!header) return;
        if (header.includes("ROL")) colMap.rol = c;
        if (header.includes("CODIGO")) colMap.codigo = c;
        if (header.includes("SEDE") && !header.includes("ROL") && !header.includes("CODIGO")) colMap.sede = c;
        if (header.includes("DIRECTOR") || header.includes("RESPONSABLE") || (header.includes("NOMBRE") && !header.includes("SEDE"))) colMap.nombre = c;
        if (header.includes("CELULAR") || header.includes("TELEFONO") || (header.includes("WHATSAPP") && !header.includes("NOTIFICAR"))) colMap.celular = c;
        if (header.includes("CORREO") && !header.includes("NOTIFICAR")) colMap.correo = c;
        if (header.includes("NOTIFICAR") && header.includes("WHATSAPP")) colMap.notifWsp = c;
        if (header.includes("NOTIFICAR") && header.includes("CORREO")) colMap.notifCorreo = c;
        if (header.includes("OBSERV")) colMap.obs = c;
      });
      break;
    }
  }

  // Si no se encontró la cabecera por nombre de columna
  if (headerRowIndex === -1 || colMap.rol === undefined) {
    return [];
  }

  let usuarios = [];

  // 2. Recorrer las filas de datos debajo del encabezado
  for (let r = headerRowIndex + 1; r < allData.length; r++) {
    const row = allData[r];
    const filaTexto = row.map(c => (c || "").toString()).join(" ").toUpperCase();

    // Detener inmediatamente si se alcanza otra sección (ej. CONSOLA DE ENVÍOS DE RECORDATORIOS)
    if (filaTexto.includes("CONSOLA") || filaTexto.includes("RECORDATORIOS (ACTIVIDADES")) {
      break;
    }

    const rol = colMap.rol !== undefined ? (row[colMap.rol] || "").toString().trim().toUpperCase() : "";
    const codigoSede = colMap.codigo !== undefined ? (row[colMap.codigo] || "").toString().trim().toUpperCase() : "";
    const nombreSede = colMap.sede !== undefined ? (row[colMap.sede] || "").toString().trim() : "";
    const nombrePersona = colMap.nombre !== undefined ? (row[colMap.nombre] || "").toString().trim() : "";
    const celular = colMap.celular !== undefined ? (row[colMap.celular] || "").toString().replace(/[^0-9]/g, "") : "";
    const correoRaw = colMap.correo !== undefined ? (row[colMap.correo] || "").toString().trim() : "";
    const notifWsp = colMap.notifWsp !== undefined ? normalizarTexto(row[colMap.notifWsp]).includes("SI") : true;
    const notifCorreo = colMap.notifCorreo !== undefined ? normalizarTexto(row[colMap.notifCorreo]).includes("SI") : true;
    const obs = colMap.obs !== undefined ? (row[colMap.obs] || "").toString().trim() : "";

    // Si la fila está vacía, saltar
    if (!rol && !codigoSede && !nombrePersona && !correoRaw) continue;

    // Descartar registros inactivos
    if (codigoSede === "INACTIVO" || nombrePersona.toUpperCase() === "INACTIVO" || correoRaw.toUpperCase() === "INACTIVO") continue;

    // Extraer correos electrónicos válidos del campo Correo Electrónico
    const regexEmail = /[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/g;
    const correosEnCelda = (correoRaw.match(regexEmail) || []).map(c => c.toLowerCase().trim());
    const correosEnFila = (filaTexto.match(regexEmail) || []).map(c => c.toLowerCase().trim());

    const fuenteCorreos = correosEnCelda.length > 0 ? correosEnCelda : correosEnFila;
    const correosValidos = [];
    fuenteCorreos.forEach(c => {
      if (!correosValidos.includes(c) && !c.includes("inactivo")) {
        correosValidos.push(c);
      }
    });

    const esAccesoAdmin = (rol.includes("PROCESOS") || rol.includes("ZONAL"));

    if (correosValidos.length > 0) {
      correosValidos.forEach(em => {
        usuarios.push({
          email: em,
          nombre: nombrePersona || rol,
          rol: rol || "RESPONSABLE",
          sede: codigoSede || "CENTRAL",
          nombreSede: nombreSede || (codigoSede ? codigoSede : "Oficina Central"),
          tel: celular,
          notifWsp: notifWsp,
          notifCorreo: notifCorreo,
          obs: obs,
          tieneAccesoProcesos: esAccesoAdmin
        });
      });
    } else if (nombrePersona && !nombrePersona.toUpperCase().includes("INACTIVO")) {
      usuarios.push({
        email: "",
        nombre: nombrePersona,
        rol: rol || "RESPONSABLE",
        sede: codigoSede || "CENTRAL",
        nombreSede: nombreSede || (codigoSede ? codigoSede : "Oficina Central"),
        tel: celular,
        notifWsp: notifWsp,
        notifCorreo: notifCorreo,
        obs: obs,
        tieneAccesoProcesos: esAccesoAdmin
      });
    }
  }

  return usuarios;
}

/**
 * Resuelve el perfil, rol y permisos del usuario activo leyendo el sheet ADMIN_DIRECTORIO
 */
function resolverPerfilUsuario(emailActivo) {
  const emailNorm = (emailActivo || "").toLowerCase().trim();
  const directorio = leerDirectorioDesdeSheet();

  if (emailNorm) {
    // 1. Coincidencia exacta con correos registrados en ADMIN_DIRECTORIO
    const usuarioExacto = directorio.find(u => u.email && u.email === emailNorm);
    if (usuarioExacto) {
      return usuarioExacto;
    }

    // 2. Coincidencia por partes de nombre si el correo no está textual en la hoja
    const usuarioPorNombre = directorio.find(u => {
      if (!u.nombre) return false;
      const partes = u.nombre.toLowerCase().split(/\s+/).filter(p => p.length >= 3);
      return partes.some(p => emailNorm.includes(p));
    });
    if (usuarioPorNombre) {
      return Object.assign({}, usuarioPorNombre, { email: emailNorm });
    }

    // 3. Inferencia por palabras clave estándar en el correo
    if (emailNorm.includes("procesos") || emailNorm.includes("auditoria")) {
      return { email: emailNorm, nombre: "Gestión de Procesos", rol: "PROCESOS", sede: "CENTRAL", tieneAccesoProcesos: true };
    }
    if (emailNorm.includes("zonal")) {
      return { email: emailNorm, nombre: "Dirección Zonal", rol: "DIRECTOR ZONAL", sede: "CENTRAL", tieneAccesoProcesos: true };
    }
    if (emailNorm.includes("sistemas") || emailNorm.includes("helpdesk")) {
      return { email: emailNorm, nombre: "Equipo Sistemas", rol: "SISTEMAS", sede: "CENTRAL", tieneAccesoProcesos: false };
    }
    if (emailNorm.includes("evaluacion")) {
      return { email: emailNorm, nombre: "Evaluación y Control", rol: "EVALUACIÓN Y CONTROL", sede: "CENTRAL", tieneAccesoProcesos: false };
    }
  }

  // 4. Si no se especificó o viene anónimo: devolver el primer usuario registrado con correo
  const primerDirector = directorio.find(u => u.rol.includes("DIRECTOR") && u.email);
  if (primerDirector) {
    return Object.assign({}, primerDirector, { email: emailNorm || primerDirector.email });
  }

  const primerUsuarioConEmail = directorio.find(u => u.email);
  if (primerUsuarioConEmail) {
    return Object.assign({}, primerUsuarioConEmail, { email: emailNorm || primerUsuarioConEmail.email });
  }

  return {
    email: emailNorm || "",
    nombre: "Usuario",
    rol: "DIRECTOR DE SEDE",
    sede: "SJL",
    tieneAccesoProcesos: false
  };
}

/**
 * Obtiene la lista de destinatarios para Apertura Oficial dividida en Para y CC según ADMIN_DIRECTORIO
 */
function apiObtenerDestinatariosApertura() {
  const directorio = leerDirectorioDesdeSheet();
  const para = [];
  const cc = [];

  directorio.forEach(u => {
    if (!u.email || u.email.includes("inactivo")) return;
    if (u.rol.includes("ZONAL")) {
      if (!cc.includes(u.email)) cc.push(u.email);
    } else {
      if (!para.includes(u.email)) para.push(u.email);
    }
  });

  return { para: para, cc: cc };
}

/**
/**
 * Determina si una actividad pertenece al rol del usuario activo según ADMIN_DIRECTORIO
 */
function correspondeActividadAUsuario(actividad, usuario) {
  if (!usuario || !usuario.rol) return true;
  const rolUser = normalizarTexto(usuario.rol);
  const encargadoAct = normalizarTexto(actividad.encargado || "");

  // Roles globales que supervisan todas las actividades
  if (rolUser.includes("PROCESOS") || rolUser.includes("ZONAL") || rolUser.includes("AUDITOR")) {
    return true;
  }

  // DIRECTORES DE SEDE Y/O TUTORES (ej: Natalie Alcalá, Yhilner Benito, Margarita Chávez, etc.)
  if (rolUser.includes("DIRECTOR") || rolUser.includes("TUTOR")) {
    return encargadoAct.includes("DIRECTOR") || encargadoAct.includes("TUTOR") || encargadoAct.includes("SEDE");
  }

  // SISTEMAS (Walter Quispe)
  if (rolUser.includes("SISTEMA") || rolUser.includes("TI") || rolUser.includes("HELPDESK")) {
    return encargadoAct.includes("SISTEMA") || encargadoAct.includes("TI") || encargadoAct.includes("DRIVE") || encargadoAct.includes("INTRANET");
  }

  // EVALUACIÓN Y CONTROL (Roxana Quispe)
  if (rolUser.includes("EVALUA")) {
    return encargadoAct.includes("EVALUA");
  }

  // PRODUCCIÓN (Giancarlos Huaraca)
  if (rolUser.includes("PRODUC")) {
    return encargadoAct.includes("PRODUC");
  }

  // VENTAS (Jocelyn Pinedo)
  if (rolUser.includes("VENTA") || rolUser.includes("COMERCIAL")) {
    return encargadoAct.includes("VENTA");
  }

  // DIGITALIZACIÓN DE CONTENIDO (Margarita Amaya)
  if (rolUser.includes("DIGITALIZA")) {
    return encargadoAct.includes("DIGITALIZA");
  }

  return encargadoAct.includes(rolUser) || rolUser.includes(encargadoAct);
}

/**
 * Lee en vivo todas las actividades, fechas, horas y estados por sede directamente desde la hoja UNMSM
 */
function apiObtenerActividadesDesdeSheet(nombreHoja) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(nombreHoja || "UNMSM");
  if (!sheet) return [];

  const allData = sheet.getDataRange().getValues();
  if (!allData || allData.length < 8) return [];

  const hoy = new Date();
  const hoyCero = new Date(hoy.getFullYear(), hoy.getMonth(), hoy.getDate()).getTime();

  // 1. Detectar fila de encabezados y mapeo dinámico de columnas
  let headerRowIndex = -1;
  let colEncargado = -1;
  let colDesc = -1;
  let colFecha = -1;
  let colHora = -1;
  let colObs = -1;
  let sedesCols = [];

  for (let r = 0; r < Math.min(allData.length, 12); r++) {
    const fila = allData[r].map(normalizarTexto);
    const haySedes = fila.some(c => ["LCE", "SJL", "IND", "COM", "SAN", "PTP", "VMT", "HDM", "SJJ"].includes(c));
    if (haySedes) {
      headerRowIndex = r;
      allData[r].forEach((header, c) => {
        const hNorm = normalizarTexto(header);
        if (["LCE", "SJL", "IND", "COM", "SAN", "PTP", "VMT", "HDM", "SJJ"].includes(hNorm)) {
          sedesCols.push({ codigo: hNorm, colIdx: c });
        }
        if (hNorm.includes("ENCARGADO") || hNorm.includes("RESPONSABLE") || hNorm === "ROL") {
          colEncargado = c;
        }
        if (hNorm.includes("ACTIVIDAD") || hNorm.includes("DESCRIP") || hNorm === "DETALLE") {
          colDesc = c;
        }
        if (hNorm.includes("FECHA") && !hNorm.includes("REALIZADO") && colFecha === -1) {
          colFecha = c;
        }
        if (hNorm.includes("HORA") && colHora === -1) {
          colHora = c;
        }
        if (hNorm.includes("OBSERV") && colObs === -1) {
          colObs = c;
        }
      });
      break;
    }
  }

  // Fallbacks seguros según estructura histórica
  if (colEncargado === -1) colEncargado = 2; // Col C
  if (colDesc === -1) colDesc = 4;           // Col E
  if (colFecha === -1) colFecha = 21;        // Col V
  if (colHora === -1) colHora = 22;          // Col W
  if (sedesCols.length === 0) {
    const codigosDefault = ["LCE", "SJL", "IND", "COM", "SAN", "PTP", "VMT"];
    codigosDefault.forEach((cod, idx) => {
      sedesCols.push({ codigo: cod, colIdx: 5 + (idx * 2) });
    });
  }

  let actividades = [];
  let etapaActual = 1;
  const mesesAbr = ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "set", "oct", "nov", "dic"];
  const diasSem = ["dom", "lun", "mar", "mié", "jue", "vie", "sáb"];

  const startRow = (headerRowIndex !== -1 ? headerRowIndex + 1 : 8);
  for (let r = startRow; r < allData.length; r++) {
    const row = allData[r];
    const filaTexto = row.map(c => (c || "").toString()).join(" ").toUpperCase();

    // Detección de Etapas
    if (filaTexto.includes("ANTES DEL INICIO")) {
      etapaActual = 1;
      continue;
    }
    if (filaTexto.includes("PRIMERA SEMANA") || filaTexto.includes("DURANTE LA PRIMERA")) {
      etapaActual = 2;
      continue;
    }
    if (filaTexto.includes("5TA SEMANA") || filaTexto.includes("DESPUES DE INICIO") || filaTexto.includes("DESPUÉS DE INICIO")) {
      etapaActual = 3;
      continue;
    }

    let descRaw = (row[colDesc] || "").toString().trim();
    if (!descRaw && row[1]) descRaw = row[1].toString().trim();
    if (!descRaw || descRaw.toUpperCase().startsWith("ACTIVIDADES") || descRaw.length < 3) {
      continue;
    }

    const encargadoRaw = (row[colEncargado] || "").toString().trim() || "DIRECTOR DE SEDE  TUTORES";
    const fechaVal = row[colFecha];
    const horaVal = row[colHora];

    // Procesar fecha en vivo
    let fechaISO = "";
    let fechaTexto = "";
    let esHoy = false;
    let esVencida = false;
    let esManana = false;

    if (fechaVal instanceof Date && !isNaN(fechaVal.getTime())) {
      const anio = fechaVal.getFullYear();
      const mes = String(fechaVal.getMonth() + 1).padStart(2, '0');
      const dia = String(fechaVal.getDate()).padStart(2, '0');
      fechaISO = `${anio}-${mes}-${dia}`;
      fechaTexto = `${diasSem[fechaVal.getDay()]}, ${fechaVal.getDate()} ${mesesAbr[fechaVal.getMonth()]}`;

      const fechaCero = new Date(anio, fechaVal.getMonth(), fechaVal.getDate()).getTime();
      const diffDias = Math.round((fechaCero - hoyCero) / 86400000);
      esHoy = (diffDias === 0);
      esManana = (diffDias === 1);
      esVencida = (diffDias < 0);
    } else if (typeof fechaVal === "string" && fechaVal.trim() !== "") {
      fechaISO = fechaVal.trim();
      fechaTexto = fechaVal.trim();
      const textoLower = fechaVal.toLowerCase();
      const diaHoy = hoy.getDate().toString();
      if (textoLower.includes(diaHoy) && (textoLower.includes("oct") || textoLower.includes("10"))) {
        esHoy = true;
      }
    }

    // Procesar hora en vivo
    let horaStr = "18:00";
    let horaTexto = "6:00 p. m.";
    if (horaVal instanceof Date && !isNaN(horaVal.getTime())) {
      const hh = horaVal.getHours();
      const mm = String(horaVal.getMinutes()).padStart(2, '0');
      horaStr = `${String(hh).padStart(2, '0')}:${mm}`;
      const suf = hh >= 12 ? "p. m." : "a. m.";
      const hh12 = hh % 12 || 12;
      horaTexto = `${hh12}:${mm} ${suf}`;
    } else if (typeof horaVal === "string" && horaVal.trim() !== "") {
      horaStr = horaVal.trim();
      horaTexto = horaVal.trim();
    }

    // Estados reales por sede
    let estados = {};
    sedesCols.forEach(s => {
      const val = row[s.colIdx];
      const valNorm = normalizarTexto(val);
      const completado = (val === true || valNorm === "REALIZADO" || valNorm === "NO APLICA" || valNorm === "SI");
      estados[s.codigo] = completado;
    });

    actividades.push({
      id: actividades.length + 1,
      fila: r + 1,
      etapa: etapaActual,
      titulo: descRaw.replace(/^-\s*/, ""),
      desc: descRaw.replace(/^-\s*/, ""),
      encargado: encargadoRaw,
      fecha: fechaISO,
      fechaTexto: fechaTexto || "Por definir",
      hora: horaStr,
      horaTexto: horaTexto,
      esHoy: esHoy,
      esManana: esManana,
      esVencida: esVencida,
      estados: estados
    });
  }

  return actividades;
}

/**
 * Retorna el contexto inicial completo para el frontend leyendo en vivo UNMSM y ADMIN_DIRECTORIO
 */
function apiObtenerContextoInicial(pageParam) {
  const emailActivo = (Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail() || "").toLowerCase().trim();
  const perfil = resolverPerfilUsuario(emailActivo);
  const directorio = leerDirectorioDesdeSheet();
  const destinatariosApertura = apiObtenerDestinatariosApertura();
  const actividadesLive = apiObtenerActividadesDesdeSheet("UNMSM");

  return {
    email: emailActivo || perfil.email,
    rol: perfil.rol,
    sede: perfil.sede,
    nombre: perfil.nombre,
    tieneAccesoProcesos: perfil.tieneAccesoProcesos,
    pageParam: pageParam || "",
    directorio: directorio,
    destinatariosApertura: destinatariosApertura,
    actividades: actividadesLive
  };
}

/**
 * Punto de entrada Web App (Inyección Base64 a prueba de fallos sintácticos)
 */
function doGet(e) {
  let pageParam = "";
  if (e && e.parameter) {
    pageParam = (e.parameter.page || e.parameter.p || e.parameter.v || "").toString().trim().toUpperCase();
  }

  // Enrutamiento modular para desarrollo y vistas independientes
  if (pageParam === "CREAR" || pageParam === "PROCESOS" || pageParam === "ADMIN") {
    return HtmlService.createHtmlOutputFromFile("CREAR")
      .setTitle("Gestión y Creación de Cronogramas - Academia Vonex")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag("viewport", "width=device-width, initial-scale=1");
  }

  if (pageParam === "PARTICIPANTES" || pageParam === "PARTICIPANTE" || pageParam === "CHECKLIST") {
    return HtmlService.createHtmlOutputFromFile("PARTICIPANTES")
      .setTitle("Módulo de Participantes - Checklist Academia Vonex")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag("viewport", "width=device-width, initial-scale=1");
  }

  const contexto = apiObtenerContextoInicial(pageParam);
  let html = HtmlService.createHtmlOutputFromFile("MODULO_CHECKLIST_VONEX").getContent();
  
  // Inyección ultra segura mediante Base64 dentro del input hidden ya ubicado en el body
  const jsonStr = JSON.stringify(contexto);
  const b64 = Utilities.base64Encode(jsonStr, Utilities.Charset.UTF_8);
  html = html.replace('id="init-contexto-b64" value=""', 'id="init-contexto-b64" value="' + b64 + '"');

  return HtmlService.createHtmlOutput(html)
    .setTitle("Gestión de Inicio de Clases - Academia Vonex")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
}

/**
 * Valida el rol y permisos del usuario activo vía google.script.run
 */
function apiObtenerRolUsuario() {
  const emailActivo = (Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail() || "").toLowerCase().trim();
  return resolverPerfilUsuario(emailActivo);
}

/**
 * Registra el cumplimiento de una actividad en la hoja UNMSM
 */
function apiMarcarCumplimiento(codigoSede, filaSheet, marcarCompletado, motivoNoAplica) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("UNMSM");
  if (!sheet) throw new Error("No se encontró la hoja UNMSM");

  const allData = sheet.getDataRange().getValues();
  let colSedeIdx = -1;
  let colObsIdx = -1;
  const sedeBuscada = normalizarTexto(codigoSede);

  // Buscar dinámicamente la fila de encabezados que contiene el código de la sede y las observaciones
  for (let r = 0; r < Math.min(allData.length, 12); r++) {
    const fila = allData[r].map(normalizarTexto);
    const idxSede = fila.findIndex(c => c === sedeBuscada);
    if (idxSede !== -1) {
      colSedeIdx = idxSede + 1; // 1-indexed para getRange
      const idxObs = fila.findIndex(c => c.includes("OBSERV"));
      if (idxObs !== -1) colObsIdx = idxObs + 1;
      break;
    }
  }

  if (colSedeIdx === -1) throw new Error("No se encontró la columna de la sede: " + codigoSede);

  const celda = sheet.getRange(filaSheet, colSedeIdx);

  if (motivoNoAplica && motivoNoAplica.trim() !== "") {
    celda.setValue("NO APLICA");
    if (colObsIdx !== -1) {
      sheet.getRange(filaSheet, colObsIdx).setValue("NO APLICA: " + motivoNoAplica);
    }
  } else {
    celda.setValue(marcarCompletado ? true : false);
  }

  return { exito: true, sede: codigoSede, fila: filaSheet, estado: marcarCompletado };
}

/**
 * Guarda las fechas y horas pactadas en la hoja de checklist oficial buscando encabezados por nombre
 */
function apiGuardarFechasPactadas(actividadesPactadas) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("UNMSM");
  if (!sheet) throw new Error("No se encontró la hoja UNMSM");

  const allData = sheet.getDataRange().getValues();
  let colFechaIdx = -1;
  let colHoraIdx = -1;

  // Localizar dinámicamente las columnas de Fecha Límite y Hora Límite por sus nombres de encabezado
  for (let r = 0; r < Math.min(allData.length, 12); r++) {
    const fila = allData[r].map(normalizarTexto);
    fila.forEach((header, c) => {
      if (header.includes("FECHA") && colFechaIdx === -1) colFechaIdx = c + 1;
      if (header.includes("HORA") && colHoraIdx === -1) colHoraIdx = c + 1;
    });
    if (colFechaIdx !== -1 && colHoraIdx !== -1) break;
  }

  // Fallback si los encabezados estuviesen en formato no textual
  if (colFechaIdx === -1) colFechaIdx = 21;
  if (colHoraIdx === -1) colHoraIdx = 22;

  actividadesPactadas.forEach(item => {
    if (item.fila && item.fecha) {
      sheet.getRange(item.fila, colFechaIdx).setValue(item.fecha);
      if (item.hora && colHoraIdx !== -1) {
        sheet.getRange(item.fila, colHoraIdx).setValue(item.hora);
      }
    }
  });

  return { exito: true, totalActualizadas: actividadesPactadas.length };
}

/**
 * Consulta el estado del Trigger de recordatorios (MANUAL por defecto)
 */
function apiObtenerEstadoTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  const triggerExistente = triggers.find(t => t.getHandlerFunction() === "enviarRecordatoriosDiariosAutomatico");

  return {
    activo: !!triggerExistente,
    modo: triggerExistente ? "AUTOMÁTICO" : "MANUAL (Por Defecto)"
  };
}

/**
 * Activa o desactiva el Trigger diario de recordatorios (Default: MANUAL)
 */
function apiToggleTriggerRecordatorios(activar) {
  const triggers = ScriptApp.getProjectTriggers();
  
  // Eliminar cualquier trigger existente de recordatorios
  triggers.forEach(t => {
    if (t.getHandlerFunction() === "enviarRecordatoriosDiariosAutomatico") {
      ScriptApp.deleteTrigger(t);
    }
  });

  if (activar) {
    // Crear trigger diario a las 8:00 a. m.
    ScriptApp.newTrigger("enviarRecordatoriosDiariosAutomatico")
      .timeBased()
      .everyDays(1)
      .atHour(8)
      .create();
  }

  return {
    activo: activar,
    modo: activar ? "AUTOMÁTICO" : "MANUAL (Por Defecto)"
  };
}

/**
 * Filtra las actividades aplicables a un usuario específico y las reglas de tiempo seleccionadas
 */
function filtrarActividadesParaUsuario(actividades, usuario, reglas) {
  return actividades.filter(act => {
    // 1. Debe corresponder estrictamente al rol del usuario
    if (!correspondeActividadAUsuario(act, usuario)) return false;

    // 2. Si el usuario tiene una sede asignada (Directores de Sede), verificar si está pendiente en su sede
    if (usuario.sede && act.estados && act.estados[usuario.sede] !== undefined) {
      if (act.estados[usuario.sede] === true) return false; // Ya completó la tarea en su sede
    } else {
      // Roles transversales (Sistemas, Evaluación, Producción, etc.): si todas las sedes están realizadas, no se notifica
      const estadosValores = Object.values(act.estados || {});
      const todasCompletadas = estadosValores.length > 0 && estadosValores.every(v => v === true);
      if (todasCompletadas) return false;
    }

    // 3. Reglas de tiempo (hoy, manana, vencidas)
    let cumpleRegla = false;
    if (reglas.hoy && act.esHoy) cumpleRegla = true;
    if (reglas.manana && act.esManana) cumpleRegla = true;
    if (reglas.vencidas && act.esVencida) cumpleRegla = true;

    return cumpleRegla;
  });
}

/**
 * Función que ejecuta el trigger si el usuario decide activarlo
 */
function enviarRecordatoriosDiariosAutomatico() {
  enviarRecordatoriosCorreoConReglas({ hoy: true, manana: true, vencidas: true });
}

/**
 * Genera el cuerpo HTML del correo de Recordatorio (Estética Imagen 3 Liquid Glass)
 */
function generarHtmlRecordatorioEstetico(programa, actividadesConSedes, urlChecklist, nombreDestinatario) {
  let cardsHtml = actividadesConSedes.map(a => {
    const sedesObj = a.sedes || Object.keys(a.estados || {}).map(cod => ({ nombre: cod, completado: a.estados[cod] }));
    const sedesHtml = sedesObj.map(s => {
      const bgColor = s.completado ? "#E3FCEF" : "#FEEBEB";
      const textColor = s.completado ? "#006644" : "#BF2600";
      const icon = s.completado ? "✓" : "○";
      return `<span style="display:inline-block; background:${bgColor}; color:${textColor}; padding:3px 9px; border-radius:6px; font-size:10px; font-weight:800; margin-right:5px; margin-bottom:5px; border:1px solid ${textColor}22;">${icon} ${s.nombre}</span>`;
    }).join("");

    const badgeColor = a.esHoy ? "#EF4444" : (a.esVencida ? "#8B0000" : (a.esManana ? "#F59E0B" : "#2563EB"));
    const badgeTexto = a.esHoy ? "VENCE HOY" : (a.esVencida ? "TIEMPO VENCIDO" : (a.esManana ? "VENCE MAÑANA" : "PRÓXIMO"));
    const fechaHora = a.fechaHoraTexto || `${a.fechaTexto || ''} • ${a.horaTexto || a.hora || '18:00 h'}`;

    return `
      <div style="background:#ffffff; border-radius:18px; padding:20px; margin-bottom:16px; box-shadow:0 8px 24px rgba(0,0,0,0.05); border:1px solid #f1f5f9;">
        <table width="100%" cellpadding="0" cellspacing="0" border="0" style="margin-bottom:8px;">
          <tr>
            <td>
              <span style="background:${badgeColor}; color:#ffffff; padding:3px 10px; border-radius:50px; font-size:10px; font-weight:800; text-transform:uppercase; letter-spacing:0.5px;">${badgeTexto}</span>
              <span style="color:#64748b; font-size:11px; font-weight:700; margin-left:8px;">${a.encargado || 'GESTIÓN'}</span>
            </td>
            <td align="right" style="color:#1e293b; font-size:12px; font-weight:800;">
              ⏰ ${fechaHora}
            </td>
          </tr>
        </table>
        
        <div style="font-size:15px; font-weight:700; color:#0f172a; line-height:1.4; margin-bottom:6px;">
          ${a.titulo}
        </div>
        <div style="font-size:12px; color:#64748b; margin-bottom:14px; line-height:1.4;">
          ${a.desc || ''}
        </div>

        <div style="border-top:1px solid #f1f5f9; padding-top:10px;">
          <div style="color:#94a3b8; font-size:9.5px; font-weight:800; text-transform:uppercase; margin-bottom:6px; letter-spacing:0.5px;">Estado por Sedes:</div>
          <div>${sedesHtml}</div>
        </div>
      </div>
    `;
  }).join("");

  const saludo = nombreDestinatario ? `Buen día <strong>${nombreDestinatario}</strong>, tienes las siguientes actividades asignadas a tu rol próximas a vencer:` : `Tienes las siguientes actividades próximas a vencer:`;

  return `
    <div style="background-color:#F4F7FA; padding:35px 12px; font-family:'Segoe UI', Roboto, Helvetica, Arial, sans-serif;">
      <div style="max-width:580px; margin:auto;">
        
        <!-- HEADER VONEX -->
        <div style="background:linear-gradient(135deg, #004691 0%, #002d62 100%); padding:28px 24px; border-radius:24px; text-align:center; margin-bottom:20px; box-shadow:0 12px 28px rgba(0,70,145,0.22);">
          <div style="color:#FFCC00; font-size:12px; font-weight:800; text-transform:uppercase; letter-spacing:1px; margin-bottom:4px;">
            ${programa} - RECORDATORIO DE GESTIÓN 🔔
          </div>
          <h1 style="color:#FFFFFF; margin:0; font-size:20px; font-weight:800;">
            Checklist Inicio de Clases
          </h1>
          <p style="color:#e0f2fe; margin:8px 0 0 0; font-size:12.5px;">${saludo}</p>
        </div>

        <!-- CARDS DE ACTIVIDADES -->
        <div style="margin-bottom:24px;">
          ${cardsHtml}
        </div>

        <!-- BOTÓN REVISAR CHECKLIST (AMARILLO VONEX) -->
        <div style="text-align:center; background:#ffffff; padding:24px; border-radius:24px; box-shadow:0 8px 20px rgba(0,0,0,0.04);">
          <a href="${urlChecklist}" target="_blank" 
             style="background:#FFCC00; color:#004691; padding:16px 38px; border-radius:50px; text-decoration:none; font-weight:900; font-size:15px; display:inline-block; box-shadow:0 6px 18px rgba(255,204,0,0.4); letter-spacing:0.5px;">
            REVISAR MI CHECKLIST
          </a>
          <p style="color:#94a3b8; font-size:10px; margin-top:16px; margin-bottom:0;">
            &copy; 2026 Academia Vonex | Sistema de Control Operativo
          </p>
        </div>

      </div>
    </div>
  `;
}

/**
 * Envia recordatorios por correo filtrados individualmente por rol según UNMSM y ADMIN_DIRECTORIO
 */
function enviarRecordatoriosCorreoConReglas(reglas) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const urlChecklist = ss.getUrl();
  const actividadesLive = apiObtenerActividadesDesdeSheet("UNMSM");
  const directorio = leerDirectorioDesdeSheet();

  let totalEnviados = 0;
  const fechaHoyTexto = Utilities.formatDate(new Date(), "GMT-5", "dd/MM/yyyy");

  // Si se ingresó un destino específico (modo prueba o envío manual dirigido)
  if (reglas && reglas.destinoEmail && reglas.destinoEmail.trim() !== "") {
    const destinoEmail = reglas.destinoEmail.trim().toLowerCase();
    const usuarioDestino = directorio.find(u => u.email === destinoEmail) || {
      email: destinoEmail,
      nombre: "Estimado(a)",
      rol: "DIRECTOR DE SEDE  TUTORES",
      sede: "LCE"
    };

    const actsParaUser = filtrarActividadesParaUsuario(actividadesLive, usuarioDestino, reglas);
    if (actsParaUser.length > 0) {
      const htmlBody = generarHtmlRecordatorioEstetico("UNMSM 2026-II", actsParaUser, urlChecklist, usuarioDestino.nombre);
      GmailApp.sendEmail(destinoEmail, `🚨 RECORDATORIO DE GESTIÓN: Actividades Próximas a Vencer - UNMSM (${fechaHoyTexto})`, "", {
        htmlBody: htmlBody,
        name: "Gestión de Procesos Vonex"
      });
      totalEnviados++;
    }
    return { exito: true, totalEnviados: totalEnviados, enviadoA: destinoEmail };
  }

  // Envío masivo personalizado por rol a todos los usuarios activos con notifCorreo = SÍ
  directorio.forEach(user => {
    if (!user.email || user.email.includes("inactivo") || !user.notifCorreo) return;

    const actsParaUser = filtrarActividadesParaUsuario(actividadesLive, user, reglas);
    if (actsParaUser.length > 0) {
      const htmlBody = generarHtmlRecordatorioEstetico("UNMSM 2026-II", actsParaUser, urlChecklist, user.nombre);
      GmailApp.sendEmail(user.email, `🚨 RECORDATORIO DE GESTIÓN: Actividades Próximas a Vencer - UNMSM (${fechaHoyTexto})`, "", {
        htmlBody: htmlBody,
        name: "Gestión de Procesos Vonex"
      });
      totalEnviados++;
    }
  });

  return { exito: true, totalEnviados: totalEnviados };
}

/**
 * Envío manual de recordatorios por correo con las reglas seleccionadas
 */
function apiEnviarRecordatoriosCorreoManual(reglas) {
  return enviarRecordatoriosCorreoConReglas(reglas || { hoy: true, manana: false, vencidas: true });
}

/**
 * Genera la vista previa y envía el Reporte Ejecutivo consolidado
 */
function apiGenerarPreviewReporteEjecutivo() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const urlChecklist = ss.getUrl();

  const htmlPreview = `
    <div style="font-family:'Segoe UI', Roboto, sans-serif; padding:20px; background:#f8fafc; border-radius:14px; border:1px solid #e2e8f0;">
      <div style="background:#004691; color:white; padding:18px; border-radius:12px; margin-bottom:16px;">
        <h3 style="margin:0; font-size:16px;">REPORTE EJECUTIVO DE CUMPLIMIENTO • UNMSM</h3>
        <p style="margin:4px 0 0 0; font-size:11px; opacity:0.85;">Inicio de Clases: 26/10/2026 • Estado en tiempo real</p>
      </div>

      <div style="display:grid; grid-template-columns:repeat(3, 1fr); gap:10px; margin-bottom:16px; text-align:center;">
        <div style="background:white; padding:12px; border-radius:10px; border:1px solid #e2e8f0;">
          <div style="font-size:20px; font-weight:800; color:#004691;">28.5%</div>
          <div style="font-size:10px; font-weight:700; color:#64748b; text-transform:uppercase;">Avance Global</div>
        </div>
        <div style="background:white; padding:12px; border-radius:10px; border:1px solid #e2e8f0;">
          <div style="font-size:20px; font-weight:800; color:#10b981;">14</div>
          <div style="font-size:10px; font-weight:700; color:#64748b; text-transform:uppercase;">Realizadas</div>
        </div>
        <div style="background:white; padding:12px; border-radius:10px; border:1px solid #e2e8f0;">
          <div style="font-size:20px; font-weight:800; color:#ef4444;">5</div>
          <div style="font-size:10px; font-weight:700; color:#64748b; text-transform:uppercase;">Vencen Hoy</div>
        </div>
      </div>

      <div style="background:white; padding:14px; border-radius:10px; border:1px solid #e2e8f0; font-size:11.5px; line-height:1.6;">
        <b>Avance por Sedes:</b><br>
        • VMT: 60% (Al día)<br>
        • COM: 35% (2 pendientes hoy)<br>
        • SJL: 25% (5 pendientes hoy)<br>
        • LCE, IND, SAN, PTP: 20% (5 pendientes hoy)<br>
      </div>
      
      <p style="font-size:11px; color:#64748b; margin-top:12px; text-align:center;">
        📎 Se adjunta el checklist en formato PDF para revisión de dirección.
      </p>
    </div>
  `;

  return { html: htmlPreview, urlChecklist: urlChecklist };
}

/**
 * Envío oficial del correo de apertura leyendo destinatarios desde ADMIN_DIRECTORIO
 */
function apiEnviarCorreoAperturaOficial(asuntoPersonalizado) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const urlChecklist = ss.getUrl();
  const asunto = asuntoPersonalizado || "GP | CHECK LIST DE ACTIVIDADES OCTUBRE 2026 - GESTIÓN DE INICIO DE CLASES";
  
  const destinatarios = apiObtenerDestinatariosApertura();
  if (!destinatarios.para || destinatarios.para.length === 0) {
    throw new Error("No hay destinatarios registrados en la hoja ADMIN_DIRECTORIO.");
  }

  const htmlBody = `
    <div style="font-family:'Segoe UI', sans-serif; max-width:640px; margin:auto; background:#ffffff; border:1px solid #e2e8f0; border-radius:18px; padding:30px; box-shadow:0 10px 25px rgba(0,70,145,0.06);">
      <div style="border-bottom:2px solid #004691; padding-bottom:14px; margin-bottom:18px;">
        <span style="background:#004691; color:#FFCC00; font-weight:800; font-size:15px; padding:4px 10px; border-radius:8px;">V</span>
        <span style="color:#004691; font-size:17px; font-weight:800; margin-left:8px;">ACADEMIA VONEX</span>
        <div style="color:#64748b; font-size:11px; font-weight:600; text-transform:uppercase; margin-top:3px;">Gestión de Procesos y Operaciones</div>
      </div>

      <p style="font-size:14.5px; color:#1e293b;">Buen día, estimados(as) 👋</p>
      <p style="font-size:13.5px; color:#334155; line-height:1.6;">
        Por medio del presente, se comparte el enlace del <strong>CHECK LIST DE ACTIVIDADES</strong> del proceso de <strong>GESTIÓN DE INICIO DE CLASES</strong>, a fin de que cada responsable pueda ir registrando el avance correspondiente a las actividades asignadas:
      </p>

      <div style="background:#f8fafc; border:1px solid #e2e8f0; border-radius:14px; padding:16px 20px; margin:20px 0;">
        <p style="margin:0 0 6px 0; font-size:13.5px;">🎓 <strong>Universidad:</strong> UNMSM</p>
        <p style="margin:0 0 10px 0; font-size:13.5px;">📅 <strong>Fecha de inicio de clases:</strong> 26/10/2026</p>
        <p style="margin:0; font-size:13.5px;">
          🔗 <a href="${urlChecklist}" target="_blank" style="color:#004691; font-weight:800; text-decoration:underline;">
            CHECK LIST – GESTIÓN DE INICIO DE CLASES OCTUBRE 2026
          </a>
        </p>
      </div>

      <p style="font-size:13px; color:#475569; line-height:1.6;">
        Este check list será monitoreado de manera continua con el objetivo de asegurar el cumplimiento oportuno de cada una de las actividades dentro de los plazos establecidos.
      </p>
      <p style="font-size:13px; color:#475569; margin-top:14px;">Quedamos atentos ante cualquier consulta 💬</p>

      <div style="margin-top:24px; padding-top:16px; border-top:1px solid #f1f5f9; font-size:12px; color:#64748b; line-height:1.5;">
        Saludos,<br><br>
        --<br>
        <span style="color:#004691; font-weight:900; font-size:16px;">vonex</span><br>
        <strong style="color:#1e293b; font-size:13px;">Gustavo Vallejo Contreras</strong><br>
        <span style="color:#64748b; font-size:12px;">Analista de Gestión de Procesos</span><br>
        <span style="color:#004691; font-weight:700;">+51 924 240 970</span> | <a href="mailto:gvallejo@hnc.pe" style="color:#004691; text-decoration:none;">gvallejo@hnc.pe</a><br>
        <span style="color:#94a3b8; font-size:11px;">Sede Independencia Admin. - Rufino Macedo 204, Independencia</span>
      </div>
    </div>
  `;

  const opciones = {
    htmlBody: htmlBody,
    name: "Gestión de Procesos Vonex"
  };
  if (destinatarios.cc && destinatarios.cc.length > 0) {
    opciones.cc = destinatarios.cc.join(",");
  }

  GmailApp.sendEmail(destinatarios.para.join(","), asunto, "", opciones);

  return { exito: true, total: destinatarios.para.length, ccTotal: (destinatarios.cc || []).length };
}

/**
 * =========================================================================
 * MÓDULO CREAR.HTML: LÓGICA DE DATOS CON HOJAS "PLANTILLA" Y "BD"
 * =========================================================================
 */

/**
 * Obtiene o inicializa la hoja PLANTILLA (Catálogo Maestro de 25 actividades)
 */
function obtenerOInicializarHojaPlantilla() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("PLANTILLA");
  if (sheet) return sheet;

  sheet = ss.insertSheet("PLANTILLA");
  const headers = [
    "ID_TEMPLATE", "ORDEN", "ETAPA", "RESPONSABLE", "ACTIVIDAD", 
    "DESCRIPCION", "DIAS_RELATIVOS", "HORA_DEFAULT", "ACTIVO", "OBSERVACION_TEMPLATE"
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setBackground("#004691")
    .setFontColor("#FFFFFF")
    .setFontWeight("bold");

  const catalogoBase = [
    { id: "TPL-001", orden: 1, etapa: "Antes del inicio de ciclo", resp: "Gestión de Procesos - Vonex", act: "Actualización de carpeta MATRICULA ALUMNOS", desc: "Reglamento interno, políticas de reembolso, términos y croquis", dias: -7, hora: "18:00" },
    { id: "TPL-002", orden: 2, etapa: "Antes del inicio de ciclo", resp: "Director de Sede", act: "Definir tutores", desc: "Envío Formato de asignación de tutores a correos institucionales", dias: -3, hora: "18:00" },
    { id: "TPL-003", orden: 3, etapa: "Antes del inicio de ciclo", resp: "Sistemas", act: "Creación de carpetas en Drive y Grupos Chat", desc: "Creación de carpetas, Espacios de Google Chat y Grupos de WhatsApp", dias: -3, hora: "13:00" },
    { id: "TPL-004", orden: 4, etapa: "Antes del inicio de ciclo", resp: "Ventas", act: "Tutores que apoyan en campaña", desc: "Enviar correo con lista de tutores de apoyo al área de Ventas", dias: -3, hora: "13:00" },
    { id: "TPL-005", orden: 5, etapa: "Antes del inicio de ciclo", resp: "Director de Sede", act: "Actualización de Croquis", desc: "Actualizar el croquis con la numeración de cada aula y salón", dias: -3, hora: "18:00" },
    { id: "TPL-006", orden: 6, etapa: "Antes del inicio de ciclo", resp: "Director de Sede", act: "Colocar Croquis / lista de salones", desc: "Colocar mural del croquis en la puerta de la sede con aulas", dias: -3, hora: "18:00" },
    { id: "TPL-007", orden: 7, etapa: "Antes del inicio de ciclo", resp: "Sistemas", act: "Actualización de la Lista de Tutores en Intranet", desc: "Actualizar Lista de Tutores en base al listado indicado por sede", dias: -3, hora: "18:00" },
    { id: "TPL-008", orden: 8, etapa: "Antes del inicio de ciclo", resp: "Sistemas", act: "Dar acceso a tutores", desc: "Agregar rol de tutor y asesor de ventas para accesos a Drive", dias: -3, hora: "18:00" },
    { id: "TPL-009", orden: 9, etapa: "Antes del inicio de ciclo", resp: "Evaluación y Control", act: "Dar acceso a carpetas y formatos", desc: "Actualizar carpetas de Ev. Desempeño según tutor", dias: -3, hora: "13:00" },
    { id: "TPL-010", orden: 10, etapa: "Antes del inicio de ciclo", resp: "Director de Sede", act: "Confirmar accesos", desc: "Responder al correo de Asignación de tutor indicando verificación", dias: -3, hora: "18:00" },
    { id: "TPL-011", orden: 11, etapa: "Antes del inicio de ciclo", resp: "Producción", act: "Envío del material", desc: "Enviar libros, boletines y guías de clase previo al inicio", dias: -3, hora: "18:00" },
    { id: "TPL-012", orden: 12, etapa: "Antes del inicio de ciclo", resp: "Tutores", act: "Entrega del material al docente", desc: "Enviar al docente las guías de clase por WhatsApp", dias: -3, hora: "18:00" },
    { id: "TPL-013", orden: 13, etapa: "Antes del inicio de ciclo", resp: "Tutores", act: "Enviar información al docente", desc: "Enviar por WhatsApp al docente datos de aula, ciclo y enlace", dias: -3, hora: "18:00" },
    { id: "TPL-014", orden: 14, etapa: "Antes del inicio de ciclo", resp: "Tutores", act: "Agregar alumnos a Grupos de Chat", desc: "Agendar y agregar a padres en Comunidades de WhatsApp", dias: -3, hora: "18:00" },
    { id: "TPL-015", orden: 15, etapa: "Antes del inicio de ciclo", resp: "Tutores", act: "Mensaje de Bienvenida", desc: "Dar mensaje de bienvenida a padres y alumnos en grupos", dias: 0, hora: "09:00" },
    { id: "TPL-016", orden: 16, etapa: "Durante la primera semana", resp: "Tutores", act: "Entrega del material al alumno (Virtual)", desc: "Enviar recordatorio y link de material digitalizado por chat", dias: 0, hora: "09:00" },
    { id: "TPL-017", orden: 17, etapa: "Durante la primera semana", resp: "Tutores", act: "Entrega de material impreso (Presencial)", desc: "Entregar material impreso al alumno en puerta/aula", dias: 0, hora: "09:00" },
    { id: "TPL-018", orden: 18, etapa: "Durante la primera semana", resp: "Director de Sede y Tutores", act: "Orientación de padres y alumnos", desc: "Orientar los 2 primeros días desde las 07:05 am en puerta", dias: 0, hora: "09:00" },
    { id: "TPL-019", orden: 19, etapa: "Durante la primera semana", resp: "Tutores", act: "Agregar alumnos a Grupos (Tarde)", desc: "Actualizar contactos todos los días a las 8:00 am y 4:30 pm", dias: 7, hora: "09:00" },
    { id: "TPL-020", orden: 20, etapa: "Durante la primera semana", resp: "Tutores", act: "Dar charla de sistema de trabajo", desc: "Sistema VONEX, Normas de convivencia y Study Time", dias: 7, hora: "09:00" },
    { id: "TPL-021", orden: 21, etapa: "Durante la primera semana", resp: "Tutores", act: "Mostrar el video tutorial del campus", desc: "Mostrar el video tutorial del campus el día lunes en clase", dias: 7, hora: "09:00" },
    { id: "TPL-022", orden: 22, etapa: "Durante la primera semana", resp: "Tutores", act: "Entrevistas individuales para alumnos", desc: "Llevar a cabo conversaciones con alumnos no asistentes", dias: 7, hora: "09:00" },
    { id: "TPL-023", orden: 23, etapa: "Durante la primera semana", resp: "Tutores", act: "Realizar la evaluación diagnóstica", desc: "Segunda semana: destinar un día a la prueba de entrada", dias: 18, hora: "09:00" },
    { id: "TPL-024", orden: 24, etapa: "Después de inicio de ciclo", resp: "Ventas", act: "Tutores que apoyan en campaña (Cierre)", desc: "Enviar correo para quitar accesos a tutores de apoyo", dias: 14, hora: "12:00" },
    { id: "TPL-025", orden: 25, etapa: "Después de inicio de ciclo", resp: "Tutores", act: "Reunión de padres", desc: "Citar a reunión de padres en 5ta semana", dias: 18, hora: "20:30" }
  ];

  const filas = catalogoBase.map(c => [
    c.id, c.orden, c.etapa, c.resp, c.act, c.desc, c.dias, c.hora, "SI", ""
  ]);

  sheet.getRange(2, 1, filas.length, headers.length).setValues(filas);
  sheet.autoResizeColumns(1, headers.length);
  return sheet;
}

/**
 * Obtiene o inicializa la hoja BD (Historial y Registro de Checklists)
 */
function obtenerOInicializarHojaBD() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("BD");
  if (sheet) return sheet;

  sheet = ss.insertSheet("BD");
  const headers = [
    "ID_CHECKLIST", "UNIVERSIDAD", "FECHA_INICIO_CICLO", "FECHA_CREACION",
    "CREADOR", "ESTADO", "TOTAL_ACTIVIDADES", "ACTIVIDADES_EXCLUIDAS",
    "FECHA_PACTACION", "DETALLE_JSON"
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setBackground("#004691")
    .setFontColor("#FFFFFF")
    .setFontWeight("bold");
  sheet.autoResizeColumns(1, headers.length);
  return sheet;
}

/**
 * Formatea una fecha a YYYY-MM-DD
 */
function formatearFechaISO(d) {
  const anio = d.getFullYear();
  const mes = String(d.getMonth() + 1).padStart(2, "0");
  const dia = String(d.getDate()).padStart(2, "0");
  return `${anio}-${mes}-${dia}`;
}

/**
 * API: Crea un nuevo Checklist a partir de PLANTILLA y lo guarda en BD
 * Estado inicial: "Borrador / Pendiente de Pactar" (NO envía correo)
 */
function apiCrearNuevoChecklist(universidad, fechaInicioCicloStr) {
  if (!universidad || !fechaInicioCicloStr) {
    throw new Error("Debe seleccionar una universidad y fecha de inicio de ciclo válida.");
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetPlantilla = obtenerOInicializarHojaPlantilla();
  const sheetBD = obtenerOInicializarHojaBD();

  // Partir fecha ancla asegurando hora local sin desajuste de zona
  const partes = fechaInicioCicloStr.split("-");
  const fechaAncla = new Date(parseInt(partes[0], 10), parseInt(partes[1], 10) - 1, parseInt(partes[2], 10), 12, 0, 0);

  // Generar ID único: CHK-YYYYMMDD-XX
  const yyyy = partes[0];
  const mm = partes[1];
  const dd = partes[2];
  const prefijoId = `CHK-${yyyy}${mm}${dd}`;

  const dataBD = sheetBD.getDataRange().getValues();
  let correlativo = 1;
  for (let i = 1; i < dataBD.length; i++) {
    const idExistente = (dataBD[i][0] || "").toString();
    if (idExistente.startsWith(prefijoId)) {
      correlativo++;
    }
  }
  const idChecklist = `${prefijoId}-${String(correlativo).padStart(2, "0")}`;

  // Leer actividades maestras de PLANTILLA
  const datosPlantilla = sheetPlantilla.getDataRange().getValues();
  const actividadesCalculadas = [];
  const diasSem = ["dom", "lun", "mar", "mié", "jue", "vie", "sáb"];
  const mesesAbr = ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "set", "oct", "nov", "dic"];

  for (let r = 1; r < datosPlantilla.length; r++) {
    const fila = datosPlantilla[r];
    const activo = (fila[8] || "SI").toString().trim().toUpperCase();
    if (activo !== "SI") continue;

    const offsetDias = parseInt(fila[6] || 0, 10);
    let horaDefault = (fila[7] || "18:00").toString().trim();
    if (fila[7] instanceof Date) {
      horaDefault = Utilities.formatDate(fila[7], Session.getScriptTimeZone() || "GMT-5", "HH:mm");
    } else if (horaDefault.includes("p. m.") || horaDefault.includes("a. m.") || horaDefault.includes("PM") || horaDefault.includes("AM")) {
      const match = horaDefault.match(/(\d+):(\d+)/);
      if (match) {
        let h = parseInt(match[1], 10);
        const m = match[2];
        if ((horaDefault.includes("p. m.") || horaDefault.includes("PM")) && h < 12) h += 12;
        if ((horaDefault.includes("a. m.") || horaDefault.includes("AM")) && h === 12) h = 0;
        horaDefault = `${String(h).padStart(2, "0")}:${m}`;
      }
    } else if (!horaDefault.match(/^\d{2}:\d{2}$/)) {
      horaDefault = "18:00";
    }

    // Calcular fecha con offset
    const fechaObj = new Date(fechaAncla.getTime() + (offsetDias * 24 * 60 * 60 * 1000));
    const fechaCalculadaISO = formatearFechaISO(fechaObj);
    const fechaTexto = `${diasSem[fechaObj.getDay()]}, ${fechaObj.getDate()} de ${mesesAbr[fechaObj.getMonth()]}`;

    actividadesCalculadas.push({
      id: fila[0],
      orden: fila[1],
      etapa: fila[2],
      responsable: fila[3],
      actividad: fila[4],
      descripcion: fila[5],
      offsetDias: offsetDias,
      fechaCalculada: fechaCalculadaISO,
      fechaTexto: fechaTexto,
      horaCalculada: horaDefault,
      noSeRealizara: false,
      observaciones: ""
    });
  }

  const creador = Session.getActiveUser().getEmail() || "director.zonal@vonex.edu.pe";
  const ahora = new Date();
  const ahoraTexto = Utilities.formatDate(ahora, Session.getScriptTimeZone() || "GMT-5", "dd/MM/yyyy HH:mm:ss");

  // Registrar en BD
  sheetBD.appendRow([
    idChecklist,
    universidad,
    fechaInicioCicloStr,
    ahoraTexto,
    creador,
    "Borrador / Pendiente de Pactar",
    actividadesCalculadas.length,
    0,
    "",
    JSON.stringify(actividadesCalculadas)
  ]);

  return {
    idChecklist: idChecklist,
    universidad: universidad,
    fechaInicioCiclo: fechaInicioCicloStr,
    creador: creador,
    fechaCreacion: ahoraTexto,
    estado: "Borrador / Pendiente de Pactar",
    actividades: actividadesCalculadas
  };
}

/**
 * API: Obtiene todos los checklists registrados en BD para el monitoreo
 */
function apiObtenerChecklistsBD() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetBD = obtenerOInicializarHojaBD();
  const allData = sheetBD.getDataRange().getValues();
  if (allData.length < 2) return [];

  const lista = [];
  for (let i = 1; i < allData.length; i++) {
    const row = allData[i];
    if (!row[0]) continue;

    let fInicio = "";
    if (row[2] instanceof Date) {
      fInicio = formatearFechaISO(row[2]);
    } else {
      fInicio = (row[2] || "").toString().trim();
    }

    let fCreacion = "";
    if (row[3] instanceof Date) {
      fCreacion = Utilities.formatDate(row[3], Session.getScriptTimeZone() || "GMT-5", "dd/MM/yyyy HH:mm:ss");
    } else {
      fCreacion = (row[3] || "").toString().trim();
    }

    let fPactacion = "";
    if (row[8] instanceof Date) {
      fPactacion = Utilities.formatDate(row[8], Session.getScriptTimeZone() || "GMT-5", "dd/MM/yyyy HH:mm:ss");
    } else {
      fPactacion = (row[8] || "").toString().trim();
    }

    lista.push({
      idChecklist: (row[0] || "").toString().trim(),
      universidad: (row[1] || "").toString().trim(),
      fechaInicioCiclo: fInicio,
      fechaCreacion: fCreacion,
      creador: (row[4] || "").toString().trim(),
      estado: (row[5] || "Borrador / Pendiente de Pactar").toString().trim(),
      totalActividades: parseInt(row[6] || 25, 10),
      actividadesExcluidas: parseInt(row[7] || 0, 10),
      fechaPactacion: fPactacion,
      detalleJSON: (row[9] || "[]").toString()
    });
  }
  return lista.reverse(); // Mostrar los más recientes primero
}

/**
 * API: Pactar fechas/horas y enviar notificación por correo a ADMIN_DIRECTORIO
 */
function apiPactarChecklist(idChecklist, actividadesModificadas) {
  if (!idChecklist || !Array.isArray(actividadesModificadas)) {
    throw new Error("Datos de pactación incompletos.");
  }

  // 1. Validar regla de excepción obligatoria:
  // Si "noSeRealizara" está activo, observaciones es OBLIGATORIO
  const excluidas = [];
  actividadesModificadas.forEach(act => {
    if (act.noSeRealizara) {
      if (!act.observaciones || act.observaciones.trim() === "") {
        throw new Error(`La actividad #${act.orden} ("${act.actividad}") fue marcada como "NO SE REALIZARÁ" pero no cuenta con la observación/justificación obligatoria.`);
      }
      excluidas.push(act);
    }
  });

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetBD = obtenerOInicializarHojaBD();
  const data = sheetBD.getDataRange().getValues();
  let rowIndex = -1;
  let filaEncontrada = null;

  for (let r = 1; r < data.length; r++) {
    if (data[r][0] === idChecklist) {
      rowIndex = r + 1; // 1-based index para getRange
      filaEncontrada = data[r];
      break;
    }
  }

  if (rowIndex === -1) {
    throw new Error(`Checklist ${idChecklist} no encontrado en la base de datos.`);
  }

  const universidad = filaEncontrada[1];
  const fechaInicioCiclo = filaEncontrada[2] instanceof Date ? formatearFechaISO(filaEncontrada[2]) : filaEncontrada[2];
  const creador = filaEncontrada[4];
  const fechaCreacion = filaEncontrada[3];

  const ahora = new Date();
  const fechaPactacion = Utilities.formatDate(ahora, Session.getScriptTimeZone() || "GMT-5", "dd/MM/yyyy HH:mm:ss");

  // Actualizar columnas en BD: ESTADO (col 6), ACTIVIDADES_EXCLUIDAS (col 8), FECHA_PACTACION (col 9), DETALLE_JSON (col 10)
  sheetBD.getRange(rowIndex, 6).setValue("Pactado / En Ejecución");
  sheetBD.getRange(rowIndex, 8).setValue(excluidas.length);
  sheetBD.getRange(rowIndex, 9).setValue(fechaPactacion);
  sheetBD.getRange(rowIndex, 10).setValue(JSON.stringify(actividadesModificadas));

  // 2. Disparar Correo Oficial de Notificación a ADMIN_DIRECTORIO
  const destinatarios = apiObtenerDestinatariosApertura();
  const urlChecklist = ss.getUrl();
  const asunto = `[NUEVO CHECKLIST PACTADO] - ${universidad} | ${idChecklist}`;

  let bloqueExcluidasHtml = "";
  if (excluidas.length > 0) {
    bloqueExcluidasHtml = `
      <div style="background:#fef2f2; border:1px solid #fecaca; border-radius:12px; padding:14px; margin:16px 0;">
        <b style="color:#b91c1c; font-size:13px;">⚠️ Actividades Excluidas (No se realizarán):</b>
        <ul style="margin:8px 0 0 16px; padding:0; font-size:12px; color:#7f1d1d;">
          ${excluidas.map(e => `<li><strong>#${e.orden} ${e.actividad}:</strong> <em>"${e.observaciones}"</em></li>`).join("")}
        </ul>
      </div>
    `;
  }

  const cuerpoHtml = `
    <div style="font-family:'Segoe UI', Roboto, Helvetica, sans-serif; max-width:680px; margin:0 auto; background:#ffffff; border-radius:16px; border:1px solid #e2e8f0; overflow:hidden; box-shadow:0 8px 30px rgba(0,0,0,0.06); padding:28px;">
      
      <!-- ENCABEZADO INSTITUCIONAL -->
      <div style="display:flex; justify-content:space-between; align-items:center; border-bottom:2px solid #004691; padding-bottom:14px; margin-bottom:20px;">
        <div>
          <span style="font-family:'Arial Black', sans-serif; font-size:22px; color:#004691; letter-spacing:-0.5px;">VONEX</span>
          <div style="font-size:11px; color:#64748b; font-weight:700; text-transform:uppercase;">Gestión de Procesos & Dirección Zonal</div>
        </div>
        <span style="background:#10b981; color:#ffffff; font-weight:800; font-size:11px; padding:5px 12px; border-radius:9999px; text-transform:uppercase; letter-spacing:0.5px;">
          PACTADO / EN EJECUCIÓN
        </span>
      </div>

      <h2 style="font-size:17px; color:#0f172a; margin:0 0 12px 0;">
        Checklist de Inicio de Clases Pactado Oficialmente
      </h2>

      <p style="font-size:13.5px; color:#334155; line-height:1.6; margin-bottom:18px;">
        Se ha completado la validación y pactación de fechas límite para el ciclo académico. A continuación se detallan los parámetros establecidos:
      </p>

      <!-- FICHA TÉCNICA -->
      <div style="background:#f8fafc; border:1px solid #e2e8f0; border-radius:12px; padding:16px 20px; font-size:13px; line-height:1.7; color:#1e293b;">
        <div>🎓 <strong>Universidad / Sede:</strong> ${universidad}</div>
        <div>🆔 <strong>ID Checklist:</strong> ${idChecklist}</div>
        <div>📅 <strong>Fecha de Inicio de Ciclo:</strong> ${fechaInicioCiclo}</div>
        <div>👤 <strong>Creado y Pactado por:</strong> ${creador}</div>
        <div>⏰ <strong>Fecha de Pactación:</strong> ${fechaPactacion}</div>
        <div>📋 <strong>Total Actividades Programadas:</strong> ${actividadesModificadas.length - excluidas.length} de ${actividadesModificadas.length}</div>
      </div>

      ${bloqueExcluidasHtml}

      <!-- BOTON CTA -->
      <div style="text-align:center; margin:28px 0 20px 0;">
        <a href="${urlChecklist}" target="_blank" style="background:#004691; color:#ffffff; text-decoration:none; font-weight:800; font-size:13px; padding:12px 28px; border-radius:10px; display:inline-block; box-shadow:0 4px 14px rgba(0,70,145,0.35);">
          Abrir Hoja de Monitoreo & Ejecución ↗
        </a>
      </div>

      <p style="font-size:12px; color:#64748b; text-align:center; margin-top:20px; border-top:1px solid #f1f5f9; padding-top:14px;">
        Este checklist será auditado periódicamente para verificar el cumplimiento por sede.<br>
        <strong>Gestión de Procesos | Academia Vonex</strong>
      </p>
    </div>
  `;

  if (destinatarios.para && destinatarios.para.length > 0) {
    GmailApp.sendEmail(destinatarios.para.join(","), asunto, "", {
      htmlBody: cuerpoHtml,
      cc: destinatarios.cc.length > 0 ? destinatarios.cc.join(",") : undefined,
      name: "Academia Vonex - Gestión de Procesos"
    });
  }

  return {
    exito: true,
    idChecklist: idChecklist,
    estado: "Pactado / En Ejecución",
    totalActividades: actividadesModificadas.length,
    excluidas: excluidas.length,
    notificados: (destinatarios.para ? destinatarios.para.length : 0) + (destinatarios.cc ? destinatarios.cc.length : 0)
  };
}

/**
 * =========================================================================
 * ENDPOINTS PARA PESTAÑA PROCESOS (WHATSAPP, CORREO COLECTIVO Y REPORTE)
 * =========================================================================
 */

/**
 * Obtiene la lista de Directores de Sede con sus actividades pendientes que vencen HOY
 * y genera el texto formateado de WhatsApp listo para copiar o abrir en WhatsApp Web
 */
function apiObtenerConsolaWhatsAppProcesos(nombreHojaChecklist) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaNombre = nombreHojaChecklist || "UNMSM";
  const sheetChecklist = ss.getSheetByName(hojaNombre);
  const sheetId = sheetChecklist ? sheetChecklist.getSheetId() : 0;
  const urlChecklist = ss.getUrl() + "#gid=" + sheetId;

  const directorio = leerDirectorioDesdeSheet();
  const directoresSede = directorio.filter(u => u.rol && u.rol.includes("DIRECTOR") && !u.rol.includes("ZONAL"));
  const actividadesLive = apiObtenerActividadesDesdeSheet(hojaNombre);

  // Filtrar actividades que vencen hoy o están vencidas
  const actsHoy = actividadesLive.filter(a => a.esHoy || a.esVencida);

  const emojisNum = ["1️⃣", "2️⃣", "3️⃣", "4️⃣", "5️⃣", "6️⃣", "7️⃣", "8️⃣", "9️⃣", "🔟"];

  const consola = directoresSede.map(d => {
    // Buscar tareas pendientes para la sede de este director
    const tareasPendientes = [];
    actsHoy.forEach(a => {
      const estadoSede = a.estados[d.sede];
      if (estadoSede !== true) { // No realizado
        tareasPendientes.push(a);
      }
    });

    let lineasTareas = "";
    if (tareasPendientes.length === 0) {
      lineasTareas = "✅ ¡Felicitaciones! Su sede se encuentra al día con todas las actividades programadas para hoy.\n";
    } else {
      tareasPendientes.forEach((t, idx) => {
        const num = emojisNum[idx] || `*${idx + 1}.*`;
        let textoItem = (t.titulo || t.actividad || "").replace(/^-\s*/, "").trim();

        if (textoItem.toLowerCase().includes("actualizar lista de salones")) {
          lineasTareas += `${num} *Actualizar lista de salones*\n   └ Incluir Nombre del tutor / Ciclo / Aula al costado del croquis\n`;
        } else if (textoItem.toLowerCase().includes("colocar un mural del croquis")) {
          lineasTareas += `${num} *Colocar un mural del croquis en la puerta de la sede*\n   └ Con los números de aulas correspondientes de cada ciclo\n`;
        } else {
          lineasTareas += `${num} *${textoItem}*\n`;
        }
      });
    }

    const mensajeWsp = 
`🚨 *RECORDATORIO | CHECKLIST INICIO DE CLASES* 🚨

Buen día ${d.nombre || 'Director(a)'}, le recordamos que *⏰HOY a las 6:00 p. m.* vence el plazo para el cumplimiento de las siguientes actividades:

📌 *Actividades por cumplir hoy:*
${lineasTareas}
🔗 *Marque su cumplimiento directamente aquí:*
👉 ${urlChecklist}

*Gestión de Procesos | Academia Vonex*`;

    let celularLimpio = (d.celular || "").toString().replace(/[^0-9]/g, "");
    if (celularLimpio && !celularLimpio.startsWith("51") && celularLimpio.length === 9) {
      celularLimpio = "51" + celularLimpio;
    }

    const urlWsp = celularLimpio ? `https://api.whatsapp.com/send?phone=${celularLimpio}&text=${encodeURIComponent(mensajeWsp)}` : "";

    return {
      sede: d.sede,
      nombreSede: d.nombreSede || d.sede,
      director: d.nombre,
      celular: celularLimpio,
      email: d.email,
      tareasPendientesCount: tareasPendientes.length,
      mensaje: mensajeWsp,
      urlWhatsApp: urlWsp
    };
  });

  return consola;
}

/**
 * Genera el Reporte Ejecutivo de Cierre consolidado y lo envía a Dirección
 */
function apiEnviarReporteEjecutivoCierre(destinatarioPersonalizado) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const urlChecklist = ss.getUrl();
  const actividadesLive = apiObtenerActividadesDesdeSheet("UNMSM");
  const directorio = leerDirectorioDesdeSheet();

  const total = actividadesLive.length;
  let totalCompletadasGlobal = 0;
  let totalPendientesHoy = 0;
  let totalVencidas = 0;

  const sedes = ["LCE", "SJL", "IND", "COM", "SAN", "PTP", "VMT"];
  const sedesResumen = sedes.map(s => {
    let cumplidas = 0;
    let pendHoy = 0;
    actividadesLive.forEach(a => {
      if (a.estados[s] === true) cumplidas++;
      if ((a.esHoy || a.esVencida) && a.estados[s] !== true) pendHoy++;
    });
    const pct = total > 0 ? Math.round((cumplidas / total) * 100) : 0;
    return { sede: s, completadas: cumplidas, pendientesHoy: pendHoy, porcentaje: pct };
  });

  sedesResumen.sort((a, b) => b.porcentaje - a.porcentaje);

  actividadesLive.forEach(a => {
    if (a.esHoy) totalPendientesHoy++;
    if (a.esVencida) totalVencidas++;
    let todasSedesOk = sedes.every(s => a.estados[s] === true);
    if (todasSedesOk) totalCompletadasGlobal++;
  });

  const pctGlobal = total > 0 ? Math.round((totalCompletadasGlobal / total) * 100) : 0;
  const hoyTexto = Utilities.formatDate(new Date(), "GMT-5", "dd/MM/yyyy HH:mm");

  const asunto = `📊 [REPORTE EJECUTIVO] Cierre de Checklist Inicio de Clases • UNMSM (${Utilities.formatDate(new Date(), "GMT-5", "dd/MM/yyyy")})`;

  const htmlBody = `
    <div style="font-family:'Segoe UI', Roboto, Helvetica, Arial, sans-serif; max-width:680px; margin:auto; background:#ffffff; border-radius:18px; border:1px solid #e2e8f0; overflow:hidden; box-shadow:0 10px 30px rgba(0,0,0,0.06);">
      
      <!-- HEADER EJECUTIVO -->
      <div style="background:linear-gradient(135deg, #004691 0%, #002d62 100%); padding:28px 24px; text-align:center; color:white;">
        <span style="font-family:'Arial Black', sans-serif; font-size:24px; color:#FFE001; letter-spacing:-0.5px;">VONEX</span>
        <h2 style="margin:8px 0 4px 0; font-size:18px; font-weight:800; color:#ffffff;">
          REPORTE EJECUTIVO DE CUMPLIMIENTO • UNMSM
        </h2>
        <p style="margin:0; font-size:11.5px; color:#e0f2fe;">
          Estado de control de hitos al ${hoyTexto} • Dirección Zonal & Gestión de Procesos
        </p>
      </div>

      <!-- 3 METRICAS HERO -->
      <div style="display:grid; grid-template-columns:repeat(3, 1fr); gap:12px; padding:24px 20px; background:#f8fafc; border-bottom:1px solid #e2e8f0; text-align:center;">
        <div style="background:#ffffff; padding:16px 12px; border-radius:14px; border:1px solid #e2e8f0; box-shadow:0 2px 8px rgba(0,0,0,0.04);">
          <div style="font-size:24px; font-weight:900; color:#004691;">${pctGlobal}%</div>
          <div style="font-size:10px; font-weight:800; color:#64748b; text-transform:uppercase; margin-top:2px;">Avance Consolidado</div>
        </div>
        <div style="background:#ffffff; padding:16px 12px; border-radius:14px; border:1px solid #e2e8f0; box-shadow:0 2px 8px rgba(0,0,0,0.04);">
          <div style="font-size:24px; font-weight:900; color:#10b981;">${totalCompletadasGlobal}</div>
          <div style="font-size:10px; font-weight:800; color:#64748b; text-transform:uppercase; margin-top:2px;">Hitos al 100%</div>
        </div>
        <div style="background:#ffffff; padding:16px 12px; border-radius:14px; border:1px solid #e2e8f0; box-shadow:0 2px 8px rgba(0,0,0,0.04);">
          <div style="font-size:24px; font-weight:900; color:#ef4444;">${totalPendientesHoy}</div>
          <div style="font-size:10px; font-weight:800; color:#64748b; text-transform:uppercase; margin-top:2px;">Vencen Hoy / Riesgo</div>
        </div>
      </div>

      <!-- TABLA COMPARATIVA POR SEDES -->
      <div style="padding:24px 20px;">
        <h3 style="margin:0 0 14px 0; font-size:13px; font-weight:800; color:#0f172a; text-transform:uppercase; letter-spacing:0.5px;">
          Ranking de Cumplimiento por Sede:
        </h3>
        <table width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse; font-size:12px;">
          <thead>
            <tr style="background:#f1f5f9; text-align:left; color:#64748b; font-size:10px; font-weight:800; text-transform:uppercase;">
              <th style="padding:10px 12px; border-radius:8px 0 0 8px;">Sede</th>
              <th style="padding:10px 12px; text-align:center;">Avance</th>
              <th style="padding:10px 12px; text-align:center;">Pendientes Hoy</th>
              <th style="padding:10px 12px; text-align:right; border-radius:0 8px 8px 0;">Estado</th>
            </tr>
          </thead>
          <tbody>
            ${sedesResumen.map(s => {
              const colorBarra = s.porcentaje >= 75 ? "#10b981" : (s.porcentaje >= 40 ? "#3b82f6" : "#f59e0b");
              return `
                <tr style="border-bottom:1px solid #f1f5f9;">
                  <td style="padding:12px; font-weight:800; color:#0f172a;">${s.sede}</td>
                  <td style="padding:12px; text-align:center;">
                    <div style="display:inline-flex; align-items:center; gap:8px;">
                      <div style="width:80px; height:7px; background:#e2e8f0; border-radius:9999px; overflow:hidden;">
                        <div style="width:${s.porcentaje}%; height:100%; background:${colorBarra};"></div>
                      </div>
                      <span style="font-weight:800; font-size:11px;">${s.porcentaje}%</span>
                    </div>
                  </td>
                  <td style="padding:12px; text-align:center; font-weight:700; color:${s.pendientesHoy > 0 ? '#ef4444' : '#10b981'};">
                    ${s.pendientesHoy > 0 ? s.pendientesHoy + ' pendientes' : 'Al día ✓'}
                  </td>
                  <td style="padding:12px; text-align:right;">
                    <span style="padding:3px 8px; border-radius:9999px; font-size:9.5px; font-weight:800; text-transform:uppercase; background:${s.porcentaje >= 75 ? '#dcfce7; color:#15803d;' : (s.porcentaje >= 40 ? '#dbeafe; color:#1d4ed8;' : '#fef3c7; color:#b45309;')};">
                      ${s.porcentaje >= 75 ? 'Excelente' : (s.porcentaje >= 40 ? 'En Proceso' : 'Atención')}
                    </span>
                  </td>
                </tr>
              `;
            }).join("")}
          </tbody>
        </table>

        <!-- BOTON ACCESO SHEET -->
        <div style="text-align:center; margin-top:28px;">
          <a href="${urlChecklist}" target="_blank" style="background:#004691; color:#ffffff; padding:12px 28px; border-radius:10px; font-weight:800; font-size:12.5px; text-decoration:none; display:inline-block; box-shadow:0 4px 14px rgba(0,70,145,0.25);">
            Abrir Checklist UNMSM en Tiempo Real ↗
          </a>
        </div>
      </div>

      <!-- FOOTER -->
      <div style="background:#f8fafc; border-top:1px solid #e2e8f0; padding:16px 20px; text-align:center; font-size:11px; color:#64748b;">
        Este informe fue emitido automáticamente desde la consola de Gestión de Procesos.<br>
        <strong>Academia Vonex | Dirección de Operaciones & Calidad</strong>
      </div>

    </div>
  `;

  // Destinatarios: Director Zonal y Procesos por defecto
  let destinos = ["marevalo@vonex.edu.pe", "gvallejo@hnc.pe"];
  if (destinatarioPersonalizado && destinatarioPersonalizado.includes("@")) {
    destinos = [destinatarioPersonalizado.trim()];
  } else {
    // Agregar correos con rol Zonal o Procesos de ADMIN_DIRECTORIO
    directorio.forEach(u => {
      if (u.email && (u.rol.includes("ZONAL") || u.rol.includes("PROCESOS"))) {
        if (!destinos.includes(u.email)) destinos.push(u.email);
      }
    });
  }

  GmailApp.sendEmail(destinos.join(","), asunto, "", {
    htmlBody: htmlBody,
    name: "Gestión de Procesos Vonex"
  });

  return {
    exito: true,
    destinatarios: destinos.join(", "),
    totalDestinatarios: destinos.length,
    avanceGlobal: pctGlobal + "%",
    htmlPreview: htmlBody
  };
}

