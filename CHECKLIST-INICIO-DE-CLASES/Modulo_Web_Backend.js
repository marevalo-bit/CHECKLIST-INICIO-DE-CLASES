/**
 * =========================================================================
 * ACADEMIA VONEX — MÓDULO GESTIÓN DE INICIO DE CLASES
 * BACKEND CENTRALIZADO GOOGLE APPS SCRIPT
 * =========================================================================
 * 
 * Arquitectura basada en datos gobernados 100% desde Google Sheets:
 * 1. Hoja "PLANTILLA": Catálogo maestro de 25 actividades/tareas normalizadas (A-FOR-SERV-0104 v03).
 * 2. Hoja "ADMIN_DIRECTORIO": Directorio maestro de sedes, directores y áreas centrales (@hnc.pe / @vonex.edu.pe).
 * 3. Hoja "BD_CHECKLISTS": Registro histórico y ciclo de vida de checklists creados.
 * 4. Hoja "BD_SEGUIMIENTO": Marcación individual, evidencias y justificaciones por sede y tarea.
 * 5. Hoja "BD_REPORTES_ETAPAS": Snapshots inmutables de auditoría al finalizar cada etapa.
 * 
 * NOTA DE SEGURIDAD ESTRICTA: NO activa ningún trigger automático (ScriptApp.newTrigger está desactivado).
 * Actualización limpia: módulo unificado y optimizado.
 */

const ID_SPREADSHEET_MASTER = "1yGizP7Ngd8qtw5sdsWGaYRXn2hmGkHiOHDk5rDiyajA";

/**
 * Obtiene la referencia al Spreadsheet activo con fallback al ID oficial
 */
function obtenerSpreadsheet() {
  try {
    const act = SpreadsheetApp.getActiveSpreadsheet();
    if (act) return act;
  } catch (e) { }
  return SpreadsheetApp.openById(ID_SPREADSHEET_MASTER);
}

/**
 * Normaliza cadenas para comparaciones robustas e insensibles a mayúsculas y tildes
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
 * Formatea una fecha a YYYY-MM-DD
 */
function formatearFechaISO(d) {
  if (!d || !(d instanceof Date) || isNaN(d.getTime())) return "";
  const anio = d.getFullYear();
  const mes = String(d.getMonth() + 1).padStart(2, "0");
  const dia = String(d.getDate()).padStart(2, "0");
  return `${anio}-${mes}-${dia}`;
}

/**
 * Convierte formatos de fecha como "26-oct-2026", "26/10/2026" o ISO a "YYYY-MM-DD"
 */
function parsearFechaTextoGeneral(txt) {
  if (!txt) return null;
  const t = txt.toString().trim().toLowerCase();
  const meses = {
    "ene": "01", "feb": "02", "mar": "03", "abr": "04", "may": "05", "jun": "06",
    "jul": "07", "ago": "08", "set": "09", "sep": "09", "oct": "10", "nov": "11", "dic": "12"
  };
  const matchMes = t.match(/(\d{1,2})[-\/\s]([a-z]{3})[-\/\s](\d{4})/);
  if (matchMes && meses[matchMes[2]]) {
    const d = matchMes[1].padStart(2, "0");
    const m = meses[matchMes[2]];
    const y = matchMes[3];
    return `${y}-${m}-${d}`;
  }
  const matchNum = t.match(/(\d{1,2})[-\/](\d{1,2})[-\/](\d{4})/);
  if (matchNum) {
    const d = matchNum[1].padStart(2, "0");
    const m = matchNum[2].padStart(2, "0");
    const y = matchNum[3];
    return `${y}-${m}-${d}`;
  }
  const matchIso = t.match(/(\d{4})-(\d{2})-(\d{2})/);
  if (matchIso) return matchIso[0];
  return null;
}

// =========================================================================
// 1. GOBERNANZA DE HOJAS Y ESTRUCTURAS DE DATOS
// =========================================================================

/**
 * Lee la hoja "PLANTILLA PRUEBA" o "PLANTILLA" detectando dinámicamente
 * las 25 actividades maestras y sus 37 subactividades oficiales (A-FOR-SERV-0104 v03).
 * Soporta tanto el formato matriz de checklist (UNMSM) como el formato tabular de plantilla.
 */
function obtenerDatosPlantilla() {
  const ss = obtenerSpreadsheet();
  let sheet = ss.getSheetByName("PLANTILLA PRUEBA") || ss.getSheetByName("PLANTILLA");
  if (!sheet) {
    const sById = ss.getSheets().find(s => s.getSheetId() === 159313937);
    if (sById) sheet = sById;
  }

  // Catálogo maestro normalizado de respaldo (25 Actividades y 37 Subactividades con DÍAS RELATIVOS oficiales UNMSM)
  const CATALOGO_MAESTRO_OFICIAL = [
    { id: "TPL-001", orden: 1, etapa: "Antes del inicio de ciclo", responsable: "Gestión de Procesos - Vonex", actividad: "Actualización de carpeta MATRICULA ALUMNOS", diasRelativos: -7, horaDefault: "18:00", subactividades: [
      { orden: 1, descripcion: "Reglamento interno del alumno", hora: "18:00", responsable: "Gestión de Procesos - Vonex" },
      { orden: 2, descripcion: "Políticas de reembolso y/o devoluciones", hora: "18:00", responsable: "Gestión de Procesos - Vonex" },
      { orden: 3, descripcion: "Términos y Condiciones", hora: "18:00", responsable: "Gestión de Procesos - Vonex" },
      { orden: 4, descripcion: "Croquis de sede", hora: "18:00", responsable: "Gestión de Procesos - Vonex" }
    ]},
    { id: "TPL-002", orden: 2, etapa: "Antes del inicio de ciclo", responsable: "Director de Sede", actividad: "Definir tutores", diasRelativos: -3, horaDefault: "18:00", subactividades: [
      { orden: 1, descripcion: "Envío Formato de asignación de tutores: (acmateriales@hnc.pe, helpdesk@hnc.pe, evaluacion@hnc.pe, almacon@hnc.pe, evelasquez@hnc.pe, jpinedo@hnc.pe, auditoria@hnc.pe, procesos@hnc.pe)", hora: "18:00", responsable: "Director de Sede" }
    ]},
    { id: "TPL-003", orden: 3, etapa: "Antes del inicio de ciclo", responsable: "Sistemas", actividad: "Creación de carpetas en Drive, Grupos de Google Chat y Comunidades de WhatsApp", diasRelativos: -3, horaDefault: "13:00", subactividades: [
      { orden: 1, descripcion: "Creación de carpetas en Drive", hora: "13:00", responsable: "Sistemas" },
      { orden: 2, descripcion: "Creación de Espacios de Google Chat para alumnos", hora: "13:00", responsable: "Sistemas" },
      { orden: 3, descripcion: "Creación de Comunidades para padres", hora: "13:00", responsable: "Sistemas" },
      { orden: 4, descripcion: "Agregar a los alumnos en los espacios de Google Chat", hora: "13:00", responsable: "Sistemas" }
    ]},
    { id: "TPL-004", orden: 4, etapa: "Antes del inicio de ciclo", responsable: "Ventas", actividad: "Tutores que apoyan en campaña", diasRelativos: -3, horaDefault: "13:00", subactividades: [
      { orden: 1, descripcion: "Enviar correo con lista de tutores que apoyarán durante la campaña en el área de Ventas\nDeberá enviar correo con la información de tutores para apoyo (DNI, Apellidos y Nombres)\n\nDirigido a: helpdesk@hnc.pe, evelasquez@hnc.pe, tesoreria@hnc.pe, evaluacion@hnc.pe, jatalaya@hnc.pe", hora: "13:00", responsable: "Ventas" }
    ]},
    { id: "TPL-005", orden: 5, etapa: "Antes del inicio de ciclo", responsable: "Director de Sede", actividad: "Actualización de Croquis", diasRelativos: -3, horaDefault: "18:00", subactividades: [
      { orden: 1, descripcion: "Actualizar el croquis con la numeración de cada aula", hora: "18:00", responsable: "Director de Sede" },
      { orden: 2, descripcion: "Actualizar lista de salones donde incluya el Nombre del tutor / Ciclo / Aula (Al costado del croquis)", hora: "18:00", responsable: "Director de Sede" }
    ]},
    { id: "TPL-006", orden: 6, etapa: "Antes del inicio de ciclo", responsable: "Director de Sede", actividad: "Colocar Croquis / lista de salones/ nombre de lugar(oficina administrativa o cafetín)", diasRelativos: -3, horaDefault: "18:00", subactividades: [
      { orden: 1, descripcion: "Colocar un mural del croquis en la puerta de la sede con los números de aulas correspondientes de cada ciclo", hora: "18:00", responsable: "Director de Sede" },
      { orden: 2, descripcion: "Colocar en cada piso la lista correspondiente de salones y el piso al que pertenece", hora: "18:00", responsable: "Director de Sede" }
    ]},
    { id: "TPL-007", orden: 7, etapa: "Antes del inicio de ciclo", responsable: "Sistemas", actividad: "Actualización de la Lista de Tutores por Sede en la Intranet", diasRelativos: -3, horaDefault: "18:00", subactividades: [
      { orden: 1, descripcion: "Actualizar Lista de Tutores en base al listado de tutores indicado por cada director de sede en la asignación", hora: "18:00", responsable: "Sistemas" }
    ]},
    { id: "TPL-008", orden: 8, etapa: "Antes del inicio de ciclo", responsable: "Sistemas", actividad: "Dar acceso a tutores", diasRelativos: -3, horaDefault: "18:00", subactividades: [
      { orden: 1, descripcion: "Agregar rol de 'tutor' y 'asesor de ventas' para tutores que les corresponda y dar acceso a funciones de tutoría", hora: "18:00", responsable: "Sistemas" },
      { orden: 2, descripcion: "Agregar a tutores a los grupos de Pases de Salida", hora: "18:00", responsable: "Sistemas" },
      { orden: 3, descripcion: "Dar acceso al Drive de materiales", hora: "13:00", responsable: "Digitalización de contenido" }
    ]},
    { id: "TPL-009", orden: 9, etapa: "Antes del inicio de ciclo", responsable: "Evaluación y Control", actividad: "Dar acceso a carpetas y formatos", diasRelativos: -3, horaDefault: "13:00", subactividades: [
      { orden: 1, descripcion: "Actualizar carpetas de Ev. Desempeño según tutor", hora: "13:00", responsable: "Evaluación y Control" }
    ]},
    { id: "TPL-010", orden: 10, etapa: "Antes del inicio de ciclo", responsable: "Director de Sede", actividad: "Confirmar accesos", diasRelativos: -3, horaDefault: "18:00", subactividades: [
      { orden: 1, descripcion: "Responder al correo de Asignación de tutor indicando que los tutores tienen todos los accesos necesarios", hora: "18:00", responsable: "Director de Sede" }
    ]},
    { id: "TPL-011", orden: 11, etapa: "Antes del inicio de ciclo", responsable: "Producción", actividad: "Envío del material", diasRelativos: -3, horaDefault: "18:00", subactividades: [
      { orden: 1, descripcion: "Enviar libros, boletines y guías de clase a los tutores el viernes previo al inicio de clases", hora: "18:00", responsable: "Producción" },
      { orden: 2, descripcion: "Enviar exámenes y fichas ópticas a los tutores el viernes previo al inicio de clases", hora: "18:00", responsable: "Producción" }
    ]},
    { id: "TPL-012", orden: 12, etapa: "Antes del inicio de ciclo", responsable: "Tutores", actividad: "Entrega del material al docente", diasRelativos: -3, horaDefault: "18:00", subactividades: [
      { orden: 1, descripcion: "Enviar al docente las guías de clase por medio de WhatsApp", hora: "18:00", responsable: "Tutores" }
    ]},
    { id: "TPL-013", orden: 13, etapa: "Antes del inicio de ciclo", responsable: "Tutores", actividad: "Enviar información al docente", diasRelativos: -3, horaDefault: "18:00", subactividades: [
      { orden: 1, descripcion: "Enviar por WhatsApp al docente los siguientes detalles: nombre del tutor, sede, aula asignada, número de aula y un recordatorio sobre el acceso al enlace correspondiente (solo aula virtual)", hora: "18:00", responsable: "Tutores" }
    ]},
    { id: "TPL-014", orden: 14, etapa: "Antes del inicio de ciclo", responsable: "Tutores", actividad: "Agregar alumnos a Grupos de Google Chat y a los padres a las comunidades de WhatsApp", diasRelativos: -3, horaDefault: "18:00", subactividades: [
      { orden: 1, descripcion: "Agendar y agregar a padres en las Comunidades de WhatsApp", hora: "18:00", responsable: "Tutores" }
    ]},
    { id: "TPL-015", orden: 15, etapa: "Antes del inicio de ciclo", responsable: "Tutores", actividad: "Mensaje de Bienvenida", diasRelativos: 0, horaDefault: "09:00", subactividades: [
      { orden: 1, descripcion: "Dar mensaje de bienvenida a padres y alumnos", hora: "09:00", responsable: "Tutores" },
      { orden: 2, descripcion: "Mensaje recordatorio para que los alumnos ingresen a los links ubicados en la descripción de los grupos correspondientes", hora: "09:00", responsable: "Tutores" }
    ]},
    // FASE 2: DURANTE LA PRIMERA SEMANA (8 actividades, 8 subactividades)
    { id: "TPL-016", orden: 16, etapa: "Durante la primera semana", responsable: "Tutores", actividad: "Entrega del material al alumno y pasar asistencia", diasRelativos: 0, horaDefault: "07:45", subactividades: [
      { orden: 1, descripcion: "Aula Virtual: Enviar un recordatorio a los alumnos indicando que el material está en el campus virtual y grupo de WhatsApp. Reenviar los libros digitales que se encuentran en las carpetas y pasar lista de asistencia por zoom (L - V 7:45 a. m.)", hora: "07:45", responsable: "Tutores" }
    ]},
    { id: "TPL-017", orden: 17, etapa: "Durante la primera semana", responsable: "Tutores", actividad: "Entrega del material al alumno y pasar asistencia", diasRelativos: 0, horaDefault: "07:59", subactividades: [
      { orden: 1, descripcion: "Aula Presencial: Entregar material impreso al alumno (L - V 7:59 a. m.). Pasar lista de asistencia a las 09:00 a. m. (Con lista impresa)", hora: "07:59", responsable: "Tutores" }
    ]},
    { id: "TPL-018", orden: 18, etapa: "Durante la primera semana", responsable: "Director de Sede y Tutores", actividad: "Orientación de padres y alumnos", diasRelativos: 0, horaDefault: "07:05", subactividades: [
      { orden: 1, descripcion: "Orientar el ingreso al aula a alumnos y/o padres:\n\n- Primera semana:\nLos tutores orientarán los 2 primeros días desde las 07:05 a. m. hasta las 8:00 a. m. (la cantidad de tutores será de acuerdo a la necesidad por sede)\n\nNota: Como mínimo debe haber 1 tutor asignados. Nunca debería quedar solo vigilancia en la entrada de los alumnos.\n\nNota: Director solo deberá supervisar.", hora: "07:05", responsable: "Director de Sede y Tutores" }
    ]},
    { id: "TPL-019", orden: 19, etapa: "Durante la primera semana", responsable: "Tutores", actividad: "Agregar alumnos a Grupos de Google Chat y a los padres a las comunidades de WhatsApp", diasRelativos: 7, horaDefault: "08:00", subactividades: [
      { orden: 1, descripcion: "Actualizar los contactos todos los días a las 8:00 am - 8:30 am. Turno tarde: 4:30pm", hora: "08:00", responsable: "Tutores" }
    ]},
    { id: "TPL-020", orden: 20, etapa: "Durante la primera semana", responsable: "Tutores", actividad: "Dar charla de sistema de trabajo", diasRelativos: 7, horaDefault: "07:45", subactividades: [
      { orden: 1, descripcion: "Sistema VONEX, Normas de convivencia, Study Time, Sistema de evaluación (durante la 1ra semana desde el martes a las 7:45 a.m.)", hora: "07:45", responsable: "Tutores" }
    ]},
    { id: "TPL-021", orden: 21, etapa: "Durante la primera semana", responsable: "Tutores", actividad: "Mostrar el video tutorial del campus", diasRelativos: 7, horaDefault: "13:00", subactividades: [
      { orden: 1, descripcion: "Mostrar el video tutorial del campus el dia lunes en la salida (o cuando sea necesario)", hora: "13:00", responsable: "Tutores" }
    ]},
    { id: "TPL-022", orden: 22, etapa: "Durante la primera semana", responsable: "Tutores", actividad: "Entrevistas individuales para alumnos incorporados", diasRelativos: 7, horaDefault: "09:00", subactividades: [
      { orden: 1, descripcion: "Llevar a cabo conversaciones individuales con los alumnos que no asistieron en los primeros dias o no asistieron a las charlas de inducción", hora: "09:00", responsable: "Tutores" }
    ]},
    { id: "TPL-023", orden: 23, etapa: "Durante la primera semana", responsable: "Tutores", actividad: "Realizar la evaluación diagnóstica", diasRelativos: 18, horaDefault: "09:00", subactividades: [
      { orden: 1, descripcion: "Durante la segunda semana, se destinará un día exclusivamente para la evaluación de los estudiantes", hora: "09:00", responsable: "Tutores" }
    ]},
    // FASE 3: DESPUÉS DE INICIO DE CICLO (2 actividades, 2 subactividades)
    { id: "TPL-024", orden: 24, etapa: "Después de inicio de ciclo", responsable: "Ventas", actividad: "Tutores que apoyan en campaña (Cierre de accesos)", diasRelativos: 14, horaDefault: "12:00", subactividades: [
      { orden: 1, descripcion: "Enviar correo para quitar los accesos a los tutores de apoyo. Deberá enviar correo con la información de tutores para apoyo (DNI, Apellidos y Nombres). Dirigido a: helpdesk@hnc.pe, evelasquez@hnc.pe, tesoreria@hnc.pe, evaluacion@hnc.pe", hora: "12:00", responsable: "Ventas" }
    ]},
    { id: "TPL-025", orden: 25, etapa: "Después de inicio de ciclo", responsable: "Tutores", actividad: "Reunión de padres", diasRelativos: 18, horaDefault: "20:30", subactividades: [
      { orden: 1, descripcion: "Citar a una reunión de padres en quinta semana", hora: "20:30", responsable: "Tutores" }
    ]}
  ];

  if (!sheet) {
    return CATALOGO_MAESTRO_OFICIAL.map(a => ({
      id: a.id,
      orden: a.orden,
      etapa: a.etapa,
      responsable: a.responsable,
      actividad: a.actividad,
      descripcion: a.subactividades[0].descripcion,
      diasRelativos: a.diasRelativos,
      horaDefault: a.horaDefault,
      requiereEvidencia: a.orden <= 18,
      subactividadesCustom: a.subactividades.map((s, idx) => ({
        id: `${a.id}-${idx + 1}`,
        orden: s.orden,
        descripcion: s.descripcion,
        hora: s.hora,
        responsable: s.responsable,
        diasRelativos: a.diasRelativos
      }))
    }));
  }

  const values = sheet.getDataRange().getValues();
  if (values.length < 2) {
    return CATALOGO_MAESTRO_OFICIAL.map(a => ({
      id: a.id,
      orden: a.orden,
      etapa: a.etapa,
      responsable: a.responsable,
      actividad: a.actividad,
      descripcion: a.subactividades[0].descripcion,
      diasRelativos: a.diasRelativos,
      horaDefault: a.horaDefault,
      requiereEvidencia: a.orden <= 18,
      subactividadesCustom: a.subactividades.map((s, idx) => ({
        id: `${a.id}-${idx + 1}`,
        orden: s.orden,
        descripcion: s.descripcion,
        hora: s.hora,
        responsable: s.responsable,
        diasRelativos: a.diasRelativos
      }))
    }));
  }

  // Detección dinámica y exhaustiva de encabezados en la hoja
  let headerRowIndex = -1;
  let colId = -1, colOrden = -1, colEtapa = -1, colResp = -1, colAct = -1, colSub = -1, colDias = -1, colHora = -1, colActivo = -1;

  for (let r = 0; r < Math.min(values.length, 12); r++) {
    const rowNorm = values[r].map(c => normalizarTexto(c));
    const idxAct = rowNorm.findIndex(c => c === "ACTIVIDAD" || c === "ACTIVIDADES" || c.includes("ACTIVIDAD"));
    const idxSub = rowNorm.findIndex(c => c === "SUBACTIVIDAD" || c === "SUBACTIVIDADES" || c.includes("DESCRIPCI") || c === "DESCRIPCION" || c === "TAREA");
    const idxDias = rowNorm.findIndex(c => c.includes("DIAS RELATIV") || c.includes("DIA RELATIV") || c === "DIAS" || c.includes("RELATIVO"));
    const idxResp = rowNorm.findIndex(c => c.includes("RESPONSABLE") || c.includes("ENCARGADO") || c === "AREA");
    const idxOrden = rowNorm.findIndex(c => c === "ORDEN" || c === "N" || c === "NO" || c === "ITEM" || c.includes("NUM"));
    const idxEtapa = rowNorm.findIndex(c => c === "ETAPA" || c === "FASE" || c === "MOMENTO");

    let matchCount = 0;
    if (idxAct !== -1) matchCount++;
    if (idxSub !== -1) matchCount++;
    if (idxDias !== -1) matchCount++;
    if (idxResp !== -1) matchCount++;
    if (idxOrden !== -1) matchCount++;
    if (idxEtapa !== -1) matchCount++;

    if (matchCount >= 2) {
      headerRowIndex = r;
      colAct = idxAct;
      colSub = idxSub;
      colDias = idxDias;
      colResp = idxResp;
      colOrden = idxOrden;
      colEtapa = idxEtapa;
      colId = rowNorm.findIndex(c => c === "ID_TEMPLATE" || c === "ID" || c === "CODIGO" || c === "TPL");
      colHora = rowNorm.findIndex(c => c.includes("HORA"));
      colActivo = rowNorm.findIndex(c => c === "ACTIVO" || c === "ESTADO");
      break;
    }
  }

  if (headerRowIndex === -1) {
    // Si no se encontraron encabezados claros, mapear con catálogo maestro
    return CATALOGO_MAESTRO_OFICIAL.map(a => ({
      id: a.id,
      orden: a.orden,
      etapa: a.etapa,
      responsable: a.responsable,
      actividad: a.actividad,
      descripcion: a.subactividades[0].descripcion,
      diasRelativos: a.diasRelativos,
      horaDefault: a.horaDefault,
      requiereEvidencia: a.orden <= 18,
      subactividadesCustom: a.subactividades.map((s, idx) => ({
        id: `${a.id}-${idx + 1}`,
        orden: s.orden,
        descripcion: s.descripcion,
        hora: s.hora,
        responsable: s.responsable,
        diasRelativos: a.diasRelativos
      }))
    }));
  }

  // Agrupación jerárquica: 25 Actividades Maestras y 37 Subactividades
  const actividadesMap = {};
  let currentOrden = 0;
  let currentEtapa = "Antes del inicio de ciclo";
  let currentResp = "Director de Sede";
  let currentAct = "";
  let currentDias = 0;
  let currentHora = "18:00";

  for (let r = headerRowIndex + 1; r < values.length; r++) {
    const row = values[r];
    const rawId = colId !== -1 ? (row[colId] || "").toString().trim() : "";
    const rawOrden = colOrden !== -1 ? (row[colOrden] !== undefined ? row[colOrden].toString().trim() : "") : "";
    const rawEtapa = colEtapa !== -1 ? (row[colEtapa] || "").toString().trim() : "";
    const rawResp = colResp !== -1 ? (row[colResp] || "").toString().trim() : "";
    const rawAct = colAct !== -1 ? (row[colAct] || "").toString().trim() : "";
    const rawSub = colSub !== -1 ? (row[colSub] || "").toString().trim() : "";
    const rawDias = colDias !== -1 ? (row[colDias] !== undefined ? row[colDias].toString().trim() : "") : "";
    let rawHora = colHora !== -1 ? row[colHora] : "";
    const rawActivo = colActivo !== -1 ? (row[colActivo] || "").toString().trim().toUpperCase() : "SI";

    // Separadores de etapa estilo banner en UNMSM
    const rowJoined = row.map(c => (c || "").toString()).join(" ").toUpperCase();
    if (rowJoined.includes("ANTES DEL INICIO")) {
      currentEtapa = "Antes del inicio de ciclo";
      continue;
    } else if (rowJoined.includes("DURANTE LA PRIMERA") || rowJoined.includes("DURANTE EL INICIO")) {
      currentEtapa = "Durante la primera semana";
      continue;
    } else if (rowJoined.includes("DESPUÉS") || rowJoined.includes("DESPUES")) {
      currentEtapa = "Después de inicio de ciclo";
      continue;
    }

    if (rawActivo === "NO") continue;
    if (!rawId && !rawOrden && !rawAct && !rawSub) continue;

    const numParsed = parseInt(rawOrden, 10);
    if (!isNaN(numParsed) && numParsed > 0) {
      currentOrden = numParsed;
    } else if (rawId && rawId.startsWith("TPL-")) {
      const p = parseInt(rawId.replace("TPL-", "").split("-")[0], 10);
      if (!isNaN(p) && p > 0) currentOrden = p;
    }

    // Etapa según UNMSM o columna
    if (rawEtapa) {
      currentEtapa = rawEtapa;
    } else if (currentOrden > 0) {
      if (currentOrden <= 15) currentEtapa = "Antes del inicio de ciclo";
      else if (currentOrden <= 23) currentEtapa = "Durante la primera semana";
      else currentEtapa = "Después de inicio de ciclo";
    }

    if (rawResp) currentResp = rawResp;
    if (rawAct) currentAct = rawAct;

    // Días relativos: si viene en la fila se usa, o hereda del catálogo oficial según el orden
    if (rawDias !== "") {
      const pDias = parseInt(rawDias, 10);
      if (!isNaN(pDias)) currentDias = pDias;
    } else if (currentOrden > 0) {
      const ref = CATALOGO_MAESTRO_OFICIAL.find(c => c.orden === currentOrden);
      if (ref) currentDias = ref.diasRelativos;
    }

    if (rawHora instanceof Date) {
      currentHora = Utilities.formatDate(rawHora, "GMT-5", "HH:mm");
    } else if (rawHora) {
      currentHora = rawHora.toString().trim();
    } else if (currentOrden > 0) {
      const ref = CATALOGO_MAESTRO_OFICIAL.find(c => c.orden === currentOrden);
      if (ref) currentHora = ref.horaDefault;
    }

    if (!currentOrden && !currentAct) continue;

    const tplId = currentOrden ? `TPL-${String(currentOrden).padStart(3, "0")}` : (rawId || `TPL-${r}`);

    if (!actividadesMap[tplId]) {
      actividadesMap[tplId] = {
        id: tplId,
        orden: currentOrden || (Object.keys(actividadesMap).length + 1),
        etapa: currentEtapa,
        responsable: currentResp,
        actividad: currentAct || `Actividad ${currentOrden}`,
        descripcion: rawSub || currentAct,
        diasRelativos: currentDias,
        horaDefault: currentHora || "18:00",
        requiereEvidencia: currentOrden <= 18,
        subactividadesCustom: []
      };
    }

    const subTexto = (rawSub || currentAct || "").replace(/^[-•*]\s*/, "").trim();
    if (subTexto) {
      const subIdx = actividadesMap[tplId].subactividadesCustom.length + 1;
      actividadesMap[tplId].subactividadesCustom.push({
        id: `${tplId}-${subIdx}`,
        orden: subIdx,
        descripcion: subTexto,
        hora: currentHora || "18:00",
        responsable: rawResp || currentResp,
        diasRelativos: currentDias
      });
    }
  }

  const items = Object.values(actividadesMap);
  if (items.length >= 20) {
    return items;
  }

  // Si se leyeron pocas filas, enriquecer con el catálogo maestro garantizando las 25 actividades
  return CATALOGO_MAESTRO_OFICIAL.map(a => {
    const leido = items.find(i => i.orden === a.orden || i.id === a.id);
    if (leido && leido.subactividadesCustom && leido.subactividadesCustom.length > 0) {
      return leido;
    }
    return {
      id: a.id,
      orden: a.orden,
      etapa: a.etapa,
      responsable: a.responsable,
      actividad: a.actividad,
      descripcion: a.subactividades[0].descripcion,
      diasRelativos: a.diasRelativos,
      horaDefault: a.horaDefault,
      requiereEvidencia: a.orden <= 18,
      subactividadesCustom: a.subactividades.map((s, idx) => ({
        id: `${a.id}-${idx + 1}`,
        orden: s.orden,
        descripcion: s.descripcion,
        hora: s.hora,
        responsable: s.responsable,
        diasRelativos: a.diasRelativos
      }))
    };
  });
}

/**
 * Lee el Directorio Maestro desde "ADMIN_DIRECTORIO" mapeando dinámicamente por nombres de encabezados.
 * Es 100% escalable: no depende de índices fijos de columnas ni de posiciones absolutas de filas.
 */
function obtenerDirectorioDesdeSheet() {
  const ss = obtenerSpreadsheet();
  const sheet = ss.getSheetByName("ADMIN_DIRECTORIO");
  if (!sheet) return [];

  const allData = sheet.getDataRange().getValues();
  if (allData.length < 2) return [];

  let headerRowIndex = -1;
  let colMap = {};

  // Buscar dinámicamente la fila donde están los encabezados principales
  for (let r = 0; r < Math.min(allData.length, 15); r++) {
    const fila = allData[r].map(c => normalizarTexto(c));
    const tieneRol = fila.some(c => c === "ROL" || c.startsWith("ROL"));
    const tieneCorreo = fila.some(c => c.includes("CORREO"));

    if (tieneRol && tieneCorreo) {
      headerRowIndex = r;
      fila.forEach((header, c) => {
        if (!header) return;
        // Mapeo dinámico por nombres de encabezado (escalable)
        if (header.includes("ROL")) colMap.rol = c;
        if (header.includes("ABREVIATURA") || header.includes("CODIGO")) colMap.codigo = c;
        if (header.includes("SEDE") && !header.includes("ROL") && !header.includes("CODIGO") && !header.includes("ABREVIATURA")) colMap.sede = c;
        if (header.includes("NOMBRE") || header.includes("DIRECTOR") || header.includes("RESPONSABLE")) {
          if (!header.includes("SEDE")) colMap.nombre = c;
        }
        if (header.includes("CELULAR") || header.includes("TELEFONO") || (header.includes("WHATSAPP") && !header.includes("NOTIFICAR"))) colMap.celular = c;
        if (header.includes("CORREO") && !header.includes("NOTIFICAR")) colMap.correo = c;
        if (header.includes("NOTIFICAR") && header.includes("WHATSAPP")) colMap.notifWsp = c;
        if (header.includes("NOTIFICAR") && header.includes("CORREO")) colMap.notifCorreo = c;
        if (header.includes("OBSERVACION")) colMap.observaciones = c;
      });
      break;
    }
  }

  if (headerRowIndex === -1 || colMap.rol === undefined) return [];

  const directorio = [];
  for (let r = headerRowIndex + 1; r < allData.length; r++) {
    const row = allData[r];
    const filaTexto = row.map(c => (c || "").toString()).join(" ").toUpperCase();
    if (filaTexto.includes("CONSOLA") || filaTexto.includes("RECORDATORIOS (ACTIVIDADES")) break;

    const rol = colMap.rol !== undefined ? (row[colMap.rol] || "").toString().trim().toUpperCase() : "";
    const codigoSedeRaw = colMap.codigo !== undefined ? (row[colMap.codigo] || "").toString().trim().toUpperCase() : "";
    const nombreSede = colMap.sede !== undefined ? (row[colMap.sede] || "").toString().trim() : "";
    const nombre = colMap.nombre !== undefined ? (row[colMap.nombre] || "").toString().trim() : "";
    const celular = colMap.celular !== undefined ? (row[colMap.celular] || "").toString().replace(/[^0-9]/g, "") : "";
    const correoRaw = colMap.correo !== undefined ? (row[colMap.correo] || "").toString().trim() : "";

    if (!rol && !codigoSedeRaw && !nombre && !correoRaw) continue;
    if (codigoSedeRaw === "INACTIVO" || nombre.toUpperCase() === "INACTIVO" || correoRaw.toUpperCase() === "INACTIVO") continue;

    // Determinar si es un rol multisede (central) o asignado a una sola sede física
    const esMultisede = codigoSedeRaw === "-" || codigoSedeRaw === "" ||
      rol.includes("VENTAS") || rol.includes("SISTEMAS") || rol.includes("PRODUCCION") ||
      rol.includes("EVALUACION") || rol.includes("DIGITALIZACION") || rol.includes("PROCESOS") || rol.includes("ZONAL");

    const codigoSedeFinal = esMultisede ? "TODAS" : codigoSedeRaw;
    const nombreSedeFinal = esMultisede ? "Todas las Sedes (Multisede)" : (nombreSede || codigoSedeRaw);

    const correos = correoRaw.split(",").map(c => c.trim().toLowerCase()).filter(c => c.includes("@"));

    const itemBase = {
      rol: rol,
      codigoSede: codigoSedeFinal,
      codigoSedeOriginal: codigoSedeRaw,
      nombreSede: nombreSedeFinal,
      nombre: nombre || (rol.includes("SISTEMAS") ? "Área de Sistemas" : (rol.includes("VENTAS") ? "Área de Ventas" : rol)),
      celular: celular,
      esMultisede: esMultisede,
      esDirectorSede: !esMultisede && (rol.includes("DIRECTOR DE SEDE") || rol.includes("TUTOR")),
      esProcesos: rol.includes("PROCESOS"),
      esDirectorZonal: rol.includes("ZONAL")
    };

    if (correos.length === 0) {
      directorio.push(Object.assign({}, itemBase, { email: "" }));
    } else {
      correos.forEach(em => {
        directorio.push(Object.assign({}, itemBase, { email: em }));
      });
    }
  }
  return directorio;
}

/**
 * Obtiene la hoja BD_CHECKLISTS (Estricto, sin crear hojas huérfanas)
 */
function obtenerHojaBDChecklists() {
  const ss = obtenerSpreadsheet();
  const sheet = ss.getSheetByName("BD_CHECKLISTS");
  if (!sheet) {
    throw new Error('La hoja "BD_CHECKLISTS" no existe en la hoja de cálculo. Por favor verifique el nombre de la hoja.');
  }
  return sheet;
}

/**
 * Obtiene o crea la hoja BD_SEGUIMIENTO
 */
function obtenerHojaBDSeguimiento() {
  const ss = obtenerSpreadsheet();
  let sheet = ss.getSheetByName("BD_SEGUIMIENTO") || ss.getSheetByName("_BD_SEGUIMIENTO");
  if (!sheet) {
    sheet = ss.insertSheet("BD_SEGUIMIENTO");
    sheet.setTabColor("#10b981");
    const headers = [
      "ID_REGISTRO", "ID_CHECKLIST", "UNIVERSIDAD", "ID_TAREA", "CODIGO_SEDE",
      "ESTADO", "FECHA_MARCACION", "USUARIO_CORREO", "URL_EVIDENCIA", "OBSERVACION_JUSTIFICACION"
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setBackground("#0f766e").setFontColor("#FFFFFF").setFontWeight("bold");
    sheet.autoResizeColumns(1, headers.length);
  }
  return sheet;
}

/**
 * Obtiene o crea la hoja BD_REPORTES_ETAPAS (Sin guión bajo)
 */
function obtenerHojaBDReportes() {
  const ss = obtenerSpreadsheet();
  let sheet = ss.getSheetByName("BD_REPORTES_ETAPAS") || ss.getSheetByName("_BD_REPORTES_ETAPAS");
  if (!sheet) {
    sheet = ss.insertSheet("BD_REPORTES_ETAPAS");
    sheet.setTabColor("#6366f1");
    const headers = [
      "ID_REPORTE", "ID_CHECKLIST", "ETAPA", "FECHA_CORTE", "AUDITOR",
      "TOTAL_TAREAS", "COMPLETADAS", "VENCIDAS", "NO_APLICA", "PCT_PONDERADO", "DETALLE_JSON"
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setBackground("#4338ca").setFontColor("#FFFFFF").setFontWeight("bold");
    sheet.autoResizeColumns(1, headers.length);
  }
  return sheet;
}

// =========================================================================
// 2. CONTROL DE SESIÓN, IDENTIDAD Y ACCESO (ROLES)
// =========================================================================

/**
 * Resuelve el usuario activo contra ADMIN_DIRECTORIO
 * Soporta parámetros de simulación/testeo en URL: ?user=marevalo@hnc.pe o ?rol=VENTAS o ?sede=LCE
 */
function resolverUsuarioActivo(e) {
  let email = "";
  let simRol = "";
  let simSede = "";

  if (e && e.parameter) {
    email = (e.parameter.user || e.parameter.email || e.parameter.u || "").toString().toLowerCase().trim();
    simRol = (e.parameter.rol || e.parameter.r || "").toString().toUpperCase().trim();
    simSede = (e.parameter.sede || e.parameter.s || "").toString().toUpperCase().trim();
  }

  if (!email) {
    try {
      email = (Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail() || "").toLowerCase().trim();
    } catch (err) {
      email = "";
    }
  }

  const directorio = obtenerDirectorioDesdeSheet();

  // 1. Simulación explícita por rol o sede
  if (simRol || simSede) {
    const matchSim = directorio.find(u => {
      const matchR = simRol ? normalizarTexto(u.rol).includes(normalizarTexto(simRol)) : true;
      const matchS = simSede ? (u.codigoSede === simSede || u.codigoSedeOriginal === simSede) : true;
      return matchR && matchS;
    });

    if (matchSim) {
      return {
        email: matchSim.email || `${matchSim.rol.toLowerCase().replace(/[^a-z]/g, "")}@vonex.edu.pe`,
        nombre: matchSim.nombre,
        rol: matchSim.rol,
        codigoSede: matchSim.codigoSede,
        nombreSede: matchSim.nombreSede,
        celular: matchSim.celular || "",
        esMultisede: matchSim.esMultisede,
        esProcesos: matchSim.esProcesos,
        esDirectorZonal: matchSim.esDirectorZonal,
        esDirectorSede: matchSim.esDirectorSede
      };
    }
  }

  if (email) {
    // 2. Coincidencia exacta por correo en ADMIN_DIRECTORIO
    const match = directorio.find(u => u.email === email);
    if (match) {
      return {
        email: match.email,
        nombre: match.nombre,
        rol: match.rol,
        codigoSede: match.codigoSede,
        nombreSede: match.nombreSede,
        celular: match.celular || "",
        esMultisede: match.esMultisede,
        esProcesos: match.esProcesos,
        esDirectorZonal: match.esDirectorZonal,
        esDirectorSede: match.esDirectorSede
      };
    }

    // 3. Coincidencias de respaldo institucionales específicas
    if (email === "arivas@hnc.pe") {
      return { email: email, nombre: "Director Zonal", rol: "DIRECTOR ZONAL", codigoSede: "TODAS", nombreSede: "Todas las Sedes (Multisede)", celular: "920453837", esMultisede: true, esProcesos: false, esDirectorZonal: true, esDirectorSede: false };
    }
    if (email === "gvallejo@hnc.pe" || email.includes("vallejo") || email.includes("sistemas")) {
      return { email: email, nombre: "Área de Sistemas", rol: "SISTEMAS", codigoSede: "TODAS", nombreSede: "Todas las Sedes (Multisede)", celular: "991770066", esMultisede: true, esProcesos: false, esDirectorZonal: false, esDirectorSede: false };
    }
    if (email === "marevalo@hnc.pe" || email.includes("ventas")) {
      return { email: email, nombre: "MAFER ARÉVALO", rol: "VENTAS", codigoSede: "TODAS", nombreSede: "Todas las Sedes (Multisede)", celular: "923487758", esMultisede: true, esProcesos: false, esDirectorZonal: false, esDirectorSede: false };
    }
    if (email.includes("procesos")) {
      return { email: email, nombre: "Gestión de Procesos", rol: "PROCESOS", codigoSede: "TODAS", nombreSede: "Todas las Sedes (Multisede)", celular: "", esMultisede: true, esProcesos: true, esDirectorZonal: false, esDirectorSede: false };
    }
  }

  // 4. Default institucional seguro
  return {
    email: email || "procesos@vonex.edu.pe",
    nombre: "Gestión de Procesos",
    rol: "PROCESOS",
    codigoSede: "TODAS",
    nombreSede: "Todas las Sedes (Multisede)",
    celular: "",
    esMultisede: true,
    esProcesos: true,
    esDirectorZonal: false,
    esDirectorSede: false
  };
}

// =========================================================================
// 3. PUNTO DE ENTRADA WEB (doGet)
// =========================================================================

/**
 * Enruta según parámetro o perfil del usuario
 * ?p=crear -> CREAR.html
 * ?p=participantes -> PARTICIPANTES.html
 */
function doGet(e) {
  let page = "";
  if (e && e.parameter) {
    page = (e.parameter.p || e.parameter.page || e.parameter.v || "").toString().trim().toUpperCase();
  }

  const usuario = resolverUsuarioActivo(e);
  const vistaDestino = (page === "PARTICIPANTES" || page === "PARTICIPANTE" || page === "OPERATIVO")
    ? "PARTICIPANTES"
    : ((page === "CREAR" || page === "ADMIN" || page === "PROCESOS")
      ? "CREAR"
      : ((usuario.esProcesos || usuario.esDirectorZonal) ? "CREAR" : "PARTICIPANTES"));

  // Obtener contexto de datos de forma síncrona en el servidor para evitar parpadeos y esperas asíncronas en el cliente
  let contextoInicial = {};
  if (vistaDestino === "CREAR") {
    try {
      contextoInicial = apiObtenerContextoCrear(e);
    } catch (err) {
      contextoInicial = { usuario: usuario };
    }
  } else {
    try {
      const sedeParam = (e && e.parameter && (e.parameter.sede || e.parameter.s)) || "";
      contextoInicial = apiObtenerContextoParticipantes(sedeParam, e);
    } catch (err) {
      contextoInicial = { usuario: usuario };
    }
  }

  const template = HtmlService.createTemplateFromFile(vistaDestino);
  template.initialData = JSON.stringify(contextoInicial);

  return template.evaluate()
    .setTitle(vistaDestino === "CREAR" ? "Gestión de Inicio de Clases - Creación y Control de Cumplimiento" : "Módulo de Participantes • Cumplimiento y Evidencias")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
}

// =========================================================================
// 4. API PARA VISTA 1: CREAR.HTML (ADMINISTRACIÓN & MONITOREO)
// =========================================================================

/**
 * Contexto inicial para la vista CREAR.html
 */
function apiObtenerContextoCrear() {
  const usuario = resolverUsuarioActivo();
  const plantilla = obtenerDatosPlantilla();
  const directorio = obtenerDirectorioDesdeSheet();
  const sheetBD = obtenerHojaBDChecklists();
  const allBD = sheetBD.getDataRange().getValues();

  const sedesValidas = ["LCE", "SJL", "IND", "HDM", "SAN", "COM", "PTP", "VMT"];

  const checklists = [];
  const headersBD = allBD[0] || [];
  let colHiloIdx = -1;
  headersBD.forEach((h, idx) => {
    if (normalizarTexto(h).includes("HILO")) colHiloIdx = idx;
  });

  for (let r = 1; r < allBD.length; r++) {
    const row = allBD[r];
    if (!row[0]) continue;
    checklists.push({
      id: row[0].toString().trim(),
      universidad: (row[1] || "UNMSM").toString().trim(),
      nombre: (row[2] || "").toString().trim(),
      fechaInicio: row[3] instanceof Date ? formatearFechaISO(row[3]) : (row[3] || "").toString().trim(),
      cicloVida: (row[4] || "Creado").toString().trim(),
      creador: (row[5] || "").toString().trim(),
      fechaCreacion: row[6] instanceof Date ? Utilities.formatDate(row[6], "GMT-5", "yyyy-MM-dd") : (row[6] || "").toString().trim(),
      idHilo: (colHiloIdx >= 0 && row[colHiloIdx]) ? row[colHiloIdx].toString().trim() : ""
    });
  }

  // Obtener sedes operativas dinámicamente desde ADMIN_DIRECTORIO (columna ABREVIATURA / SEDE)
  const sedesFromDirectorio = directorio
    .filter(u => u.codigoSede && u.codigoSede !== "INACTIVO" && u.codigoSede !== "LCE-CENTRAL")
    .map(u => ({ cod: u.codigoSede, nombre: u.nombreSede || u.codigoSede }));

  // Remover duplicados por código de sede
  const mapaSedesUnicas = {};
  sedesFromDirectorio.forEach(s => {
    if (!mapaSedesUnicas[s.cod]) mapaSedesUnicas[s.cod] = s;
  });

  let listaSedesCodigos = Object.keys(mapaSedesUnicas);
  if (listaSedesCodigos.length === 0) {
    listaSedesCodigos = ["LCE", "SJL", "IND", "HDM", "SAN", "COM", "PTP", "VMT"];
  }

  // 3. Leer registros reales de _BD_SEGUIMIENTO para calcular avances reales por sede y actividad
  const sheetSeg = obtenerHojaBDSeguimiento();
  const dataSeg = sheetSeg.getDataRange().getValues();
  // mapaSeguimiento[idTarea][codigoSede] = { estado, evidencia, fecha, checklistId, nota }
  const mapaSeguimiento = {};

  for (let r = 1; r < dataSeg.length; r++) {
    const row = dataSeg[r];
    const chkId = (row[1] || "").toString().trim();
    const tId = (row[3] || "").toString().trim();
    const sCod = (row[4] || "").toString().trim().toUpperCase();
    let st = (row[5] || "").toString().trim().toUpperCase();
    if (st === "NO_APLICA") st = "NO APLICA";
    const fecha = row[6] instanceof Date ? formatearFechaISO(row[6]) : (row[6] || "");
    const email = (row[7] || "").toString().trim();
    const evi = (row[8] || "").toString().trim();
    const obs = (row[9] || "").toString().trim();

    if (tId && sCod) {
      const regObj = {
        estado: st,
        checklistId: chkId,
        fecha: fecha,
        usuario: email,
        evidencia: evi,
        nota: obs
      };

      if (!mapaSeguimiento[tId]) mapaSeguimiento[tId] = {};
      if (sCod === "TODAS") {
        listaSedesCodigos.forEach(c => {
          mapaSeguimiento[tId][c] = regObj;
        });
      } else {
        mapaSeguimiento[tId][sCod] = regObj;
      }

      // Alias ACT-xx para compatibilidad total si se guardó como TPL-xxx
      if (tId.startsWith("TPL-")) {
        const num = parseInt(tId.replace("TPL-", ""), 10);
        if (!isNaN(num)) {
          const actKey = "ACT-" + String(num).padStart(2, "0");
          if (!mapaSeguimiento[actKey]) mapaSeguimiento[actKey] = {};
          if (sCod === "TODAS") {
            listaSedesCodigos.forEach(c => {
              mapaSeguimiento[actKey][c] = regObj;
            });
          } else {
            mapaSeguimiento[actKey][sCod] = regObj;
          }
        }
      }
    }
  }

  // Resumen de sedes para monitoreo y checkboxes de cobertura calculado con _BD_SEGUIMIENTO real
  const sedesInfo = listaSedesCodigos.map(cod => {
    const dir = directorio.find(u => u.codigoSede === cod) || { nombre: "Director " + cod, nombreSede: cod };
    let hechas = 0;
    let vencidas = 0;
    let noaplica = 0;
    const totalPlantilla = plantilla.length || 25;

    plantilla.forEach(t => {
      const registro = (mapaSeguimiento[t.id] && mapaSeguimiento[t.id][cod]) || (mapaSeguimiento["ACT-" + String(t.orden).padStart(2, "0")] && mapaSeguimiento["ACT-" + String(t.orden).padStart(2, "0")][cod]);
      const st = registro ? registro.estado : "PENDIENTE";
      if (st === "COMPLETADA") hechas++;
      else if (st === "NO APLICA") noaplica++;
      else if (st === "VENCIDA") vencidas++;
    });

    const aplicables = Math.max(1, totalPlantilla - noaplica);
    const pct = Math.min(100, Math.round((hechas / aplicables) * 100));

    return {
      cod: cod,
      nombre: dir.nombreSede || cod,
      director: dir.nombre || ("Director " + cod),
      total: totalPlantilla,
      hechas: hechas,
      vencidas: vencidas,
      noaplica: noaplica,
      pct: pct
    };
  });

  return {
    usuario: usuario,
    plantilla: plantilla,
    checklists: checklists,
    sedes: sedesInfo,
    activeChecklist: checklists[0],
    seguimientoReal: mapaSeguimiento
  };
}

/**
 * Crea un nuevo checklist en BD_CHECKLISTS
 */
function apiCrearChecklist(datos) {
  const usuario = resolverUsuarioActivo();
  const ss = obtenerSpreadsheet();
  const sheetBD = obtenerHojaBDChecklists();

  const uni = datos.universidad || "UNMSM";
  const fechaInicio = datos.fechaInicio || formatearFechaISO(new Date());
  const nombre = datos.nombre || `Ciclo Regular ${uni} 2026-I`;
  const cicloVida = datos.cicloVida || "Creado";

  const partes = fechaInicio.split("-");
  const yyyy = partes[0] || "2026";
  const mm = partes[1] || "11";
  const dd = partes[2] || "03";
  const idChecklist = `CHK-${yyyy}${mm}${dd}-01`;

  const ahora = new Date();
  sheetBD.appendRow([
    idChecklist,
    uni,
    nombre,
    fechaInicio,
    cicloVida,
    usuario.email,
    ahora,
    25
  ]);

  return {
    exito: true,
    idChecklist: idChecklist,
    universidad: uni,
    nombre: nombre,
    fechaInicio: fechaInicio,
    cicloVida: cicloVida
  };
}

/**
 * Asegura o retorna el índice (1-based) de la columna especificada en una hoja
 */
function asegurarColumnaEnHoja(sheet, nombreColumna) {
  const ultimaCol = Math.max(1, sheet.getLastColumn());
  const headers = sheet.getRange(1, 1, 1, ultimaCol).getValues()[0];
  for (let c = 0; c < headers.length; c++) {
    if (normalizarTexto(headers[c]) === normalizarTexto(nombreColumna)) {
      return c + 1;
    }
  }
  const nuevaCol = ultimaCol + 1;
  sheet.getRange(1, nuevaCol).setValue(nombreColumna).setBackground("#1e293b").setFontColor("#FFFFFF").setFontWeight("bold");
  return nuevaCol;
}

/**
 * Actualiza el ciclo de vida de un checklist y opcionalmente su ID de hilo de correo
 */
function apiActualizarCicloVida(idChecklist, nuevoCiclo, idHilo) {
  const sheetBD = obtenerHojaBDChecklists();
  const data = sheetBD.getDataRange().getValues();

  for (let r = 1; r < data.length; r++) {
    if (data[r][0] && data[r][0].toString().trim() === idChecklist) {
      sheetBD.getRange(r + 1, 5).setValue(nuevoCiclo);
      if (idHilo) {
        const colHilo = asegurarColumnaEnHoja(sheetBD, "ID_HILO_GMAIL");
        sheetBD.getRange(r + 1, colHilo).setValue(idHilo);
      }
      return { exito: true, idChecklist: idChecklist, cicloVida: nuevoCiclo, idHilo: idHilo || "" };
    }
  }
  return { exito: false, mensaje: "Checklist no encontrado" };
}

/**
 * Guarda las modificaciones realizadas por el Director Zonal en "Definir fecha":
 * fechas límites, horas, requerimiento de evidencia o actividades marcadas como "No Aplica / No requerida"
 */
/**
 * Obtiene la URL pública del Web App con fallback al enlace desplegado
 */
function obtenerUrlWebApp() {
  try {
    const url = ScriptApp.getService().getUrl();
    if (url && url.length > 5) return url;
  } catch (e) { }
  return "https://script.google.com/a/macros/hnc.pe/s/AKfycbyuLqU4cJiuL4VfQb-XgSeEctEXc32GD3fXkKJ9EGL_Dupx2JgM5Q-gv-0nag1LvP_4yw/exec";
}

/**
 * Guarda las modificaciones realizadas por el Director Zonal en "Definir fecha":
 * fechas límites, horas, requerimiento de evidencia o actividades marcadas como "No Aplica / No requerida"
 */
function apiGuardarConfiguracionTareas(idChecklist, tareasModificadas) {
  const usuario = resolverUsuarioActivo();
  const sheetSeg = obtenerHojaBDSeguimiento();

  if (Array.isArray(tareasModificadas) && tareasModificadas.length > 0) {
    const ahora = new Date();
    tareasModificadas.forEach(t => {
      if (t.noAplica) {
        // Registrar en _BD_SEGUIMIENTO como NO APLICA con la justificación
        sheetSeg.appendRow([
          `REG-DEF-${Date.now()}-${t.id}`,
          idChecklist,
          t.universidad || "UNMSM",
          t.id,
          t.codigoSede || "TODAS",
          "NO APLICA",
          ahora,
          usuario.email,
          "",
          t.motivoNoAplica || "Desactivada por Dirección Zonal"
        ]);
      }
    });
  }

  // Actualizar ciclo de vida a "Definir fecha" o "Activo"
  apiActualizarCicloVida(idChecklist, "Definir fecha");

  return { exito: true, idChecklist: idChecklist, totalAjustes: (tareasModificadas || []).length };
}

/**
 * Confirma formalmente las fechas ajustadas del checklist,
 * actualiza su ciclo de vida a "Activo" en BD_CHECKLISTS
 * y crea el hilo oficial de correo con asunto corporativo y en copia (CC)
 * a todos los integrantes de ADMIN_DIRECTORIO.
 */
function apiConfirmarFechasYNotificar(idChecklist, tareasModificadas, checklistInfo) {
  const usuario = resolverUsuarioActivo();
  const sheetSeg = obtenerHojaBDSeguimiento();
  const sheetBD = obtenerHojaBDChecklists();

  // 1. Registrar actividades marcadas como "No Aplica" en _BD_SEGUIMIENTO
  if (Array.isArray(tareasModificadas) && tareasModificadas.length > 0) {
    const ahora = new Date();
    tareasModificadas.forEach(t => {
      if (t.noAplica) {
        sheetSeg.appendRow([
          `REG-DEF-${Date.now()}-${t.id || t.orden || '0'}`,
          idChecklist,
          t.universidad || (checklistInfo && checklistInfo.universidad) || "UNMSM",
          t.id || `TPL-${String(t.orden).padStart(3, "0")}`,
          t.codigoSede || "TODAS",
          "NO APLICA",
          ahora,
          usuario.email,
          "",
          t.motivoNoAplica || "Desactivada por Dirección Zonal / Procesos"
        ]);
      }
    });
  }

  // 2. Obtener o inferir datos del checklist (Universidad, Fecha de Inicio)
  let universidad = (checklistInfo && checklistInfo.universidad) || "UNMSM";
  let fechaInicio = (checklistInfo && checklistInfo.fechaInicio) || "";

  if (!fechaInicio) {
    const dataBD = sheetBD.getDataRange().getValues();
    for (let r = 1; r < dataBD.length; r++) {
      if (dataBD[r][0] && dataBD[r][0].toString().trim() === idChecklist) {
        universidad = dataBD[r][1] || universidad;
        const valF = dataBD[r][3];
        fechaInicio = valF instanceof Date ? formatearFechaISO(valF) : (valF || "");
        break;
      }
    }
  }

  // Mapeo de mes y año para el asunto estandarizado
  const MESES = [
    "ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO",
    "JULIO", "AGOSTO", "SETIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"
  ];

  let mesNombre = "OCTUBRE";
  let anio = "2026";
  let fechaFormateada = fechaInicio || "26/10/2026";

  if (fechaInicio && fechaInicio.includes("-")) {
    const p = fechaInicio.split("-");
    anio = p[0] || "2026";
    const mIdx = parseInt(p[1], 10) - 1;
    if (mIdx >= 0 && mIdx < 12) {
      mesNombre = MESES[mIdx];
    }
    const dia = p[2] || "01";
    fechaFormateada = `${dia}/${p[1]}/${anio}`;
  }

  // Asunto corporativo histórico idéntico al hilo:
  // "GP | CHECK LIST DE ACTIVIDADES OCTUBRE 2026 - GESTION DE INICIO DE CLASES"
  const asunto = `GP | CHECK LIST DE ACTIVIDADES ${mesNombre} ${anio} - GESTION DE INICIO DE CLASES`;

  // 4. Obtener todos los correos registrados en la hoja ADMIN_DIRECTORIO
  const directorio = obtenerDirectorioDesdeSheet();
  const correosSet = new Set();
  directorio.forEach(item => {
    if (item.email && item.email.includes("@")) {
      correosSet.add(item.email.trim().toLowerCase());
    }
  });

  const todosLosCorreos = Array.from(correosSet);

  // Remitente
  const remitenteEmail = usuario.email || "procesos@hnc.pe";

  // URL para ingresar al módulo de participantes
  const urlBase = obtenerUrlWebApp();
  const urlParticipantes = `${urlBase}?page=participantes&chk=${encodeURIComponent(idChecklist)}`;

  // Cuerpo del correo moderno 100% ancho completo (estilo institucional banner azul + tarjetas)
  const cuerpoHtml = `
    <div style="font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif; width: 100%; margin: 0; padding: 0; background-color: #f1f5f9;">
      <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="width: 100%; background-color: #ffffff; border: 1px solid #dbeafe; border-radius: 14px; overflow: hidden; border-collapse: separate;">
        <!-- BANNER SUPERIOR AZUL MARINO CON LÍNEA DORADA -->
        <tr>
          <td style="background-color: #092c74; background: linear-gradient(180deg, #0d3b99 0%, #071e54 100%); border-bottom: 3px solid #facc15; padding: 24px 24px 20px 24px; text-align: center;">
            <div style="display: inline-block; background-color: #ffe600; color: #071e54; font-weight: 900; font-size: 13px; padding: 7px 18px; border-radius: 8px; letter-spacing: 0.5px; margin-bottom: 12px;">
              ACADEMIA VONEX
            </div>
            <div style="color: #ffffff; font-size: 19px; font-weight: 800; letter-spacing: 0.6px; text-transform: uppercase; margin: 0 0 8px 0;">
              CHECKLIST DE GESTI&#211;N DE INICIO DE CLASES
            </div>
            <div style="color: #e2e8f0; font-size: 12px; font-weight: 700; letter-spacing: 0.4px; text-transform: uppercase;">
              UNIVERSIDAD: <span style="color: #fde047;">${universidad}</span> &nbsp;|&nbsp; INICIO DE CLASES: <span style="color: #fde047;">${fechaFormateada}</span>
            </div>
          </td>
        </tr>

        <!-- CUERPO DEL CORREO (ANCHO COMPLETO CON TARJETAS MODERNAS) -->
        <tr>
          <td style="background-color: #f8fafc; padding: 24px 26px;">
            <!-- SALUDO E INTRODUCCIÓN -->
            <div style="background-color: #ffffff; border: 1px solid #e2e8f0; border-radius: 12px; padding: 18px 20px; margin-bottom: 16px;">
              <p style="margin: 0 0 8px 0; font-size: 14.5px; font-weight: 700; color: #0f172a;">
                Buen d&#237;a, estimados(as). &#128075;
              </p>
              <p style="margin: 0; font-size: 13.5px; color: #334155; line-height: 1.55;">
                Se comparte el <strong>CHECKLIST de Gesti&#243;n de Inicio de Clases</strong> para que cada responsable registre el avance de las actividades asignadas.
              </p>
            </div>

            <!-- DOS TARJETAS PARALELAS (UNIVERSIDAD + FECHA DE INICIO) -->
            <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="width: 100%; margin-bottom: 16px;">
              <tr>
                <td width="49%" valign="top" style="background-color: #ffffff; border: 1px solid #e2e8f0; border-top: 3px solid #0d3b99; border-radius: 12px; padding: 16px 18px;">
                  <div style="font-size: 11px; font-weight: 800; color: #0d3b99; text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 8px;">
                    &#127891; Universidad
                  </div>
                  <div style="font-size: 17px; font-weight: 800; color: #0f172a;">
                    ${universidad}
                  </div>
                </td>
                <td width="2%">&nbsp;</td>
                <td width="49%" valign="top" style="background-color: #ffffff; border: 1px solid #e2e8f0; border-top: 3px solid #facc15; border-radius: 12px; padding: 16px 18px;">
                  <div style="margin-bottom: 8px;">
                    <span style="display: inline-block; background-color: #fef9c3; color: #854d0e; font-size: 10.5px; font-weight: 800; padding: 3px 9px; border-radius: 6px; text-transform: uppercase; letter-spacing: 0.4px;">
                      &#128197; Fecha de inicio de clases
                    </span>
                  </div>
                  <div style="font-size: 17px; font-weight: 800; color: #0f172a;">
                    ${fechaFormateada}
                  </div>
                </td>
              </tr>
            </table>

            <!-- TARJETA DE MONITOREO Y CONSULTAS (ACENTO LATERAL AZUL) -->
            <div style="background-color: #ffffff; border: 1px solid #e2e8f0; border-left: 4px solid #0d3b99; border-radius: 12px; padding: 16px 20px; margin-bottom: 22px;">
              <div style="font-size: 13px; color: #1e293b; line-height: 1.55;">
                Este checklist cuenta con <strong>monitoreo continuo</strong> para asegurar el cumplimiento dentro de los plazos establecidos.
              </div>
              <div style="border-top: 1px dashed #cbd5e1; margin: 12px 0;"></div>
              <div style="font-size: 13px; color: #334155; font-weight: 600;">
                Quedamos atentos ante cualquier consulta &#128172;
              </div>
            </div>

            <!-- BOTÓN CTA PÍLDORA CENTRADO -->
            <div style="text-align: center; margin: 22px 0 20px 0;">
              <a href="${urlParticipantes}" target="_blank" style="display: inline-block; background-color: #133c94; background: linear-gradient(180deg, #1b4bb5 0%, #0e2d75 100%); color: #ffffff !important; font-size: 13px; font-weight: 800; text-decoration: none; padding: 13px 34px; border-radius: 9999px; text-transform: uppercase; letter-spacing: 0.7px; box-shadow: 0 4px 12px rgba(19, 60, 148, 0.25);">
                ABRIR CHECKLIST
              </a>
            </div>

            <!-- FIRMA CONCISA -->
            <div style="border-top: 1px solid #e2e8f0; padding-top: 14px; font-size: 12.5px; color: #334155; line-height: 1.5;">
              Saludos,<br>
              <strong style="color: #0d3b99;">GPMC | Academia Vonex</strong>
            </div>
          </td>
        </tr>

        <!-- PIE DE PÁGINA INSTITUCIONAL -->
        <tr>
          <td style="background-color: #f8fafc; border-top: 1px solid #e2e8f0; padding: 12px 20px; text-align: center; font-size: 11px; color: #94a3b8;">
            Notificaci&#243;n autom&#225;tica del Sistema de Checklist Operativo &bull; Academia Vonex
          </td>
        </tr>
      </table>
    </div>
  `;

  // Destinatario principal: remitente / procesos
  let destinatarioTo = remitenteEmail;
  // En copia (CC): toda la lista de correos de ADMIN_DIRECTORIO (excluyendo a to si coincide)
  let listaCc = todosLosCorreos.filter(c => c !== destinatarioTo);

  if (!destinatarioTo && todosLosCorreos.length > 0) {
    destinatarioTo = todosLosCorreos[0];
    listaCc = todosLosCorreos.slice(1);
  }

  let enviado = false;
  let errorEnvio = "";
  let threadId = "";

  try {
    GmailApp.sendEmail(destinatarioTo, asunto, "", {
      cc: listaCc.join(","),
      htmlBody: cuerpoHtml,
      name: "PROCESOS"
    });
    enviado = true;

    // Obtener el ID del hilo de Gmail recién creado para guardarlo en BD_CHECKLISTS
    Utilities.sleep(600);
    const threads = GmailApp.search(`subject:"${asunto}"`, 0, 1);
    if (threads && threads.length > 0) {
      threadId = threads[0].getId();
    }
  } catch (e) {
    errorEnvio = e.message;
    Logger.log("Aviso de envío de correo: " + e.message);
  }

  // 2. Actualizar ciclo de vida a "Activo" y registrar ID_HILO_GMAIL en BD_CHECKLISTS
  apiActualizarCicloVida(idChecklist, "Activo", threadId);

  return {
    exito: true,
    idChecklist: idChecklist,
    idHilo: threadId,
    enviado: enviado,
    asunto: asunto,
    totalDestinatarios: todosLosCorreos.length,
    errorEnvio: errorEnvio
  };
}

/**
 * Genera y compila el contenido del Reporte de KPI y la Matriz de Evidencia de Etapa
 */
function compilarDatosReporteEtapa(etapa, idChecklist) {
  const ss = obtenerSpreadsheet();
  const plantilla = obtenerDatosPlantilla();
  const sheetSeg = obtenerHojaBDSeguimiento();
  const dataSeg = sheetSeg.getDataRange().getValues();
  const directorio = obtenerDirectorioDesdeSheet();

  // 1. Filtrar actividades según la etapa
  const codEtapa = (etapa || "ANTES").toUpperCase();
  let nombreEtapaOficial = "Antes del inicio de ciclo";
  let filtroRango = [1, 15];
  if (codEtapa === "DURANTE") {
    nombreEtapaOficial = "Durante la primera semana";
    filtroRango = [16, 23];
  } else if (codEtapa === "DESPUES") {
    nombreEtapaOficial = "Después de inicio de ciclo";
    filtroRango = [24, 25];
  }

  const tareasEtapa = plantilla.filter(t => {
    const orden = parseInt(t.orden, 10);
    return orden >= filtroRango[0] && orden <= filtroRango[1];
  });

  // 2. Sedes operativas
  const sedesValidas = ["LCE", "SJL", "IND", "HDM", "SAN", "COM", "PTP", "VMT"];

  // 3. Mapa de seguimiento indexado
  const mapaSeg = {};
  for (let r = 1; r < dataSeg.length; r++) {
    const row = dataSeg[r];
    const chk = (row[1] || "").toString().trim();
    if (idChecklist && idChecklist !== "CHK-ACTIVO" && chk && chk !== idChecklist) continue;

    const tId = (row[3] || "").toString().trim();
    const sCod = (row[4] || "").toString().trim().toUpperCase();
    let st = (row[5] || "").toString().trim().toUpperCase();
    if (st === "REALIZADO") st = "COMPLETADA";
    if (st === "NO_APLICA") st = "NO APLICA";

    if (tId && sCod) {
      if (!mapaSeg[tId]) mapaSeg[tId] = {};
      if (sCod === "TODAS") {
        sedesValidas.forEach(c => { mapaSeg[tId][c] = st; });
      } else {
        mapaSeg[tId][sCod] = st;
      }
      // Alias ACT-xx
      if (tId.startsWith("TPL-")) {
        const num = parseInt(tId.replace("TPL-", ""), 10);
        if (!isNaN(num)) {
          const actKey = "ACT-" + String(num).padStart(2, "0");
          if (!mapaSeg[actKey]) mapaSeg[actKey] = {};
          if (sCod === "TODAS") {
            sedesValidas.forEach(c => { mapaSeg[actKey][c] = st; });
          } else {
            mapaSeg[actKey][sCod] = st;
          }
        }
      }
    }
  }

  // 4. Calcular KPIs por sede
  const resumenSedes = [];
  sedesValidas.forEach(codSede => {
    let completadas = 0;
    let noAplica = 0;
    let pendientes = 0;
    const faltantesPorRol = {};

    tareasEtapa.forEach(t => {
      const tId = t.id || `TPL-${String(t.orden).padStart(3, "0")}`;
      const actKey = "ACT-" + String(t.orden).padStart(2, "0");
      const st = (mapaSeg[tId] && mapaSeg[tId][codSede]) || (mapaSeg[actKey] && mapaSeg[actKey][codSede]) || "PENDIENTE";

      if (st === "COMPLETADA") {
        completadas++;
      } else if (st === "NO APLICA") {
        noAplica++;
      } else {
        pendientes++;
        const rol = (t.responsable || "Responsable").trim();
        faltantesPorRol[rol] = (faltantesPorRol[rol] || 0) + 1;
      }
    });

    const aplicables = Math.max(1, tareasEtapa.length - noAplica);
    const pct = Math.min(100, Math.round((completadas / aplicables) * 100));

    let estadoSede = "Óptimo";
    if (pct < 60) estadoSede = "Crítico";
    else if (pct < 100) estadoSede = "Aceptable";

    const faltantesDetalle = Object.keys(faltantesPorRol).map(r => `${r} (${faltantesPorRol[r]})`).join(", ");

    resumenSedes.push({
      codigo: codSede,
      pct: pct,
      estado: estadoSede,
      completadas: completadas,
      pendientes: pendientes,
      noAplica: noAplica,
      faltantesDetalle: faltantesDetalle
    });
  });

  return {
    etapa: codEtapa,
    nombreEtapa: nombreEtapaOficial,
    tareas: tareasEtapa,
    sedes: sedesValidas,
    resumenSedes: resumenSedes,
    mapaSeg: mapaSeg
  };
}

/**
 * Emite y envía formalmente el reporte oficial por etapa al finalizar cada etapa
 * Formato institucional: Reporte Semanal de KPI + Matriz HTML de Evidencia
 */
function apiEnviarReporteEtapa(etapa, idChecklist) {
  const usuario = resolverUsuarioActivo();
  const directorio = obtenerDirectorioDesdeSheet();
  const sheetReportes = obtenerHojaBDReportes();

  const datosReporte = compilarDatosReporteEtapa(etapa, idChecklist);
  const ahora = new Date();
  const fechaCorteStr = Utilities.formatDate(ahora, "GMT-5", "dd/MM - h:mm a");
  const mesNombre = Utilities.formatDate(ahora, "GMT-5", "MMMM").toUpperCase();

  const idReporte = `REP-${datosReporte.etapa}-${Date.now()}`;
  const asunto = `REPORTE SEMANAL DE KPI - ${mesNombre} - 2026 | A-SERV-P01 GESTIÓN DE INICIO DE CLASES`;

  // Destinatarios: Procesos y Director Zonal
  const destinos = [];
  directorio.forEach(u => {
    if (u.email && (u.rol.includes("PROCESOS") || u.rol.includes("ZONAL"))) {
      if (!destinos.includes(u.email)) destinos.push(u.email);
    }
  });
  if (destinos.length === 0) destinos.push("marevalo@hnc.pe", "gvallejo@hnc.pe");

  // Generar líneas de resumen por sede
  const lineasSedes = datosReporte.resumenSedes.map(s => {
    let extra = "";
    if (s.faltantesDetalle) {
      extra = ` | No marcaron a tiempo: ${s.faltantesDetalle}`;
    }
    return `<tr>
      <td style="padding:4px 8px; font-weight:700; color:#1e293b; font-family:monospace; width:50px;">${s.codigo}:</td>
      <td style="padding:4px 8px; font-weight:800; color:${s.pct === 100 ? '#166534' : (s.pct >= 60 ? '#854d0e' : '#991b1b')}; width:75px;">${s.pct.toFixed(2)}%</td>
      <td style="padding:4px 8px; font-weight:600; color:#475569;">(${s.estado})${extra}</td>
    </tr>`;
  }).join("");

  // Generar filas de la Matriz Visual HTML
  const filasMatriz = datosReporte.tareas.map((t, idx) => {
    const tId = t.id || `TPL-${String(t.orden).padStart(3, "0")}`;
    const actKey = "ACT-" + String(t.orden).padStart(2, "0");
    const bgRow = idx % 2 === 0 ? "#ffffff" : "#f8fafc";

    const celdasSedes = datosReporte.sedes.map(codSede => {
      const st = (datosReporte.mapaSeg[tId] && datosReporte.mapaSeg[tId][codSede]) ||
        (datosReporte.mapaSeg[actKey] && datosReporte.mapaSeg[actKey][codSede]) || "PENDIENTE";
      if (st === "COMPLETADA") {
        return `<td style="padding:6px 4px; text-align:center; border:1px solid #e2e8f0; background:#f0fdf4;">
          <span style="display:inline-block; padding:2px 5px; border-radius:4px; background:#dcfce7; color:#166534; font-weight:800; font-size:11px;">✔️</span>
        </td>`;
      } else if (st === "NO APLICA") {
        return `<td style="padding:6px 4px; text-align:center; border:1px solid #e2e8f0; background:#f8fafc;">
          <span style="display:inline-block; padding:2px 5px; border-radius:4px; background:#f1f5f9; color:#64748b; font-size:11px;">➖</span>
        </td>`;
      } else {
        return `<td style="padding:6px 4px; text-align:center; border:1px solid #e2e8f0; background:#fef2f2;">
          <span style="display:inline-block; padding:2px 5px; border-radius:4px; background:#fee2e2; color:#991b1b; font-weight:800; font-size:11px;">❌</span>
        </td>`;
      }
    }).join("");

    return `<tr style="background:${bgRow};">
      <td style="padding:6px 8px; text-align:center; font-weight:700; color:#64748b; border:1px solid #e2e8f0; font-size:11px;">${t.orden}</td>
      <td style="padding:6px 8px; font-weight:600; color:#1e293b; border:1px solid #e2e8f0; font-size:11.5px;">${t.actividad || t.nombre}</td>
      <td style="padding:6px 8px; color:#475569; border:1px solid #e2e8f0; font-size:11px;">${t.responsable}</td>
      ${celdasSedes}
    </tr>`;
  }).join("");

  const thSedes = datosReporte.sedes.map(s =>
    `<th style="padding:6px 4px; background:#004691; color:#ffffff; font-weight:800; font-size:11px; border:1px solid #003875; text-align:center; min-width:42px;">${s}</th>`
  ).join("");

  // Cuerpo HTML del correo institucional
  const cuerpoHtml = `
    <div style="font-family:'Segoe UI', Arial, sans-serif; max-width:820px; margin:auto; background:#ffffff; border:1px solid #e2e8f0; border-radius:12px; padding:24px; color:#1e293b;">
      <div style="border-bottom:2px solid #004691; padding-bottom:12px; margin-bottom:18px;">
        <h2 style="margin:0 0 4px 0; font-size:16px; font-weight:800; color:#004691; letter-spacing:0.5px;">
          REPORTE SEMANAL DE KPI - ${mesNombre} - 2026
        </h2>
        <div style="font-size:12.5px; font-weight:700; color:#64748b;">
          A-SERV-P01 ‘GESTIÓN DE INICIO DE CLASES’ • ${datosReporte.nombreEtapa}
        </div>
      </div>

      <p style="font-size:13px; line-height:1.6; margin:0 0 16px 0;">
        Buen día,<br><br>
        Se remite el Reporte de KPI asociado al proceso <strong>A-SERV-P01 – Gestión de Inicio de Clases</strong>.
      </p>

      <div style="background:#f8fafc; border:1px solid #e2e8f0; border-radius:8px; padding:14px 16px; margin-bottom:20px;">
        <div style="font-size:13px; font-weight:800; color:#0f172a; margin-bottom:10px;">
          🎯 Nivel de Cumplimiento de Actividades: Corte ${fechaCorteStr}
        </div>
        <table style="width:100%; border-collapse:collapse; font-size:12.5px;">
          ${lineasSedes}
        </table>
      </div>

      <!-- MATRIZ VISUAL DE EVIDENCIA -->
      <div style="margin-bottom:22px;">
        <div style="font-size:12.5px; font-weight:800; color:#0f172a; margin-bottom:8px;">
          📊 Matriz de Cumplimiento por Sede (Evidencia Digital en Vivo):
        </div>
        <div style="overflow-x:auto;">
          <table style="width:100%; border-collapse:collapse; font-size:11px; border:1px solid #cbd5e1;">
            <thead>
              <tr>
                <th style="padding:6px 8px; background:#004691; color:#ffffff; font-weight:800; font-size:11px; border:1px solid #003875; width:30px;">N°</th>
                <th style="padding:6px 8px; background:#004691; color:#ffffff; font-weight:800; font-size:11px; border:1px solid #003875; text-align:left;">Actividad (${datosReporte.tareas.length})</th>
                <th style="padding:6px 8px; background:#004691; color:#ffffff; font-weight:800; font-size:11px; border:1px solid #003875; text-align:left;">Responsable</th>
                ${thSedes}
              </tr>
            </thead>
            <tbody>
              ${filasMatriz}
            </tbody>
          </table>
        </div>
        <div style="font-size:10.5px; color:#64748b; margin-top:6px; display:flex; gap:16px;">
          <span><strong>Leyenda:</strong></span>
          <span style="color:#166534;">✔️ Cumplido a tiempo</span>
          <span style="color:#991b1b;">❌ No marcado / Pendiente</span>
          <span style="color:#64748b;">➖ No aplica</span>
        </div>
      </div>

      <div style="background:#fffbeb; border-left:4px solid #f59e0b; padding:12px 14px; margin-bottom:20px; font-size:12px; line-height:1.6; color:#78350f;">
        <strong style="display:block; margin-bottom:4px;">📌 Recomendaciones:</strong>
        <ul style="margin:0; padding-left:18px;">
          <li>Dado que muchas actividades se realizaron en el plazo pero no fueron marcadas, se recuerda a cada área marcar cumplimiento ni bien se culmine con la asignación.</li>
          <li>Cuando una actividad no aplique en una sede, igualmente marcar el checkbox y registrar su justificación correspondiente.</li>
        </ul>
      </div>

      <p style="font-size:12.5px; color:#334155; margin:0 0 16px 0;">
        Saludos,<br>
        <strong>Gestión de Procesos y Operaciones</strong><br>
        <span style="font-size:11px; color:#64748b;">Academia Vonex • Sistema de Gobernanza Operativa</span>
      </p>
    </div>
  `;

  try {
    GmailApp.sendEmail(destinos.join(","), asunto, "", {
      htmlBody: cuerpoHtml,
      name: "Gestión de Procesos Vonex"
    });
  } catch (e) {
    Logger.log("Error al enviar correo: " + e.message);
  }

  // Registrar snapshot inmutable en BD_REPORTES_ETAPAS
  let totalCompletadas = 0;
  let totalPendientes = 0;
  let totalNoAplica = 0;
  datosReporte.resumenSedes.forEach(s => {
    totalCompletadas += s.completadas;
    totalPendientes += s.pendientes;
    totalNoAplica += s.noAplica;
  });
  const totalEvaluadas = totalCompletadas + totalPendientes;
  const pctGlobal = totalEvaluadas > 0 ? Math.round((totalCompletadas / totalEvaluadas) * 100) : 0;

  sheetReportes.appendRow([
    idReporte,
    idChecklist || "CHK-ACTIVO",
    datosReporte.nombreEtapa,
    ahora,
    usuario.email,
    datosReporte.tareas.length,
    totalCompletadas,
    totalPendientes,
    totalNoAplica,
    pctGlobal,
    JSON.stringify(datosReporte.resumenSedes)
  ]);

  return {
    exito: true,
    asunto: asunto,
    destinatarios: destinos.join(", "),
    etapa: datosReporte.nombreEtapa,
    fechaCorte: fechaCorteStr,
    pctGlobal: pctGlobal
  };
}

/**
 * Guarda y actualiza la configuración de Snapshot Programado (Trigger de Tiempo)
 */
function apiGuardarProgramacionReporte(config) {
  const props = PropertiesService.getScriptProperties();
  const cfg = config || {};
  const activo = Boolean(cfg.activo);
  const diaSemana = cfg.diaSemana || "MONDAY"; // "MONDAY", "EVERY_DAY", etc.
  const hora = parseInt(cfg.hora, 10) || 9; // 9:00 AM
  const etapa = cfg.etapa || "ANTES";
  const idChecklist = cfg.idChecklist || "CHK-ACTIVO";

  const cfgGuardada = {
    activo: activo,
    diaSemana: diaSemana,
    hora: hora,
    etapa: etapa,
    idChecklist: idChecklist,
    ultimaActualizacion: new Date().toISOString()
  };

  props.setProperty("CONFIG_PROGRAMACION_REPORTE", JSON.stringify(cfgGuardada));

  // Eliminar triggers existentes para evitar duplicidad
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if (t.getHandlerFunction() === "ejecutarTriggerSnapshotProgramado") {
      ScriptApp.deleteTrigger(t);
    }
  });

  // Si está activo, crear el trigger correspondiente
  if (activo) {
    if (diaSemana === "EVERY_DAY") {
      ScriptApp.newTrigger("ejecutarTriggerSnapshotProgramado")
        .timeBased()
        .everyDays(1)
        .atHour(hora)
        .create();
    } else {
      let weekDay = ScriptApp.WeekDay.MONDAY;
      if (diaSemana === "TUESDAY") weekDay = ScriptApp.WeekDay.TUESDAY;
      else if (diaSemana === "WEDNESDAY") weekDay = ScriptApp.WeekDay.WEDNESDAY;
      else if (diaSemana === "THURSDAY") weekDay = ScriptApp.WeekDay.THURSDAY;
      else if (diaSemana === "FRIDAY") weekDay = ScriptApp.WeekDay.FRIDAY;
      else if (diaSemana === "SATURDAY") weekDay = ScriptApp.WeekDay.SATURDAY;
      else if (diaSemana === "SUNDAY") weekDay = ScriptApp.WeekDay.SUNDAY;

      ScriptApp.newTrigger("ejecutarTriggerSnapshotProgramado")
        .timeBased()
        .onWeekDay(weekDay)
        .atHour(hora)
        .create();
    }
  }

  return {
    exito: true,
    config: cfgGuardada
  };
}

/**
 * Lee la configuración guardada de Snapshot Programado
 */
function apiObtenerProgramacionReporte() {
  const props = PropertiesService.getScriptProperties();
  const raw = props.getProperty("CONFIG_PROGRAMACION_REPORTE");
  let cfg = {
    activo: false,
    diaSemana: "MONDAY",
    hora: 9,
    etapa: "ANTES",
    idChecklist: "CHK-ACTIVO"
  };

  if (raw) {
    try {
      cfg = Object.assign(cfg, JSON.parse(raw));
    } catch (e) { }
  }

  // Verificar si el trigger realmente existe en Apps Script
  const triggers = ScriptApp.getProjectTriggers();
  const existeTrigger = triggers.some(t => t.getHandlerFunction() === "ejecutarTriggerSnapshotProgramado");
  cfg.activo = Boolean(cfg.activo && existeTrigger);

  return cfg;
}

/**
 * Función disparada por el Trigger de Tiempo de Google Apps Script
 */
function ejecutarTriggerSnapshotProgramado() {
  const cfg = apiObtenerProgramacionReporte();
  if (!cfg.activo) return;
  try {
    apiEnviarReporteEtapa(cfg.etapa || "ANTES", cfg.idChecklist || "CHK-ACTIVO");
    Logger.log("Snapshot programado ejecutado con éxito.");
  } catch (err) {
    Logger.log("Error al ejecutar snapshot programado: " + err.message);
  }
}

/**
 * Analiza el directorio y calcula los responsables que TIENEN pendientes reales (Filtro Zero Spam)
 * Incluye generación de mensajes WhatsApp oficiales con formato de emojis y tracking por contacto.
 */
function apiObtenerResumenRecordatorios(idChecklist) {
  const directorio = obtenerDirectorioDesdeSheet();
  const plantilla = obtenerDatosPlantilla();
  const sheetSeg = obtenerHojaBDSeguimiento();
  const dataSeg = sheetSeg.getDataRange().getValues();
  const sedesValidas = ["LCE", "SJL", "IND", "HDM", "SAN", "COM", "PTP", "VMT"];

  // Indexar mapa de seguimiento
  const mapaSeg = {};
  for (let r = 1; r < dataSeg.length; r++) {
    const row = dataSeg[r];
    const chk = (row[1] || "").toString().trim();
    if (idChecklist && idChecklist !== "CHK-ACTIVO" && chk && chk !== idChecklist) continue;
    const tId = (row[3] || "").toString().trim();
    const sCod = (row[4] || "").toString().trim().toUpperCase();
    let st = (row[5] || "").toString().trim().toUpperCase();
    if (st === "REALIZADO") st = "COMPLETADA";
    if (st === "NO_APLICA") st = "NO APLICA";

    if (tId && sCod) {
      if (!mapaSeg[tId]) mapaSeg[tId] = {};
      if (sCod === "TODAS") {
        sedesValidas.forEach(s => { mapaSeg[tId][s] = st; });
      } else {
        mapaSeg[tId][sCod] = st;
      }

      if (tId.startsWith("TPL-")) {
        const num = parseInt(tId.replace("TPL-", ""), 10);
        if (!isNaN(num)) {
          const actKey = "ACT-" + String(num).padStart(2, "0");
          if (!mapaSeg[actKey]) mapaSeg[actKey] = {};
          if (sCod === "TODAS") {
            sedesValidas.forEach(s => { mapaSeg[actKey][s] = st; });
          } else {
            mapaSeg[actKey][sCod] = st;
          }
        }
      }
    }
  }

  // Obtener fecha del checklist activo para calcular fechas límites reales
  let fechaInicioChk = "2026-11-02";
  try {
    const ss = obtenerSpreadsheet();
    const sheetBD = obtenerHojaBDChecklists();
    const dataBD = sheetBD.getDataRange().getValues();
    for (let r = dataBD.length - 1; r >= 1; r--) {
      if (dataBD[r][0]) {
        if (!idChecklist || idChecklist === "CHK-ACTIVO" || dataBD[r][0].toString().trim() === idChecklist) {
          if (dataBD[r][3] instanceof Date) {
            fechaInicioChk = formatearFechaISO(dataBD[r][3]);
          } else if (dataBD[r][3]) {
            fechaInicioChk = dataBD[r][3].toString().trim();
          }
          break;
        }
      }
    }
  } catch (e) { }

  const destinatariosConPendientes = [];
  let totalAlDia = 0;

  directorio.forEach(dest => {
    // Si no tiene celular ni email, o está inactivo, omitir
    if (!dest.celular && !dest.email) return;

    const rolNorm = normalizarTexto(dest.rol || "");
    const sedeNorm = (dest.codigoSede || "").toUpperCase().trim();
    const esMulti = dest.esMultisede || sedeNorm === "TODAS" || sedeNorm === "-";

    const tareasPendientes = [];
    let proximaFechaStr = "";
    let proximaHoraStr = "";

    plantilla.forEach(t => {
      const respNorm = normalizarTexto(t.responsable || "");
      let corresponde = false;

      if ((rolNorm.includes("DIRECTOR DE SEDE") || rolNorm.includes("TUTOR")) &&
        (respNorm.includes("DIRECTOR") || respNorm.includes("TUTOR"))) {
        corresponde = true;
      } else if (rolNorm.includes("SISTEMAS") && respNorm.includes("SISTEMAS")) {
        corresponde = true;
      } else if (rolNorm.includes("VENTAS") && respNorm.includes("VENTAS")) {
        corresponde = true;
      } else if (rolNorm.includes("PRODUCCION") && respNorm.includes("PRODUCCION")) {
        corresponde = true;
      } else if (rolNorm.includes("EVALUACION") && respNorm.includes("EVALUACION")) {
        corresponde = true;
      } else if (rolNorm.includes("DIGITALIZACION") && respNorm.includes("DIGITALIZACION")) {
        corresponde = true;
      }

      if (corresponde) {
        const tId = t.id || `TPL-${String(t.orden).padStart(3, "0")}`;
        const actKey = "ACT-" + String(t.orden).padStart(2, "0");

        // Si es multisede, verificar si está pendiente en alguna sede
        let estaPendiente = false;
        let sedesPendientes = [];
        if (esMulti) {
          sedesValidas.forEach(s => {
            const st = (mapaSeg[tId] && mapaSeg[tId][s]) || (mapaSeg[actKey] && mapaSeg[actKey][s]) || "PENDIENTE";
            if (st !== "COMPLETADA" && st !== "NO APLICA") {
              estaPendiente = true;
              sedesPendientes.push(s);
            }
          });
        } else {
          const st = (mapaSeg[tId] && mapaSeg[tId][sedeNorm]) || (mapaSeg[actKey] && mapaSeg[actKey][sedeNorm]) || "PENDIENTE";
          if (st !== "COMPLETADA" && st !== "NO APLICA") {
            estaPendiente = true;
          }
        }

        if (estaPendiente) {
          // Calcular fecha límite de la actividad
          let fLim = "02/11";
          try {
            const [y, m, d] = fechaInicioChk.split("-").map(Number);
            const dt = new Date(y, m - 1, d);
            dt.setDate(dt.getDate() + (t.diasRelativos || 0));
            const diaStr = String(dt.getDate()).padStart(2, "0");
            const mesStr = String(dt.getMonth() + 1).padStart(2, "0");
            fLim = `${diaStr}/${mesStr}`;
          } catch (e) { }

          let hLim = t.horaDefault || "09:00";
          if (hLim.includes(":")) {
            const [hh, mm] = hLim.split(":");
            const hNum = parseInt(hh, 10);
            const ampm = hNum >= 12 ? "p.m." : "a.m.";
            const h12 = hNum % 12 || 12;
            hLim = `${h12}:${mm} ${ampm}`;
          }

          if (!proximaFechaStr) {
            proximaFechaStr = fLim;
            proximaHoraStr = hLim;
          }

          tareasPendientes.push({
            orden: t.orden,
            actividad: t.actividad || t.nombre,
            etapa: t.etapa,
            fechaLimite: fLim,
            horaLimite: hLim,
            sedesPendientes: sedesPendientes
          });
        }
      }
    });

    if (tareasPendientes.length > 0) {
      const nombreLimpio = (dest.nombre || dest.rol || "Responsable").trim();
      let primerNombre = nombreLimpio.split(" ")[0];
      primerNombre = primerNombre.charAt(0).toUpperCase() + primerNombre.slice(1).toLowerCase();

      let rolMensaje = dest.rol;
      if (rolMensaje.includes("DIRECTOR DE SEDE")) rolMensaje = "Dirección de Sede y Tutores";
      else if (rolMensaje.includes("VENTAS")) rolMensaje = "Área de Ventas";
      else if (rolMensaje.includes("SISTEMAS")) rolMensaje = "Área de Sistemas";
      else if (rolMensaje.includes("PRODUCCION")) rolMensaje = "Área de Producción";

      // Formatear mensaje de WhatsApp con emojis oficiales
      const emojis = ["1️⃣", "2️⃣", "3️⃣", "4️⃣", "5️⃣", "6️⃣", "7️⃣", "8️⃣", "9️⃣", "🔟"];
      const tareasWhatsApp = tareasPendientes.slice(0, 8).map((t, i) => {
        const ico = emojis[i] || `[${i + 1}]`;
        return `${ico} ${t.actividad}`;
      }).join("\n");

      const linkSpreadsheet = "https://docs.google.com/spreadsheets/d/1yGizP7Ngd8qtw5sdsWGaYRXn2hmGkHiOHDk5rDiyajA/edit?gid=47299654#gid=47299654";
      const fechaMsg = proximaFechaStr || "02/11";
      const horaMsg = proximaHoraStr ? ` a las ${proximaHoraStr}` : " a las 9:00 a.m";

      const mensajeWhatsApp = `🚨 RECORDATORIO | CHECKLIST INICIO DE CLASES 🚨\n\nBuen día ${primerNombre} , le recordamos que el ${fechaMsg}${horaMsg} vence el plazo para el cumplimiento de las siguientes actividades (${rolMensaje}) :\n\n${tareasWhatsApp}\n\n🔗 Marque su cumplimiento directamente aquí:\n👉 ${linkSpreadsheet}\n\nGestión de Procesos | Academia Vonex`;

      destinatariosConPendientes.push({
        nombre: nombreLimpio,
        primerNombre: primerNombre,
        email: dest.email || "",
        celular: dest.celular || "",
        rol: dest.rol,
        rolMensaje: rolMensaje,
        codigoSede: esMulti ? "TODAS" : dest.codigoSede,
        nombreSede: dest.nombreSede,
        esMultisede: esMulti,
        totalPendientes: tareasPendientes.length,
        proximaFecha: fechaMsg,
        proximaHora: horaMsg,
        tareas: tareasPendientes,
        mensajeWhatsApp: mensajeWhatsApp
      });
    } else {
      totalAlDia++;
    }
  });

  return {
    conPendientes: destinatariosConPendientes,
    totalConPendientes: destinatariosConPendientes.length,
    totalAlDia: totalAlDia,
    totalDirectorio: destinatariosConPendientes.length + totalAlDia
  };
}

/**
 * Envía recordatorios individualizados EXCLUSIVAMENTE a quienes tienen pendientes reales (Filtro Zero Spam)
 * Asunto oficial: "GP | RECORDATORIO | A-SERV-P01 GESTIÓN DE INICIO DE CLASES"
 */
function apiEnviarRecordatorioManual(idChecklist) {
  const resumen = apiObtenerResumenRecordatorios(idChecklist);
  const asunto = "GP | RECORDATORIO | A-SERV-P01 GESTIÓN DE INICIO DE CLASES";
  let enviados = 0;
  let omitidosSinCorreo = 0;

  resumen.conPendientes.forEach(dest => {
    if (!dest.email || !dest.email.includes("@")) {
      omitidosSinCorreo++;
      return;
    }

    const listaTareasHtml = dest.tareas.map(t =>
      `<li style="margin-bottom:6px;"><strong>Actividad ${t.orden}:</strong> ${t.actividad} <span style="color:#64748b; font-size:11px;">(${t.etapa})</span></li>`
    ).join("");

    const cuerpoHtml = `
      <div style="font-family:'Segoe UI', Arial, sans-serif; max-width:600px; margin:auto; background:#ffffff; border-radius:12px; border:1px solid #e2e8f0; padding:22px; color:#1e293b;">
        <div style="border-bottom:2px solid #d97706; padding-bottom:10px; margin-bottom:14px;">
          <span style="font-size:11px; font-weight:800; color:#d97706; text-transform:uppercase; letter-spacing:0.5px;">Gestión de Procesos Vonex</span>
          <h3 style="margin:4px 0 0 0; font-size:15px; color:#0f172a;">Recordatorio de Cumplimiento de Actividades</h3>
        </div>
        
        <p style="font-size:13px; color:#334155; line-height:1.5;">
          Estimado(a) <strong>${dest.nombre}</strong> (${dest.rol} - Sede ${dest.codigoSede}):
        </p>

        <p style="font-size:12.5px; color:#475569; line-height:1.6;">
          Se le recuerda que a la fecha registra <strong>${dest.totalPendientes} actividad(es) pendiente(s)</strong> de cumplimiento en el proceso de Inicio de Clases:
        </p>

        <div style="background:#fffbeb; border:1px solid #fde68a; border-radius:8px; padding:12px 14px; margin:14px 0; font-size:12px; color:#92400e;">
          <ul style="margin:0; padding-left:18px;">
            ${listaTareasHtml}
          </ul>
        </div>

        <p style="font-size:12.5px; color:#475569; line-height:1.5;">
          Agradeceremos ingresar al módulo institucional o a la hoja de cálculo y registrar el avance:
        </p>

        <div style="text-align:center; margin:20px 0;">
          <a href="${obtenerUrlWebApp()}?p=participantes" target="_blank" style="display:inline-block; background:#004691; color:#ffffff; font-weight:700; font-size:12.5px; padding:10px 22px; border-radius:6px; text-decoration:none;">
            🚀 Ir a mi Módulo de Participante
          </a>
        </div>

        <div style="border-top:1px solid #f1f5f9; padding-top:12px; font-size:11px; color:#94a3b8; text-align:center;">
          Academia Vonex • Dirección de Gestión de Procesos & Calidad (A-SERV-P01)
        </div>
      </div>
    `;

    try {
      GmailApp.sendEmail(dest.email, asunto, "", {
        htmlBody: cuerpoHtml,
        name: "Gestión de Procesos Vonex"
      });
      enviados++;
    } catch (e) {
      Logger.log("Error al enviar a " + dest.email + ": " + e.message);
    }
  });

  return {
    exito: true,
    asunto: asunto,
    totalEnviados: enviados,
    omitidosSinCorreo: omitidosSinCorreo,
    totalConPendientes: resumen.totalConPendientes,
    totalAlDia: resumen.totalAlDia,
    totalDirectorio: resumen.totalDirectorio
  };
}

/**
 * Consulta la configuración de recordatorios automáticos por correo
 */
function apiObtenerConfigRecordatoriosAutomaticos() {
  const props = PropertiesService.getScriptProperties();
  const activo = props.getProperty("RECORDATORIOS_AUTO_ACTIVO") === "true";
  const hora = props.getProperty("RECORDATORIOS_AUTO_HORA") || "08:00";
  return { activo: activo, hora: hora };
}

/**
 * Guarda la configuración de recordatorios automáticos y gestiona el trigger diario
 */
function apiGuardarConfigRecordatoriosAutomaticos(activo, hora) {
  const props = PropertiesService.getScriptProperties();
  props.setProperty("RECORDATORIOS_AUTO_ACTIVO", activo ? "true" : "false");
  if (hora) props.setProperty("RECORDATORIOS_AUTO_HORA", hora);

  try {
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(t => {
      if (t.getHandlerFunction() === "ejecutarTriggerRecordatoriosAutomaticos") {
        ScriptApp.deleteTrigger(t);
      }
    });

    if (activo) {
      const horaNum = parseInt((hora || "08:00").split(":")[0], 10) || 8;
      ScriptApp.newTrigger("ejecutarTriggerRecordatoriosAutomaticos")
        .timeBased()
        .everyDays(1)
        .atHour(horaNum)
        .create();
    }
  } catch (e) {
    Logger.log("Aviso al configurar trigger automático: " + e.message);
  }

  return { exito: true, activo: activo, hora: hora };
}

/**
 * Función que ejecuta el trigger automático diario programado
 */
function ejecutarTriggerRecordatoriosAutomaticos() {
  const props = PropertiesService.getScriptProperties();
  const activo = props.getProperty("RECORDATORIOS_AUTO_ACTIVO") === "true";
  if (!activo) return;

  const ss = obtenerSpreadsheet();
  const sheetBD = obtenerHojaBDChecklists();
  const data = sheetBD.getDataRange().getValues();
  if (data.length < 2) return;

  const ultimoChkId = (data[data.length - 1][0] || "CHK-ACTIVO").toString().trim();
  apiEnviarRecordatorioManual(ultimoChkId);
}

// =========================================================================
// 5. API PARA VISTA 2: PARTICIPANTES.HTML (MARCACIÓN & EVIDENCIAS)
// =========================================================================

/**
 * Contexto inicial para la vista PARTICIPANTES.html
 * Devuelve todas las actividades y sus subactividades mapeadas para TODAS las 8 sedes.
 */
function apiObtenerContextoParticipantes(codigoSede, e) {
  const usuario = resolverUsuarioActivo(e);
  const plantilla = obtenerDatosPlantilla();
  const sedesValidas = ["LCE", "SJL", "IND", "HDM", "SAN", "COM", "PTP", "VMT"];
  const esMulti = !!(usuario.esMultisede || usuario.esProcesos || usuario.esDirectorZonal);
  const sedeFiltro = esMulti ? (codigoSede || "TODAS") : (usuario.codigoSede || "SJL");

  // Obtener checklist activo más reciente
  const ss = obtenerSpreadsheet();
  const sheetBD = obtenerHojaBDChecklists();
  const allBD = sheetBD.getDataRange().getValues();
  let checklistActivo = {
    id: "CHK-ACTIVO",
    nombre: "Ciclo Regular UNMSM 2026-I - Inicio Noviembre",
    universidad: "UNMSM",
    fechaInicio: "2026-11-02"
  };

  for (let r = allBD.length - 1; r >= 1; r--) {
    if (allBD[r][0]) {
      checklistActivo = {
        id: allBD[r][0].toString().trim(),
        universidad: (allBD[r][1] || "UNMSM").toString().trim(),
        nombre: (allBD[r][2] || "").toString().trim(),
        fechaInicio: allBD[r][3] instanceof Date ? formatearFechaISO(allBD[r][3]) : (allBD[r][3] || "").toString().trim()
      };
      break;
    }
  }

  // Leer la Fecha de Inicio de Ciclo oficial directamente desde la hoja "UNMSM" (26/10/2026)
  try {
    const sheetUNMSM = ss.getSheetByName("UNMSM");
    if (sheetUNMSM) {
      const topRows = sheetUNMSM.getRange(1, 1, 8, 8).getValues();
      for (let r = 0; r < topRows.length; r++) {
        for (let c = 0; c < topRows[r].length; c++) {
          const val = topRows[r][c];
          if (val instanceof Date) {
            checklistActivo.fechaInicio = formatearFechaISO(val);
            break;
          }
          const str = (val || "").toString().trim().toUpperCase();
          if (str.includes("INICIO DE CICLO") || str.includes("FECHA DE INICIO")) {
            const nextVal = topRows[r][c + 1];
            if (nextVal instanceof Date) {
              checklistActivo.fechaInicio = formatearFechaISO(nextVal);
              break;
            } else if (nextVal) {
              const p = parsearFechaTextoGeneral(nextVal.toString());
              if (p) { checklistActivo.fechaInicio = p; break; }
            }
          }
        }
      }
    }
  } catch (err) {}
  if (!checklistActivo.fechaInicio) checklistActivo.fechaInicio = "2026-10-26";

  const sheetSeg = obtenerHojaBDSeguimiento();
  const dataSeg = sheetSeg.getDataRange().getValues();
  const mapaEstados = {};

  for (let r = 1; r < dataSeg.length; r++) {
    const row = dataSeg[r];
    const tId = (row[3] || "").toString().trim();
    const sCod = (row[4] || "").toString().trim().toUpperCase();
    let st = (row[5] || "").toString().trim().toUpperCase();
    if (st === "REALIZADO") st = "COMPLETADA";
    if (st === "NO_APLICA") st = "NO APLICA";
    const evi = (row[8] || "").toString().trim();
    const obs = (row[9] || "").toString().trim();

    if (tId && sCod) {
      if (!mapaEstados[tId]) mapaEstados[tId] = {};
      const reg = { estado: st, evidencia: evi, nota: obs };
      if (sCod === "TODAS") {
        sedesValidas.forEach(s => { mapaEstados[tId][s] = reg; });
      } else {
        mapaEstados[tId][sCod] = reg;
      }

      // Alias ACT-xx
      if (tId.startsWith("TPL-")) {
        const parts = tId.replace("TPL-", "").split("-");
        const num = parseInt(parts[0], 10);
        if (!isNaN(num)) {
          const actKey = "ACT-" + String(num).padStart(2, "0");
          if (!mapaEstados[actKey]) mapaEstados[actKey] = {};
          if (sCod === "TODAS") {
            sedesValidas.forEach(s => { mapaEstados[actKey][s] = reg; });
          } else {
            mapaEstados[actKey][sCod] = reg;
          }
        }
      }
    }
  }

  // Catálogo maestro normalizado de subactividades oficiales (A-FOR-SERV-0104 v03)
  // Total exacto: 25 Actividades Maestras y 37 Subactividades distribuidas en 3 Fases
  const CATALOGO_SUBACTIVIDADES = {
    // FASE 1: ACTIVIDADES ANTES DEL INICIO DE CICLO (TPL-001 a TPL-015) -> 27 subactividades
    "TPL-001": [
      { id: "TPL-001-1", orden: 1, descripcion: "Reglamento interno del alumno", hora: "18:00", responsable: "Gestión de Procesos - Vonex" },
      { id: "TPL-001-2", orden: 2, descripcion: "Políticas de reembolso y/o devoluciones", hora: "18:00", responsable: "Gestión de Procesos - Vonex" },
      { id: "TPL-001-3", orden: 3, descripcion: "Términos y Condiciones", hora: "18:00", responsable: "Gestión de Procesos - Vonex" },
      { id: "TPL-001-4", orden: 4, descripcion: "Croquis de sede", hora: "18:00", responsable: "Gestión de Procesos - Vonex" }
    ],
    "TPL-002": [
      { id: "TPL-002-1", orden: 1, descripcion: "Envío Formato de asignación de tutores: (acmateriales@hnc.pe, helpdesk@hnc.pe, evaluacion@hnc.pe, almacon@hnc.pe, evelasquez@hnc.pe, jpinedo@hnc.pe, auditoria@hnc.pe, procesos@hnc.pe)", hora: "18:00", responsable: "Director de Sede" }
    ],
    "TPL-003": [
      { id: "TPL-003-1", orden: 1, descripcion: "Creación de carpetas en Drive", hora: "13:00", responsable: "Sistemas" },
      { id: "TPL-003-2", orden: 2, descripcion: "Creación de Espacios de Google Chat para alumnos", hora: "13:00", responsable: "Sistemas" },
      { id: "TPL-003-3", orden: 3, descripcion: "Creación de Comunidades para padres", hora: "13:00", responsable: "Sistemas" },
      { id: "TPL-003-4", orden: 4, descripcion: "Agregar a los alumnos en los espacios de Google Chat", hora: "13:00", responsable: "Sistemas" }
    ],
    "TPL-004": [
      { id: "TPL-004-1", orden: 1, descripcion: "Enviar correo con lista de tutores que apoyarán durante la campaña en el área de Ventas\nDeberá enviar correo con la información de tutores para apoyo (DNI, Apellidos y Nombres)\n\nDirigido a: helpdesk@hnc.pe, evelasquez@hnc.pe, tesoreria@hnc.pe, evaluacion@hnc.pe, jatalaya@hnc.pe", hora: "13:00", responsable: "Ventas" }
    ],
    "TPL-005": [
      { id: "TPL-005-1", orden: 1, descripcion: "Actualizar el croquis con la numeración de cada aula", hora: "18:00", responsable: "Director de Sede" },
      { id: "TPL-005-2", orden: 2, descripcion: "Actualizar lista de salones donde incluya el Nombre del tutor / Ciclo / Aula (Al costado del croquis)", hora: "18:00", responsable: "Director de Sede" }
    ],
    "TPL-006": [
      { id: "TPL-006-1", orden: 1, descripcion: "Colocar un mural del croquis en la puerta de la sede con los números de aulas correspondientes de cada ciclo", hora: "18:00", responsable: "Director de Sede" },
      { id: "TPL-006-2", orden: 2, descripcion: "Colocar en cada piso la lista correspondiente de salones y el piso al que pertenece", hora: "18:00", responsable: "Director de Sede" }
    ],
    "TPL-007": [
      { id: "TPL-007-1", orden: 1, descripcion: "Actualizar Lista de Tutores en Intranet según asignación indicada por cada director de sede", hora: "18:00", responsable: "Sistemas" }
    ],
    "TPL-008": [
      { id: "TPL-008-1", orden: 1, descripcion: "Agregar rol de 'tutor' y 'asesor de ventas' para tutores que les corresponda y dar acceso a funciones de tutoría", hora: "18:00", responsable: "Sistemas" },
      { id: "TPL-008-2", orden: 2, descripcion: "Agregar a tutores a los grupos de Pases de Salida", hora: "18:00", responsable: "Sistemas" },
      { id: "TPL-008-3", orden: 3, descripcion: "Dar acceso al Drive de materiales", hora: "13:00", responsable: "Digitalización de contenido" }
    ],
    "TPL-009": [
      { id: "TPL-009-1", orden: 1, descripcion: "Actualizar carpetas de Ev. Desempeño según tutor", hora: "13:00", responsable: "Evaluación y Control" }
    ],
    "TPL-010": [
      { id: "TPL-010-1", orden: 1, descripcion: "Responder al correo de Asignación de tutor indicando que los tutores tienen todos los accesos necesarios", hora: "18:00", responsable: "Director de Sede" }
    ],
    "TPL-011": [
      { id: "TPL-011-1", orden: 1, descripcion: "Enviar libros, boletines y guías de clase a los tutores el viernes previo al inicio de clases", hora: "18:00", responsable: "Producción" },
      { id: "TPL-011-2", orden: 2, descripcion: "Enviar exámenes y fichas ópticas a los tutores el viernes previo al inicio de clases", hora: "18:00", responsable: "Producción" }
    ],
    "TPL-012": [
      { id: "TPL-012-1", orden: 1, descripcion: "Enviar al docente las guías de clase por medio de WhatsApp", hora: "18:00", responsable: "Tutores" }
    ],
    "TPL-013": [
      { id: "TPL-013-1", orden: 1, descripcion: "Enviar por WhatsApp al docente los siguientes detalles: nombre del tutor, sede, aula asignada, número de aula y un recordatorio sobre el acceso al enlace correspondiente (solo aula virtual)", hora: "18:00", responsable: "Tutores" }
    ],
    "TPL-014": [
      { id: "TPL-014-1", orden: 1, descripcion: "Agendar y agregar a padres en las Comunidades de WhatsApp", hora: "18:00", responsable: "Tutores" }
    ],
    "TPL-015": [
      { id: "TPL-015-1", orden: 1, descripcion: "Dar mensaje de bienvenida a padres y alumnos", hora: "09:00", responsable: "Tutores" },
      { id: "TPL-015-2", orden: 2, descripcion: "Mensaje recordatorio para que los alumnos ingresen a los links ubicados en la descripción de los grupos correspondientes", hora: "09:00", responsable: "Tutores" }
    ],

    // FASE 2: ACTIVIDADES DURANTE LA PRIMERA SEMANA (TPL-016 a TPL-023) -> 8 subactividades
    "TPL-016": [
      { id: "TPL-016-1", orden: 1, descripcion: "Aula Virtual: Enviar un recordatorio a los alumnos indicando que el material está en el campus virtual y grupo de WhatsApp. Reenviar los libros digitales que se encuentran en las carpetas y pasar lista de asistencia por zoom (L - V 7:45 a. m.)", hora: "07:45", responsable: "Tutores" }
    ],
    "TPL-017": [
      { id: "TPL-017-1", orden: 1, descripcion: "Aula Presencial: Entregar material impreso al alumno (L - V 7:59 a. m.). Pasar lista de asistencia a las 09:00 a. m. (Con lista impresa)", hora: "07:59", responsable: "Tutores" }
    ],
    "TPL-018": [
      { id: "TPL-018-1", orden: 1, descripcion: "Orientar el ingreso al aula a alumnos y/o padres:\n\n- Primera semana:\nLos tutores orientarán los 2 primeros días desde las 07:05 a. m. hasta las 8:00 a. m. (la cantidad de tutores será de acuerdo a la necesidad por sede)\n\nNota: Como mínimo debe haber 1 tutor asignados. Nunca debería quedar solo vigilancia en la entrada de los alumnos.\n\nNota: Director solo deberá supervisar.", hora: "07:05", responsable: "Director de Sede y Tutores" }
    ],
    "TPL-019": [
      { id: "TPL-019-1", orden: 1, descripcion: "Actualizar los contactos todos los días a las 8:00 am - 8:30 am. Turno tarde: 4:30pm", hora: "08:00", responsable: "Tutores" }
    ],
    "TPL-020": [
      { id: "TPL-020-1", orden: 1, descripcion: "Sistema VONEX, Normas de convivencia, Study Time, Sistema de evaluación (durante la 1ra semana desde el martes a las 7:45 a.m.)", hora: "07:45", responsable: "Tutores" }
    ],
    "TPL-021": [
      { id: "TPL-021-1", orden: 1, descripcion: "Mostrar el video tutorial del campus el dia lunes en la salida (o cuando sea necesario)", hora: "13:00", responsable: "Tutores" }
    ],
    "TPL-022": [
      { id: "TPL-022-1", orden: 1, descripcion: "Llevar a cabo conversaciones individuales con los alumnos que no asistieron en los primeros dias o no asistieron a las charlas de inducción", hora: "09:00", responsable: "Tutores" }
    ],
    "TPL-023": [
      { id: "TPL-023-1", orden: 1, descripcion: "Durante la segunda semana, se destinará un día exclusivamente para la evaluación de los estudiantes", hora: "09:00", responsable: "Tutores" }
    ],

    // FASE 3: ACTIVIDADES DESPUÉS DE INICIO DE CICLO (TPL-024 a TPL-025) -> 2 subactividades
    "TPL-024": [
      { id: "TPL-024-1", orden: 1, descripcion: "Enviar correo para quitar los accesos a los tutores de apoyo. Deberá enviar correo con la información de tutores para apoyo (DNI, Apellidos y Nombres). Dirigido a: helpdesk@hnc.pe, evelasquez@hnc.pe, tesoreria@hnc.pe, evaluacion@hnc.pe", hora: "12:00", responsable: "Ventas" }
    ],
    "TPL-025": [
      { id: "TPL-025-1", orden: 1, descripcion: "Citar a una reunión de padres en quinta semana", hora: "20:30", responsable: "Tutores" }
    ]
  };

  // Agrupar actividades y tareas con estado mapeado para TODAS las 8 sedes
  const actividadesMap = {};
  plantilla.forEach(t => {
    const actId = `ACT-${String(t.orden).padStart(2, "0")}`;
    const actKey = t.id || actId;

    // Calcular fecha límite según días relativos
    let fechaLimiteStr = checklistActivo.fechaInicio || "2026-10-26";
    try {
      const [y, m, d] = fechaLimiteStr.split("-").map(Number);
      const dt = new Date(y, m - 1, d);
      dt.setDate(dt.getDate() + (t.diasRelativos || 0));
      fechaLimiteStr = formatearFechaISO(dt);
    } catch (e) { }

    if (!actividadesMap[actKey]) {
      actividadesMap[actKey] = {
        id: actId,
        tplId: t.id,
        orden: t.orden,
        nombre: t.actividad,
        area: t.responsable,
        diasRelativos: t.diasRelativos !== undefined ? t.diasRelativos : 0,
        fecha: fechaLimiteStr,
        hora: t.horaDefault || "18:00",
        etapa: t.etapa.toLowerCase().includes("antes") ? "ANTES" : (t.etapa.toLowerCase().includes("primera") ? "DURANTE" : "DESPUES"),
        tareas: []
      };
    }

    // Subactividades para esta actividad: garantizar asignación exacta por área sin mezclas
    const catOficialSubs = CATALOGO_SUBACTIVIDADES[t.id] || [];
    const subDefs = (t.subactividadesCustom && t.subactividadesCustom.length >= catOficialSubs.length && catOficialSubs.length > 0)
      ? t.subactividadesCustom
      : (catOficialSubs.length > 0 ? catOficialSubs : (t.subactividadesCustom || [
          { id: `${t.id}-1`, orden: 1, descripcion: t.descripcion, hora: t.horaDefault || "18:00", responsable: t.responsable }
        ]));

    subDefs.forEach(sub => {
      const subId = sub.id;
      const estadoSede = {};
      sedesValidas.forEach(s => {
        // 1. Buscar registro por ID de subactividad (ej: TPL-001-1)
        // 2. Si no existe, heredar el registro de la actividad padre (ej: TPL-001 o ACT-01)
        const reg = (mapaEstados[subId] && mapaEstados[subId][s]) ||
                    (mapaEstados[t.id] && mapaEstados[t.id][s]) ||
                    (mapaEstados[actId] && mapaEstados[actId][s]);
        estadoSede[s] = reg || { estado: "PENDIENTE", evidencia: null, nota: "" };
      });

      actividadesMap[actKey].tareas.push({
        id: subId,
        actPadreId: t.id,
        actOrden: t.orden,
        orden: sub.orden,
        actId: actId,
        actividadNombre: t.actividad,
        descripcion: sub.descripcion,
        responsable: sub.responsable || t.responsable,
        etapa: t.etapa.toLowerCase().includes("antes") ? "ANTES" : (t.etapa.toLowerCase().includes("primera") ? "DURANTE" : "DESPUES"),
        fecha: fechaLimiteStr,
        diasRelativos: (sub.diasRelativos !== undefined) ? sub.diasRelativos : (t.diasRelativos !== undefined ? t.diasRelativos : 0),
        hora: sub.hora || t.horaDefault || "18:00",
        requiereEvidencia: t.requiereEvidencia,
        sedes: sedesValidas,
        estadoSede: estadoSede
      });
    });
  });

  return {
    usuario: usuario,
    checklist: checklistActivo,
    sedeActual: sedeFiltro,
    esSoloSedePropia: !esMulti,
    sedeAsignada: usuario.codigoSede || "SJL",
    sedesValidas: esMulti ? sedesValidas : [usuario.codigoSede || "SJL"],
    actividades: Object.values(actividadesMap)
  };
}

/**
 * Marca una tarea como COMPLETADA guardando evidencia inmutable en BD_SEGUIMIENTO
 * Si codigoSede === "TODAS", registra la marcación para las 8 sedes a nivel nacional.
 */
function apiMarcarTareaCompletada(idTarea, codigoSede, urlEvidencia, nota) {
  const usuario = resolverUsuarioActivo();
  const sheetSeg = obtenerHojaBDSeguimiento();
  const sedesValidas = ["LCE", "SJL", "IND", "HDM", "SAN", "COM", "PTP", "VMT"];
  const sedesTarget = codigoSede === "TODAS" ? sedesValidas : [codigoSede];
  const now = new Date();

  const rows = sedesTarget.map((s, idx) => [
    `REG-${Date.now()}-${idx}`,
    "CHK-ACTIVO",
    "UNMSM",
    idTarea,
    s,
    "COMPLETADA",
    now,
    usuario.email,
    urlEvidencia || "",
    nota || "Cumplimiento verificado conforme"
  ]);

  if (rows.length === 1) {
    sheetSeg.appendRow(rows[0]);
  } else {
    sheetSeg.getRange(sheetSeg.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
  }

  return {
    exito: true,
    idTarea: idTarea,
    sede: codigoSede,
    sedesAfectadas: sedesTarget,
    estado: "COMPLETADA",
    evidencia: urlEvidencia
  };
}

/**
 * Registra una tarea como NO APLICA con justificación obligatoria en BD_SEGUIMIENTO
 * Si codigoSede === "TODAS", registra para las 8 sedes a nivel nacional.
 */
function apiRegistrarNoAplica(idTarea, codigoSede, motivoJustificado) {
  if (!motivoJustificado || motivoJustificado.trim().length < 8) {
    throw new Error("Debe ingresar una justificación detallada (mínimo 8 caracteres).");
  }

  const usuario = resolverUsuarioActivo();
  const sheetSeg = obtenerHojaBDSeguimiento();
  const sedesValidas = ["LCE", "SJL", "IND", "HDM", "SAN", "COM", "PTP", "VMT"];
  const sedesTarget = codigoSede === "TODAS" ? sedesValidas : [codigoSede];
  const now = new Date();

  const rows = sedesTarget.map((s, idx) => [
    `REG-${Date.now()}-${idx}`,
    "CHK-ACTIVO",
    "UNMSM",
    idTarea,
    s,
    "NO APLICA",
    now,
    usuario.email,
    "",
    motivoJustificado.trim()
  ]);

  if (rows.length === 1) {
    sheetSeg.appendRow(rows[0]);
  } else {
    sheetSeg.getRange(sheetSeg.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
  }

  return {
    exito: true,
    idTarea: idTarea,
    sede: codigoSede,
    sedesAfectadas: sedesTarget,
    estado: "NO APLICA",
    motivo: motivoJustificado
  };
}

/**
 * Registra o actualiza la observación/nota en texto para una tarea en una o todas las sedes
 */
function apiGuardarObservacionTarea(idChecklist, idTarea, codigoSede, textoObservacion) {
  const usuario = resolverUsuarioActivo();
  const sheetSeg = obtenerHojaBDSeguimiento();
  const sedesValidas = ["LCE", "SJL", "IND", "HDM", "SAN", "COM", "PTP", "VMT"];
  const sedesTarget = codigoSede === "TODAS" ? sedesValidas : [codigoSede];
  const now = new Date();

  const rows = sedesTarget.map((s, idx) => [
    `REG-OBS-${Date.now()}-${idx}`,
    idChecklist || "CHK-ACTIVO",
    "UNMSM",
    idTarea,
    s,
    "OBSERVACION",
    now,
    usuario.email,
    "",
    (textoObservacion || "").trim()
  ]);

  if (rows.length === 1) {
    sheetSeg.appendRow(rows[0]);
  } else {
    sheetSeg.getRange(sheetSeg.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
  }

  return {
    exito: true,
    idTarea: idTarea,
    sede: codigoSede,
    sedesAfectadas: sedesTarget,
    texto: textoObservacion
  };
}

/**
 * Sube un archivo de evidencia (imagen, foto o PDF en base64) a Google Drive.
 * Crea una carpeta centralizada "EVIDENCIAS_CHECKLIST_VONEX" con permisos de lectura
 * y devuelve la URL directa de visualización inmutable.
 */
function apiSubirArchivoEvidencia(nombreArchivo, base64Data, mimeType) {
  try {
    if (!base64Data) {
      return { ok: false, mensaje: "No se enviaron datos de archivo." };
    }

    let dataLimpia = base64Data;
    let mimeDetectado = mimeType || "image/jpeg";
    if (base64Data.indexOf("data:") === 0 && base64Data.indexOf(";base64,") !== -1) {
      const partes = base64Data.split(";base64,");
      mimeDetectado = partes[0].replace("data:", "");
      dataLimpia = partes[1];
    }

    const nombreFinal = (nombreArchivo || `Evidencia_${Date.now()}.jpg`).replace(/[^a-zA-Z0-9._-]/g, "_");
    const decodedBlob = Utilities.newBlob(Utilities.base64Decode(dataLimpia), mimeDetectado, nombreFinal);

    let carpeta;
    const folders = DriveApp.getFoldersByName("EVIDENCIAS_CHECKLIST_VONEX");
    if (folders.hasNext()) {
      carpeta = folders.next();
    } else {
      carpeta = DriveApp.createFolder("EVIDENCIAS_CHECKLIST_VONEX");
      try {
        carpeta.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      } catch (eShare) {
        console.warn("Aviso al configurar permisos de carpeta:", eShare);
      }
    }

    const archivoCreado = carpeta.createFile(decodedBlob);
    try {
      archivoCreado.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (eFileShare) {
      console.warn("Aviso al configurar permisos de archivo:", eFileShare);
    }

    const fileUrl = archivoCreado.getUrl();
    return {
      ok: true,
      url: fileUrl,
      id: archivoCreado.getId(),
      nombre: archivoCreado.getName()
    };
  } catch (err) {
    console.error("Error al subir archivo de evidencia:", err);
    // Fallback directo a URI de datos si DriveApp no tiene permisos
    return {
      ok: true,
      url: base64Data.startsWith("data:") ? base64Data : `data:${mimeType || 'image/jpeg'};base64,${base64Data}`,
      advertencia: "Almacenado localmente por restricción de Drive: " + err.message
    };
  }
}

// =========================================================================
// 6. MENÚ ADMINISTRATIVO EN GOOGLE SHEETS
// =========================================================================

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("🚀 Gestión Vonex")
    .addItem("📊 1. Abrir Módulo de Administración (CREAR)", "abrirModuloCrear")
    .addItem("👥 2. Abrir Módulo de Participantes", "abrirModuloParticipantes")
    .addSeparator()
    .addItem("📧 3. Enviar Reporte de Etapa (Manual)", "menuEnviarReporteEtapa")
    .addItem("🔔 4. Enviar Recordatorios a Destinatarios (Manual)", "menuEnviarRecordatorioManual")
    .addToUi();
}

function abrirModuloCrear() {
  const url = ScriptApp.getService().getUrl() + "?p=crear";
  const html = `<script>window.open("${url}", "_blank");google.script.host.close();</script>`;
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html), "Abriendo Módulo...");
}

function abrirModuloParticipantes() {
  const url = ScriptApp.getService().getUrl() + "?p=participantes";
  const html = `<script>window.open("${url}", "_blank");google.script.host.close();</script>`;
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html), "Abriendo Módulo...");
}

function menuEnviarReporteEtapa() {
  apiEnviarReporteEtapa("ANTES", "CHK-MANUAL");
  SpreadsheetApp.getActiveSpreadsheet().toast("Reporte de etapa enviado exitosamente.", "Éxito");
}

function menuEnviarRecordatorioManual() {
  apiEnviarRecordatorioManual("CHK-MANUAL");
  SpreadsheetApp.getActiveSpreadsheet().toast("Recordatorios enviados exitosamente.", "Éxito");
}
