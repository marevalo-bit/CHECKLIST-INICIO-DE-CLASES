const DIRECTORES = ["marevalo@vonex.edu.pe"];

function enviarReporteEjecutivoBubble3D() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojas = ["UNMSM-INTENSIVO", "UNMSM", "UNI", "UNHEVAL"];
  const hoy = new Date();
  const hoyCero = new Date(hoy.setHours(0,0,0,0));

  hojas.forEach(nombreHoja => {

    const sheet = ss.getSheetByName(nombreHoja);
    if (!sheet) return;

    const data = sheet.getDataRange().getValues();
    const colSedesIndices = [5,7,9,11,13,15,17,19];

    let total = 0;
    let realizadas = 0;
    let vencidas = 0;
    let riesgo = 0;

    for (let i = 8; i < data.length; i++) {

      const descripcion = (data[i][4] || "").toString().trim();
      if (!descripcion) continue;
      if (descripcion.toUpperCase().startsWith("ACTIVIDADES")) continue;

      const fechaRaw = data[i][21];
      if (!(fechaRaw instanceof Date)) continue;

      const diffDias = Math.floor(
        (new Date(fechaRaw).setHours(0,0,0,0) - hoyCero) / (1000 * 60 * 60 * 24)
      );

      colSedesIndices.forEach(colIdx => {

        let estado = data[i][colIdx];

        if (estado === true) estado = "REALIZADO";
        if (estado === false) estado = "";

        estado = (estado || "").toString().toUpperCase().trim();

        if (!estado) return;

        total++;

        if (estado === "REALIZADO" || estado === "NO APLICA") {
          realizadas++;
        }

        if (estado !== "REALIZADO" && estado !== "NO APLICA") {

          if (diffDias < 0) vencidas++;
          if (diffDias === 1) riesgo++;

        }

      });

    }

    Logger.log("===== " + nombreHoja + " =====");
    Logger.log("TOTAL CELDAS: " + total);
    Logger.log("REALIZADAS: " + realizadas);
    Logger.log("VENCIDAS: " + vencidas);
    Logger.log("RIESGO: " + riesgo);

  });

}


function enviarReporteEjecutivoVonex() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojas = ["UNMSM-INTENSIVO", "UNMSM", "UNI", "UNHEVAL"];
  const hoy = new Date();
  const hoyCero = new Date(hoy.setHours(0,0,0,0));

  hojas.forEach(nombreHoja => {

    const sheet = ss.getSheetByName(nombreHoja);
    if (!sheet) return;

    const data = sheet.getDataRange().getValues();
    const sedesCols = detectarColumnasSedes(sheet);

    let total = 0;
    let realizadas = 0;
    let vencidas = 0;
    let riesgo = 0;

    let mapaSedes = {};
    let encargadosSet = new Set();

    for (let i = 8; i < data.length; i++) {

      const descripcion = (data[i][4] || "").toString();
      if (descripcion.toUpperCase().startsWith("ACTIVIDADES")) continue;

      const fecha = data[i][21];
      if (!(fecha instanceof Date)) continue;

      const diffDias = Math.floor((new Date(fecha).setHours(0,0,0,0) - hoyCero)/(1000*60*60*24));

      const correoEncargado = (data[i][3] || "").toString().trim();

      if (correoEncargado.includes("@") && correoEncargado.includes(".")) {
       encargadosSet.add(correoEncargado);
      }

      sedesCols.forEach(col => {

        let estado = (data[i][col] || "").toString().toUpperCase().trim();
        if (!estado) return;

        total++;

        if (!mapaSedes[col]) mapaSedes[col] = {total:0, realizadas:0};

        mapaSedes[col].total++;

        if (estado === "REALIZADO" || estado === "NO APLICA") {
          realizadas++;
          mapaSedes[col].realizadas++;
        }

        if (estado === "PENDIENTE") {
          if (diffDias < 0) vencidas++;
          if (diffDias === 1) riesgo++;
        }

      });

    }

    const avance = total ? Math.round((realizadas/total)*100) : 0;

    let cardsSedes = "";

    sedesCols.forEach(col => {

      const nombreSede = sheet.getRange(7,col+1).getValue();
      const sedeData = mapaSedes[col] || {total:0, realizadas:0};

      const porcentaje = sedeData.total ? Math.round((sedeData.realizadas/sedeData.total)*100) : 0;

      cardsSedes += `
      <div style="background:white;padding:15px;border-radius:15px;margin-bottom:10px;
      box-shadow:0 8px 20px rgba(0,0,0,0.05);">
        <b>${nombreSede}</b><br>
        Avance: ${porcentaje}%
      </div>`;
    });

    const body = `
    <div style="font-family:Segoe UI;padding:30px;background:#f4f7fa;">
      <h2>${nombreHoja}</h2>

      <div style="background:white;padding:20px;border-radius:20px;
      box-shadow:0 10px 30px rgba(0,0,0,0.1);margin-bottom:20px;">
        <h3>Avance Global: ${avance}%</h3>
        <p>Realizadas: ${realizadas}</p>
        <p>Vencidas: ${vencidas}</p>
        <p>En riesgo: ${riesgo}</p>
      </div>

      <h3>Avance por Sede</h3>
      ${cardsSedes}
    </div>
    `;

    enviarEnMismoHilo(
      DIRECTORES.join(","),
      "REPORTE EJECUTIVO CHECKLIST " + nombreHoja,
      "",
      {
        htmlBody: body,
        cc: Array.from(encargadosSet).join(",")
      }
    );

  });

}

function detectarColumnasSedes(sheet) {

  const encabezados = sheet.getRange(7,1,1,sheet.getLastColumn()).getValues()[0];
  let columnasSedes = [];

  encabezados.forEach((valor, index) => {

    const texto = (valor || "").toString().toUpperCase().trim();
    if (!texto) return;
    if (texto.includes("FECHA")) return;

    if (index >= 5) columnasSedes.push(index);

  });

  return columnasSedes;
}
function enviarEnMismoHilo(destino, asunto, htmlBody, cc) {

  const threads = GmailApp.search('subject:"' + asunto + '"');

  if (threads.length > 0) {

    threads[0].reply("", {
      htmlBody: htmlBody,
      cc: cc
    });

  } else {

    GmailApp.sendEmail(destino, asunto, "", {
      htmlBody: htmlBody,
      cc: cc
    });

  }

}