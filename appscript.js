// =====================================================
//  PRODE MUNDIAL 2026 — App Script completo corregido
// =====================================================

function doGet(e) {
  let accion = e.parameter.accion;
  if (accion == "login")                    return login(e);
  if (accion == "partidos")                 return obtenerPartidos();
  if (accion == "apostar")                  return guardarPronostico(e);
  if (accion == "resultado")                return guardarResultado(e);
  if (accion == "posiciones")               return obtenerPosiciones();
  if (accion == "misPronosticos")           return obtenerPronosticos(e);
  if (accion == "eliminatorias")            return obtenerEliminatorias();
  if (accion == "setGanador")               return setGanador(e);
  if (accion == "inicializarEliminatorias") return inicializarEliminatorias();
  if (accion == "actualizarSlot")           return actualizarSlot(e);
  if (accion == "resetearTodo")             return resetearTodo();
  if (accion == "sincronizarResultados")    return sincronizarResultados();
  return ContentService.createTextOutput("API PRODE");
}

// ─── LOGIN ────────────────────────────────────────────
function login(e) {
  let usuario  = e.parameter.usuario;
  let password = e.parameter.password;
  let hoja     = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Usuarios");
  let datos    = hoja.getDataRange().getValues();

  for (let i = 1; i < datos.length; i++) {
    if (datos[i][1] == usuario && datos[i][2] == password) {
      return jsonOutput({ ok: true, rol: datos[i][3] });
    }
  }
  return jsonOutput({ ok: false });
}

// ─── PARTIDOS ─────────────────────────────────────────
function obtenerPartidos() {
  let hoja   = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Partidos");
  let datos  = hoja.getDataRange().getValues();
  let partidos = [];

  for (let i = 1; i < datos.length; i++) {
    let fechaObj  = new Date(datos[i][4]);
    let horaStr   = String(datos[i][7]);
    let horaParts = horaStr.split(":");
    let inicio    = new Date(
      fechaObj.getFullYear(), fechaObj.getMonth(), fechaObj.getDate(),
      parseInt(horaParts[0]), parseInt(horaParts[1])
    );
    let ahora = new Date();
    let fin   = new Date(inicio.getTime() + 2 * 60 * 60 * 1000);
    let estado = ahora < inicio ? "proximo" : ahora <= fin ? "en_juego" : "finalizado";

    partidos.push({
      id:            datos[i][0],
      grupo:         datos[i][1],
      local:         datos[i][2],
      visitante:     datos[i][3],
      fecha:         datos[i][4],
      hora:          String(datos[i][7]),
      estado:        estado,
      localFlag:     getFlag(datos[i][2]),
      visitanteFlag: getFlag(datos[i][3]),
    });
  }
  return jsonOutput(partidos);
}

// ─── GUARDAR PRONÓSTICO ───────────────────────────────
function guardarPronostico(e) {
  let usuario    = e.parameter.usuario;
  let partido    = e.parameter.partido;
  let pronostico = e.parameter.pronostico;
  let hoja       = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Pronosticos");
  let datos      = hoja.getDataRange().getValues();
  let encontrado = false;

  for (let i = 1; i < datos.length; i++) {
    if (datos[i][0] == usuario && datos[i][1] == partido) {
      hoja.getRange(i + 1, 3).setValue(pronostico);
      encontrado = true;
      break;
    }
  }
  if (!encontrado) hoja.appendRow([usuario, partido, pronostico]);
  return jsonOutput({ ok: true });
}

// ─── GUARDAR RESULTADO ────────────────────────────────
function guardarResultado(e) {
  let partido    = e.parameter.partido;
  let resultado  = e.parameter.resultado;
  let hoja       = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Resultados");
  let datos      = hoja.getDataRange().getValues();
  let encontrado = false;

  for (let i = 1; i < datos.length; i++) {
    if (datos[i][0] == partido) {
      hoja.getRange(i + 1, 2).setValue(resultado);
      encontrado = true;
      break;
    }
  }
  if (!encontrado) hoja.appendRow([partido, resultado]);
  recalcularPosiciones();
  return jsonOutput({ ok: true });
}

// ─── RECALCULAR POSICIONES ────────────────────────────
function recalcularPosiciones() {
  let ss              = SpreadsheetApp.getActiveSpreadsheet();
  let hojaPronosticos = ss.getSheetByName("Pronosticos");
  let hojaResultados  = ss.getSheetByName("Resultados");
  let hojaPosiciones  = ss.getSheetByName("Posiciones");
  let pronosticos     = hojaPronosticos.getDataRange().getValues();
  let resultados      = hojaResultados.getDataRange().getValues();
  let puntos          = {};

  for (let i = 1; i < pronosticos.length; i++) {
    let usuario      = pronosticos[i][0];
    let partido      = pronosticos[i][1];
    let apuesta      = pronosticos[i][2];
    let resultadoReal = "";

    for (let j = 1; j < resultados.length; j++) {
      if (resultados[j][0] == partido) { resultadoReal = resultados[j][1]; break; }
    }
    if (!puntos[usuario]) puntos[usuario] = 0;
    if (apuesta == resultadoReal && resultadoReal != "") puntos[usuario]++;
  }

  // Borrar filas físicamente para evitar duplicados
  const lastRow = hojaPosiciones.getLastRow();
  if (lastRow > 1) hojaPosiciones.deleteRows(2, lastRow - 1);
  hojaPosiciones.getRange(1, 1, 1, 2).setValues([["Usuario", "Puntos"]]);

  const ordenados = Object.entries(puntos).sort((a, b) => b[1] - a[1]);
  if (ordenados.length > 0) {
    hojaPosiciones.getRange(2, 1, ordenados.length, 2).setValues(ordenados.map(([u, p]) => [u, p]));
  }
}

// ─── OBTENER POSICIONES ───────────────────────────────
function obtenerPosiciones() {
  let hoja  = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Posiciones");
  let datos = hoja.getDataRange().getValues();
  let pos   = [];
  for (let i = 1; i < datos.length; i++) pos.push({ usuario: datos[i][0], puntos: datos[i][1] });
  pos.sort((a, b) => b.puntos - a.puntos);
  return jsonOutput(pos);
}

// ─── OBTENER PRONÓSTICOS ──────────────────────────────
function obtenerPronosticos(e) {
  let usuario = e.parameter.usuario;
  let hoja    = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Pronosticos");
  let datos   = hoja.getDataRange().getValues();
  let pron    = {};
  for (let i = 1; i < datos.length; i++) {
    if (datos[i][0] == usuario) pron[datos[i][1]] = datos[i][2];
  }
  return jsonOutput(pron);
}

// ─── OBTENER ELIMINATORIAS ────────────────────────────
function obtenerEliminatorias() {
  let hoja    = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Eliminatorias");
  let datos   = hoja.getDataRange().getValues();
  let partidos = [];
  for (let i = 1; i < datos.length; i++) {
    partidos.push({
      id:        datos[i][0],
      fase:      datos[i][1],
      local:     datos[i][2],
      visitante: datos[i][3],
      ganador:   datos[i][4],
      fecha:     datos[i][7] || "",
      sede:      datos[i][8] || ""
    });
  }
  return jsonOutput(partidos);
}

// ─── SET GANADOR (con avance automático por slot) ─────
function setGanador(e) {
  let id      = e.parameter.id;
  let ganador = e.parameter.ganador;
  let perdedor = e.parameter.perdedor || "";

  let hoja  = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Eliminatorias");
  let datos = hoja.getDataRange().getValues();

  for (let i = 1; i < datos.length; i++) {
    if (datos[i][0] == id) {
      // Guardar ganador (col E = índice 4)
      hoja.getRange(i + 1, 5).setValue(ganador);

      let sigId    = datos[i][5]; // col F: SiguienteId ganador
      let sigSlot  = datos[i][6]; // col G: L o V
      let perdId   = datos[i][7] || ""; // col H: SiguienteId perdedor (solo semis)
      let perdSlot = datos[i][8] || ""; // col I: slot perdedor

      // Determinar perdedor si no vino en el parámetro
      if (!perdedor) {
        perdedor = (ganador == datos[i][2]) ? datos[i][3] : datos[i][2];
      }

      // Avanzar ganador al siguiente partido
      if (sigId) avanzarEquipo(hoja, sigId, ganador, sigSlot);

      // Avanzar perdedor (3er puesto para semis)
      if (perdId) avanzarEquipo(hoja, perdId, perdedor, perdSlot);

      break;
    }
  }
  return jsonOutput({ ok: true });
}

function avanzarEquipo(hoja, sigId, equipo, slot) {
  let datos = hoja.getDataRange().getValues();
  for (let j = 1; j < datos.length; j++) {
    if (datos[j][0] == sigId) {
      // col C (índice 2) = local, col D (índice 3) = visitante
      let col = (slot == "L") ? 3 : 4;
      hoja.getRange(j + 1, col).setValue(equipo);
      break;
    }
  }
}

// ─── SINCRONIZAR RESULTADOS AUTOMÁTICAMENTE ───────────
// Requiere: guardar tu API key de football-data.org en
// Apps Script → Configuración del proyecto → Propiedades de script
// Clave: FOOTBALL_API_KEY  Valor: tu_api_key

function sincronizarResultados() {
  const apiKey = PropertiesService.getScriptProperties().getProperty("FOOTBALL_API_KEY");
  if (!apiKey) return jsonOutput({ ok: false, error: "Falta FOOTBALL_API_KEY en propiedades del script" });

  // Mapeo nombres en inglés (API) → español (tu sheet)
  const NOMBRES = {
    "Mexico": "México", "South Africa": "Sudáfrica", "South Korea": "Corea del Sur",
    "Czech Republic": "República Checa", "Czechia": "República Checa",
    "Canada": "Canadá", "Bosnia and Herzegovina": "Bosnia y Herzegovina",
    "Qatar": "Qatar", "Switzerland": "Suiza", "Brazil": "Brasil",
    "Morocco": "Marruecos", "Haiti": "Haití", "Scotland": "Escocia",
    "United States": "Estados Unidos", "USA": "Estados Unidos",
    "Paraguay": "Paraguay", "Australia": "Australia", "Turkey": "Turquía",
    "Germany": "Alemania", "Curaçao": "Curazao", "Curacao": "Curazao",
    "Ivory Coast": "Costa de Marfil", "Côte d'Ivoire": "Costa de Marfil",
    "Ecuador": "Ecuador", "Netherlands": "Países Bajos", "Japan": "Japón",
    "Sweden": "Suecia", "Tunisia": "Túnez", "Belgium": "Bélgica",
    "Egypt": "Egipto", "Iran": "Irán", "New Zealand": "Nueva Zelanda",
    "Spain": "España", "Cape Verde": "Cabo Verde", "Saudi Arabia": "Arabia Saudita",
    "Uruguay": "Uruguay", "France": "Francia", "Senegal": "Senegal",
    "Iraq": "Irak", "Norway": "Noruega", "Argentina": "Argentina",
    "Algeria": "Argelia", "Austria": "Austria", "Jordan": "Jordania",
    "Portugal": "Portugal", "DR Congo": "RD de Congo",
    "Uzbekistan": "Uzbekistán", "Colombia": "Colombia",
    "England": "Inglaterra", "Croatia": "Croacia", "Ghana": "Ghana", "Panama": "Panamá"
  };

  // Traer partidos finalizados del Mundial
  let response;
  try {
    response = UrlFetchApp.fetch(
      "https://api.football-data.org/v4/competitions/WC/matches?status=FINISHED",
      { headers: { "X-Auth-Token": apiKey }, muteHttpExceptions: true }
    );
  } catch(e) {
    return jsonOutput({ ok: false, error: "Error de red: " + e.message });
  }

  if (response.getResponseCode() !== 200) {
    return jsonOutput({ ok: false, error: "API error " + response.getResponseCode() + ": " + response.getContentText() });
  }

  const data    = JSON.parse(response.getContentText());
  const matches = data.matches || [];

  // Cargar partidos del sheet para buscar por equipos
  const ss          = SpreadsheetApp.getActiveSpreadsheet();
  const hojaPartidos = ss.getSheetByName("Partidos");
  const partidos    = hojaPartidos.getDataRange().getValues();

  let actualizados = 0;

  matches.forEach(m => {
    const score = m.score && m.score.fullTime;
    if (!score || score.home === null || score.away === null) return;

    const localApi     = NOMBRES[m.homeTeam.name] || m.homeTeam.shortName || m.homeTeam.name;
    const visitanteApi = NOMBRES[m.awayTeam.name] || m.awayTeam.shortName || m.awayTeam.name;

    // Determinar L / E / V
    let resultado;
    if      (score.home > score.away) resultado = "L";
    else if (score.home < score.away) resultado = "V";
    else                               resultado = "E";

    // Buscar el partido en el sheet por nombres de equipos
    let partidoId = null;
    for (let i = 1; i < partidos.length; i++) {
      const localSheet     = String(partidos[i][2]).trim();
      const visitanteSheet = String(partidos[i][3]).trim();
      if (localSheet === localApi && visitanteSheet === visitanteApi) {
        partidoId = partidos[i][0];
        break;
      }
    }

    if (!partidoId) return; // no encontrado

    // Guardar resultado
    const hojaResultados = ss.getSheetByName("Resultados");
    const resultados     = hojaResultados.getDataRange().getValues();
    let encontrado = false;

    for (let i = 1; i < resultados.length; i++) {
      if (resultados[i][0] == partidoId) {
        hojaResultados.getRange(i + 1, 2).setValue(resultado);
        encontrado = true;
        break;
      }
    }
    if (!encontrado) hojaResultados.appendRow([partidoId, resultado]);
    actualizados++;
  });

  if (actualizados > 0) recalcularPosiciones();

  return jsonOutput({ ok: true, actualizados: actualizados, total: matches.length });
}

// ─── CREAR TRIGGER AUTOMÁTICO (ejecutar UNA vez desde el editor) ──
// Llama a sincronizarResultados() cada hora automáticamente
function crearTriggerSincronizacion() {
  // Borrar triggers anteriores del mismo nombre
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === "sincronizarResultados") ScriptApp.deleteTrigger(t);
  });
  // Crear trigger cada hora
  ScriptApp.newTrigger("sincronizarResultados")
    .timeBased()
    .everyHours(1)
    .create();
  Logger.log("✅ Trigger creado: sincronizarResultados cada 1 hora");
}

// ─── RESETEAR TODO ────────────────────────────────────
function resetearTodo() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Limpiar Resultados (mantener cabecera)
  const hojaRes = ss.getSheetByName("Resultados");
  if (hojaRes.getLastRow() > 1) hojaRes.deleteRows(2, hojaRes.getLastRow() - 1);

  // Limpiar Pronosticos (mantener cabecera)
  const hojaPron = ss.getSheetByName("Pronosticos");
  if (hojaPron.getLastRow() > 1) hojaPron.deleteRows(2, hojaPron.getLastRow() - 1);

  // Limpiar Posiciones (mantener cabecera)
  const hojaPos = ss.getSheetByName("Posiciones");
  if (hojaPos.getLastRow() > 1) hojaPos.deleteRows(2, hojaPos.getLastRow() - 1);

  // Reinicializar Eliminatorias con placeholders
  inicializarEliminatorias();

  return jsonOutput({ ok: true, mensaje: "Todo reseteado correctamente" });
}

// ─── ACTUALIZAR SLOT (local o visitante de un partido) ─
function actualizarSlot(e) {
  let id     = e.parameter.id;
  let slot   = e.parameter.slot;   // "L" o "V"
  let equipo = e.parameter.equipo;

  let hoja  = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Eliminatorias");
  let datos = hoja.getDataRange().getValues();

  for (let i = 1; i < datos.length; i++) {
    if (datos[i][0] == id) {
      let col = (slot === "L") ? 3 : 4; // col C = local, col D = visitante
      hoja.getRange(i + 1, col).setValue(equipo);
      break;
    }
  }
  return jsonOutput({ ok: true });
}

// ─── INICIALIZAR ELIMINATORIAS ────────────────────────
// Ejecutar UNA vez desde el editor de Apps Script para poblar la hoja
function inicializarEliminatorias() {
  let ss   = SpreadsheetApp.getActiveSpreadsheet();
  let hoja = ss.getSheetByName("Eliminatorias");
  hoja.clearContents();
  hoja.appendRow([
    "Id","Fase","Local","Visitante","Ganador",
    "SiguienteId","SlotSiguiente","SiguienteIdPerdedor","SlotPerdedor",
    "Fecha","Sede"
  ]);

  // [Id, Fase, Local, Visitante, Ganador, SigId, SigSlot, PerdId, PerdSlot, Fecha, Sede]
  let partidos = [
    // ── 16AVOS ──
    ["P73","16avos de Final","2° Grupo A",         "2° Grupo B",              "","P90","L","","","28/06","Los Ángeles"],
    ["P74","16avos de Final","1° Grupo E",         "3° Grupo A/B/C/D/F",      "","P89","L","","","29/06","Boston"],
    ["P75","16avos de Final","1° Grupo F",         "2° Grupo C",              "","P90","V","","","29/06","Monterrey"],
    ["P76","16avos de Final","1° Grupo C",         "2° Grupo F",              "","P91","L","","","29/06","Houston"],
    ["P77","16avos de Final","1° Grupo I",         "3° Grupo C/D/F/G/H",      "","P89","V","","","30/06","Nueva Jersey"],
    ["P78","16avos de Final","2° Grupo E",         "2° Grupo I",              "","P91","V","","","30/06","Dallas"],
    ["P79","16avos de Final","1° Grupo A",         "3° Grupo C/E/F/H/I",      "","P92","L","","","30/06","Ciudad de México"],
    ["P80","16avos de Final","1° Grupo L",         "3° Grupo E/H/I/J/K",      "","P92","V","","","01/07","Atlanta"],
    ["P81","16avos de Final","1° Grupo D",         "3° Grupo B/E/F/I/J",      "","P94","L","","","01/07","San Francisco"],
    ["P82","16avos de Final","1° Grupo G",         "3° Grupo A/E/H/I/J",      "","P94","V","","","01/07","Seattle"],
    ["P83","16avos de Final","2° Grupo K",         "2° Grupo L",              "","P93","L","","","02/07","Toronto"],
    ["P84","16avos de Final","1° Grupo H",         "2° Grupo J",              "","P93","V","","","02/07","Los Ángeles"],
    ["P85","16avos de Final","1° Grupo B",         "3° Grupo E/F/G/I/J",      "","P96","L","","","02/07","Vancouver"],
    ["P86","16avos de Final","1° Grupo J",         "2° Grupo H",              "","P95","L","","","03/07","Miami"],
    ["P87","16avos de Final","1° Grupo K",         "3° Grupo D/E/I/J/L",      "","P96","V","","","03/07","Kansas City"],
    ["P88","16avos de Final","2° Grupo D",         "2° Grupo G",              "","P95","V","","","03/07","Dallas"],
    // ── OCTAVOS ──
    ["P89","Octavos de Final","Gan. P74",          "Gan. P77",                "","P97","L","","","04/07","Philadelphia"],
    ["P90","Octavos de Final","Gan. P73",          "Gan. P75",                "","P97","V","","","04/07","Houston"],
    ["P91","Octavos de Final","Gan. P76",          "Gan. P78",                "","P99","L","","","05/07","Nueva Jersey"],
    ["P92","Octavos de Final","Gan. P79",          "Gan. P80",                "","P99","V","","","05/07","Ciudad de México"],
    ["P93","Octavos de Final","Gan. P83",          "Gan. P84",                "","P98","L","","","06/07","Dallas"],
    ["P94","Octavos de Final","Gan. P81",          "Gan. P82",                "","P98","V","","","06/07","Seattle"],
    ["P95","Octavos de Final","Gan. P86",          "Gan. P88",                "","P100","L","","","07/07","Atlanta"],
    ["P96","Octavos de Final","Gan. P85",          "Gan. P87",                "","P100","V","","","07/07","Vancouver"],
    // ── CUARTOS ──
    ["P97","Cuartos de Final","Gan. P89",          "Gan. P90",                "","P101","L","","","09/07","Boston"],
    ["P98","Cuartos de Final","Gan. P93",          "Gan. P94",                "","P101","V","","","10/07","Los Ángeles"],
    ["P99","Cuartos de Final","Gan. P91",          "Gan. P92",                "","P102","L","","","11/07","Miami"],
    ["P100","Cuartos de Final","Gan. P95",         "Gan. P96",                "","P102","V","","","11/07","Kansas City"],
    // ── SEMIS ──
    ["P101","Semifinal","Gan. P97",                "Gan. P98",                "","P104","L","P103","L","14/07","Dallas"],
    ["P102","Semifinal","Gan. P99",                "Gan. P100",               "","P104","V","P103","V","15/07","Atlanta"],
    // ── 3ER PUESTO Y FINAL ──
    ["P103","Tercer Puesto","Perd. P101",          "Perd. P102",              "","","","","","18/07","Miami"],
    ["P104","Final","Gan. P101",                   "Gan. P102",               "","","","","","19/07","Nueva Jersey"],
  ];

  partidos.forEach(p => hoja.appendRow(p));
  return jsonOutput({ ok: true, mensaje: "Eliminatorias inicializadas: " + partidos.length + " partidos" });
}

// ─── HELPER ───────────────────────────────────────────
function jsonOutput(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function getFlag(country) {
  const flags = {
    "México": "🇲🇽", "Sudáfrica": "🇿🇦", "Corea del Sur": "🇰🇷",
    "República Checa": "🇨🇿", "Canadá": "🇨🇦", "Bosnia y Herzegovina": "🇧🇦",
    "Qatar": "🇶🇦", "Suiza": "🇨🇭", "Brasil": "🇧🇷", "Marruecos": "🇲🇦",
    "Haití": "🇭🇹", "Escocia": "🏴󠁧󠁢󠁳󠁣󠁴󠁿", "Estados Unidos": "🇺🇸", "Paraguay": "🇵🇾",
    "Australia": "🇦🇺", "Turquía": "🇹🇷", "Alemania": "🇩🇪", "Curazao": "🇨🇼",
    "Costa de Marfil": "🇨🇮", "Ecuador": "🇪🇨", "Países Bajos": "🇳🇱",
    "Japón": "🇯🇵", "Suecia": "🇸🇪", "Túnez": "🇹🇳", "Bélgica": "🇧🇪",
    "Egipto": "🇪🇬", "Irán": "🇮🇷", "Nueva Zelanda": "🇳🇿", "España": "🇪🇸",
    "Cabo Verde": "🇨🇻", "Arabia Saudita": "🇸🇦", "Uruguay": "🇺🇾",
    "Francia": "🇫🇷", "Senegal": "🇸🇳", "Irak": "🇮🇶", "Noruega": "🇳🇴",
    "Argentina": "🇦🇷", "Austria": "🇦🇹", "Jordania": "🇯🇴", "Portugal": "🇵🇹",
    "RD de Congo": "🇨🇩", "Uzbekistán": "🇺🇿", "Colombia": "🇨🇴",
    "Inglaterra": "🏴󠁧󠁢󠁥󠁮󠁧󠁿", "Croacia": "🇭🇷", "Ghana": "🇬🇭",
    "Panamá": "🇵🇦", "Argelia": "🇩🇿"
  };
  return flags[country] || "";
}
