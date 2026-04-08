// ============================================================
//  IpercCode.gs — ARCHIVO UNICO (Apps Script Web App)
//
//  Unifica Doget webapp.gs + IpercCode.gs en un solo archivo.
//  Junto con Iperc.html forma el proyecto completo:
//    - IpercCode.gs  (este archivo, backend)
//    - Iperc.html    (frontend)
// ============================================================

// ─────────────────────────────────────────
//  CONSTANTES DE HOJA / COLUMNAS
// ─────────────────────────────────────────
const SHEET_DATOS    = "DATOS";
const FILA_INICIO    = 6;

const COL_PROCESO    = 2;   // B
const COL_AREA       = 3;   // C
const COL_TAREA      = 4;   // D
const COL_PUESTO     = 5;   // E
const COL_RUTINARIO  = 6;   // F
const COL_PELIGROS   = 7;   // G
const COL_RIESGO     = 8;   // H
const COL_PROB       = 9;   // I
const COL_SEV        = 10;  // J
// K (11) Score inicial (formula), L (12) Nivel inicial (formula)
const COL_ELIMINAC   = 13;  // M
const COL_SUSTIT     = 14;  // N
const COL_ING        = 15;  // O
const COL_ADM        = 16;  // P
const COL_EPP        = 17;  // Q
// R (18) Prob Res, S (19) Sev Res, T (20) Score Res, U (21) Nivel Res
// V (22) Accion, W (23) Responsable

// ─────────────────────────────────────────
//  doGet — SIRVE LA PAGINA WEB
//  OBLIGATORIO para Web Apps
// ─────────────────────────────────────────
function doGet(e) {
  return HtmlService
    .createHtmlOutputFromFile("Iperc")
    .setTitle("IPERC Linea Base — Sistema IA")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag("viewport", "width=device-width, initial-scale=1.0");
}

// ─────────────────────────────────────────
//  MATRIZ CONFIG — Standalone
// ─────────────────────────────────────────
function matrizFCXDefault_() {
  return {
    nombre: "Matriz FCX 4x4", size: 4, maxProb: 4, maxSev: 4,
    probs: [
      { v:1, l:"Improbable",   d:"Muy improbable durante la vida de la operacion" },
      { v:2, l:"Posible",      d:"Puede ocurrir durante la vida de la operacion" },
      { v:3, l:"Probable",     d:"Puede ocurrir menos de una vez al anio" },
      { v:4, l:"Casi Seguro",  d:"Evento recurrente o mas de una vez al anio" }
    ],
    sevs: [
      { v:1, l:"Menor",         d:"Lesion minima o primeros auxilios" },
      { v:2, l:"Moderado",      d:"Tratamiento medico o labores restringidas" },
      { v:3, l:"Significativo", d:"Fatalidades o discapacidades permanentes" },
      { v:4, l:"Catastrofico",  d:"Fatalidades multiples" }
    ],
    nivs: [
      { l:"BAJO",    de:1,  ha:2,  c:"#3d9e3d", t:"#ffffff" },
      { l:"MEDIO",   de:3,  ha:4,  c:"#d4a017", t:"#ffffff" },
      { l:"ALTO",    de:6,  ha:8,  c:"#e05c00", t:"#ffffff" },
      { l:"CRITICO", de:9,  ha:16, c:"#cc0000", t:"#ffffff" }
    ]
  };
}

function leerMatrizConfig() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return matrizFCXDefault_();
  var sheet = ss.getSheetByName("Configuracion Matriz");
  if (!sheet) return matrizFCXDefault_();
  try {
    var maxProb = Number(sheet.getRange("D3").getValue()) || 4;
    var maxSev  = Number(sheet.getRange("F3").getValue()) || 4;
    var probs = [];
    for (var i = 0; i < maxProb; i++) {
      var r = sheet.getRange(5 + i, 1, 1, 3).getValues()[0];
      probs.push({ v: Number(r[0]) || (i+1), l: String(r[1] || "Nivel "+(i+1)), d: String(r[2] || "") });
    }
    var sevs = [];
    for (var j = 0; j < maxSev; j++) {
      var r2 = sheet.getRange(5 + j, 5, 1, 3).getValues()[0];
      sevs.push({ v: Number(r2[0]) || (j+1), l: String(r2[1] || "Nivel "+(j+1)), d: String(r2[2] || "") });
    }
    var nivs = [];
    for (var k = 5; k <= 12; k++) {
      var r3 = sheet.getRange(k, 9, 1, 5).getValues()[0];
      if (!r3[0]) break;
      nivs.push({ l: String(r3[0]), de: Number(r3[1]) || 1, ha: Number(r3[2]) || 4, c: String(r3[3] || "#3d9e3d"), t: String(r3[4] || "#ffffff") });
    }
    var cfg = matrizFCXDefault_();
    return {
      nombre:  String(sheet.getRange("B3").getValue() || cfg.nombre),
      size:    Math.max(maxProb, maxSev),
      maxProb: maxProb,
      maxSev:  maxSev,
      probs: probs.length > 0 ? probs : cfg.probs,
      sevs:  sevs.length  > 0 ? sevs  : cfg.sevs,
      nivs:  nivs.length  > 0 ? nivs  : cfg.nivs
    };
  } catch(e) {
    Logger.log("Error leyendo matriz: " + e);
    return matrizFCXDefault_();
  }
}

// ─────────────────────────────────────────
//  HELPERS DE NIVEL Y FORMULAS
// ─────────────────────────────────────────
function calcularNivelRiesgoByScore(score, cfg) {
  var nivs = (cfg && cfg.nivs) || [];
  for (var i = 0; i < nivs.length; i++) {
    var n = nivs[i];
    if (score >= n.de && score <= n.ha) {
      return { label: n.l, color: n.c, colorTexto: n.t };
    }
  }
  if (nivs.length) {
    var last = nivs[nivs.length - 1];
    return { label: last.l, color: last.c, colorTexto: last.t };
  }
  return null;
}

function buildFormulaNivel(fila, col, cfg) {
  var nivs = (cfg && cfg.nivs) || [];
  if (!nivs.length) return '=""';
  var cell = col + fila;
  var formula = '=IF(' + cell + '="","",';
  var closures = 1;
  for (var i = 0; i < nivs.length; i++) {
    var n = nivs[i];
    if (i === nivs.length - 1) {
      formula += '"' + n.l + '"';
    } else {
      formula += 'IF(AND(' + cell + '>=' + n.de + ',' + cell + '<=' + n.ha + '),"' + n.l + '",';
      closures++;
    }
  }
  for (var k = 0; k < closures; k++) formula += ')';
  return formula;
}

// Stub opcional: si tienes una hoja RESUMEN puedes ampliarla aqui.
function actualizarResumen() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return;
    var sheet = ss.getSheetByName("RESUMEN");
    if (!sheet) return;
    sheet.getRange("B2").setValue(new Date());
  } catch (e) {
    Logger.log("actualizarResumen: " + e);
  }
}

// ─────────────────────────────────────────
//  PUENTE HTML <-> SHEET
// ─────────────────────────────────────────

/** Guarda todas las filas IPERC al Sheet */
function guardarFilasAlSheet(filasJSON) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return { ok: false, error: "Sin Spreadsheet activo" };
  const sheet = ss.getSheetByName(SHEET_DATOS);
  if (!sheet) return { ok: false, error: "Hoja " + SHEET_DATOS + " no encontrada" };

  const filas = JSON.parse(filasJSON);
  const cfg   = leerMatrizConfig();

  // Limpiar filas de datos anteriores (mantener encabezados)
  const ultimaFila = Math.max(sheet.getLastRow(), FILA_INICIO + filas.length - 1);
  if (ultimaFila >= FILA_INICIO) {
    sheet.getRange(FILA_INICIO, 1, ultimaFila - FILA_INICIO + 1, 24).clearContent();
    sheet.getRange(FILA_INICIO, 1, ultimaFila - FILA_INICIO + 1, 24).setBackground("#FFFFFF");
  }

  filas.forEach((f, i) => {
    if (!f.proceso) return;
    const fila = FILA_INICIO + i;

    sheet.getRange(fila, 1).setValue(i + 1);
    sheet.getRange(fila, COL_PROCESO).setValue(f.proceso   || "");
    sheet.getRange(fila, COL_AREA   ).setValue(f.area      || "");
    sheet.getRange(fila, COL_TAREA  ).setValue(f.tarea     || "");
    sheet.getRange(fila, COL_PUESTO ).setValue(f.puesto    || "");
    sheet.getRange(fila, COL_RUTINARIO).setValue(f.rnr     || "R");

    if (f.peligros) {
      sheet.getRange(fila, COL_PELIGROS).setValue(f.peligros    || "");
      sheet.getRange(fila, COL_RIESGO  ).setValue(f.riesgo      || "");
      sheet.getRange(fila, COL_PROB    ).setValue(Number(f.prob) || "");
      sheet.getRange(fila, COL_SEV     ).setValue(Number(f.sev)  || "");
      sheet.getRange(fila, COL_ELIMINAC).setValue(f.eliminacion  || "");
      sheet.getRange(fila, COL_SUSTIT  ).setValue(f.sustitucion  || "");
      sheet.getRange(fila, COL_ING     ).setValue(f.ing          || "");
      sheet.getRange(fila, COL_ADM     ).setValue(f.adm          || "");
      sheet.getRange(fila, COL_EPP     ).setValue(f.epp          || "");
      sheet.getRange(fila, COL_PELIGROS, 1, COL_EPP - COL_PELIGROS + 1).setBackground("#FFFACD");

      // Formulas y colores
      sheet.getRange(fila, 11).setFormula(`=IF(I${fila}*J${fila}=0,"",I${fila}*J${fila})`);
      sheet.getRange(fila, 12).setFormula(buildFormulaNivel(fila, "K", cfg));
      if (f.prob && f.sev) {
        const nv = calcularNivelRiesgoByScore(Number(f.prob) * Number(f.sev), cfg);
        if (nv) sheet.getRange(fila, 12).setBackground(nv.color).setFontColor(nv.colorTexto).setFontWeight("bold").setHorizontalAlignment("center");
      }
    }

    if (f.probRes) sheet.getRange(fila, 18).setValue(Number(f.probRes) || "");
    if (f.sevRes)  sheet.getRange(fila, 19).setValue(Number(f.sevRes)  || "");
    sheet.getRange(fila, 20).setFormula(`=IF(R${fila}*S${fila}=0,"",R${fila}*S${fila})`);
    sheet.getRange(fila, 21).setFormula(buildFormulaNivel(fila, "T", cfg));
    if (f.accion)      sheet.getRange(fila, 22).setValue(f.accion      || "");
    if (f.responsable) sheet.getRange(fila, 23).setValue(f.responsable || "");

    sheet.getRange(fila, 1, 1, 23).setWrap(true).setVerticalAlignment("top");
    sheet.setRowHeight(fila, 85);
  });

  actualizarResumen();
  return { ok: true, filas: filas.filter(f => f.proceso).length };
}

/** Lee filas del Sheet y las devuelve al HTML */
function leerFilasDelSheet() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return JSON.stringify([]);
  const sheet = ss.getSheetByName(SHEET_DATOS);
  if (!sheet) return JSON.stringify([]);

  const uf = sheet.getLastRow();
  if (uf < FILA_INICIO) return JSON.stringify([]);

  const datos = sheet.getRange(FILA_INICIO, 1, uf - FILA_INICIO + 1, 24).getValues();
  const filas = datos
    .filter(r => r[COL_PROCESO - 1])
    .map(r => ({
      proceso:    r[COL_PROCESO - 1],
      area:       r[COL_AREA    - 1],
      tarea:      r[COL_TAREA   - 1],
      puesto:     r[COL_PUESTO  - 1],
      rnr:        r[COL_RUTINARIO - 1],
      peligros:   r[COL_PELIGROS - 1],
      riesgo:     r[COL_RIESGO  - 1],
      prob:       r[COL_PROB    - 1],
      sev:        r[COL_SEV     - 1],
      eliminacion:r[COL_ELIMINAC - 1],
      sustitucion:r[COL_SUSTIT  - 1],
      ing:        r[COL_ING     - 1],
      adm:        r[COL_ADM     - 1],
      epp:        r[COL_EPP     - 1],
      probRes:    r[17],
      sevRes:     r[18],
      accion:     r[21],
      responsable:r[22]
    }));

  return JSON.stringify(filas);
}

/** Devuelve la configuracion de la matriz al HTML */
function leerMatrizConfigJSON() {
  return JSON.stringify(leerMatrizConfig());
}

/** Guarda la configuracion de empresa desde el HTML */
function guardarConfigEmpresaDesdeHTML(cfgJSON) {
  PropertiesService.getScriptProperties().setProperty("CONFIG_EMPRESA", cfgJSON);
  return { ok: true };
}

/** Lee la configuracion de empresa */
function leerConfigEmpresa() {
  return PropertiesService.getScriptProperties().getProperty("CONFIG_EMPRESA") || "{}";
}

/** Guarda la API Key */
function guardarAPIKeyDesdeHTML(key) {
  if (!key || key.length < 20) return { ok: false, error: "Clave muy corta" };
  PropertiesService.getScriptProperties().setProperty("GEMINI_API_KEY", key);
  return { ok: true };
}

/** Verifica si hay API Key guardada */
function tieneAPIKey() {
  return !!PropertiesService.getScriptProperties().getProperty("GEMINI_API_KEY");
}

/** Obtiene la API Key guardada, null si no existe */
function obtenerAPIKey() {
  var key = PropertiesService.getScriptProperties().getProperty("GEMINI_API_KEY");
  if (!key || key.trim().length < 20) return null;
  return key.trim();
}

// ─────────────────────────────────────────
//  ANALISIS CON GEMINI — AUTOCONTENIDO
//  Recibe: {proceso, area, tarea, puesto, rnr, maxProb, maxSev, escalaProb, escalaSev, escalaNivs}
// ─────────────────────────────────────────
function analizarFilaConIAv2(filaJSON) {
  var apiKey = obtenerAPIKey();
  if (!apiKey) return JSON.stringify({ error: "Sin API Key configurada. Ve a la pestana API Key." });

  var f = JSON.parse(filaJSON);
  var modelo = PropertiesService.getScriptProperties().getProperty("GEMINI_MODELO") || "gemini-2.0-flash";
  var url = "https://generativelanguage.googleapis.com/v1beta/models/" + modelo + ":generateContent?key=" + apiKey;

  var tipo = (f.rnr === "NR") ? "No Rutinario" : "Rutinario";
  var escP = f.escalaProb || "1=Improbable, 2=Posible, 3=Probable, 4=Casi Seguro";
  var escS = f.escalaSev  || "1=Menor, 2=Moderado, 3=Significativo, 4=Catastrofico";
  var escN = f.escalaNivs || "1-2=BAJO, 3-4=MEDIO, 6-8=ALTO, 9-16=CRITICO";
  var mP   = f.maxProb || 4;
  var mS   = f.maxSev  || 4;

  var prompt =
    "Eres un experto SSOMA en mineria en Peru. Normas: DS-024-2016-EM, ISO 45001:2018, NIOSH.\n" +
    "TAREA:\n- Proceso: " + f.proceso + "\n- Area: " + f.area +
    "\n- Tarea: " + f.tarea + "\n- Puesto: " + f.puesto + "\n- Tipo: " + tipo + "\n\n" +
    "PROBABILIDAD (entero 1 al " + mP + "): " + escP + "\n" +
    "CONSECUENCIA (entero 1 al " + mS + "): " + escS + "\n" +
    "NIVELES (ProbxCons): " + escN + "\n\n" +
    "REGLAS ESTRICTAS:\n" +
    "1. Responde SOLO con JSON valido. Sin texto extra, sin backticks, sin markdown.\n" +
    "2. En listas usa ' | ' como separador (NO saltos de linea).\n" +
    "3. probabilidad = entero 1-" + mP + ", severidad = entero 1-" + mS + ".\n" +
    "4. prob_residual y sev_residual deben dar producto MENOR al inicial.\n" +
    "5. Terminologia tecnica minera peruana.\n" +
    "JSON a devolver:\n" +
    '{"peligros":"1. Peligro A | 2. Peligro B | 3. Peligro C",' +
    '"riesgo":"1. Riesgo A | 2. Riesgo B",' +
    '"probabilidad":2,"prob_justificacion":"razon breve",' +
    '"severidad":3,"sev_justificacion":"razon breve",' +
    '"eliminacion":"medida o N/A","sustitucion":"medida o N/A",' +
    '"ing_controles":"1. Control 1 | 2. Control 2",' +
    '"adm_controles":"1. PETS | 2. Capacitacion | 3. Supervision",' +
    '"epp":"1. Casco | 2. Lentes | 3. Guantes | 4. Botas punta acero",' +
    '"prob_residual":1,"sev_residual":2,' +
    '"accion_mejora":"accion concreta y medible"}';

  var payload = {
    contents: [{ parts: [{ text: prompt }] }],
    generationConfig: {
      temperature: 0.1,
      maxOutputTokens: 1200,
      responseMimeType: "application/json"
    }
  };

  var options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    var response = UrlFetchApp.fetch(url, options);
    var code = response.getResponseCode();
    var body = response.getContentText();

    if (code !== 200) {
      var errMsg = "HTTP " + code;
      try {
        var errObj = JSON.parse(body);
        errMsg = errObj.error && errObj.error.message ? errObj.error.message.substring(0, 120) : errMsg;
      } catch(_) {}
      return JSON.stringify({ error: errMsg });
    }

    var data = JSON.parse(body);
    var content = data.candidates &&
                  data.candidates[0] &&
                  data.candidates[0].content &&
                  data.candidates[0].content.parts &&
                  data.candidates[0].content.parts[0] &&
                  data.candidates[0].content.parts[0].text;

    if (!content) {
      var reason = (data.candidates && data.candidates[0] && data.candidates[0].finishReason) || "desconocida";
      return JSON.stringify({ error: "Sin contenido. Razon: " + reason });
    }

    // Sanitizar: limpiar markdown y saltos de linea dentro de strings JSON
    var limpio = content
      .replace(/```json/gi, "")
      .replace(/```/g, "")
      .trim();

    limpio = limpio.replace(/("(?:[^"\\]|\\.)*")/g, function(m) {
      return m.replace(/\r\n/g, " ").replace(/\n/g, " ").replace(/\r/g, " ");
    });

    var resultado = JSON.parse(limpio);
    return JSON.stringify(resultado);

  } catch(e) {
    return JSON.stringify({ error: "Error servidor: " + e.toString().substring(0, 100) });
  }
}
