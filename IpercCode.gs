// ============================================================
//  ARCHIVO: Codigo.gs  (pega esto en tu proyecto Apps Script)
//  Este archivo maneja la Web App y el puente con el Sheet
// ============================================================

// ─────────────────────────────────────────
//  doGet — SIRVE LA PÁGINA WEB
//  Esta función es OBLIGATORIA para Web Apps
// ─────────────────────────────────────────
function doGet(e) {
  return HtmlService
    .createHtmlOutputFromFile("Iperc app")   // ← nombre del archivo .html en el proyecto
    .setTitle("IPERC Línea Base — Sistema IA")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag("viewport", "width=device-width, initial-scale=1.0");
}

// ─────────────────────────────────────────
//  doPost — RECIBE ACCIONES DESDE EL HTML
//  El HTML llama a google.script.run.* en su lugar
// ─────────────────────────────────────────

// ─────────────────────────────────────────
//  PUENTE HTML ↔ SHEET
//  Estas funciones son llamadas desde el HTML
//  con: google.script.run.nombreFuncion(params)
// ─────────────────────────────────────────

/** Guarda todas las filas IPERC al Sheet */
function guardarFilasAlSheet(filasJSON) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_DATOS);
  if (!sheet) return { ok: false, error: "Hoja DATOS no encontrada" };

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

      // Fórmulas y colores
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

    // Wrap
    sheet.getRange(fila, 1, 1, 23).setWrap(true).setVerticalAlignment("top");
    sheet.setRowHeight(fila, 85);
  });

  actualizarResumen();
  return { ok: true, filas: filas.filter(f => f.proceso).length };
}

/** Lee filas del Sheet y las devuelve al HTML */
function leerFilasDelSheet() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
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

/** Devuelve la configuración de la matriz al HTML */
function leerMatrizConfigJSON() {
  return JSON.stringify(leerMatrizConfig());
}

/** Guarda la configuración de empresa desde el HTML */
function guardarConfigEmpresaDesdeHTML(cfgJSON) {
  const cfg = JSON.parse(cfgJSON);
  PropertiesService.getScriptProperties().setProperty("CONFIG_EMPRESA", cfgJSON);
  return { ok: true };
}

/** Lee la configuración de empresa */
function leerConfigEmpresa() {
  return PropertiesService.getScriptProperties().getProperty("CONFIG_EMPRESA") || "{}";
}

/** Guarda la API Key */
function guardarAPIKeyDesdeHTML(key) {
  if (!key || key.length < 20) return { ok: false, error: "Clave muy corta" };
  PropertiesService.getScriptProperties().setProperty("GEMINI_API_KEY", key);
  return { ok: true };
}

/** Verifica si hay API Key guardada (devuelve solo si existe, no la key) */
function tieneAPIKey() {
  return !!PropertiesService.getScriptProperties().getProperty("GEMINI_API_KEY");
}


/** Analiza UNA fila con Gemini (llamado fila por fila desde el HTML) */
function analizarFilaConIA(filaJSON) {
  const apiKey = obtenerAPIKey();
  if (!apiKey) return JSON.stringify({ error: "Sin API Key. Configúrala en el menú ⚙️" });

  const f   = JSON.parse(filaJSON);
  const cfg = leerMatrizConfig();

  try {
    const res = llamarGemini(apiKey, cfg, f.proceso, f.area, f.tarea, f.puesto, f.rnr || "R");
    return JSON.stringify(res);
  } catch(e) {
    return JSON.stringify({ error: e.toString() });
  }
}
/**
 * Obtiene la API Key guardada en ScriptProperties.
 * Retorna null si no existe o está vacía.
 */
function obtenerAPIKey() {
  var key = PropertiesService.getScriptProperties().getProperty("GEMINI_API_KEY");
  if (!key || key.trim().length < 20) return null;
  return key.trim();
}

// Helper para calcular nivel por score
function calcularNivelRiesgoByScore(score, cfg) {
  for (const n of cfg.niveles) {
    if (score >= n.desde && score <= n.hasta) return n;
  }
  return cfg.niveles[cfg.niveles.length - 1];
}