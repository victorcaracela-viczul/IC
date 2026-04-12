/* ============================================================================
   🎮 CONFIGURACIÓN DE IDs (CEREBRO DEL SISTEMA)
   ============================================================================ */

// 1️⃣ ID DE LA HOJA DE PERSONAL (SSOMA) - Para validar DNI
const ID_HOJA_SSOMA  = "1NDDHlTfWxmObgm8JZu5WAnCECB3gU6e_k7o_sFcMrkw"; 

// 2️⃣ ID DE LA HOJA DEL JUEGO (Preguntas, Temas, Registros)
const ID_HOJA_KAHOOT = "1POBqECeHddeOcYbJR4zIPQoaorFkjl8E8cIMIEy-jGA"; 


/* ============================================================================
   🚀 FUNCIONES PRINCIPALES
   ============================================================================ */

function doGet() {
  return HtmlService.createHtmlOutputFromFile("KAHOOT")
    .setTitle("Portal de Capacitación SSOMA")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// 🛠️ FUNCIÓN DE INSTALACIÓN (Solo se usa si las hojas no existen)
function configurarSistema() {
  const ss = SpreadsheetApp.openById(ID_HOJA_KAHOOT);
  
  // 1. TEMAS
  let hojaTemas = ss.getSheetByName("Temas");
  if (!hojaTemas) {
    hojaTemas = ss.insertSheet("Temas");
    hojaTemas.appendRow(["TÍTULO CAPACITACIÓN", "Descripción", "Estado", "ImagenURL", "FechaCreacion"]);
    hojaTemas.appendRow(["Demo Seguridad", "Prueba del sistema", "ACTIVO", "https://cdn-icons-png.flaticon.com/512/1165/1165674.png", new Date()]);
  }
  
  // 2. PREGUNTAS
  let hojaPreguntas = ss.getSheetByName("Preguntas");
  if (!hojaPreguntas) hojaPreguntas = ss.insertSheet("Preguntas");
  
  // 3. REGISTROS (Actualizado con columna ALIAS)
  let hojaRegistros = ss.getSheetByName("Registros");
  if (!hojaRegistros) {
    hojaRegistros = ss.insertSheet("Registros");
    // ID | USUARIO | ALIAS | CAPACITACIÓN | PUNTAJE | TIEMPO | FECHA | HORA
    hojaRegistros.appendRow(["ID_INTENTO", "Usuario", "Alias", "CAPACITACIÓN", "Puntaje", "Tiempo", "Fecha", "Hora"]);
  }

  // 4. ACCESO
  let hojaAcceso = ss.getSheetByName("Acceso");
  if (!hojaAcceso) {
    hojaAcceso = ss.insertSheet("Acceso");
    hojaAcceso.appendRow(["Usuario", "Contraseña"]);
    hojaAcceso.appendRow(["admin", "123456"]);
  }
}

/* ─── 🔐 1. VALIDACIÓN POR DNI (TU LÓGICA BLINDADA) ─── */
function validarUsuarioPorDNI(dniInput) {
  try {
    const dniBuscado = String(dniInput).trim(); 
    
    // Abrir Hoja SSOMA (Externa)
    const ss = SpreadsheetApp.openById(ID_HOJA_SSOMA);
    const hoja = ss.getSheetByName("PERSONAL");
    
    if (!hoja) return { exito: false, error: "No existe hoja PERSONAL" };

    // Leer Datos: DNI (Col B) y NOMBRE (Col C)
    const ultFila = hoja.getLastRow();
    if (ultFila < 2) return { exito: false, error: "Base de datos vacía" };

    const datos = hoja.getRange(2, 2, ultFila - 1, 2).getValues(); 

    // Buscar DNI
    const usuarioEncontrado = datos.find(fila => String(fila[0]).trim() === dniBuscado);

    if (usuarioEncontrado) {
      const nombreCompleto = String(usuarioEncontrado[1]).trim().toUpperCase();
      
      // Crear Alias (Primer nombre)
      let partes = nombreCompleto.split(" ");
      let alias = partes[0];
      if (partes.length > 1 && partes[0].length < 3) alias += " " + partes[1]; 

      return { exito: true, nombreReal: nombreCompleto, alias: alias };
    } else {
      return { exito: false };
    }

  } catch (e) {
    return { exito: false, error: e.toString() };
  }
}

/* ─── 📊 2. OBTENER TEMAS (DASHBOARD) ─── */
function obtenerDatosUsuario(nombreReal) {
  try {
    const ss = SpreadsheetApp.openById(ID_HOJA_KAHOOT);
    const hojaTemas = ss.getSheetByName("Temas");
    const hojaRegistros = ss.getSheetByName("Registros");
    
    if (!hojaTemas || hojaTemas.getLastRow() < 2) return [];
    
    // Leer Temas
    const temas = hojaTemas.getRange(2, 1, hojaTemas.getLastRow()-1, 5).getValues();
    
    // Leer Registros
    let registros = [];
    if (hojaRegistros && hojaRegistros.getLastRow() > 1) {
      registros = hojaRegistros.getRange(2, 1, hojaRegistros.getLastRow()-1, 8).getValues();
    }
    
    const hoy = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");
    const usuarioNormalizado = String(nombreReal).trim().toUpperCase();

    return temas.map(t => {
      const titulo = t[0]; 
      
      // Verificamos si ya jugó hoy
      const jugoHoy = registros.some(r => {
        if(!r[6]) return false; // r[6] es la FECHA (Columna G)
        try {
          let uReg = String(r[1]).trim().toUpperCase(); // Col B: Usuario
          let tReg = r[3]; // Col D: Título (Desplazado por Alias)
          let fReg = Utilities.formatDate(new Date(r[6]), Session.getScriptTimeZone(), "dd/MM/yyyy");
          return uReg == usuarioNormalizado && tReg == titulo && fReg == hoy;
        } catch(e) { return false; }
      });

      // Calculamos mejor puntaje histórico
      const misIntentos = registros.filter(r => String(r[1]).trim().toUpperCase() == usuarioNormalizado && r[3] == titulo);
      const mejorPuntaje = misIntentos.length > 0 ? Math.max(...misIntentos.map(r => Number(r[4]) || 0)) : "-";
      
      let estado = "DISPONIBLE";
      if (t[2] !== "ACTIVO") estado = "CERRADO";
      else if (jugoHoy) estado = "COMPLETADO_HOY";
      
      return { id: titulo, titulo: titulo, desc: t[1], estado: estado, img: t[3], mejorPuntaje: mejorPuntaje };
    });
  } catch (e) {
    throw new Error("Error en Backend: " + e.message);
  }
}

/* ─── 🎲 3. OBTENER PREGUNTAS (ALEATORIAS) ─── */
function obtenerPreguntasJuego(titulo) {
  const ss = SpreadsheetApp.openById(ID_HOJA_KAHOOT);
  const hoja = ss.getSheetByName("Preguntas");
  
  if (!hoja || hoja.getLastRow() < 2) return [];
  
  const numCols = hoja.getLastColumn();
  if (numCols < 7) return []; 
  
  const datos = hoja.getRange(2, 1, hoja.getLastRow()-1, numCols).getValues();
  
  var preguntas = datos.filter(fila => fila[0] === titulo).map(fila => {
    var opcionesOriginales = [fila[2], fila[3], fila[4], fila[5]];
    var respuestaCorrectaTexto = opcionesOriginales[(Number(fila[6]) || 1) - 1];
    
    // Mezclar opciones
    var opcionesMezcladas = opcionesOriginales.slice().sort(() => Math.random() - 0.5);
    var nuevoIndiceRespuesta = opcionesMezcladas.indexOf(respuestaCorrectaTexto);
    
    return {
      pregunta: fila[1], 
      opciones: opcionesMezcladas, 
      respuesta: nuevoIndiceRespuesta,
      gifCorrecto: fila[7] || "", 
      gifIncorrecto: fila[8] || ""
    };
  });
  
  // Mezclar orden de preguntas
  return preguntas.sort(() => Math.random() - 0.5);
}

/* ─── 💾 4. GUARDAR INTENTO (CON ALIAS Y VALIDACIÓN) ─── */
function guardarIntento(usuario, alias, titulo, aciertos, total, tiempo) {
  const ss = SpreadsheetApp.openById(ID_HOJA_KAHOOT);
  let hoja = ss.getSheetByName("Registros");
  if(!hoja) hoja = ss.insertSheet("Registros");
  
  const fecha = new Date();
  const hoyStr = Utilities.formatDate(fecha, Session.getScriptTimeZone(), "dd/MM/yyyy");
  
  // --- VALIDACIÓN ANTI-DOBLE JUEGO ---
  if (hoja.getLastRow() > 1) {
    const datosReg = hoja.getRange(2, 1, hoja.getLastRow()-1, 7).getValues();
    const yaJugo = datosReg.some(r => {
      try {
        let uReg = String(r[1]).trim().toUpperCase();
        let tReg = r[3]; // Col D: Título
        let fReg = Utilities.formatDate(new Date(r[6]), Session.getScriptTimeZone(), "dd/MM/yyyy");
        return uReg === usuario.toUpperCase() && tReg === titulo && fReg === hoyStr;
      } catch(e) { return false; }
    });
    
    if (yaJugo) {
      return { puntaje: Math.round((aciertos / total) * 100), ranking: obtenerRanking(titulo) };
    }
  }
  
  const puntaje = total > 0 ? Math.round((aciertos / total) * 100) : 0;
  
  // Guardamos: ID | USUARIO | ALIAS | TITULO | PUNTAJE | TIEMPO | FECHA | HORA
  hoja.appendRow([
    Utilities.getUuid(), 
    usuario.toUpperCase(), 
    alias,                 
    titulo,                
    puntaje, 
    tiempo, 
    fecha, 
    Utilities.formatDate(fecha, Session.getScriptTimeZone(), "HH:mm:ss")
  ]);
  
  return { puntaje: puntaje, ranking: obtenerRanking(titulo) };
}

/* ─── 🏆 5. RANKING MEJORADO (PUNTOS → TIEMPO → INTENTOS) ─── */
function obtenerRanking(titulo) {
  const ss = SpreadsheetApp.openById(ID_HOJA_KAHOOT);
  const hoja = ss.getSheetByName("Registros");
  if (!hoja || hoja.getLastRow() < 2) return [];
  
  // Leemos hasta la columna 6 (Tiempo = F)
  // Estructura: A=ID, B=User, C=Alias, D=Titulo, E=Ptos, F=Tiempo
  const datos = hoja.getRange(2, 1, hoja.getLastRow()-1, 6).getValues();
  const intentosTema = datos.filter(r => r[3] === titulo); // Col D (3) es Título
  
  const estadisticas = {};
  
  intentosTema.forEach(r => {
    const nombreReal = String(r[1]).toUpperCase(); // Col B
    const alias = r[2]; // Col C (Alias)
    const p = Number(r[4]); // Col E (Puntaje)
    const t = Number(r[5]); // Col F (Tiempo)
    
    if (!estadisticas[nombreReal]) {
      estadisticas[nombreReal] = { 
        nombreMostrar: alias, 
        mejorPuntaje: 0, 
        mejorTiempo: 999999,
        intentos: 0 
      };
    }
    
    // Contar intentos
    estadisticas[nombreReal].intentos++;
    
    // Actualizar mejor puntaje y tiempo
    // Si tiene mejor puntaje, actualiza todo
    if (p > estadisticas[nombreReal].mejorPuntaje) {
      estadisticas[nombreReal].mejorPuntaje = p;
      estadisticas[nombreReal].mejorTiempo = t;
      estadisticas[nombreReal].nombreMostrar = alias;
    } 
    // Si empata en puntos pero tiene mejor tiempo
    else if (p === estadisticas[nombreReal].mejorPuntaje && t < estadisticas[nombreReal].mejorTiempo) {
      estadisticas[nombreReal].mejorTiempo = t;
      estadisticas[nombreReal].nombreMostrar = alias;
    }
  });
  
  // Convertir a array y ordenar correctamente:
  // 1° PUNTOS (mayor a menor)
  // 2° TIEMPO (menor a mayor - desempate)
  // 3° INTENTOS (solo informativo, no afecta ranking)
  return Object.values(estadisticas)
    .map(u => ({ 
      nombre: u.nombreMostrar, 
      puntos: u.mejorPuntaje, 
      tiempo: u.mejorTiempo,
      intentos: u.intentos 
    }))
    .sort((a, b) => {
      // Primero por PUNTOS (descendente)
      if (b.puntos !== a.puntos) return b.puntos - a.puntos;
      // Si empatan en puntos, por TIEMPO (ascendente - el más rápido gana)
      return a.tiempo - b.tiempo;
    })
    .slice(0, 10);
}

/* ─── ⚙️ 6. FUNCIONES ADMIN ─── */
function verificarCredencialesAdmin(u, p) {
  const ss = SpreadsheetApp.openById(ID_HOJA_KAHOOT);
  const hoja = ss.getSheetByName("Acceso");
  if(!hoja) return false;
  const datos = hoja.getDataRange().getValues();
  for(let i=1; i<datos.length; i++) if(String(datos[i][0])===u && String(datos[i][1])===p) return true;
  return false;
}

function adminObtenerTemas() {
  const ss = SpreadsheetApp.openById(ID_HOJA_KAHOOT);
  const s = ss.getSheetByName("Temas");
  if(!s || s.getLastRow() < 2) return [];
  return s.getRange(2, 1, s.getLastRow()-1, 3).getValues().map((r, i) => ({ fila: i + 2, titulo: r[0], estado: r[2] }));
}

function adminCambiarEstado(f, e) {
  const ss = SpreadsheetApp.openById(ID_HOJA_KAHOOT);
  ss.getSheetByName("Temas").getRange(f, 3).setValue(e);
}

/* ─── 🏆 7. RANKING DETALLADO (CON FANTASMAS) ─── */
function obtenerRankingDetallado(titulo) {
  const ssJuego = SpreadsheetApp.openById(ID_HOJA_KAHOOT);
  const ssPersonal = SpreadsheetApp.openById(ID_HOJA_SSOMA);
  
  const hojaReg = ssJuego.getSheetByName("Registros");
  const hojaUsers = ssPersonal.getSheetByName("PERSONAL"); 
  
  // 1. PROCESAR RANKING (Jugadores)
  let rankingFinal = [];
  let jugaronSet = new Set(); 

  if (hojaReg && hojaReg.getLastRow() > 1) {
    const datos = hojaReg.getRange(2, 1, hojaReg.getLastRow()-1, 6).getValues();
    const intentosTema = datos.filter(r => r[3] === titulo);
    
    const stats = {};

    intentosTema.forEach(r => {
      const realName = String(r[1]).trim().toUpperCase();
      const alias = r[2];
      const puntos = Number(r[4]);
      const tiempo = Number(r[5]);

      jugaronSet.add(realName); 

      if (!stats[realName]) {
        stats[realName] = { nombre: alias, mejorPuntaje: 0, mejorTiempo: 999999, intentos: 0 };
      }
      stats[realName].intentos++;
      
      // Lógica de mejor puntaje
      if (puntos > stats[realName].mejorPuntaje) {
        stats[realName].mejorPuntaje = puntos;
        stats[realName].mejorTiempo = tiempo;
        stats[realName].nombre = alias;
      } else if (puntos === stats[realName].mejorPuntaje && tiempo < stats[realName].mejorTiempo) {
        stats[realName].mejorTiempo = tiempo;
      }
    });

    rankingFinal = Object.values(stats).sort((a, b) => {
      if (b.mejorPuntaje !== a.mejorPuntaje) return b.mejorPuntaje - a.mejorPuntaje;
      return a.mejorTiempo - b.mejorTiempo;
    });
  }

  // 2. DETECTAR FANTASMAS (Quienes faltan)
  let faltantes = [];
  if (hojaUsers && hojaUsers.getLastRow() > 1) {
    const todosUsuarios = hojaUsers.getRange(2, 3, hojaUsers.getLastRow()-1, 1).getValues();
    
    todosUsuarios.forEach(u => {
      let nombreDb = String(u[0]).trim();
      if (nombreDb && !jugaronSet.has(nombreDb.toUpperCase())) {
        let partes = nombreDb.split(" ");
        let corto = partes[0];
        if(partes.length > 1) corto += " " + partes[1].charAt(0) + ".";
        faltantes.push(corto);
      }
    });
  }

  return {
    ranking: rankingFinal,
    faltantes: faltantes
  };
}