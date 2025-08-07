/**
 * @fileoverview Script con funciones para manejar datos de terminales portuarias.
 * Optimizado para evitar repetición de código y datos.
 */

/**
 * Array constante con los datos maestros de las terminales.
 * Definido globalmente para ser accesible por todas las funciones del script.
 * @const {Array<Object>}
 */
const TERMINAL_DATA = Object.freeze([
  {
    port: "Los Angeles",
    terminalName: "APM Terminals Los Angeles (Pier 400)",
    address: "2500 Navy Way, Terminal Island",
    firmsCode: "W185",
    city: "Los Angeles",
    state: "CA",
    postal: "90731",
    country: "US",
  },
  {
    port: "Los Angeles",
    terminalName: "Fenix Marine Services (Pier 300)",
    address: "614 Terminal Way, Terminal Island",
    firmsCode: "Y257",
    city: "Los Angeles",
    state: "CA",
    postal: "90731",
    country: "US",
  },
  {
    port: "Los Angeles",
    terminalName: "Everport Terminal Services (Evergreen)",
    address: "389 Terminal Island Way, Terminal Island",
    firmsCode: "Y124",
    city: "Los Angeles",
    state: "CA",
    postal: "90731",
    country: "US",
  },
  {
    port: "Los Angeles",
    terminalName: "TraPac Los Angeles",
    address: "630 West Harry Bridges Blvd, Wilmington",
    firmsCode: "Y258",
    city: "Los Angeles",
    state: "CA",
    postal: "90744",
    country: "US",
  },
  {
    port: "Los Angeles",
    terminalName: "West Basin Container Terminal",
    address: "2050 John S. Gibson Blvd, San Pedro",
    firmsCode: "Y773",
    city: "Los Angeles",
    state: "CA",
    postal: "90731",
    country: "US",
  },
  {
    port: "Los Angeles",
    terminalName: "Yusen Terminals Inc. (YTI)",
    address: "701 New Dock Street, Terminal Island",
    firmsCode: "Y790",
    city: "Los Angeles",
    state: "CA",
    postal: "90731",
    country: "US",
  },
  {
    port: "Long Beach",
    terminalName: "International Transportation Service (ITS)",
    address: "1281 Pier G Way, Long Beach",
    firmsCode: "Y309",
    city: "Long Beach",
    state: "CA",
    postal: "90802",
    country: "US",
  },
  {
    port: "Long Beach",
    terminalName: "Long Beach Container Terminal (LBCT) Pier F",
    address: "201 S. Pico Avenue",
    firmsCode: "W183",
    city: "Long Beach",
    state: "CA",
    postal: "90802",
    country: "US",
  },
  {
    port: "Long Beach",
    terminalName: "Long Beach Container Terminal (LBCT) Pier E (Middle Harbor)",
    address: "201 South Pico Avenue",
    firmsCode: "WAC8",
    city: "Long Beach",
    state: "CA",
    postal: "90802",
    country: "US",
  },
  {
    port: "Long Beach",
    terminalName: "Pacific Container Terminal (PCT)",
    address: "1521 Pier J Avenue",
    firmsCode: "W182",
    city: "Long Beach",
    state: "CA",
    postal: "90802",
    country: "US",
  },
  {
    port: "Long Beach",
    terminalName: "SSA Marine Terminal (Pier A)",
    address: "700 Pier A Plaza",
    firmsCode: "Z978",
    city: "Long Beach",
    state: "CA",
    postal: "90813",
    country: "US",
  },
  {
    port: "Long Beach",
    terminalName: "Matson Terminal (Pier C)",
    address: "1320 Pier C Street",
    firmsCode: "Z611",
    city: "Long Beach",
    state: "CA",
    postal: "90802",
    country: "US",
  },
  {
    port: "Long Beach",
    terminalName: "Total Terminals International (TTI)",
    address: "301 Mediterranean Ave",
    firmsCode: "Z952",
    city: "Long Beach",
    state: "CA",
    postal: "90731",
    country: "US",
  },
]);
// const TERMINAL_DATA = Object.freeze([
//   { port: "Los Angeles", terminalName: "APM Terminals Los Angeles (Pier 400)", address: "2500 Navy Way, Terminal Island", firmsCode: "W185", city: "Los Angeles" ,state: "CA",postal: "90731",country: "US"  },
//   { port: "Los Angeles", terminalName: "Fenix Marine Services (Pier 300)", address: "614 Terminal Way, Terminal Island", firmsCode: "Y257", city: "Los Angeles",state: "CA",postal: "90731",country: "US" },
//   { port: "Los Angeles", terminalName: "Everport Terminal Services (Evergreen)", address: "389 Terminal Island Way, Terminal Island", firmsCode: "Y124", city: "Los Angeles",state: "CA",postal: "90731",country: "US"  },
//   { port: "Los Angeles", terminalName: "TraPac Los Angeles", address: "630 West Harry Bridges Blvd, Wilmington", firmsCode: "Y258", city: "Los Angeles",state: "CA",postal: "90744",country: "US"  },
//   { port: "Los Angeles", terminalName: "West Basin Container Terminal", address: "2050 John S. Gibson Blvd, San Pedro", firmsCode: "Y773", city: "Los Angeles", state: "CA",postal: "90731",country: "US" },
//   { port: "Los Angeles", terminalName: "Yusen Terminals Inc. (YTI)", address: "701 New Dock Street, Terminal Island", firmsCode: "Y790", city: "Los Angeles" , state: "CA",postal: "90731",country: "US"},
//   { port: "Long Beach", terminalName: "International Transportation Service (ITS)", address: "1281 Pier G Way, Long Beach", firmsCode: "Y309", city: "Long Beach", state: "CA",postal: "90802",country: "US"},
//   { port: "Long Beach", terminalName: "Long Beach Container Terminal (LBCT) Pier F", address: "201 S. Pico Avenue", firmsCode: "W183", city: "Long Beach",  state: "CA",postal: "90802",country: "US" },
//   { port: "Long Beach", terminalName: "Long Beach Container Terminal (LBCT) Pier E (Middle Harbor)", address: "201 South Pico Avenue", firmsCode: "WAC8", city: "Long Beach", state: "CA",postal: "90802",country: "US"},  
//   { port: "Long Beach", terminalName: "Pacific Container Terminal (PCT)", address: "1521 Pier J Avenue", firmsCode: "W182", city: "Long Beach", state: "CA",postal: "90802",country: "US" },
//   { port: "Long Beach", terminalName: "SSA Marine Terminal (Pier A)", address: "700 Pier A Plaza", firmsCode: "Z978", city: "Long Beach", state: "CA",postal: "90813",country: "US" },
//   { port: "Long Beach", terminalName: "Matson Terminal (Pier C)", address: "1320 Pier C Street", firmsCode: "Z611", city: "Long Beach",state: "CA",postal: "90802",country: "US"  },    
//   { port: "Long Beach", terminalName: "Total Terminals International (TTI)", address: "301 Mediterranean Ave", firmsCode: "Z952", city: "Long Beach",state: "CA",postal: "90731",country: "US" },

// ]);

/**
 * Array constante con todos los Firms Codes válidos, derivado de TERMINAL_DATA.
 * Se calcula una sola vez al cargar el script.
 * @const {Array<string>}
 */
const VALID_FIRMS_CODES = Object.freeze(TERMINAL_DATA.map(terminal => terminal.firmsCode));

// --- Funciones Utilitarias ---

/**
 * Busca un firmsCode válido dentro de un string de texto.
 * Utiliza la lista precalculada VALID_FIRMS_CODES.
 *
 * @param {string} inputText El string donde buscar el firmsCode.
 * @return {string|null} El primer firmsCode válido encontrado en el texto, o null si no se encuentra ninguno.
 * @customfunction
 */
function getFirmsCodeFromString(inputText) {
  // Verificamos si el inputText es realmente un string
  if (typeof inputText !== 'string' || !inputText) {
    // Logger.log('La entrada proporcionada no es un string válido.'); // Opcional: Log para depuración
    return null;
  }

  // Iteramos sobre cada firmsCode válido precalculado
  for (const code of VALID_FIRMS_CODES) {
    // Verificamos si el string de entrada contiene el firmsCode actual
    // Usamos includes() que es sensible a mayúsculas/minúsculas
    if (inputText.includes(code)) {
      return code; // Si lo encuentra, lo devuelve inmediatamente
    }
  }

  // Si el bucle termina sin encontrar ningún firmsCode válido
  return null; // Devuelve null si no se encontró ninguna coincidencia
}

/**
 * Obtiene el objeto completo de la terminal basado en su Firms Code.
 * Reutiliza la constante global TERMINAL_DATA.
 *
 * @param {string} firmsCode El Firms Code a buscar.
 * @return {Object|null} El objeto de la terminal si se encuentra, o null si no existe.
 */
function getTerminalByFirmsCode(firmsCode) {
  if (!firmsCode || typeof firmsCode !== 'string') {
    return null; // Entrada inválida
  }
  // Usamos el método find() que es eficiente para encontrar el primer elemento que cumple la condición
  const terminal = TERMINAL_DATA.find(terminal => terminal.firmsCode === firmsCode);
  return terminal || null; // Retorna el objeto encontrado o null si find() devuelve undefined
}

// --- Ejemplos de Uso ---

function testTerminalFunctions() {
  Logger.log("--- Pruebas getFirmsCodeFromString ---");
  const texto1 = "HANJIN SHIPPING CO, BERTHS T132-140 Z952.";
  const texto2 = "Por favor revisar el estado en WAC8 - Long Beach.";
  const texto3 = "Este texto no contiene un código válido.";
  const texto4 = "Aquí hay múltiples códigos W185 y también Y790.";
  const texto5 = null; // Entrada inválida

  const code1 = getFirmsCodeFromString(texto1);
  const code2 = getFirmsCodeFromString(texto2);
  const code3 = getFirmsCodeFromString(texto3);
  const code4 = getFirmsCodeFromString(texto4);
  const code5 = getFirmsCodeFromString(texto5);

  Logger.log(`Texto 1: "${texto1}" -> FirmsCode: ${code1}`); // Esperado: Y257
  Logger.log(`Texto 2: "${texto2}" -> FirmsCode: ${code2}`); // Esperado: WAC8
  Logger.log(`Texto 3: "${texto3}" -> FirmsCode: ${code3}`); // Esperado: null
  Logger.log(`Texto 4: "${texto4}" -> FirmsCode: ${code4}`); // Esperado: W185 (el primero en TERMINAL_DATA que coincide)
  Logger.log(`Texto 5: "${texto5}" -> FirmsCode: ${code5}`); // Esperado: null

  Logger.log("\n--- Pruebas getTerminalByFirmsCode ---");
  const terminal1 = getTerminalByFirmsCode(code1); // Usa el resultado de la prueba anterior (Y257)
  const terminal2 = getTerminalByFirmsCode("Z978");
  const terminal3 = getTerminalByFirmsCode("CODIGO_INEXISTENTE");
  const terminal4 = getTerminalByFirmsCode(null); // Entrada inválida

  Logger.log(`Terminal para ${code1}: ${terminal1 ? terminal1.terminalName : 'No encontrada'}`); // Esperado: Fenix Marine Services (Pier 300)
  Logger.log(`Terminal para Z978: ${terminal2 ? terminal2.terminalName : 'No encontrada'}`);   // Esperado: SSA Marine Terminal (Pier A)
  Logger.log(`Terminal para CODIGO_INEXISTENTE: ${terminal3 ? terminal3.terminalName : 'No encontrada'}`); // Esperado: No encontrada
  Logger.log(`Terminal para null: ${terminal4 ? terminal4.terminalName : 'No encontrada'}`); // Esperado: No encontrada

  if (terminal1) {
     Logger.log(`Dirección de ${terminal1.terminalName}: ${terminal1.address}`);
  }
}