// /**
//  * Retrieves terminal information based on a FIRMS code.
//  *
//  * @param {string} firmsCode The FIRMS code to search for.
//  * @returns {object|null} An object containing the terminal name and address details
//  *                         if found, or null if no matching terminal is found.
//  *                         Returns null if input is invalid.
//  */
// function getTerminalInfo(firmsCode) {
//   // Input validation.
//   if (typeof firmsCode !== 'string' || firmsCode.trim() === '') {
//     return null;
//   }

//   const searchFirmsCode = firmsCode.trim().toUpperCase();

//   const terminalData = [
//     { port: "Los Angeles", terminalName: "APM Terminals Los Angeles (Pier 400)", address: "2500 Navy Way, Terminal Island, CA 90731", firmsCode: "W185", city: "Los Angeles" }, //Added city
//     { port: "Los Angeles", terminalName: "Fenix Marine Services (Pier 300)", address: "614 Terminal Way, Terminal Island, CA 90731", firmsCode: "Y257", city: "Los Angeles" },
//     { port: "Los Angeles", terminalName: "Everport Terminal Services (Evergreen)", address: "389 Terminal Island Way, Terminal Island, CA 90731", firmsCode: "Y124", city: "Los Angeles" },
//     { port: "Los Angeles", terminalName: "TraPac Los Angeles", address: "630 West Harry Bridges Blvd, Wilmington, CA 90744", firmsCode: "Y258", city: "Los Angeles" },
//     { port: "Los Angeles", terminalName: "West Basin Container Terminal", address: "2050 John S. Gibson Blvd, San Pedro, CA 90731", firmsCode: "Y773", city: "Los Angeles" },
//     { port: "Los Angeles", terminalName: "Yusen Terminals Inc. (YTI)", address: "701 New Dock Street, Terminal Island, CA 90731", firmsCode: "Y790", city: "Los Angeles" },
//     { port: "Long Beach", terminalName: "International Transportation Service (ITS)", address: "1281 Pier G Way, Long Beach, CA 90802", firmsCode: "Y309", city: "Long Beach" },
//     { port: "Long Beach", terminalName: "Long Beach Container Terminal (LBCT) – Pier F", address: "201 S. Pico Avenue, Long Beach, CA 90802", firmsCode: "W183", city: "Long Beach" },
//     { port: "Long Beach", terminalName: "Long Beach Container Terminal (LBCT) – Pier E (Middle Harbor)", address: "201 South Pico Avenue, Long Beach, CA 90802", firmsCode: "WAC8", city: "Long Beach" },
//     { port: "Long Beach", terminalName: "Pacific Container Terminal (PCT)", address: "1521 Pier J Avenue, Long Beach, CA 90802", firmsCode: "W182", city: "Long Beach" },
//     { port: "Long Beach", terminalName: "SSA Marine Terminal (Pier A)", address: "700 Pier A Plaza, Long Beach, CA 90813", firmsCode: "Z978", city: "Long Beach" },
//     { port: "Long Beach", terminalName: "Matson Terminal (Pier C)", address: "1320 Pier C Street, Long Beach, CA 90802", firmsCode: "Z611", city: "Long Beach" },
//     { port: "Long Beach", terminalName: "Total Terminals International (TTI)", address: "301 Mediterranean Ave, Long Beach, CA 90802", firmsCode: "Z952", city: "Long Beach" },
//   ];

//   const terminal = terminalData.find(item => item.firmsCode === searchFirmsCode);

//   if (terminal) {
//     return parseAddress(terminal);  // Use the helper function.
//   } else {
//     return null;
//   }
// }


// /**
//  * Parses the address string into its components.
//  *
//  * @param {object} terminal The terminal object containing the full address string.
//  * @returns {object} An object containing the parsed address components.
//  */
// function parseAddress(terminal) {
//     const addressString = terminal.address;
//     const addressParts = addressString.split(','); // Split by comma.

//     //Robust Address Parsing
//     let streetAddress = "";
//     let city = "";
//     let state = "";
//     let zipCode = "";

//     // Extract ZIP code (last part)
//     if (addressParts.length > 0) {
//       const lastPart = addressParts[addressParts.length - 1].trim();
//       const zipMatch = lastPart.match(/\d{5}(-\d{4})?/); // Match 5 or 5-4 digit ZIP
//       if (zipMatch) {
//         zipCode = zipMatch[0];
//         addressParts[addressParts.length - 1] = lastPart.replace(zipCode, '').trim(); // Remove ZIP from last part
//       }
//     }

//    // Extract state (now part of the last element)
//     if (addressParts.length > 0) {
//         const lastPart = addressParts[addressParts.length - 1].trim();
//         const stateMatch = lastPart.match(/[A-Z]{2}/); // Match two capital letters
//         if (stateMatch) {
//             state = stateMatch[0];
//             addressParts[addressParts.length - 1] = lastPart.replace(state, '').trim(); //Remove the state
//         }
//     }
    
//     // Get city from 'port' information (more reliable)
//     city = terminal.city;
    
//     // Remaining parts are the street address
//     streetAddress = addressParts.join(',').trim();

//     return {
//       terminalName: terminal.terminalName,
//       streetAddress: streetAddress,
//       city: city,
//       state: state,
//       zipCode: zipCode,
//       country: "USA"  // Add the country.
//     };
// }

// // Example Usage (for testing - keeps previous examples and adds new ones)
// function testGetTerminalInfo() {
//   // Test case 1: Valid FIRMS code (exists)
//   const terminal1 = getTerminalInfo("Y258");
//   console.log("Test 1 (Valid - Y258):", terminal1);

//   // Test case 2: Valid FIRMS code (exists, different case)
//   const terminal2 = getTerminalInfo("w185");
//   console.log("Test 2 (Valid - w185):", terminal2);

//   // Test case 3: Valid FIRMS code with extra spaces
//   const terminal3 = getTerminalInfo("  Z952  ");
//   console.log("Test 3 (Valid with spaces - Z952):", terminal3);

//   // Test case 4: Invalid FIRMS code (does not exist)
//   const terminal4 = getTerminalInfo("XXXX");
//   console.log("Test 4 (Invalid - XXXX):", terminal4);

//   // Test case 5: Invalid input (empty string)
//   const terminal5 = getTerminalInfo("");
//   console.log("Test 5 (Invalid - empty):", terminal5);

//   // Test case 6: Invalid input (null)
//   const terminal6 = getTerminalInfo(null);
//   console.log("Test 6 (Invalid - null):", terminal6);

//     // Test case 7: Invalid input (number)
//   const terminal7 = getTerminalInfo(1234);
//   console.log("Test 7 (Invalid - number):", terminal7);

//       // Test case 8: Invalid input (undefined)
//   const terminal8 = getTerminalInfo(undefined);
//   console.log("Test 8 (Invalid - undefined):", terminal8);

//     // Test case 9: WAC8 - checking a different state
//   const terminal9 = getTerminalInfo("WAC8");
//   console.log("Test 9 (WAC8) - edge case:", terminal9);
// }


// /**
//  * Busca un firmsCode válido dentro de un string de texto.
//  * El firmsCode debe existir en la lista predefinida de terminalData.
//  *
//  * @param {string} inputText El string donde buscar el firmsCode.
//  * @return {string|null} El primer firmsCode válido encontrado en el texto, o null si no se encuentra ninguno.
//  * @customfunction
//  */
// function getFirmsCodeFromString(inputText) {
//   // Definimos el array con los datos de las terminales
//   const terminalData = [
//     { port: "Los Angeles", terminalName: "APM Terminals Los Angeles (Pier 400)", address: "2500 Navy Way, Terminal Island, CA 90731", firmsCode: "W185", city: "Los Angeles" },
//     { port: "Los Angeles", terminalName: "Fenix Marine Services (Pier 300)", address: "614 Terminal Way, Terminal Island, CA 90731", firmsCode: "Y257", city: "Los Angeles" },
//     { port: "Los Angeles", terminalName: "Everport Terminal Services (Evergreen)", address: "389 Terminal Island Way, Terminal Island, CA 90731", firmsCode: "Y124", city: "Los Angeles" },
//     { port: "Los Angeles", terminalName: "TraPac Los Angeles", address: "630 West Harry Bridges Blvd, Wilmington, CA 90744", firmsCode: "Y258", city: "Los Angeles" },
//     { port: "Los Angeles", terminalName: "West Basin Container Terminal", address: "2050 John S. Gibson Blvd, San Pedro, CA 90731", firmsCode: "Y773", city: "Los Angeles" },
//     { port: "Los Angeles", terminalName: "Yusen Terminals Inc. (YTI)", address: "701 New Dock Street, Terminal Island, CA 90731", firmsCode: "Y790", city: "Los Angeles" },
//     { port: "Long Beach", terminalName: "International Transportation Service (ITS)", address: "1281 Pier G Way, Long Beach, CA 90802", firmsCode: "Y309", city: "Long Beach" },
//     { port: "Long Beach", terminalName: "Long Beach Container Terminal (LBCT) – Pier F", address: "201 S. Pico Avenue, Long Beach, CA 90802", firmsCode: "W183", city: "Long Beach" },
//     { port: "Long Beach", terminalName: "Long Beach Container Terminal (LBCT) – Pier E (Middle Harbor)", address: "201 South Pico Avenue, Long Beach, CA 90802", firmsCode: "WAC8", city: "Long Beach" },
//     { port: "Long Beach", terminalName: "Pacific Container Terminal (PCT)", address: "1521 Pier J Avenue, Long Beach, CA 90802", firmsCode: "W182", city: "Long Beach" },
//     { port: "Long Beach", terminalName: "SSA Marine Terminal (Pier A)", address: "700 Pier A Plaza, Long Beach, CA 90813", firmsCode: "Z978", city: "Long Beach" },
//     { port: "Long Beach", terminalName: "Matson Terminal (Pier C)", address: "1320 Pier C Street, Long Beach, CA 90802", firmsCode: "Z611", city: "Long Beach" },
//     { port: "Long Beach", terminalName: "Total Terminals International (TTI)", address: "301 Mediterranean Ave, Long Beach, CA 90802", firmsCode: "Z952", city: "Long Beach" },
//   ];

//   // Extraemos solo los firmsCode válidos a un array
//   const validFirmsCodes = terminalData.map(terminal => terminal.firmsCode);

//   // Verificamos si el inputText es realmente un string
//   if (typeof inputText !== 'string') {
//     Logger.log('La entrada proporcionada no es un string.');
//     return null; // O podrías lanzar un error si prefieres
//   }

//   // Iteramos sobre cada firmsCode válido
//   for (const code of validFirmsCodes) {
//     // Verificamos si el string de entrada contiene el firmsCode actual
//     if (inputText.includes(code)) {
//       return code; // Si lo encuentra, lo devuelve inmediatamente
//     }
//   }

//   // Si el bucle termina sin encontrar ningún firmsCode válido
//   return null; // Devuelve null si no se encontró ninguna coincidencia
// }

// // --- Ejemplos de Uso ---
// function testGetFirmsCode() {
//   const texto1 = "La información del contenedor para Y257 está lista.";
//   const texto2 = "TRANSPAC CONT SVS (WIL 136/139) Y258";
//   const texto3 = "Este texto no contiene un código válido.";
//   const texto4 = "Aquí hay múltiples códigos W185 y también Y790.";
//   const texto5 = 12345; // No es un string

//   Logger.log(`Texto 1: "${texto1}" -> FirmsCode: ${getFirmsCodeFromString(texto1)}`); // Esperado: Y257
//   Logger.log(`Texto 2: "${texto2}" -> FirmsCode: ${getFirmsCodeFromString(texto2)}`); // Esperado: WAC8
//   Logger.log(`Texto 3: "${texto3}" -> FirmsCode: ${getFirmsCodeFromString(texto3)}`); // Esperado: null
//   Logger.log(`Texto 4: "${texto4}" -> FirmsCode: ${getFirmsCodeFromString(texto4)}`); // Esperado: W185 (devuelve el primero que encuentra en la lista `terminalData`)
//   Logger.log(`Texto 5: "${texto5}" -> FirmsCode: ${getFirmsCodeFromString(texto5)}`); // Esperado: null
// }


