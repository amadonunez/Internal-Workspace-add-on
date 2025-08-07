/**
 * Crea un pie de página estándar para las tarjetas del complemento.
 *
 * Esta función genera un pie de página fijo que contiene un botón con un enlace a la página de instrucciones del complemento.  El enlace a las instrucciones se define mediante la constante `INSTRUCTIONS_URL`.
 *
 * Ejemplo de uso:
 *
 * function miTarjeta() {
 *   var builder = CardService.newCardBuilder();
 *   // ... (resto de la construcción de la tarjeta) ...
 *   builder.setFixedFooter(standardAddonFooter());
 *   return builder.build();
 * }
  
 * @author mario.estrella@amadotrucking.com
 * @lastModified 2025-02-13 (YYYY-MM-DD)
 * @return {CardService.FixedFooter} El pie de página fijo. --- 
 */
function standardAddonFooter(){
  const INSTRUCTIONS_URL = "https://sites.google.com/amadotrucking.com/intranet/workspace-add-on-help-page"

  return CardService.newFixedFooter()
              .setPrimaryButton(
                    CardService.newTextButton().setText("Instrucciones").setOpenLink(CardService.newOpenLink().setUrl(INSTRUCTIONS_URL))
                )
}




/**
 * Extrae la dirección de correo electrónico de una cadena que puede contener el nombre y la dirección.
 *
 * Esta función utiliza una expresión regular para extraer la dirección de correo electrónico que se encuentra 
 * dentro de los corchetes angulares ("<" y ">") de una cadena que representa una dirección de correo electrónico.
 * Si no se encuentran corchetes angulares, la función devuelve la cadena original sin modificar. Esto permite
 * manejar casos en los que la cadena solo contiene la dirección de correo electrónico sin el nombre.
 *
 * Ejemplo de uso:
 *
 * ```javascript
 * const direccionCompleta1 = "Nombre Apellido <[email address removed]>";
 * const direccion1 = extractEmailAddress(direccionCompleta1); // Devuelve "[email address removed]"
 *
 * const direccionCompleta2 = "[email address removed]";
 * const direccion2 = extractEmailAddress(direccionCompleta2); // Devuelve "[email address removed]"
 * ```
 *
 * @author mario.estrella@amadotrucking.com
 * @lastModified 2024-02-13 (YYYY-MM-DD)
 * @param {string} fullAddress La cadena que contiene la dirección de correo electrónico (puede incluir el nombre).
 * @return {string} La dirección de correo electrónico extraída.
 */
function extractEmailAddress(fullAddress) {
  const match = fullAddress.match(/<([^>]+)>/); // Extracts the content inside "< >"
  return match ? match[1] : fullAddress; // If no match, return the full string (handles cases where no "< >")
}
