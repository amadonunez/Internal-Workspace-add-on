/**
 * Crea una tarjeta de página de inicio para el complemento/auxiliar de Workspace.
 * Esta función construye una tarjeta que se muestra en la página de inicio de Workspace 
 * cuando el add-on está instalado. La tarjeta proporciona una breve descripción 
 * del add-on, enlaces a Gmail y Drive, e instrucciones básicas para su uso.
 *
 * Esta función está referenciada en el archivo `appscript.json` bajo la sección
 * `addOns.home.homepageTrigger` de la siguiente manera:
 *
 * ```json
 * "addOns": {
 *   "home": {
 *     "homepageTrigger": {
 *       "runFunction": "createCommonHomepageCard"
 *     }
 *   }
 * }
 * ```
 *
 * @author mario.estrella@amadotrucking.com
 * @lastModified 2025-02-13 (YYYY-MM-DD)
 * @return {CardService.Card} La tarjeta de página de inicio construida.
 */
function createCommonHomepageCard(e) {
  return CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader().setTitle("Auxiliar Workspace")) // Add-on Name
    .addSection(
      CardService.newCardSection()
        .addWidget(CardService.newDecoratedText()
                .setText("Automatiza tareas en Gmail y Google Drive.").setTopLabel("Propósito").setWrapText(true))
        .addWidget(CardService.newDivider())
        .addWidget(CardService.newTextButton().setText("Abrir Gmail").setOpenLink(CardService.newOpenLink().setUrl("https://mail.google.com")))
        .addWidget(CardService.newTextButton().setText("Abrir Drive").setOpenLink(CardService.newOpenLink().setUrl("https://drive.google.com")))
    )
    .setFixedFooter(standardAddonFooter())
    .build();
}

/**
 * Determina el contexto en el que se está ejecutando el complemento.
 *
 * Esta función analiza el evento (e) para determinar si el complemento se está ejecutando 
 * en Gmail, Drive o en la página de inicio de Workspace.
 *
 * - Si el evento contiene `e.gmail`, se considera contexto de Gmail.
 * - Si el evento contiene `e.drive`, se considera contexto de Drive.
 * - Si no hay evento o no coincide con Gmail o Drive, intenta acceder a GmailApp.getActiveMessage(). 
 *   Si tiene éxito, se considera contexto de Gmail. Si falla (por falta de autorización o 
 *   porque no está en Gmail), se considera contexto de la página de inicio.
 *
 * Esta función es útil para adaptar el comportamiento del complemento según el contexto 
 * en el que se esté utilizando.
 *
 * @author mario.estrella@amadotrucking.com
 * @lastModified 2025-02-13 (YYYY-MM-DD)
 * @param {Object} e El evento que activó la función.
 * @return {string} El contexto del complemento: "gmail", "drive" o "homepage".
 */
function getContext(e) {
  if (e && e.gmail) {  // Check if it's a Gmail add-on event
    return "gmail";
  } else if (e && e.drive) { // Check if it's a Drive add-on event
    return "drive";
  } else {
      // Check if the add-on was opened in a Gmail context (not necessarily by an event)
      // This is a bit of a hack, but it works
      try {
          GmailApp.getActiveMessage(); // Will throw an error if not in Gmail context
          return "gmail";
      } catch (error) {
          // If the error is not related to authorization, it's not a Gmail context
          if (!error.message.includes('Authorization is required to perform that action')) {
              return "homepage"; // Or "other" if you have other contexts
          }
      }

    return "homepage"; // Default to homepage if no specific context is detected

  }
}