/**
 * Crea una tarjeta de página de inicio para el complemento/auxiliar de Gmail.
 * Esta función construye una tarjeta que se muestra en la página de inicio de Gmail 
 * cuando el add-on está instalado. La tarjeta proporciona una breve descripción 
 * del add-on e instrucciones básicas para su uso.
 *
 * Esta función está referenciada en el archivo `appscript.json` bajo la sección
 * `addOns.common.homepageTrigger` de la siguiente manera:
 *
 * ```json
 * "addOns": {
 *   "common": {
 *     "homepageTrigger": {
 *       "runFunction": "createGmailHomepageCard"
 *     }
 *   }
 * }
 * ```
 *
 * @author mario.estrella@amadotrucking.com
 * @lastModified 2025-02-13 (YYYY-MM-DD)
 * @return {CardService.Card} La tarjeta de página de inicio construida.
 */
function createGmailHomepageCard(e) {
  return[
    createGenericCard()
  ]
}



/**
 * Crea una tarjeta en la interfaz de Gmail que permite al usuario interactuar con el correo electrónico actual.
 * Esta función está configurada en el archivo `appscript.json` para que se active automáticamente
 * cuando se abre un correo electrónico en Gmail (trigger contextual unconditional).
 *
 * Esta función realiza las siguientes acciones:
 *
 * 1. **Obtiene el ID del mensaje y el token de acceso** del evento de Gmail.
 * 2. **Autentica la aplicación** utilizando el token de acceso.
 * 3. **Obtiene el hilo del mensaje** a partir del ID del mensaje.
 * 4. **Extrae información del mensaje**, como el asunto, el remitente y el cuerpo.
 * 5. **Determina qué tipo de tarjeta mostrar** en función de criterios específicos, como el asunto del correo electrónico o el remitente.
 * 6. **Crea la tarjeta correspondiente**:
 *    - Si el asunto coincide con "Salsas Castillo", crea una tarjeta con casillas de verificación para los archivos adjuntos.
 *    - Si el remitente es de un dominio específico, crea una tarjeta personalizada con información relevante.
 *    - Si el asunto coincide con "Otro Asunto Importante", crea otra tarjeta personalizada.
 *    - Si no se cumple ningún criterio, crea una tarjeta genérica.
 * 7. **Devuelve la tarjeta creada** para que se muestre en la interfaz de Gmail.
 *
 * En caso de error, registra el mensaje de error en el registro de la aplicación.
 *
 * @author mario.estrella@amadotrucking.com
 * @lastModified 2025-02-13 (YYYY-MM-DD)
 * 
 * @param {Gmail.GmailThread} thread El hilo de Gmail.
 * @return {CardService.Card} La tarjeta.
 */
// function createOpenGmailCard(e) {

//   var messageId = e.gmail.messageId;
//   var accessToken = e.gmail.accessToken;
//   GmailApp.setCurrentMessageAccessToken(accessToken);

//   try {
//     const accessToken = e.gmail.accessToken;
//     Logger.log("Access Token: " + accessToken);

//     if (!accessToken) {
//       throw new Error("Access token is missing!");
//     }

//     GmailApp.setCurrentMessageAccessToken(accessToken);
//     var message = GmailApp.getMessageById(messageId);
//     var subject = message.getThread().getFirstMessageSubject();

//     const threadId = e.gmail.threadId;
//     Logger.log("Thread ID: " + threadId);

//     if (!threadId) {
//       throw new Error("Thread ID is missing!");
//     }

//     const thread = message.getThread();

//     if (!thread) {
//       throw new Error("Failed to fetch the thread.");
//     }

//     Logger.log("Thread: " + thread.getId());

//     const messages = thread.getMessages();
//     Logger.log("Messages: " + JSON.stringify(messages));
//     Logger.log("First email: " + JSON.stringify(messages[0]))

//     const senderEmail = messages[0].getFrom();
//     Logger.log("Sender Email: " + senderEmail);
//     const actualEmail = extractEmailAddress(senderEmail);



//   // Criterios (en orden de prioridad)
//   const criteria = [
//     {
//       condition: /Salsas Castillo/i.test(subject),
//       cardFunction: () => createCardWithAttachmentCheckboxesAndEmailOption(thread) // Updated function call
//     },
//     {
//       condition: /T0/i.test(subject), // Checks if "T0" appears in the subject
//       cardFunction: () => createCardWithAttachmentbuttonILS(thread)
//     },
//     {
//       condition: /LCS/i.test(subject), // Checks if "LCS" appears in the subject
//       cardFunction: () => createCardWithAttachmentbuttonTaylor(thread)

//     },
//     {
//       condition: /Berlin/i.test(subject), // Checks if "Berlin" appears in the subject
//       cardFunction: () => createCardWithAttachmentbuttonBerlin(thread)
//     },
//     {
//       condition: /Otro Asunto Importante/.test(subject),
//       cardFunction: () => createCustomCard(thread, "Otro Asunto", "Este correo tiene otro asunto importante.")
//     },
//   ];

//  // Encuentra el primer criterio que se cumple
//     const matchingCriterion = criteria.find(c => c.condition);

//     // Llama a la función de tarjeta correspondiente o la genérica
//     let card;
//     if (matchingCriterion) {
//       Logger.log("Criterio encontrado, llamando a: " + matchingCriterion.cardFunction.toString());
//       card = matchingCriterion.cardFunction();
//     } else {
//       Logger.log("Ningún criterio coincide, creando tarjeta genérica.");
//       card = createGenericCard(thread); // Asumiendo que createGenericCard existe y puede aceptar thread
//     }

//     return [card];
//   } catch (error) {
//     Logger.log("Error: " + error.message);
//   }

// }


/**
 * Crea una tarjeta en la interfaz de Gmail que permite al usuario interactuar con el correo electrónico actual.
 *
 * [El resto de la documentación sigue siendo la misma...]
 */
function createOpenGmailCard(e) {

  var messageId = e.gmail.messageId;
  var accessToken = e.gmail.accessToken;
  GmailApp.setCurrentMessageAccessToken(accessToken);

  try {
    const accessToken = e.gmail.accessToken;

    if (!accessToken) {
      throw new Error("Access token is missing!");
    }

    GmailApp.setCurrentMessageAccessToken(accessToken);
    var message = GmailApp.getMessageById(messageId);
    var subject = message.getThread().getFirstMessageSubject();

    const threadId = e.gmail.threadId;

    if (!threadId) {
      throw new Error("Thread ID is missing!");
    }

    const thread = message.getThread();

    if (!thread) {
      throw new Error("Failed to fetch the thread.");
    }

    const messages = thread.getMessages();
    const senderEmail = messages[0].getFrom();
    const actualEmail = extractEmailAddress(senderEmail);

    // *** NUEVO CÓDIGO: Obtener las etiquetas del hilo ***
    const labels = thread.getLabels();
    const labelNames = labels.map(label => label.getName()); // Extraer nombres de etiquetas
    Logger.log("Labels: " + JSON.stringify(labelNames));

    // Criterios (en orden de prioridad)
    const criteria = [
      {
        condition: labelNames.includes("Salsas Castillo"), // Verifica si la etiqueta "Salsas Castillo" está presente
        cardFunction: () => createCardWithAttachmentCheckboxesAndEmailOption(thread) // Updated function call
      },
      {
        condition: /T0/i.test(subject), // Checks if "T0" appears in the subject
        cardFunction: () => createCardWithAttachmentbuttonILS(thread)
      },
      {
        condition: /LCS/i.test(subject), // Checks if "LCS" appears in the subject
        cardFunction: () => createCardWithAttachmentbuttonTaylor(thread)

      },
      {
        condition: /Berlin/i.test(subject), // Checks if "Berlin" appears in the subject
        cardFunction: () => createCardWithAttachmentbuttonBerlin(thread)
      },
      {
        condition: /Otro Asunto Importante/.test(subject),
        cardFunction: () => createCustomCard(thread, "Otro Asunto", "Este correo tiene otro asunto importante.")
      },
    ];

    // Encuentra el primer criterio que se cumple
    const matchingCriterion = criteria.find(c => c.condition);

    // Llama a la función de tarjeta correspondiente o la genérica
    let card;
    if (matchingCriterion) {
      Logger.log("Criterio encontrado, llamando a: " + matchingCriterion.cardFunction.toString());
      card = matchingCriterion.cardFunction();
    } else {
      Logger.log("Ningún criterio coincide, creando tarjeta genérica.");
      card = createGenericCard(thread); // Asumiendo que createGenericCard existe y puede aceptar thread
    }

    return [card];
  } catch (error) {
    Logger.log("Error: " + error.message);
  }

}




/**
 * Crea la tarjeta genérica que se muestra cuando el asunto del correo electrónico no coincide con ningún criterio específico.
 * Esta tarjeta proporciona una descripción general de las funcionalidades del complemento y un botón para iniciar el procesamiento de correos electrónicos.
 *
 * La tarjeta incluye un encabezado, una sección de descripción con un icono de correo, una sección de actividades con una lista de tareas y sus iconos (aunque los iconos no se muestran directamente en la documentación), una sección con un botón para iniciar el procesamiento de correos electrónicos y un pie de página estándar.
 *
 * Ejemplo de uso:
 * ```javascript
 * function mostrarTarjetaGenerica() {
 *   return createGenericCard();
 * }
 * ```
 *
 * @author mario.estrella@amadotrucking.com
 * @lastModified 2024-02-13 (YYYY-MM-DD)
 * @return {CardService.Card} La tarjeta genérica.
 */
function createGenericCard() {
  return CardService.newCardBuilder()
  .setHeader(CardService.newCardHeader().setTitle("Auxiliar para Gmail  -  Version 1.0"))
  .setPeekCardHeader(CardService.newCardHeader()
    .setTitle("Auxiliar para correos y cadenas")
    .setSubtitle("Volver al auxilar de tareas para correos y cadenas"))
  .addSection(
      CardService.newCardSection()
      .addWidget(CardService.newDecoratedText()
        .setStartIcon(CardService.newIconImage()
          .setMaterialIcon(CardService.newMaterialIcon()
            .setName('mail')
            .setWeight(500)))
        .setTopLabel("Descripción")
        .setText("Automatiza tareas de procesamiento de correo-e")
        .setWrapText(true))
    )
  .addSection(
      CardService.newCardSection()
      .addWidget(CardService.newTextParagraph().setText("Actividades:"))
      .addWidget(CardService.newTextParagraph().setText("• Procesamiento para SALSAS CASTILLO"))
      .addWidget(CardService.newTextParagraph().setText("• Creacion de ordenes ILS"))
      .addWidget(CardService.newTextParagraph().setText("• Creacion de ordenes Taylor"))
      .addWidget(CardService.newTextParagraph().setText("• Creacion de ordenes Berlin"))
    )
  .setFixedFooter(standardAddonFooter())
  .build();
}

/**
 * Crea una tarjeta personalizada con un título y mensaje.
 *
 * Esta función genera una tarjeta con un título y un mensaje especificados. 
 * La tarjeta se basa en un constructor de tarjetas base (`createBaseCardBuilder`) 
 * que presumiblemente se define en otro lugar del código.
 *
 * Ejemplo de uso:
 *
 * ```javascript
 * const miTarjeta = createCustomCard(thread, "Título de la tarjeta", "Este es el mensaje de la tarjeta.");
 * ```
 *
 * @author mario.estrella@amadotrucking.com
 * @lastModified 2024-02-13 (YYYY-MM-DD)
 * @param {Gmail.GmailThread} thread El hilo de Gmail.
 * @param {string} title El título de la tarjeta.
 * @param {string} message El mensaje de la tarjeta.
 * @return {CardService.Card} La tarjeta.
 */
function createCustomCard(thread, title, message) {
  return BaseCardBuilder(`Auxiliar para Gmail (${title})`)
  .addSection(CardService.newCardSection().addWidget(CardService.newTextParagraph().setText(message)))
  .build();
}

/**
 * Crea un constructor de tarjetas base con el encabezado y el pie de página comunes.
 *
 * Esta función genera un objeto `CardBuilder` preconfigurado con un encabezado que contiene el título 
 * especificado y un pie de página estándar (`standardAddonFooter`).  Este constructor de tarjetas base 
 * se puede utilizar para crear tarjetas personalizadas añadiendo más contenido (secciones, widgets, etc.).
 *
 * Ejemplo de uso:
 *
 * ```javascript
 * const cardBuilder = createBaseCardBuilder("Título de la tarjeta");
 * cardBuilder.addSection(CardService.newCardSection().addWidget(...)); // Añadir contenido a la tarjeta
 * const tarjeta = cardBuilder.build(); // Construir la tarjeta final
 * ```
 *
 * @author Tu Nombre (o el nombre del autor)
 * @lastModified 2024-08-28 (YYYY-MM-DD)
 * @param {string} title El título de la tarjeta.
 * @return {CardService.CardBuilder} El constructor de tarjetas.
 */
function BaseCardBuilder(title) {
  return CardService.newCardBuilder()
  .setHeader(CardService.newCardHeader().setTitle(title))
  .setPeekCardHeader(CardService.newCardHeader()
    .setTitle("Auxiliar para correos y cadenas")
    .setSubtitle("Volver al auxilar de tareas para correos y cadenas"))
  .setFixedFooter(standardAddonFooter());
}