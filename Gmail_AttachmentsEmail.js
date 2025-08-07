
/**
 * Crea una tarjeta con casillas de verificación para los archivos adjuntos PDF, 
 * limitando los widgets a un máximo de MAX_WIDGETS.
 *
 * Esta función genera una tarjeta que muestra casillas de verificación para cada archivo adjunto PDF 
 * encontrado en un hilo de Gmail. Permite al usuario seleccionar los archivos PDF que desea 
 * y luego enviarlos por correo electrónico. 
 *
 * Se incluye un campo de entrada de texto para que el usuario especifique la dirección de correo 
 * electrónico del destinatario.
 *
 * La tarjeta también incluye un botón "Enviar el correo-e con los archivos seleccionados" que, al 
 * hacer clic, ejecuta la función `sendSelectedPDFsEmail`.
 *
 * Ejemplo de uso:
 *
 * ```javascript
 * function mostrarTarjetaConAdjuntos(thread) {
 *   return createCardWithAttachmentCheckboxesAndEmailOption(thread);
 * }
 * ```
 *
 * @author mario.estrella@amadotrucking.com
 * @lastModified 2024-02-13 (YYYY-MM-DD)
 * @param {Gmail.GmailThread} thread El hilo de Gmail.
 * @return {CardService.Card} La tarjeta construida.
 */
function createCardWithAttachmentCheckboxesAndEmailOption(thread) {
  const attachmentsSection = createAttachmentsSection(thread);

  // Agregar un campo de entrada para la dirección de correo electrónico del destinatario
  attachmentsSection.addWidget(
    CardService.newTextInput()
      .setTitle("Destinatario: ")
      .setValue("dirección de correo-e")
      .setValidation(CardService.newValidation().setInputType(CardService.InputType.EMAIL))
      .setFieldName("recipient_email") // Nombre del campo de entrada de correo electrónico
  );

  const anAction = CardService.newAction().setFunctionName('sendSelectedPDFsEmail');

  // NEW: Add a text input for the email subject
  attachmentsSection.addWidget(
    CardService.newTextInput()
      .setTitle("Asunto:") // Subject line field
      .setFieldName("email_subject")
      .setValue("se adjuntan documentos solicitados") // Set a default subject
  );

  // Botón para enviar los archivos seleccionados
  attachmentsSection.addWidget(
    CardService.newTextButton()
      .setText("Enviar")
      .setMaterialIcon(CardService.newMaterialIcon().setName('send').setWeight(500))
      .setOnClickAction(anAction)
  );

  return CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader().setTitle("Auxiliar para Gmail "))
    .setPeekCardHeader(
      CardService.newCardHeader()
        .setTitle("Auxiliar para correos y cadenas")
        .setSubtitle("Volver al auxiliar de tareas para correos y cadenas")
    )
    .addSection(attachmentsSection)
    .setFixedFooter(standardAddonFooter())
    .build();
}

/**
 * Crea una sección de tarjeta con una lista de archivos adjuntos en formato PDF e imágenes.
 *
 * @param {GmailThread} thread - El hilo de correo electrónico del cual obtener los archivos adjuntos.
 * @return {CardService.CardSection} - Una sección de tarjeta con casillas de verificación para seleccionar archivos adjuntos.
 */
function createAttachmentsSection(thread) {
  const MAX_WIDGETS = 80;
  let messages = thread.getMessages();
  let attachmentsSection = CardService.newCardSection().addWidget(CardService.newTextParagraph().setText("Adjuntos"));
  let widgetCount = 0;
  let pdfAttachments = [];

  for (let message of messages) {
    let attachments = message.getAttachments();
    for (let attachment of attachments) {
      if (
        (attachment.getContentType().indexOf("application/pdf") !== -1 || attachment.getContentType().indexOf("image/") !== -1) &&
        !attachment.getName().startsWith("Outlook-") // Filtrar archivos adjuntos creados por Outlook
      ) {
        if (widgetCount < MAX_WIDGETS) {
          pdfAttachments.push({
            title: attachment.getName(),
            value: attachment.getName(),
            selected: true
          });
          widgetCount++;
        } else {
          attachmentsSection.addWidget(CardService.newTextParagraph().setText("Too many PDF attachments to display."));
          break;
        }
      }
    }
    if (widgetCount >= MAX_WIDGETS) {
      break;
    }
  }

  if (pdfAttachments.length > 0) {
    let checkboxGroup = CardService.newSelectionInput()
      .setType(CardService.SelectionInputType.CHECK_BOX)
      .setTitle('Seleccionar')
      .setFieldName('pdf_attachments'); // Removed setOnChangeAction

    pdfAttachments.forEach(item => {
      checkboxGroup.addItem(item.title, item.value, item.selected);
    });

    attachmentsSection.addWidget(checkboxGroup);
  } else {
    attachmentsSection.addWidget(CardService.newTextParagraph().setText("No PDF attachments found."));
  }

  return attachmentsSection;
}

/**
 * Envía los archivos PDF seleccionados por correo electrónico.
 *
 * Esta función se ejecuta cuando el usuario hace clic en el botón "Enviar el correo-e con los archivos seleccionados" 
 * en la tarjeta creada por `createCardWithAttachmentCheckboxes`. 
 *
 * Obtiene la lista de archivos PDF seleccionados del evento (`e`), 
 * construye un correo electrónico con los archivos adjuntos y lo envía al usuario actual.
 *
 * @author mario.estrella@amadotrucking.com
 * @lastModified 2024-02-13 (YYYY-MM-DD)
 * @param {Object} e El evento del formulario de la tarjeta.
 */
function sendSelectedPDFsEmail(e) {

  var messageId = e.gmail.messageId;
  var accessToken = e.gmail.accessToken;
  GmailApp.setCurrentMessageAccessToken(accessToken);
  const threadId = e.gmail.threadId;
  var message = GmailApp.getMessageById(messageId);
  const thread = message.getThread();

  const formInputs = e.commonEventObject.formInputs;

  // Corrected line: Get the selected PDF filenames from the form input
  const selectedPDFs = formInputs.pdf_attachments.stringInputs.value || {}; 

  if (selectedPDFs.length > 0) {
    let recipient = Session.getActiveUser().getEmail();
    let subject = formInputs.email_subject ? formInputs.email_subject.stringInputs.value[0] : "Selected PDF Attachments";

    let emailBody = "The following PDF attachments were selected:\n\n";

    // Corrected line: Initialize the attachments array
    let attachments = []; 

    if (Array.isArray(selectedPDFs)) {
      selectedPDFs.forEach(pdfName => {
        emailBody += pdfName + "\n";
        let attachment = thread.getMessages().flatMap(message => {
          return message.getAttachments().filter(att => att.getName() === pdfName);
        });
        attachments.push(...attachment);
      });
    } else {
      emailBody += selectedPDFs + "\n";
      let attachment = thread.getMessages().flatMap(message => {
        return message.getAttachments().filter(att => att.getName() === selectedPDFs);
      });
      attachments.push(...attachment);
    }

  // Get the recipient email
  const recipientEmail = formInputs.recipient_email ? formInputs.recipient_email.stringInputs.value[0] : "";
  Logger.log("Recipient Email: " + recipientEmail);

  if (!recipientEmail) {
    Logger.log("No recipient email provided.");
    return;
  }

    GmailApp.sendEmail(recipientEmail, subject, emailBody, { attachments: attachments });
  } else {
    Logger.log("No PDF attachments selected.");
  }
}

/********************************************************************************************/
/**Amado N.
 * Crea una tarjeta con una sección que muestra botones para los archivos PDF adjuntos,
 * un encabezado, un pie de página y una acción asociada.
 *
 * @param {GoogleAppsScript.Gmail.GmailThread} thread El hilo de correo electrónico del cual extraer los adjuntos.
 * @return {GoogleAppsScript.CardService.Card} La tarjeta construida con los elementos especificados.
 */
function createCardWithAttachmentbuttonILS(thread) {
  // Crea la sección de la tarjeta que mostrará los archivos PDF adjuntos como botones.
  // Utiliza la función 'createAttachmentsSectionILS' para generar esta sección.
  const attachmentsSection = createAttachmentsSectionILS(thread);

  // Crea una acción que se ejecutará cuando se haga clic en un botón específico.
  // En este caso, la acción llamará a la función 'sendSelectedPDFsEmail'.
  const anAction = CardService.newAction().setFunctionName('sendSelectedPDFsEmail');

  // Construye la tarjeta principal utilizando un CardBuilder.
  return CardService.newCardBuilder()
    // Establece el encabezado principal de la tarjeta.
    .setHeader(CardService.newCardHeader().setTitle("Auxiliar para crear órdenes en Rose Rocket con guías de ILS"))
    // Establece el encabezado de la tarjeta que se muestra cuando la tarjeta está minimizada (peek card).
    .setPeekCardHeader(
      CardService.newCardHeader()
        .setTitle("Auxiliar para correos y cadenas") // Título del encabezado de la tarjeta minimizada.
        .setSubtitle("Volver al auxiliar de tareas para correos y cadenas") // Subtítulo del encabezado de la tarjeta minimizada.
    )
  .addSection(
      CardService.newCardSection()
      .addWidget(CardService.newTextParagraph().setText("<b>Instrucciones:</b>"))
      .addWidget(CardService.newTextParagraph().setText("Dar click en la guía de ILS para crear la orden en Rose Rocket."))
      .addWidget(CardService.newTextParagraph().setText("La orden sera creada en Rose Rocket y recibiras un correo-e con el número de la orden y la liga para abrirla."))

    )    
    // Agrega la sección de archivos adjuntos a la tarjeta.
    .addSection(attachmentsSection)
    // Establece un pie de página fijo en la parte inferior de la tarjeta.
    // Utiliza la función 'standardAddonFooter' para generar el pie de página estándar.
    .setFixedFooter(standardAddonFooter())
    // Construye la tarjeta final y la retorna.  **IMPORTANTE: Llama a build()!**
    .build();
}
/**Amado N.
 * Crea una sección de tarjeta que muestra los archivos PDF adjuntos en un hilo de correo electrónico
 * como botones, permitiendo al usuario hacer clic para procesar cada archivo.
 *
 * @param {GoogleAppsScript.Gmail.GmailThread} thread El hilo de correo electrónico del cual extraer los adjuntos.
 * @return {GoogleAppsScript.CardService.CardSection} La sección de la tarjeta con los botones de los adjuntos.
 */
function createAttachmentsSectionILS(thread) {
  // Define el número máximo de widgets (botones) a mostrar en la sección para evitar sobrecargar la UI.
  const MAX_WIDGETS = 80;
  // Obtiene todos los mensajes dentro del hilo de correo electrónico.
  try {
    var messages = thread.getMessages();
  } catch (e) {
    Logger.log("Error al obtener los mensajes del hilo: " + e);
    return CardService.newCardSection().addWidget(CardService.newTextParagraph().setText("Error al obtener los mensajes del hilo."));
  }

  // Crea una nueva sección de tarjeta y establece el título como "Adjuntos PDF".
  let attachmentsSection = CardService.newCardSection().addWidget(CardService.newTextParagraph().setText("<b>Archivos</b>"));
  // Inicializa un contador para rastrear el número de widgets agregados a la sección.
  let widgetCount = 0;

  // Itera a través de cada mensaje en el hilo de correo electrónico.
  for (let message of messages) {
    // Obtiene todos los archivos adjuntos del mensaje actual.
    try {
      var attachments = message.getAttachments();
    } catch (e) {
      Logger.log("Error al obtener los adjuntos del mensaje: " + message.getId() + " - " + e);
      // En caso de error al obtener los adjuntos, continúa con el siguiente mensaje.
      continue;
    }
    // Itera a través de cada archivo adjunto.
    for (let attachment of attachments) {
      // Verifica si el archivo adjunto es un archivo PDF y si su nombre no comienza con "Outlook-" (para filtrar archivos creados por Outlook).
      if (
        attachment.getContentType().indexOf("application/pdf") !== -1 &&
        !attachment.getName().startsWith("Outlook-")
      ) {
        // Verifica si el número de widgets en la sección aún no ha alcanzado el máximo permitido.
        if (widgetCount < MAX_WIDGETS) {
          // Obtiene el nombre del archivo adjunto.
          let attachmentName = attachment.getName();

          // Crea una acción que se ejecutará cuando se haga clic en el botón del archivo adjunto.
          let processAction = CardService.newAction()
            .setFunctionName('processPdfAttachment') // Establece el nombre de la función que se llamará al hacer clic en el botón.
            .setParameters({ // Define los parámetros que se pasarán a la función.
              'messageId': message.getId(), // El ID del mensaje que contiene el archivo adjunto.
              'attachmentName': attachmentName // El nombre del archivo adjunto.
            });

          // Crea un botón de texto con el nombre del archivo adjunto como etiqueta y asocia la acción creada.
          let processButton = CardService.newTextButton()
            .setText(attachmentName) // Establece el texto del botón como el nombre del archivo adjunto.
            .setOnClickAction(processAction) // Asocia la acción al evento de clic del botón.
    
               // Agrega el botón a la sección de la tarjeta.
          attachmentsSection.addWidget(processButton);
          // Incrementa el contador de widgets.
          widgetCount++;
        } else {
          // Si se ha alcanzado el número máximo de widgets, agrega un mensaje indicando que hay demasiados archivos adjuntos para mostrar.
          attachmentsSection.addWidget(CardService.newTextParagraph().setText("Demasiados adjuntos PDF para mostrar."));
          // Sale del bucle interno (archivos adjuntos) ya que no se pueden agregar más botones.
          break;
        }
      }
    }
    // Si se ha alcanzado el número máximo de widgets, sale del bucle externo (mensajes) ya que no se pueden agregar más botones.
    if (widgetCount >= MAX_WIDGETS) {
      break;
    }
  }

  // Si no se encontraron archivos PDF, agrega un mensaje indicando que no se encontraron archivos adjuntos.
  if (widgetCount === 0) {
    attachmentsSection.addWidget(CardService.newTextParagraph().setText("No se encontraron adjuntos PDF."));
  }

  // Retorna la sección de la tarjeta con los botones de los archivos adjuntos.
  return attachmentsSection;
}
/*********************************************************************************************/
/**Amado N.
 * Función que se ejecuta cuando el usuario hace clic en un botón de archivo PDF adjunto.
 * Esta función recupera el archivo adjunto del mensaje de correo electrónico y permite su procesamiento.
 *
 * @param {GoogleAppsScript.CardService.ActionResponseEvent} e El evento de acción que contiene los parámetros del botón clicado.
 * @return {GoogleAppsScript.CardService.ActionResponse} Una respuesta de acción para la tarjeta.
 */
function processPdfAttachment(e) {
  try {
    // Obtiene el ID del mensaje y el nombre del archivo adjunto de los parámetros del evento.
    let messageId = e.parameters.messageId;
    let attachmentName = e.parameters.attachmentName;

    // Obtiene el mensaje de Gmail utilizando el ID del mensaje.
    try {
      var message = GmailApp.getMessageById(messageId);
    } catch (error) {
      Logger.log("Error al obtener el mensaje con ID: " + messageId + " - " + error);
      return CardService.newActionResponseBuilder()
        .setNotification(CardService.newNotification().setText("Error al obtener el mensaje. ID: " + messageId))
        .build();
    }

    // Obtiene todos los archivos adjuntos del mensaje.
    try {
      var attachments = message.getAttachments();
    } catch (error) {
      Logger.log("Error al obtener los adjuntos del mensaje: " + messageId + " - " + error);
      return CardService.newActionResponseBuilder()
        .setNotification(CardService.newNotification().setText("Error al obtener los adjuntos del mensaje."))
        .build();
    }

    // Busca el archivo adjunto específico por su nombre.
    let attachment = attachments.find(att => att.getName() === attachmentName);

    // Si no se encuentra el archivo adjunto, registra un error y devuelve una notificación al usuario.
    if (!attachment) {
      Logger.log("Archivo adjunto no encontrado: " + attachmentName);
      return CardService.newActionResponseBuilder()
        .setNotification(CardService.newNotification().setText("Archivo adjunto no encontrado."))
        .build();
    }

    let scriptProperties = PropertiesService.getScriptProperties();
    let customerID = scriptProperties.getProperty('ILS-CustomerId');

    let processor = PropertiesService.getScriptProperties();
    let apiEndpoint = processor.getProperty('ILS-processor');

    var data = ProcesadorAIdeDocumentos.processPdfWithDocumentAI(attachment, apiEndpoint);
    var responseAI = ProcesadorAIdeDocumentos.transformar_AI_a_RoseRocket(customerID, data);
    var responseRR = ProcesadorAIdeDocumentos.createRoseRocketOrder(customerID, responseAI);

    var orderId = responseRR.order.id;

    var uploadResult = RoseRocketGASLibrary.uploadFileToOrder('Amado', orderId, attachment, 'other', attachment.getName());

    var public_id = responseRR.order.public_id;

    let orgCode = PropertiesService.getScriptProperties();
    let orgCodeID = orgCode.getProperty('ILS-CustomerId');

    var link = "https://athn.roserocket.com/#/ops/orders/" + orderId;

    // Send email to the active user
    let subject = `Rose Rocket order created: ${public_id} for ILS`;
    let body = `Rose Rocket order link: ${link}`;

    sendSimpleMail(subject, body);

    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText("Archivo " + attachmentName + " procesado correctamente."))
      .build();
  } catch (error) {
    // Si ocurre un error durante el procesamiento, registra el error y devuelve una notificación al usuario.
    Logger.log("Error al procesar el archivo adjunto: " + error);
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText("Error al procesar el archivo adjunto: " + error))
      .build();
  }
}

function sendSimpleMail(subjectLine, body){

    const emailRecipient = Session.getActiveUser().getEmail() + ', ils.bot@amadotrucking.com'; 
    const emailSubject = subjectLine;
    let emailBody = body;

      try {
        // Send the email using MailApp service
        MailApp.sendEmail({
          to: emailRecipient,
          subject: emailSubject,
          htmlBody: emailBody,
          //attachments: [attachment] // Enviar el archivo
        });
        Logger.log(`Email sent successfully to ${emailRecipient}`);
      } catch (error) {
        Logger.log(`Error sending email: ${error}`);
      }

}

/**
 * Sends an HTML email to the active user.
 * 
 * @param {string} subjectLine - The subject of the email.
 * @param {string} htmlBody - The HTML content of the email.
 */
function sendHtmlEmailToActiveUser(subjectLine, htmlBody) {
    const userEmail = Session.getActiveUser().getEmail();

    if (!userEmail) {
        Logger.log("No active user email found.");
        return;
    }

    try {
        // Construct the raw email message
        const emailContent =
            "From: " + userEmail + "\r\n" +
            "To: " + userEmail + "\r\n" +
            "Subject: " + subjectLine + "\r\n" +
            "MIME-Version: 1.0\r\n" +
            "Content-Type: text/html; charset=UTF-8\r\n\r\n" +
            htmlBody;

        // Encode in Base64 as required by Gmail API
        const encodedEmail = Utilities.base64EncodeWebSafe(emailContent);

        // Get OAuth token for authorization
        const oauthToken = ScriptApp.getOAuthToken();

        // Define the Gmail API endpoint
        const url = "https://www.googleapis.com/gmail/v1/users/me/messages/send";

        // API request options
        const options = {
            method: "post",
            contentType: "application/json",
            headers: {
                Authorization: "Bearer " + oauthToken
            },
            payload: JSON.stringify({ raw: encodedEmail }),
            muteHttpExceptions: true
        };

        // Send the request
        const response = UrlFetchApp.fetch(url, options);
        const responseCode = response.getResponseCode();
        const responseBody = response.getContentText();

        if (responseCode === 200) {
            Logger.log(`Email sent successfully to ${userEmail}`);
        } else {
            Logger.log(`Error sending email: ${responseBody}`);
        }
    } catch (error) {
        Logger.log("Error sending email via Gmail API: " + error);
    }
}

function testSendHtmlEmailToActiveUser() {
    const testSubject = "Test Email from Google Apps Script";
    const testBody = "<h2>Hello!</h2><p>This is a test email sent via Google Apps Script.</p>";

    try {
        sendHtmlEmailToActiveUser(testSubject, testBody);
        Logger.log("Test email sent successfully.");
    } catch (error) {
        Logger.log("Error sending test email: " + error);
    }
}


/*****************************Taylor********************************/


/**
 * Crea una tarjeta con botones para procesar archivos adjuntos de tipo XLSX
 * @param {GoogleAppsScript.Gmail.GmailThread} thread El hilo de Gmail actual.
 * @returns {GoogleAppsScript.Card.Card} La tarjeta construida.
 */
function createCardWithAttachmentbuttonTaylor(thread) {
  // Crea la sección de la taeta que muestra los archivos XLSX adjuntos como botones.
  const attachmentsSection = createAttachmentsSectionTaylor(thread);

  // Construye la tarjeta principal utilizando un CardBuilder.
  return CardService.newCardBuilder()
    // Establece el encabezado principal de la tarjeta.
    .setHeader(CardService.newCardHeader().setTitle("Auxiliar para crear órdenes en Rose Rocket con guías de Taylor"))
    // Establece el encabezado de la tarjeta que se muestra cuando la tarjeta está minimizada (peek card).
    .setPeekCardHeader(
      CardService.newCardHeader()
        .setTitle("Auxiliar para correos y cadenas") // Título del encabezado de la tarjeta minimizada.
        .setSubtitle("Volver al auxiliar de tareas para correos y cadenas") // Subtítulo del encabezado de la tarjeta minimizada.
    )
    // Agrega una sección con instrucciones para el usuario.
    .addSection(
      CardService.newCardSection()
        .addWidget(CardService.newTextParagraph().setText("<b>Instrucciones:</b>"))
        .addWidget(CardService.newTextParagraph().setText("Dar click en la guía de Taylor para crear la orden en Rose Rocket."))
        .addWidget(CardService.newTextParagraph().setText("La orden sera creada en Rose Rocket y recibiras un correo-e con el número de la orden y la liga para abrirla."))
    )
    // Agrega la sección de archivos adjuntos a la tarjeta.
    .addSection(attachmentsSection)
    // Establece un pie de página fijo en la parte inferior de la tarjeta.
    .setFixedFooter(standardAddonFooter())
    // Construye la tarjeta final y la retorna.  **IMPORTANTE: Llama a build()!**
    .build();
}

/**
 * Procesa un archivo adjunto XLSX seleccionado por el usuario.  Extrae información de origen y destino de la UI y muestra una notificación.
 * @param {GoogleAppsScript.Events.ActionResponseEvent} e El evento de acción que contiene información sobre la solicitud.
 * @returns {GoogleAppsScript.Card.ActionResponse} La respuesta de acción para la tarjeta.
 */

function processXlsxAttachmentTaylor(e) {
  const ATTACHMENT_NOT_FOUND = "Archivo adjunto no encontrado.";
  const ROSE_ROCKET_API_ERROR = "Error al crear la orden en Rose Rocket: ";
  const PROCESSING_ERROR = "Error al procesar el archivo adjunto: ";

  let attachmentName = e.parameters.attachmentName;
  let threadId = e.parameters.threadId;
  Logger.log('processXlsxAttachmentTaylor started');

  try {
    const messageId = e.parameters.messageId;
    const originStorage = e.formInput.origin;
    const destinationStorage = e.formInput.destination;

    if (!originStorage || originStorage.trim() === "") {
      throw new Error("Origen no puede estar vacío.");
    }
    if (!destinationStorage || destinationStorage.trim() === "") {
      throw new Error("Destino no puede estar vacío.");
    }
    let message = GmailApp.getMessageById(messageId);
    Logger.log('Mensaje obtenido exitosamente');

    let attachments = message.getAttachments();
    let attachment = attachments.find(att => att.getName() === attachmentName);

    if (!attachment) {
      Logger.log(`Archivo adjunto no encontrado: ${attachmentName}`);
      return CardService.newActionResponseBuilder()
        .setNotification(CardService.newNotification().setText(ATTACHMENT_NOT_FOUND))
        .build();
    }
    Logger.log(`Archivo adjunto encontrado: ${attachmentName}`);

    const shipmentid = ProcesadorAIdeDocumentos.extractAfterSecondFor(message.getSubject());
    const scriptProperties = PropertiesService.getScriptProperties();
    const customerId = scriptProperties.getProperty('Taylor-customerId'); 
    const instance = scriptProperties.getProperty('Instance.Amado'); 


    const file = DriveApp.createFile(attachment);
    const fileId = file.getId();

    const fileconverted = ProcesadorAIdeDocumentos.convertXlsxToGoogleSheet(fileId);
    const json = ProcesadorAIdeDocumentos.crearJson(fileconverted, "Sheet1");
    const orderData = ProcesadorAIdeDocumentos.creaJsonMultistopTaylor(json, customerId, shipmentid, originStorage, destinationStorage, instance);
    let responseRR;

    try {
       responseRR = RoseRocketGASLibrary.createOrderMultiStopTaylor(instance, customerId, orderData);
      if (!responseRR || !responseRR.multistop_order || !responseRR.multistop_order.id) {
        let errorMessage = "Respuesta inválida de la API de Rose Rocket: ";
        if (responseRR && responseRR.errors) {
          errorMessage += JSON.stringify(responseRR.errors); // Incluye los errores específicos de Rose Rocket
        } else {
          errorMessage += "Estructura de respuesta inesperada.";
        }
        throw new Error(errorMessage);
      }
      
    } catch (rrError) {
      Logger.log(`Error en la API de Rose Rocket: ${rrError.message}\n${rrError.stack}`);
      return CardService.newActionResponseBuilder()
        .setNotification(CardService.newNotification().setText(ROSE_ROCKET_API_ERROR + rrError.message))
        .build();
    }

    const orderid = responseRR.multistop_order.id;
    const public_id = responseRR.multistop_order.public_id;

    const pdfBlob = ProcesadorAIdeDocumentos.convertGoogleSheetToPdf(fileconverted);
    const uploadedfiles = ProcesadorAIdeDocumentos.UploadFileToOrderRR(orderid, pdfBlob);


    var link = "https://athn.roserocket.com/#/ops/orders/" + orderid;
    let subject = `Rose Rocket order created: ${public_id} for TAYLOR`;
    let body = `Rose Rocket order link: ${link}`;
    sendSimpleMail(subject, body)


    const processedFilesString = scriptProperties.getProperty(threadId);
    let processedFiles = processedFilesString ? JSON.parse(processedFilesString) : [];

    if (!processedFiles.includes(attachmentName)) {
      processedFiles.push(attachmentName);
      scriptProperties.setProperty(threadId, JSON.stringify(processedFiles));
      Logger.log(`Nombre del archivo adjunto guardado en ScriptProperties: ${attachmentName}`);
    } else {
      Logger.log(`El nombre del archivo adjunto ya existe en ScriptProperties: ${attachmentName}`);
    }
    //Logger.log('Guardado el nombre del archivo en ScriptProperties (si es necesario)');

    const thread = GmailApp.getThreadById(threadId);
    Logger.log('Thread obtenido exitosamente');

    //const updatedCard = createCardWithAttachmentbuttonTaylor(thread);
   // Logger.log('Nueva tarjeta creada');
    // Intenta liberar el archivo convertido
    try {
      DriveApp.getFileById(fileconverted).setTrashed(true); //Enviado a la papelera
      Logger.log(`Archivo temporal ${fileconverted} enviado a la papelera.`);
    } catch (e) {
      Logger.log(`No se pudo enviar el archivo temporal ${fileconverted} a la papelera: ${e.message}`);
    }
    const actionResponse = CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText(`Archivo ${attachmentName} procesado correctamente.`))
      .build();

    Logger.log('ActionResponse construido');
    return actionResponse;

  } catch (error) {
    Logger.log(`Error al procesar el archivo adjunto: ${error.message}\n${error.stack}`);
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText(PROCESSING_ERROR + error.message))
      .build();
  }
}

/**
 * Extrae los detalles de la organización de los datos JSON proporcionados por la API de Rose Rocket.
 * @param {object} data El objeto JSON que contiene los datos de la organización.
 * @returns {Array<object>} Un array de objetos, donde cada objeto contiene los detalles de una organización. Retorna un array vacío si la estructura del JSON es inesperada o si hay un error.
 */
function extractOrganizationDetails(data) {
  try {
    if (!data?.data?.address_books) {
      Logger.log("La estructura del JSON no es la esperada (extractOrganizationDetails).");
      return []; // Return an empty array if address_books is missing
    }

    return data.data.address_books.map(addressBook => {
      return {
        orgName: addressBook?.org_name || "",
        city: addressBook?.city || "",
        postal: addressBook?.postal || "",
        state: addressBook?.state || "",
        country: addressBook?.country || "",
        address1: addressBook?.address_1 || "",
        storage: `${addressBook?.city || ""} - ${addressBook?.org_name || ""} - ${addressBook?.postal || ""} - ${addressBook?.state || ""} - ${addressBook?.country || ""} - ${addressBook?.address_1 || ""}` // Combines fields into a single string
      };
    });
  } catch (e) {
    Logger.log(`Error al extraer detalles de organizaciones: ${e}`);
    return []; // Return an empty array on error
  }
}



/**
 * Crea la sección de la tarjeta que muestra los archivos adjuntos XLSX del hilo de correo electrónico como botones. Permite seleccionar origen y destino.
 * @param {GoogleAppsScript.Gmail.GmailThread} thread El hilo de Gmail actual.
 * @returns {GoogleAppsScript.Card.CardSection} La sección de la tarjeta construida.
 */
function createAttachmentsSectionTaylor(thread) {
  const MAX_WIDGETS = 80;
  let messages;

  try {
    messages = thread.getMessages();
  } catch (e) {
    Logger.log("Error getting thread messages: " + e);
    return CardService.newCardSection().addWidget(CardService.newTextParagraph().setText("Error al obtener los mensajes del hilo."));
  }

  let attachmentsSection = CardService.newCardSection().setHeader("<b>  -      Stops      -  </b>");
  let widgetCount = 0;
  let organizationDetails = [];

  try {

    var scriptProperties1 = PropertiesService.getScriptProperties();
    var customerId = scriptProperties1.getProperty('Taylor-customerId');

    var getinstance = PropertiesService.getScriptProperties();
    var instance = getinstance.getProperty('Instance.Amado');

    const data = RoseRocketGASLibrary.obtenerDatosAddressBook(instance, customerId, ""); 
   
    organizationDetails = extractOrganizationDetails(data);
    if (organizationDetails.length === 0) {
      Logger.log("No organization details found or error extracting them.");
    }
  } catch (error) {
    Logger.log("Error fetching or extracting organization details: " + error);
  }

  let originSelectionInput = CardService.newSelectionInput()
    .setFieldName("origin")
    .setTitle("Origen")
    .setType(CardService.SelectionInputType.DROPDOWN);

  if (organizationDetails.length > 0) {
    organizationDetails.forEach(org => {
      const label = `${org.orgName} - ${org.city} - ${org.state} - ${org.postal} - ${org.country}`;  // Create a combined label
      originSelectionInput.addItem(label, org.storage, false); // Store the combined string in storage
    });
  } else {
    originSelectionInput.addItem("No organizations found", "none", true);
    originSelectionInput.setDisabled(true);
  }
  attachmentsSection.addWidget(originSelectionInput);

  // **Destino SelectionInput (Dropdown)**
  let destinationSelectionInput = CardService.newSelectionInput()
    .setFieldName("destination")
    .setTitle("Destino")
    .setType(CardService.SelectionInputType.DROPDOWN);

  if (organizationDetails.length > 0) {
    organizationDetails.forEach(org => {
      const label = `${org.orgName} - ${org.city} - ${org.state} - ${org.postal} - ${org.country}`;  // Create a combined label
      destinationSelectionInput.addItem(label, org.storage, false); // Store the combined string in storage
    });
  } else {
    destinationSelectionInput.addItem("No organizations found", "none", true);
    destinationSelectionInput.setDisabled(true);
  }
  attachmentsSection.addWidget(destinationSelectionInput);

  const scriptProperties = PropertiesService.getScriptProperties();
  const threadId = thread.getId();

  const processedFilesString = scriptProperties.getProperty(threadId);
  const processedFiles = processedFilesString ? JSON.parse(processedFilesString) : [];

  // Use a Set for faster lookups of processed files.
  const processedFilesSet = new Set(processedFiles);

  for (let message of messages) {
    let attachments;
    try {
      attachments = message.getAttachments();
    } catch (e) {
      Logger.log("Error getting attachments for message " + message.getId() + ": " + e);
      continue;
    }

    for (let attachment of attachments) {
      let attachmentName = attachment.getName();
      if (
        attachment.getContentType().indexOf("spreadsheetml.sheet") !== -1 &&
        !attachmentName.startsWith("Outlook-")
      ) {
        if (processedFilesSet.has(attachmentName)) {
          // Already processed: show a message instead of a button
          attachmentsSection.addWidget(CardService.newTextParagraph().setText(`<i>${attachmentName}</i> - Procesado ✅`));
        } else {
          if (widgetCount < MAX_WIDGETS) {
            let processAction = CardService.newAction()
              .setFunctionName('processXlsxAttachmentTaylor')
              .setParameters({
                'messageId': message.getId(),
                'attachmentName': attachmentName,
                'threadId': threadId,
                'origin': '', // These get populated from input fields
                'destination': '' // These get populated from input fields
              });

            let processButton = CardService.newTextButton()
              .setText("Crear orden-" + attachmentName)
              .setOnClickAction(processAction);

            attachmentsSection.addWidget(processButton);
            widgetCount++;
          } else {
            attachmentsSection.addWidget(CardService.newTextParagraph().setText("Demasiados adjuntos XLSX para mostrar."));
            break;
          }
        }
      }
    }
    if (widgetCount >= MAX_WIDGETS) {
      break;
    }
  }

  if (widgetCount === 0 && processedFiles.length === 0) {
    attachmentsSection.addWidget(CardService.newTextParagraph().setText("No se encontraron adjuntos XLSX."));
  }
  return attachmentsSection;
}


// /**
//  * Crea la sección de la tarjeta que muestra los archivos adjuntos XLSX del hilo de correo electrónico como botones. Permite seleccionar origen y destino.
//  * @param {GoogleAppsScript.Gmail.GmailThread} thread El hilo de Gmail actual.
//  * @returns {GoogleAppsScript.Card.CardSection} La sección de la tarjeta construida.
//  */
// function createAttachmentsSectionTaylor(thread) {
//   const MAX_WIDGETS = 80;
//   let messages;

//   try {
//     messages = thread.getMessages();
//   } catch (e) {
//     Logger.log("Error getting thread messages: " + e);
//     return CardService.newCardSection().addWidget(CardService.newTextParagraph().setText("Error al obtener los mensajes del hilo."));
//   }

//   let attachmentsSection = CardService.newCardSection().setHeader("<b>  -      Stops      -  </b>");
//   let widgetCount = 0;
//   let organizationDetails = [];

//   try {

//     var scriptProperties1 = PropertiesService.getScriptProperties();
//     var customerId = scriptProperties1.getProperty('Taylor-customerId');

//     var getinstance = PropertiesService.getScriptProperties();
//     var instance = getinstance.getProperty('Instance.Amado');

//     const data = RoseRocketGASLibrary.obtenerDatosAddressBook(instance, customerId, ""); 
   
//     organizationDetails = extractOrganizationDetails(data);
//     if (organizationDetails.length === 0) {
//       Logger.log("No organization details found or error extracting them.");
//     }
//   } catch (error) {
//     Logger.log("Error fetching or extracting organization details: " + error);
//   }

//   let originSelectionInput = CardService.newSelectionInput()
//     .setFieldName("origin")
//     .setTitle("Origen")
//     .setType(CardService.SelectionInputType.DROPDOWN);

//   if (organizationDetails.length > 0) {
//     organizationDetails.forEach(org => {
//       const label = `${org.orgName} - ${org.city} - ${org.state} - ${org.postal} - ${org.country}`;  // Create a combined label
//       originSelectionInput.addItem(label, org.storage, false); // Store the combined string in storage
//     });
//   } else {
//     originSelectionInput.addItem("No organizations found", "none", true);
//     originSelectionInput.setDisabled(true);
//   }
//   attachmentsSection.addWidget(originSelectionInput);

//   // **Destino SelectionInput (Dropdown)**
//   let destinationSelectionInput = CardService.newSelectionInput()
//     .setFieldName("destination")
//     .setTitle("Destino")
//     .setType(CardService.SelectionInputType.DROPDOWN);

//   if (organizationDetails.length > 0) {
//     organizationDetails.forEach(org => {
//       const label = `${org.orgName} - ${org.city} - ${org.state} - ${org.postal} - ${org.country}`;  // Create a combined label
//       destinationSelectionInput.addItem(label, org.storage, false); // Store the combined string in storage
//     });
//   } else {
//     destinationSelectionInput.addItem("No organizations found", "none", true);
//     destinationSelectionInput.setDisabled(true);
//   }
//   attachmentsSection.addWidget(destinationSelectionInput);

//   const scriptProperties = PropertiesService.getScriptProperties();
//   const threadId = thread.getId();

//   const processedFilesString = scriptProperties.getProperty(threadId);
//   const processedFiles = processedFilesString ? JSON.parse(processedFilesString) : [];

//   for (let message of messages) {
//     let attachments;
//     try {
//       attachments = message.getAttachments();
//     } catch (e) {
//       Logger.log("Error getting attachments for message " + message.getId() + ": " + e);
//       continue;
//     }

//     for (let attachment of attachments) {
//       let attachmentName = attachment.getName();
//       if (
//         attachment.getContentType().indexOf("spreadsheetml.sheet") !== -1 &&
//         !attachmentName.startsWith("Outlook-")
//       ) {
//         if (processedFiles.includes(attachmentName)) {
//           // Already processed: show a message instead of a button
//           attachmentsSection.addWidget(CardService.newTextParagraph().setText(`<i>${attachmentName}</i> - Procesado ✅`));
//         } else {
//           if (widgetCount < MAX_WIDGETS) {
//             let processAction = CardService.newAction()
//               .setFunctionName('processXlsxAttachmentTaylor')
//               .setParameters({
//                 'messageId': message.getId(),
//                 'attachmentName': attachmentName,
//                 'threadId': threadId,
//                 'origin': '', // These get populated from input fields
//                 'destination': '' // These get populated from input fields
//               });

//             let processButton = CardService.newTextButton()
//               .setText("Crear orden-" + attachmentName)
//               .setOnClickAction(processAction);

//             attachmentsSection.addWidget(processButton);
//             widgetCount++;
//           } else {
//             attachmentsSection.addWidget(CardService.newTextParagraph().setText("Demasiados adjuntos XLSX para mostrar."));
//             break;
//           }
//         }
//       }
//     }
//     if (widgetCount >= MAX_WIDGETS) {
//       break;
//     }
//   }

//   if (widgetCount === 0 && processedFiles.length === 0) {
//     attachmentsSection.addWidget(CardService.newTextParagraph().setText("No se encontraron adjuntos XLSX."));
//   }
//   return attachmentsSection;
// }



/****************************************************Berlin********************************************************************************/

/******************************************************************************************************************************************/

// Function to extract data from the destination string
function extractDestinationData(stopValue) {
  if (!stopValue) {
    return {
      org_name: null,
      address_1: null,
      city: null,
      state: null,
      country: null,
      postal: null
    };
  }

  const destinationParts = stopValue.split(" - ");

  if (destinationParts.length === 6) {
    return {
      city: destinationParts[0].trim(),
      org_name: destinationParts[1].trim(),
      postal: destinationParts[2].trim(),
      state: destinationParts[3].trim(),
      country: destinationParts[4].trim(),
      address_1: destinationParts[5].trim()
    };
  } else {
    console.warn(`Unexpected format: ${stopValue}`);
    return {
      org_name: null,
      address_1: null,
      city: null,
      state: null,
      country: null,
      postal: null
    };
  }
}
function extractRefNumber(subject) {
  const regex = /REF#\s*([A-Z0-9]+)/i; 
  const match = subject.match(regex);

  if (match && match[1]) {
    return match[1].trim(); 
  } else {
    return null;
  }
}
function buscarHMO(texto) {
  texto = texto.toUpperCase();
  var posicion = texto.indexOf("HMO");
  return posicion !== -1;
}


/********************************************************************************************************************/

function createCardWithAttachmentbuttonBerlin(thread) {
  // Crea la sección de la tarjeta que muestra los archivos XLSX adjuntos como botones.
  const attachmentsSection = createAttachmentsSectionBerlin(thread);

  // Construye la tarjeta principal utilizando un CardBuilder.
  return CardService.newCardBuilder()
    // Establece el encabezado principal de la tarjeta.
    .setHeader(CardService.newCardHeader().setTitle("Auxiliar para crear órdenes en Rose Rocket con guías de Berlin"))
    // Establece el encabezado de la tarjeta que se muestra cuando la tarjeta está minimizada (peek card).
    .setPeekCardHeader(
      CardService.newCardHeader()
        .setTitle("Auxiliar para correos y cadenas") // Título del encabezado de la tarjeta minimizada.
        .setSubtitle("Volver al auxiliar de tareas para correos y cadenas") // Subtítulo del encabezado de la tarjeta minimizada.
    )
    // Agrega una sección con instrucciones para el usuario.
    .addSection(
      CardService.newCardSection()
        .addWidget(CardService.newTextParagraph().setText("<b>Instrucciones:</b>"))
        .addWidget(CardService.newTextParagraph().setText("Dar click en la guía de Berlin para crear la orden en Rose Rocket."))
        .addWidget(CardService.newTextParagraph().setText("La orden sera creada en Rose Rocket y recibiras un correo-e con el número de la orden y la liga para abrirla."))
    )
    // Agrega la sección de archivos adjuntos a la tarjeta.
    .addSection(attachmentsSection)
    // Establece un pie de página fijo en la parte inferior de la tarjeta.
    .setFixedFooter(standardAddonFooter())
    // Construye la tarjeta final y la retorna.  **IMPORTANTE: Llama a build()!**
    .build();
}



/**
 * Extracts the MBL reference from the email subject.
 *
 * @param {string} subject The subject line of the email.
 * @returns {string|null} The MBL reference, or null if not found.
 */
function extractMBLReferenceFromSubject(subject) {
  // Regular expression to find "MBL: " followed by alphanumeric characters and optionally other allowed characters.
  //This regex allows a wider range of characters in the MBL, as well as handles the MLB case
  const mblRegex = /(?:MBL|MLB): ([A-Z0-9#\-]+)/i; // Case-insensitive

  const match = subject.match(mblRegex);

  if (match && match[1]) {
    return match[1].trim(); // Return the captured MBL reference, removing leading/trailing whitespace
  } else {
    return null; // MBL reference not found in the subject
  }
}

/**
 * Extracts the REF# reference from the email subject.
 *
 * @param {string} subject The subject line of the email.
 * @returns {string|null} The REF# reference, or null if not found.
 */
function extractRefReferenceFromSubject(subject) {
  // Regular expression to find "REF# " followed by alphanumeric characters and optionally other allowed characters.
  const refRegex = /REF# ([A-Za-z0-9\-]+)/;

  const match = subject.match(refRegex);

  if (match && match[1]) {
    return match[1].trim(); // Return the captured REF# reference, removing leading/trailing whitespace
  } else {
    return null; // REF# reference not found in the subject
  }
}


/*********************************************************************************************/
function processPdfAttachmentBerlin(e) {
  try {
    // Obtiene el ID del mensaje y el nombre del archivo adjunto de los parámetros del evento.
    let messageId = e.parameters.messageId;
    let attachmentName = e.parameters.attachmentName;

    // Obtiene el mensaje de Gmail utilizando el ID del mensaje.
    try {
      var message = GmailApp.getMessageById(messageId);
    } catch (error) {
      Logger.log("Error al obtener el mensaje con ID: " + messageId + " - " + error);
      return CardService.newActionResponseBuilder()
        .setNotification(CardService.newNotification().setText("Error al obtener el mensaje. ID: " + messageId))
        .build();
    }

    // Obtiene todos los archivos adjuntos del mensaje.
    try {
      var attachments = message.getAttachments();
    } catch (error) {
      Logger.log("Error al obtener los adjuntos del mensaje: " + messageId + " - " + error);
      return CardService.newActionResponseBuilder()
        .setNotification(CardService.newNotification().setText("Error al obtener los adjuntos del mensaje."))
        .build();
    }

    // Busca el archivo adjunto específico por su nombre.
    let attachment = attachments.find(att => att.getName() === attachmentName);

    // Si no se encuentra el archivo adjunto, registra un error y devuelve una notificación al usuario.
    if (!attachment) {
      Logger.log("Archivo adjunto no encontrado: " + attachmentName);
      return CardService.newActionResponseBuilder()
        .setNotification(CardService.newNotification().setText("Archivo adjunto no encontrado."))
        .build();
    }
    var scriptProperties1 = PropertiesService.getScriptProperties();
    var customerId = scriptProperties1.getProperty('Berlin-CustomerId');

    var getinstance = PropertiesService.getScriptProperties();
    var instance = getinstance.getProperty('Instance.Amado');

    const apiEndpoint = "https://us-documentai.googleapis.com/v1/projects/786016653675/locations/us/processors/cde5d4e4e1abc046:process";   
    
    var parsedData = ProcesadorAIdeDocumentos.processPdfWithDocumentAIBerlin(attachment, apiEndpoint);       

    const order = ProcesadorAIdeDocumentos.creaJsonMultilegsBerlin(parsedData,customerId,instance)
      
    const responseRR = RoseRocketGASLibrary.createOrder(instance, customerId, order);

    const orderid = responseRR.order.id;
    const public_id = responseRR.order.public_id;

    var uploadResult = RoseRocketGASLibrary.uploadFileToOrder('Amado', orderid, attachment, 'other', attachment.getName());
     
    var link = "https://athn.roserocket.com/#/ops/orders/" + orderid;
    let subject = `Rose Rocket order created: ${public_id} for Berlin`;
    let body = `Rose Rocket order link: ${link}`;
    sendSimpleMail(subject, body)


    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText("Archivo " + attachmentName + " procesado correctamente."))
      .build();
  } catch (error) {
    // Si ocurre un error durante el procesamiento, registra el error y devuelve una notificación al usuario.
    Logger.log("Error al procesar el archivo adjunto: " + error);
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText("Error al procesar el archivo adjunto: " + error))
      .build();
  }
}



// /*********************************************************************************************/
// function processPdfAttachmentBerlin(e) {
//   try {
//     // Obtiene el ID del mensaje y el nombre del archivo adjunto de los parámetros del evento.
//     let messageId = e.parameters.messageId;
//     let attachmentName = e.parameters.attachmentName;

//     // Obtiene el mensaje de Gmail utilizando el ID del mensaje.
//     try {
//       var message = GmailApp.getMessageById(messageId);
//     } catch (error) {
//       Logger.log("Error al obtener el mensaje con ID: " + messageId + " - " + error);
//       return CardService.newActionResponseBuilder()
//         .setNotification(CardService.newNotification().setText("Error al obtener el mensaje. ID: " + messageId))
//         .build();
//     }

//     // Obtiene todos los archivos adjuntos del mensaje.
//     try {
//       var attachments = message.getAttachments();
//     } catch (error) {
//       Logger.log("Error al obtener los adjuntos del mensaje: " + messageId + " - " + error);
//       return CardService.newActionResponseBuilder()
//         .setNotification(CardService.newNotification().setText("Error al obtener los adjuntos del mensaje."))
//         .build();
//     }

//     // Busca el archivo adjunto específico por su nombre.
//     let attachment = attachments.find(att => att.getName() === attachmentName);

//     // Si no se encuentra el archivo adjunto, registra un error y devuelve una notificación al usuario.
//     if (!attachment) {
//       Logger.log("Archivo adjunto no encontrado: " + attachmentName);
//       return CardService.newActionResponseBuilder()
//         .setNotification(CardService.newNotification().setText("Archivo adjunto no encontrado."))
//         .build();
//     }

//     var scriptProperties1 = PropertiesService.getScriptProperties();
//     var customerId = scriptProperties1.getProperty('Berlin-CustomerId');

//     var getinstance = PropertiesService.getScriptProperties();
//     var instance = getinstance.getProperty('Instance.Amado');

//     const apiEndpoint = "https://us-documentai.googleapis.com/v1/projects/786016653675/locations/us/processors/cde5d4e4e1abc046:process";   
   
//     var parsedData = ProcesadorAIdeDocumentos.processPdfWithDocumentAIBerlin(attachment, apiEndpoint);       
   
//     const order = ProcesadorAIdeDocumentos.creaJsonMultilegsBerlin(parsedData,customerId,instance)

//     const responseRR = RoseRocketGASLibrary.createOrder(instance, customerId, order);

//     Logger.log(responseRR);

//     return CardService.newActionResponseBuilder()
//       .setNotification(CardService.newNotification().setText("Archivo " + attachmentName + " procesado correctamente."))
//       .build();
//   } catch (error) {
//     // Si ocurre un error durante el procesamiento, registra el error y devuelve una notificación al usuario.
//     Logger.log("Error al procesar el archivo adjunto: " + error);
//     return CardService.newActionResponseBuilder()
//       .setNotification(CardService.newNotification().setText("Error al procesar el archivo adjunto: " + error))
//       .build();
//   }
// }


// /***********************************************************************************************************************/
// const MAX_WIDGETS = 80;

// function createAttachmentsSectionBerlin(thread) {
//     // Crear la sección de la tarjeta
//     let attachmentsSection = CardService.newCardSection()
//         // Considera un encabezado más descriptivo si es apropiado
//         .setHeader("<b>Lista de 7512 Departure</b>"); // Encabezado ligeramente mejorado

//     let widgetCount = 0;
//     const threadId = thread.getId(); // El threadId SÍ se usa en actionParams

//     // --- Código eliminado ---
//     // let allAttachments = []; // No se usaba para la salida
//     // const scriptProperties = PropertiesService.getScriptProperties(); // No se usaba
//     // const processedFilesString = scriptProperties.getProperty(threadId); // No se usaba
//     // const processedFiles = processedFilesString ? JSON.parse(processedFilesString) : []; // No se usaba

//     // Iterar sobre los mensajes del hilo
//     for (let message of thread.getMessages()) {
//         let attachments;
//         try {
//             attachments = message.getAttachments();
//         } catch (e) {
//             Logger.log("Error getting attachments for message " + message.getId() + ": " + e);
//             continue; // Saltar al siguiente mensaje si hay error
//         }

//         // Iterar sobre los adjuntos del mensaje
//         for (let attachment of attachments) {
//             let attachmentName = attachment.getName();

//             // Filtrar adjuntos por nombre (insensible a mayúsculas/minúsculas)
//             if (attachmentName.toLowerCase().includes("7512 departure")) {

//                 // --- Código eliminado ---
//                 // No añadir a allAttachments ya que no se usa

//                 // Añadir un botón si no hemos alcanzado el límite
//                 if (widgetCount < MAX_WIDGETS) {
//                     // Parámetros para la acción del botón
//                     const actionParams = {
//                         'messageId': message.getId(),
//                         'attachmentName': attachmentName,
//                         'threadId': threadId, // Usamos el threadId aquí
//                     };

//                     // Acción que se ejecutará al hacer clic
//                     let processAction = CardService.newAction()
//                         .setFunctionName('processPdfAttachmentBerlin')
//                         .setParameters(actionParams);

//                     // Crear el botón
//                     let processButton = CardService.newTextButton()
//                         .setText(attachmentName) // No es necesario "" +
//                         .setOnClickAction(processAction);

//                     // Añadir el botón a la sección
//                     attachmentsSection.addWidget(processButton);
//                     widgetCount++;

//                 } else {
//                     // Si se alcanza el límite, añadir mensaje y salir del bucle de adjuntos
//                     attachmentsSection.addWidget(CardService.newTextParagraph().setText("Demasiados adjuntos coincidentes para mostrar (límite: " + MAX_WIDGETS + ").")); // Mensaje ligeramente mejorado
//                     break; // Salir del bucle de adjuntos de este mensaje
//                 }
//             } // Fin del if (filtro de nombre)
//         } // Fin del bucle de adjuntos

//         // Si ya alcanzamos el límite global, salir del bucle de mensajes
//         if (widgetCount >= MAX_WIDGETS) {
//             break;
//         }
//     } // Fin del bucle de mensajes

//     // Añadir mensaje si no se encontró ningún adjunto que cumpla el criterio
//     if (widgetCount === 0) {
//         // Mensaje actualizado para no mencionar "PDF" a menos que se filtre por tipo
//         attachmentsSection.addWidget(CardService.newTextParagraph().setText("No se encontraron adjuntos con '7512 Departure' en el nombre."));
//         // --- Código eliminado ---
//         // El bloque else if era redundante y usaba processedFiles que fue eliminado
//     }

//     // Devolver la sección construida
//     return attachmentsSection;
// }

function createAttachmentsSectionBerlin(thread) {
  const MAX_WIDGETS = 80;
    let attachmentsSection = CardService.newCardSection()
        .setHeader("<b>Lista de 7512 Departure</b>");

    let widgetCount = 0;
    const threadId = thread.getId();

    // Usaremos un Set para rastrear los nombres de archivos ya procesados
    const processedAttachmentNames = new Set();

    for (let message of thread.getMessages()) {
        let attachments = message.getAttachments();

        for (let attachment of attachments) {
            let attachmentName = attachment.getName();

            if (attachmentName.toLowerCase().includes("7512 departure")) {
                // Verifica si el nombre del archivo ya fue procesado
                if (!processedAttachmentNames.has(attachmentName)) {
                    if (widgetCount < MAX_WIDGETS) {
                        let actionParams = {
                            'messageId': message.getId(),
                            'attachmentName': attachmentName,
                            'threadId': threadId,
                        };

                        let processAction = CardService.newAction()
                            .setFunctionName('processPdfAttachmentBerlin')
                            .setParameters(actionParams);

                        let processButton = CardService.newTextButton()
                            .setText(attachmentName)
                            .setOnClickAction(processAction);

                        attachmentsSection.addWidget(processButton);
                        widgetCount++;

                        // Agrega el nombre del archivo al conjunto de nombres procesados
                        processedAttachmentNames.add(attachmentName);
                    } else {
                        attachmentsSection.addWidget(CardService.newTextParagraph().setText("Demasiados adjuntos coincidentes para mostrar (límite: " + MAX_WIDGETS + ")."));
                        break;
                    }
                }
            }
        }

        if (widgetCount >= MAX_WIDGETS) {
            break;
        }
    }

    if (widgetCount === 0) {
        attachmentsSection.addWidget(CardService.newTextParagraph().setText("No se encontraron adjuntos con '7512 Departure' en el nombre."));
    }

    return attachmentsSection;
}



/*********************
 * todos los archivos desarrollo*
 * ****************** *
 */

// function processAllPdfAttachmentsBerlin(e) {
//   const attachmentsString = e.parameters.attachments;
//   const attachments = JSON.parse(attachmentsString);

//   var scriptProperties1 = PropertiesService.getScriptProperties();
//   var customerId = scriptProperties1.getProperty('Berlin-CustomerId');

//   var getinstance = PropertiesService.getScriptProperties();
//   var instance = getinstance.getProperty('Instance.Amado');

//   let processedCount = 0;
//   let errors = [];
//   let cardText = "";
//   let totalAttachments = attachments.length;
//   let progressBarLength = 20; // Longitud de la barra de progreso en caracteres

//   for (let i = 0; i < totalAttachments; i++) {
//     let attachmentInfo = attachments[i];
//     let messageId = attachmentInfo.messageId;
//     let attachmentName = attachmentInfo.attachmentName;
//     let threadId = attachmentInfo.threadId;

//     try {
//       let thread = GmailApp.getThreadById(threadId);
//       let message = thread.getMessages().find(msg => msg.getId() === messageId);
//       let attachment = message.getAttachments().find(att => att.getName() === attachmentName);

//       if (!attachment) {
//         errors.push(`No se pudo encontrar el adjunto '${attachmentName}' en el mensaje '${messageId}'.`);
//         continue;
//       }

//       if (!attachmentName.toLowerCase().includes("7512 departure")) {
//         Logger.log(`El adjunto '${attachmentName}' no contiene '7512 departure'. Se omite.`);
//         continue;
//       }

//       const apiEndpoint = "https://us-documentai.googleapis.com/v1/projects/786016653675/locations/us/processors/cde5d4e4e1abc046:process";
//       var parsedData = ProcesadorAIdeDocumentos.processPdfWithDocumentAIBerlin(attachment, apiEndpoint);
//       const order = ProcesadorAIdeDocumentos.creaJsonMultilegsBerlin(parsedData, customerId, instance);
//       Logger.log(order);
//       processedCount++;

//     } catch (error) {
//       errors.push(`Error general al procesar '${attachmentName}': ${error}`);
//       Logger.log(`Error general al procesar '${attachmentName}': ${error}`);
//     }

//     // Actualizar la tarjeta en cada iteracion
//         if (true) { // Always update card, now with progress bar
//       // Calculate progress
//       let progress = (i + 1) / totalAttachments;
//       let filledBars = Math.round(progress * progressBarLength);
//       let emptyBars = progressBarLength - filledBars;

//       // Create progress bar string using Unicode characters
//       let progressBar = '█'.repeat(filledBars) + '░'.repeat(emptyBars); // Use full block and light shade

//       cardText = `Se han procesado ${processedCount} de ${totalAttachments} adjuntos.\nProgreso: ${progressBar} (${Math.round(progress * 100)}%)`;
//       if (errors.length > 0) {
//         cardText += `\nErrores:\n${errors.slice(-3).join('\n')}`;
//       }

//       let card = CardService.newCardBuilder()
//         .setHeader(CardService.newCardHeader().setTitle("Proceso en Curso"))
//         .addSection(CardService.newCardSection().addWidget(CardService.newTextParagraph().setText(cardText)))
//         .build();

//       let navigation = CardService.newNavigation()
//         .updateCard(card);

//       let actionResponse = CardService.newActionResponseBuilder()
//         .setNavigation(navigation)
//         .build();

//       return actionResponse; // IMPORTANTE: SALIR DE LA FUNCIÓN
//     }
//   }
// }


// const MAX_WIDGETS = 80;

// function createAttachmentsSectionBerlin(thread) {
//     let attachmentsSection = CardService.newCardSection().setHeader("<b>Adjuntos</b>");
//     let widgetCount = 0;

//     const threadId = thread.getId();
//     const scriptProperties = PropertiesService.getScriptProperties();
//     const processedFilesString = scriptProperties.getProperty(threadId);
//     const processedFiles = processedFilesString ? JSON.parse(processedFilesString) : [];

//     for (let message of thread.getMessages()) {

//          let attachments;
//         try {
//             attachments = message.getAttachments();
//         } catch (e) {
//             Logger.log("Error getting attachments for message " + message.getId() + ": " + e);
//             continue;
//         }

//         for (let attachment of attachments) {
//             let attachmentName = attachment.getName();

//             // Check if the attachment name contains "7512 Departure" (case-insensitive)
//             if (attachmentName.toLowerCase().includes("7512 departure")) {  // <-- KEY CHANGE

//                 if (widgetCount < MAX_WIDGETS) {

//                     const actionParams = {
//                         'messageId': message.getId(),
//                         'attachmentName': attachmentName,
//                         'threadId': threadId,
//                     };

//                     let processAction = CardService.newAction()
//                         .setFunctionName('processPdfAttachmentBerlin')
//                         .setParameters(actionParams);

//                     let processButton = CardService.newTextButton()
//                         .setText("" + attachmentName)
//                         .setOnClickAction(processAction);

//                     attachmentsSection.addWidget(processButton);
//                     widgetCount++;
//                 } else {
//                     attachmentsSection.addWidget(CardService.newTextParagraph().setText("Demasiados adjuntos PDF para mostrar."));
//                     break;
//                 }
//             } // End of if (attachmentName.toLowerCase().includes("7512 departure"))
//         }
//         if (widgetCount >= MAX_WIDGETS) {
//             break;
//         }
//     }

//     if (widgetCount === 0 && processedFiles.length === 0) {
//         attachmentsSection.addWidget(CardService.newTextParagraph().setText("No se encontraron adjuntos PDF con '7512 Departure' en el nombre."));  // <-- KEY CHANGE
//     } else if (widgetCount === 0) {
//         attachmentsSection.addWidget(CardService.newTextParagraph().setText("No se encontraron adjuntos PDF con '7512 Departure' en el nombre."));
//     }

//     return attachmentsSection;
// }
// /***********************************************************************************************************************/
// const MAX_WIDGETS = 80;

// function createAttachmentsSectionBerlin(thread) {
//     let attachmentsSection = CardService.newCardSection().setHeader("<b>Adjuntos</b>");
//     let widgetCount = 0;

//     const threadId = thread.getId();
//     const scriptProperties = PropertiesService.getScriptProperties();
//     const processedFilesString = scriptProperties.getProperty(threadId);
//     const processedFiles = processedFilesString ? JSON.parse(processedFilesString) : [];

//     for (let message of thread.getMessages()) {

//         const subject = message.getSubject();
//         var encontrado = buscarHMO(subject);

//         if (encontrado) {
//             Logger.log(`La palabra 'HMO' se encontró en el asunto del correo electrónico con ID: ${message.getId()}`);
//         } else {
//             Logger.log(`La palabra 'HMO' no se encontró en el asunto del correo electrónico con ID: ${message.getId()}`);
//         }

//         let attachments;
//         try {
//             attachments = message.getAttachments();
//         } catch (e) {
//             Logger.log("Error getting attachments for message " + message.getId() + ": " + e);
//             continue;
//         }

//         for (let attachment of attachments) {
//             let attachmentName = attachment.getName();

//                     if (widgetCount < MAX_WIDGETS) {

//                         const actionParams = {
//                             'messageId': message.getId(),
//                             'attachmentName': attachmentName,
//                             'threadId': threadId,
//                         };

//                         let processAction = CardService.newAction()
//                             .setFunctionName('processPdfAttachmentBerlin')
//                             .setParameters(actionParams);

//                         let processButton = CardService.newTextButton()
//                             .setText("" + attachmentName)
//                             .setOnClickAction(processAction);

//                         attachmentsSection.addWidget(processButton);
//                         widgetCount++;
//                     } else {
//                         attachmentsSection.addWidget(CardService.newTextParagraph().setText("Demasiados adjuntos PDF para mostrar."));
//                         break;
//                     }
//         }
//         if (widgetCount >= MAX_WIDGETS) {
//             break;
//         }
//     }

//     if (widgetCount === 0 && processedFiles.length === 0) {
//         attachmentsSection.addWidget(CardService.newTextParagraph().setText("No se encontraron adjuntos PDF."));
//     }

//     return attachmentsSection;
// }









// /********************************************************************************************/
// /* OTROS */
// /********************************************************************************************/
// function processPdfAttachmenOtros(e) {

//  const stops = [];

//   const numStops = Object.keys(e.formInput)
//     .filter((key) => key.startsWith("stop"))
//     .length;

//   for (let i = 1; i <= numStops; i++) {
//     const stopKey = `stop${i}`;
//     const stopValue = e.formInput[stopKey];

//     const destinationData = extractDestinationData(stopValue);

//     stops.push({
//       id: stopKey,
//       org_name: destinationData.org_name,
//       address_1: destinationData.address_1,
//       city: destinationData.city,
//       state: destinationData.state,
//       country: destinationData.country,
//       postal: destinationData.postal,
//       zipCodeValid: null
//     });
//   }

//   try {
//     // Obtiene el ID del mensaje y el nombre del archivo adjunto de los parámetros del evento.
//     let messageId = e.parameters.messageId;
//     let attachmentName = e.parameters.attachmentName;

//     // Obtiene el mensaje de Gmail utilizando el ID del mensaje.
//     try {
//       var message = GmailApp.getMessageById(messageId);
//     } catch (error) {
//       Logger.log("Error al obtener el mensaje con ID: " + messageId + " - " + error);
//       return CardService.newActionResponseBuilder()
//         .setNotification(CardService.newNotification().setText("Error al obtener el mensaje. ID: " + messageId))
//         .build();
//     }

//     // Obtiene todos los archivos adjuntos del mensaje.
//     try {
//       var attachments = message.getAttachments();
//     } catch (error) {
//       Logger.log("Error al obtener los adjuntos del mensaje: " + messageId + " - " + error);
//       return CardService.newActionResponseBuilder()
//         .setNotification(CardService.newNotification().setText("Error al obtener los adjuntos del mensaje."))
//         .build();
//     }

//     // Busca el archivo adjunto específico por su nombre.
//     let attachment = attachments.find(att => att.getName() === attachmentName);

//     // Si no se encuentra el archivo adjunto, registra un error y devuelve una notificación al usuario.
//     if (!attachment) {
//       Logger.log("Archivo adjunto no encontrado: " + attachmentName);
//       return CardService.newActionResponseBuilder()
//         .setNotification(CardService.newNotification().setText("Archivo adjunto no encontrado."))
//         .build();
//     }

//     var scriptProperties1 = PropertiesService.getScriptProperties();
//     var customerId = scriptProperties1.getProperty('Berlin-CustomerId');

//     var getinstance = PropertiesService.getScriptProperties();
//     var instance = getinstance.getProperty('Instance.Amado');

//     const apiEndpoint = "https://us-documentai.googleapis.com/v1/projects/786016653675/locations/us/processors/cde5d4e4e1abc046:process";   
//     var parsedData = ProcesadorAIdeDocumentos.processPdfWithDocumentAIBerlin(attachment, apiEndpoint);       
   
//     const res = ProcesadorAIdeDocumentos.creaJsonMultistopBerlin(parsedData, stops, customerId,instance)

//     return CardService.newActionResponseBuilder()
//       .setNotification(CardService.newNotification().setText("Archivo " + attachmentName + " procesado correctamente."))
//       .build();
//   } catch (error) {
//     // Si ocurre un error durante el procesamiento, registra el error y devuelve una notificación al usuario.
//     Logger.log("Error al procesar el archivo adjunto: " + error);
//     return CardService.newActionResponseBuilder()
//       .setNotification(CardService.newNotification().setText("Error al procesar el archivo adjunto: " + error))
//       .build();
//   }
// }

// /***********************************************************************************************************************/
// const MAX_WIDGETS = 80;
// const PDF_CONTENT_TYPE = "application/pdf";
// const OUTLOOK_PREFIX = "Outlook-";
// const REQUIRED_ATTACHMENT_NUMBER = "7512"; // Define the required number here

// function createAttachmentsSectionBerlin(thread) {
//     let attachmentsSection = CardService.newCardSection().setHeader("<b>   Stops </b>");
//     let widgetCount = 0;
//     let organizationDetails = [];

//     // Display a loading indicator
//     attachmentsSection.addWidget(CardService.newTextParagraph().setText("<i>Cargando organizaciones...</i>"));

//     // Determine the number of stops dynamically.  This is KEY.
//     const scriptProperties = PropertiesService.getScriptProperties();
//     const subject = thread.getFirstMessageSubject();
//     let stopPropertyKey = subject.includes("HMO") ? "Berlin-Hmo-Stop" : "Berlin-TJ-Stop";

//     let numStops = 0; // Declare numStops with a default value.
//     const numStopsProperty = scriptProperties.getProperty(stopPropertyKey);
//     if (numStopsProperty) {
//         try {
//             numStops = parseInt(numStopsProperty, 10); // Parse as integer
//         } catch (e) {
//             Logger.log("Error parsing stop property: " + e + ". Using default value of 0.");
//             numStops = 0; // Default if parsing fails
//             attachmentsSection.addWidget(CardService.newTextParagraph().setText("Error al procesar el número de paradas.")); //User message
//         }
//     } else {
//         Logger.log("No stop property found, defaulting to 0 stops.");
//         numStops = 0; // Default if not found.
//     }

//     const stopSelectionInputs = []; // Array to hold the SelectionInput objects

//     try {
//         var scriptProperties1 = PropertiesService.getScriptProperties();
//         var customerId = scriptProperties1.getProperty('Berlin-CustomerId');

//         var getinstance = PropertiesService.getScriptProperties();
//         var instance = getinstance.getProperty('Instance.Amado');

//         const data = RoseRocketGASLibrary.obtenerDatosAddressBook(instance, customerId, "");

//         organizationDetails = extractOrganizationDetails(data);
//         if (organizationDetails.length === 0) {
//             Logger.log("No organization details found or error extracting them.");
//             //attachmentsSection.clear(); // Remove loading message //ERROR HERE
//             attachmentsSection = CardService.newCardSection().setHeader("<b>   Stops </b>"); // Re-initialize
//             attachmentsSection.addWidget(CardService.newTextParagraph().setText("No se encontraron organizaciones. Por favor, verifique la configuración y el ID del cliente."));
//         } else {
//             //attachmentsSection.clear(); //Remove the loading message //ERROR HERE
//             attachmentsSection = CardService.newCardSection().setHeader("<b>   Stops </b>"); // Re-initialize
//             // Create the stop dropdowns dynamically
//             for (let i = 1; i <= numStops; i++) {
//                 let stopSelectionInput = CardService.newSelectionInput()
//                     .setFieldName(`stop${i}`)
//                     .setTitle(`Stop ${i}`)
//                     .setType(CardService.SelectionInputType.DROPDOWN);

//                 if (organizationDetails.length > 0) {
//                     organizationDetails.forEach(org => {
//                         const label = `${org.orgName} - ${org.city} - ${org.state} - ${org.postal} - ${org.country}`;
//                         stopSelectionInput.addItem(label, org.storage, false);
//                     });
//                 } else {
//                     stopSelectionInput.addItem("No organizations found", "none", true);
//                 }
//                 stopSelectionInputs.push(stopSelectionInput);
//                 attachmentsSection.addWidget(stopSelectionInput); // Add to the section as created
//             }
//         }

//     } catch (error) {
//         Logger.log("Error fetching or extracting organization details: " + error + " Customer ID: " + customerId);
//         //attachmentsSection.clear(); // Remove loading message //ERROR HERE
//         attachmentsSection = CardService.newCardSection().setHeader("<b>   Stops </b>"); // Re-initialize

//         attachmentsSection.addWidget(CardService.newTextParagraph().setText("Error al obtener o extraer detalles de la organización. Por favor, revise la configuración."));
//         return attachmentsSection;
//     }
//     if (organizationDetails.length === 0) return attachmentsSection;

//     const threadId = thread.getId();

//     const processedFilesString = scriptProperties.getProperty(threadId);
//     const processedFiles = processedFilesString ? JSON.parse(processedFilesString) : [];

//     for (let message of thread.getMessages()) {

//         const subject = message.getSubject();
//         var encontrado = buscarHMO(subject);

//         if (encontrado) {
//             Logger.log(`La palabra 'HMO' se encontró en el asunto del correo electrónico con ID: ${message.getId()}`);
//         } else {
//             Logger.log(`La palabra 'HMO' no se encontró en el asunto del correo electrónico con ID: ${message.getId()}`);
//         }

//         let attachments;
//         try {
//             attachments = message.getAttachments();
//         } catch (e) {
//             Logger.log("Error getting attachments for message " + message.getId() + ": " + e);
//             continue;
//         }

//         for (let attachment of attachments) {
//             let attachmentName = attachment.getName();
//             // Check if the attachment name contains "7512" before proceeding
//             if (attachmentName.includes(REQUIRED_ATTACHMENT_NUMBER) &&
//                 attachment.getContentType().indexOf(PDF_CONTENT_TYPE) !== -1 &&
//                 !attachmentName.startsWith(OUTLOOK_PREFIX)) {
//                 if (processedFiles.includes(attachmentName)) {
//                     attachmentsSection.addWidget(CardService.newTextParagraph().setText(`<i>${attachmentName}</i> - Procesado ✅`));
//                 } else {
//                     if (widgetCount < MAX_WIDGETS) {

//                         const actionParams = {
//                             'messageId': message.getId(),
//                             'attachmentName': attachmentName,
//                             'threadId': threadId,
//                         };

//                         for (let i = 1; i <= numStops; i++) {
//                             actionParams[`stop${i}`] = ''; //Add the empty values
//                         }

//                         let processAction = CardService.newAction()
//                             .setFunctionName('processPdfAttachmentBerlin')
//                             .setParameters(actionParams);

//                         let processButton = CardService.newTextButton()
//                             .setText("Crear orden - " + attachmentName)
//                             .setOnClickAction(processAction);

//                         attachmentsSection.addWidget(processButton);
//                         widgetCount++;
//                     } else {
//                         attachmentsSection.addWidget(CardService.newTextParagraph().setText("Demasiados adjuntos PDF para mostrar."));
//                         break;
//                     }
//                 }
//             }
//         }
//         if (widgetCount >= MAX_WIDGETS) {
//             break;
//         }
//     }

//     if (widgetCount === 0 && processedFiles.length === 0) {
//         attachmentsSection.addWidget(CardService.newTextParagraph().setText("No se encontraron adjuntos PDF."));
//     }

//     return attachmentsSection;
// }



// function createAttachmentsSectionBerlin(thread) {
//   const MAX_WIDGETS = 80;
//   let messages;

//   try {
//     messages = thread.getMessages();
//   } catch (e) {
//     Logger.log("Error getting thread messages: " + e);
//     return CardService.newCardSection().addWidget(CardService.newTextParagraph().setText("Error al obtener los mensajes del hilo."));
//   }

//   let attachmentsSection = CardService.newCardSection().setHeader("<b>   Stops </b>");
//   let widgetCount = 0;
//   let organizationDetails = [];

//   try {
//     var scriptProperties1 = PropertiesService.getScriptProperties();
//     var customerId = scriptProperties1.getProperty('Berlin-CustomerId');

//     var getinstance = PropertiesService.getScriptProperties();
//     var instance = getinstance.getProperty('Instance.Amado');

//     const data = RoseRocketGASLibrary.obtenerDatosAddressBook(instance, customerId, "");

//     organizationDetails = extractOrganizationDetails(data);
//     if (organizationDetails.length === 0) {
//       Logger.log("No organization details found or error extracting them.");
//     }
//   } catch (error) {
//     Logger.log("Error fetching or extracting organization details: " + error);
//   }

//   // Determine the number of stops dynamically.  This is KEY.
//   let numStops = 0;
//   const scriptProperties = PropertiesService.getScriptProperties();
//   //This needs to be dynamic based on the subject line, i have omitted code to do that here.
//   const subject = thread.getFirstMessageSubject();
//   let stopPropertyKey = "";

//   if(subject.includes("HMO")){
//      stopPropertyKey = "Berlin-Hmo-Stop";
//   }else{
//      stopPropertyKey = "Berlin-TJ-Stop";
//   }
//   const numStopsProperty = scriptProperties.getProperty(stopPropertyKey);
//   if (numStopsProperty) {
//     numStops = parseInt(numStopsProperty, 10); // Parse as integer
//   } else {
//     Logger.log("No stop property found, defaulting to 0 stops.");
//     numStops = 0; // Default if not found.
//   }
//   const stopSelectionInputs = []; // Array to hold the SelectionInput objects

//   // Create the stop dropdowns dynamically
//   for (let i = 1; i <= numStops; i++) {
//     let stopSelectionInput = CardService.newSelectionInput()
//       .setFieldName(`stop${i}`)
//       .setTitle(`Stop ${i}`)
//       .setType(CardService.SelectionInputType.DROPDOWN);

//     if (organizationDetails.length > 0) {
//       organizationDetails.forEach(org => {
//         const label = `${org.orgName} - ${org.city} - ${org.state} - ${org.postal} - ${org.country}`;
//         stopSelectionInput.addItem(label, org.storage, false);
//       });
//     } else {
//       stopSelectionInput.addItem("No organizations found", "none", true);
//     }
//     stopSelectionInputs.push(stopSelectionInput);
//     attachmentsSection.addWidget(stopSelectionInput); // Add to the section as created
//   }

//   // If no organizations are found, show a message and exit early
//   if (organizationDetails.length === 0) {
//     let noOrgMessage = CardService.newTextParagraph().setText("No organizations found. Please check your configuration.");
//     attachmentsSection.addWidget(noOrgMessage);
//     return attachmentsSection; // Exit the function early
//   }
//   const threadId = thread.getId();

//   const processedFilesString = scriptProperties.getProperty(threadId);
//   const processedFiles = processedFilesString ? JSON.parse(processedFilesString) : [];

//   for (let message of messages) {
//     // **INTEGRATE HMO SEARCH HERE**
//     const subject = message.getSubject();
//     var encontrado = buscarHMO(subject);

//     if (encontrado) {
//       Logger.log(`La palabra 'HMO' se encontró en el asunto del correo electrónico con ID: ${message.getId()}`);
//     } else {
//       Logger.log(`La palabra 'HMO' no se encontró en el asunto del correo electrónico con ID: ${message.getId()}`);
//     }
//     // **END HMO SEARCH**

//     let attachments;
//     try {
//       attachments = message.getAttachments();
//     } catch (e) {
//       Logger.log("Error getting attachments for message " + message.getId() + ": " + e);
//       continue;
//     }

//     for (let attachment of attachments) {
//       let attachmentName = attachment.getName();
//       if (
//         attachment.getContentType().indexOf("application/pdf") !== -1 &&
//         !attachmentName.startsWith("Outlook-")
//       ) {
//         if (processedFiles.includes(attachmentName)) {
//           // Already processed: show a message instead of a button
//           attachmentsSection.addWidget(CardService.newTextParagraph().setText(`<i>${attachmentName}</i> - Procesado ✅`));
//         } else {
//           if (widgetCount < MAX_WIDGETS) {

//             const actionParams = {
//               'messageId': message.getId(),
//               'attachmentName': attachmentName,
//               'threadId': threadId,
//             };

//             // Dynamically add stop parameters to the action
//             for (let i = 1; i <= numStops; i++) {
//               actionParams[`stop${i}`] = ''; //Add the empty values
//             }

//             let processAction = CardService.newAction()
//               .setFunctionName('processPdfAttachmentBerlin')
//               .setParameters(actionParams);

//             let processButton = CardService.newTextButton()
//               .setText( "Crear orden - " + attachmentName)
//               .setOnClickAction(processAction);

//             attachmentsSection.addWidget(processButton);
//             widgetCount++;
//           } else {
//             attachmentsSection.addWidget(CardService.newTextParagraph().setText("Demasiados adjuntos PDF para mostrar."));
//             break;
//           }
//         }
//       }
//     }
//     if (widgetCount >= MAX_WIDGETS) {
//       break;
//     }
//   }
//   if (widgetCount === 0 && processedFiles.length === 0) {
//     attachmentsSection.addWidget(CardService.newTextParagraph().setText("No se encontraron adjuntos PDF."));
//   }
//   return attachmentsSection;
// }



function buscarHMO(texto) {
  texto = texto.toUpperCase();
  var posicion = texto.indexOf("HMO");
  return posicion !== -1;
}


// function extractTableData(messageBody) {
//  // Logger.log("extractTableData (Plain Text Version) called");

//   // Regular expression to match the entire table structure
//   const tableRegex = /\*Container #\*\s*\n\*Inbond #\*\s*\n\*MBL REF#\*\s*\n\*PKG\*\s*\n\*DELIVERY\*\s*\n\*PORT CODE\*\s*\n\*PO#\*\s*\n([\s\S]*?)Please confirm received/i;

//   // Regular expression to match each row in the table
//   const rowRegex = /\*([A-Z0-9]+)\*\s*\n\*([0-9]+)\*\s*\n\*([A-Z0-9\s]+)\*\s*\n\*([0-9]+)\*\s*\n\*([A-Z]+)\*\s*\n\*([0-9]+)\*\s*\n\*([0-9&\s]+)\*/g;

//   const tableMatch = messageBody.match(tableRegex);

//   if (!tableMatch) {
//     Logger.log("No se encontró la tabla en el cuerpo del correo.");
//     return []; // Retornar un array vacío si no se encuentra la tabla
//   }

//   const tableContent = tableMatch[1].trim(); // Extract the content between the headers and "Please confirm"
//   //Logger.log("Table Content Extracted:\n" + tableContent);

//   let rows = [];
//   let rowMatch;

//   while ((rowMatch = rowRegex.exec(tableContent)) !== null) {
//     rows.push({
//       'Container': rowMatch[1],
//       'Inbond': rowMatch[2],
//       'MBL REF': rowMatch[3].trim(),
//       'PKG': rowMatch[4],
//       'DELIVERY': rowMatch[5],
//       'PORT CODE': rowMatch[6],
//       'PO': rowMatch[7].trim()
//     });
//   }

//  // Logger.log("Extracted Rows: " + JSON.stringify(rows, null, 2));
//   return rows;
// }




