/**
 * Handles the checkbox change event.
 *
 * @param {Object} e The event object.
 */
function handleCheckboxChange(e) {
  Logger.log('Checkboxes changed:');
  Logger.log(e.formInput); // Log the form input values
  // Add your logic to handle the checkbox changes here
  let selectedAttachments = e.formInput['pdf_attachments']; // Get selected values

  if (selectedAttachments) {
    if (typeof selectedAttachments === 'string') {
      Logger.log("Selected Attachment: " + selectedAttachments);
    } else {
      selectedAttachments.forEach(attachment => {
        Logger.log("Selected Attachment: " + attachment);
      });
    }
  } else {
    Logger.log("No attachments selected");
  }
}


// /********************************************************************************************/
// /** Amado N.V.
//  * Crea una tarjeta con una sección que muestra botones para los archivos PDF adjuntos,
//  * un encabezado, un pie de página y una acción asociada.
//  *
//  * @param {GoogleAppsScript.Gmail.GmailThread} thread El hilo de correo electrónico del cual extraer los adjuntos.
//  * @return {GoogleAppsScript.CardService.Card} La tarjeta construida con los elementos especificados.
//  */
// function createCardWithAttachmentbuttonILS(thread) {
//   // Crea la sección de la tarjeta que mostrará los archivos PDF adjuntos como botones.
//   // Utiliza la función 'createAttachmentsSectionILS' para generar esta sección.
//   const attachmentsSection = createAttachmentsSectionILS(thread);

//   // Crea una acción que se ejecutará cuando se haga clic en un botón específico.
//   // En este caso, la acción llamará a la función 'sendSelectedPDFsEmail'.
//   const anAction = CardService.newAction().setFunctionName('sendSelectedPDFsEmail');

//   // Construye la tarjeta principal utilizando un CardBuilder.
//   return CardService.newCardBuilder()
//     // Establece el encabezado principal de la tarjeta.
//     .setHeader(CardService.newCardHeader().setTitle("Auxiliar para Gmail ILS"))
//     // Establece el encabezado de la tarjeta que se muestra cuando la tarjeta está minimizada (peek card).
//     .setPeekCardHeader(
//       CardService.newCardHeader()
//         .setTitle("Auxiliar para correos y cadenas") // Título del encabezado de la tarjeta minimizada.
//         .setSubtitle("Volver al auxiliar de tareas para correos y cadenas") // Subtítulo del encabezado de la tarjeta minimizada.
//     )
//     // Agrega la sección de archivos adjuntos a la tarjeta.
//     .addSection(attachmentsSection)
//     // Establece un pie de página fijo en la parte inferior de la tarjeta.
//     // Utiliza la función 'standardAddonFooter' para generar el pie de página estándar.
//     .setFixedFooter(standardAddonFooter())
//     // Construye la tarjeta final y la retorna.  **IMPORTANTE: Llama a build()!**
//     .build();
// }

// /** agregado por Amado 
//  * Crea una sección de tarjeta que muestra los archivos PDF adjuntos en un hilo de correo electrónico
//  * como botones, permitiendo al usuario hacer clic para procesar cada archivo.
//  *
//  * @param {GoogleAppsScript.Gmail.GmailThread} thread El hilo de correo electrónico del cual extraer los adjuntos.
//  * @return {GoogleAppsScript.CardService.CardSection} La sección de la tarjeta con los botones de los adjuntos.
//  */
// function createAttachmentsSectionILS(thread) {
//   // Define el número máximo de widgets (botones) a mostrar en la sección para evitar sobrecargar la UI.
//   const MAX_WIDGETS = 80;
//   // Obtiene todos los mensajes dentro del hilo de correo electrónico.
//   let messages = thread.getMessages();
//   // Crea una nueva sección de tarjeta y establece el título como "Adjuntos PDF".
//   let attachmentsSection = CardService.newCardSection().addWidget(CardService.newTextParagraph().setText("Adjuntos PDF"));
//   // Inicializa un contador para rastrear el número de widgets agregados a la sección.
//   let widgetCount = 0;

//   // Itera a través de cada mensaje en el hilo de correo electrónico.
//   for (let message of messages) {
//     // Obtiene todos los archivos adjuntos del mensaje actual.
//     let attachments = message.getAttachments();
//     // Itera a través de cada archivo adjunto.
//     for (let attachment of attachments) {
//       // Verifica si el archivo adjunto es un archivo PDF y si su nombre no comienza con "Outlook-" (para filtrar archivos creados por Outlook).
//       if (
//         attachment.getContentType().indexOf("application/pdf") !== -1 &&
//         !attachment.getName().startsWith("Outlook-")
//       ) {
//         // Verifica si el número de widgets en la sección aún no ha alcanzado el máximo permitido.
//         if (widgetCount < MAX_WIDGETS) {
//           // Obtiene el nombre del archivo adjunto.
//           let attachmentName = attachment.getName();

//           // Crea una acción que se ejecutará cuando se haga clic en el botón del archivo adjunto.
//           let processAction = CardService.newAction()
//             .setFunctionName('processPdfAttachment') // Establece el nombre de la función que se llamará al hacer clic en el botón.
//             .setParameters({ // Define los parámetros que se pasarán a la función.
//               'messageId': message.getId(), // El ID del mensaje que contiene el archivo adjunto.
//               'attachmentName': attachmentName // El nombre del archivo adjunto.
//             });

//           // Crea un botón de texto con el nombre del archivo adjunto como etiqueta y asocia la acción creada.
//           let processButton = CardService.newTextButton()
//             .setText(attachmentName) // Establece el texto del botón como el nombre del archivo adjunto.
//             .setOnClickAction(processAction) // Asocia la acción al evento de clic del botón.
    
//                // Agrega el botón a la sección de la tarjeta.
//           attachmentsSection.addWidget(processButton);
//           // Incrementa el contador de widgets.
//           widgetCount++;
//         } else {
//           // Si se ha alcanzado el número máximo de widgets, agrega un mensaje indicando que hay demasiados archivos adjuntos para mostrar.
//           attachmentsSection.addWidget(CardService.newTextParagraph().setText("Demasiados adjuntos PDF para mostrar."));
//           // Sale del bucle interno (archivos adjuntos) ya que no se pueden agregar más botones.
//           break;
//         }
//       }
//     }
//     // Si se ha alcanzado el número máximo de widgets, sale del bucle externo (mensajes) ya que no se pueden agregar más botones.
//     if (widgetCount >= MAX_WIDGETS) {
//       break;
//     }
//   }

//   // Si no se encontraron archivos PDF, agrega un mensaje indicando que no se encontraron archivos adjuntos.
//   if (widgetCount === 0) {
//     attachmentsSection.addWidget(CardService.newTextParagraph().setText("No se encontraron adjuntos PDF."));
//   }

//   // Retorna la sección de la tarjeta con los botones de los archivos adjuntos.
//   return attachmentsSection;
// }

// /**
//  * Función que se ejecuta cuando el usuario hace clic en un botón de archivo PDF adjunto.
//  * Esta función recupera el archivo adjunto del mensaje de correo electrónico y permite su procesamiento.
//  *
//  * @param {GoogleAppsScript.CardService.ActionResponseEvent} e El evento de acción que contiene los parámetros del botón clicado.
//  * @return {GoogleAppsScript.CardService.ActionResponse} Una respuesta de acción para la tarjeta.
//  */
// function processPdfAttachment(e) {
//   try {
//     // Obtiene el ID del mensaje y el nombre del archivo adjunto de los parámetros del evento.
//     let messageId = e.parameters.messageId;
//     let attachmentName = e.parameters.attachmentName;

//     // Obtiene el mensaje de Gmail utilizando el ID del mensaje.
//     let message = GmailApp.getMessageById(messageId);
//     // Obtiene todos los archivos adjuntos del mensaje.
//     let attachments = message.getAttachments();

//     // Busca el archivo adjunto específico por su nombre.
//     let attachment = attachments.find(att => att.getName() === attachmentName);

//     // Si no se encuentra el archivo adjunto, registra un error y devuelve una notificación al usuario.
//     if (!attachment) {
//       Logger.log("Archivo adjunto no encontrado: " + attachmentName);
//       return CardService.newActionResponseBuilder()
//         .setNotification(CardService.newNotification().setText("Archivo adjunto no encontrado."))
//         .build();
//     }

//     // **Aquí es donde procesas el archivo PDF**
//     // Puedes acceder al contenido del archivo adjunto usando `attachment.getDataAsString()` o `attachment.getDataAsBlob()`.
//     // Luego, puedes realizar la operación que desees con el contenido del archivo PDF.

//     // Ejemplo: Imprimir el tamaño del archivo en los logs
//     Logger.log("Tamaño del archivo " + attachmentName + ": " + attachment.getSize() + " bytes");

//     // Ejemplo: Devolver una notificación al usuario
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
