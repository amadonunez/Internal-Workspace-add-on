# Auxiliar de Gmail para Tareas Automatizadas

Este complemento de Google Workspace está diseñado para simplificar y automatizar tareas repetitivas directamente desde Gmail, ayudando a mejorar la eficiencia y la productividad del equipo.

## ¿Qué hace este complemento?

El "Auxiliar para Gmail" se integra con tu correo electrónico para automatizar la creación de órdenes de trabajo en el sistema **Rose Rocket**. Funciona extrayendo información clave de los correos electrónicos y archivos adjuntos (como archivos PDF y hojas de cálculo) que recibes de diferentes clientes.

### Funcionalidades Principales:

*   **Creación automática de órdenes:** Cuando recibes un correo electrónico con una guía de carga de clientes específicos, el complemento lo detecta y te presenta una tarjeta interactiva. Con un solo clic, puedes procesar la información del archivo adjunto y crear una nueva orden en Rose Rocket, sin necesidad de introducir los datos manualmente.

*   **Soporte para Múltiples Clientes:** El sistema está configurado para trabajar con varios clientes, entre ellos:
    *   **SALSAS CASTILLO**
    *   **ILS**
    *   **Taylor**
    *   **Berlin**

*   **Interfaz Sencilla en Gmail:** No necesitas salir de Gmail para realizar estas acciones. El complemento muestra una barra lateral con botones y opciones claras para que puedas gestionar las órdenes de forma rápida y sencilla.

*   **Notificaciones por Correo Electrónico:** Una vez que se ha creado una orden en Rose Rocket, el sistema envía automáticamente un correo electrónico de confirmación para notificar a las partes interesadas.

*   **Envío de Archivos Adjuntos:** Además de la creación de órdenes, el complemento te permite seleccionar y enviar fácilmente archivos adjuntos a otros destinatarios directamente desde la interfaz del complemento.

## ¿Cómo funciona?

1.  **Recibes un correo electrónico:** Cuando llega un correo de un cliente configurado (por ejemplo, con el asunto "Guía de Carga de Berlin"), el complemento se activa.
2.  **Aparece una tarjeta interactiva:** En la barra lateral de Gmail, verás una tarjeta con el título "Auxiliar para crear órdenes en Rose Rocket".
3.  **Procesas el archivo adjunto:** La tarjeta te mostrará los archivos adjuntos del correo. Simplemente haz clic en el archivo que deseas procesar.
4.  **Se crea la orden:** El complemento lee la información del archivo, se conecta con Rose Rocket y crea la orden de trabajo automáticamente.
5.  **Recibes una confirmación:** Se te notificará por correo electrónico que la orden ha sido creada con éxito.

## Flujos de Trabajo Detallados

El archivo `Gmail_AttachmentsEmail.js` contiene la lógica principal para manejar los diferentes tipos de correos y clientes. A continuación se describen los flujos de trabajo:

### Flujo de Trabajo General para Adjuntos

*   **`createCardWithAttachmentCheckboxesAndEmailOption(thread)`:** Esta función crea una tarjeta genérica que muestra todos los archivos adjuntos (PDF e imágenes) de un correo electrónico con casillas de verificación.
*   **Selección y Envío:** El usuario puede seleccionar los archivos que desea y especificar un destinatario y un asunto para enviarlos por correo electrónico.
*   **`sendSelectedPDFsEmail(e)`:** Esta función se activa al hacer clic en el botón de enviar, recopila los archivos seleccionados y los envía al destinatario especificado.

### Flujo de Trabajo para Clientes Específicos

El complemento utiliza funciones específicas para cada cliente, las cuales se activan según el asunto o la etiqueta del correo electrónico.

#### Cliente: ILS

*   **`createCardWithAttachmentbuttonILS(thread)`:** Crea una tarjeta específica para los correos de ILS.
*   **`createAttachmentsSectionILS(thread)`:** Muestra los archivos adjuntos PDF como botones en la tarjeta.
*   **`processPdfAttachment(e)`:** Al hacer clic en un botón de archivo adjunto, esta función:
    1.  Extrae la información del archivo PDF utilizando la API de Google Document AI.
    2.  Transforma los datos extraídos al formato requerido por Rose Rocket.
    3.  Crea una nueva orden en Rose Rocket.
    4.  Sube el archivo PDF original a la orden recién creada.
    5.  Envía un correo de notificación con el enlace a la nueva orden.

#### Cliente: Taylor

*   **`createCardWithAttachmentbuttonTaylor(thread)`:** Crea una tarjeta para los correos de Taylor.
*   **`createAttachmentsSectionTaylor(thread)`:** Muestra los archivos adjuntos XLSX como botones y proporciona menús desplegables para seleccionar el origen y el destino.
*   **`processXlsxAttachmentTaylor(e)`:** Al hacer clic en un botón de archivo adjunto:
    1.  Convierte el archivo XLSX a una hoja de cálculo de Google.
    2.  Extrae los datos de la hoja de cálculo.
    3.  Crea una orden de múltiples paradas en Rose Rocket.
    4.  Convierte la hoja de cálculo a PDF y la sube a la orden.
    5.  Envía un correo de notificación con el enlace a la nueva orden.

#### Cliente: Berlin

*   **`createCardWithAttachmentbuttonBerlin(thread)`:** Crea una tarjeta para los correos de Berlin.
*   **`createAttachmentsSectionBerlin(thread)`:** Muestra los archivos adjuntos que contienen "7512 Departure" en el nombre como botones.
*   **`processPdfAttachmentBerlin(e)`:** Al hacer clic en un botón de archivo adjunto:
    1.  Procesa el archivo PDF con Google Document AI.
    2.  Crea una orden de múltiples tramos en Rose Rocket.
    3.  Sube el archivo PDF a la orden.
    4.  Envía un correo de notificación con el enlace a la nueva orden.

## Tecnologías Utilizadas

Este complemento ha sido desarrollado con **Google Apps Script**, lo que le permite integrarse de forma nativa con los servicios de Google Workspace. Utiliza tecnologías como:

*   **CardService:** Para crear las interfaces de usuario interactivas dentro de Gmail.
*   **Gmail API:** Para leer y procesar los correos electrónicos.
*   **Rose Rocket API:** Para conectarse con el sistema de gestión de órdenes.
*   **Google Document AI:** Para extraer información de manera inteligente de los documentos PDF.

Este es un complemento interno, diseñado para optimizar los flujos de trabajo específicos de la empresa.