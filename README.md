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

## ¿Cómo funciona?

1.  **Recibes un correo electrónico:** Cuando llega un correo de un cliente configurado (por ejemplo, con el asunto "Guía de Carga de Berlin"), el complemento se activa.
2.  **Aparece una tarjeta interactiva:** En la barra lateral de Gmail, verás una tarjeta con el título "Auxiliar para crear órdenes en Rose Rocket".
3.  **Procesas el archivo adjunto:** La tarjeta te mostrará los archivos adjuntos del correo. Simplemente haz clic en el archivo que deseas procesar.
4.  **Se crea la orden:** El complemento lee la información del archivo, se conecta con Rose Rocket y crea la orden de trabajo automáticamente.
5.  **Recibes una confirmación:** Se te notificará por correo electrónico que la orden ha sido creada con éxito.

## Tecnologías Utilizadas

Este complemento ha sido desarrollado con **Google Apps Script**, lo que le permite integrarse de forma nativa con los servicios de Google Workspace. Utiliza tecnologías como:

*   **CardService:** Para crear las interfaces de usuario interactivas dentro de Gmail.
*   **Gmail API:** Para leer y procesar los correos electrónicos.
*   **Rose Rocket API:** Para conectarse con el sistema de gestión de órdenes.
*   **Google Document AI:** Para extraer información de manera inteligente de los documentos PDF.

Este es un complemento interno, diseñado para optimizar los flujos de trabajo específicos de la empresa.
