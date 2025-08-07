# Internal Workspace add-on

Este es el repositorio de un complemento interno de Google Workspace diseñado para mejorar la funcionalidad y automatizar tareas específicas dentro de un entorno de Google Workspace.

**Funcionalidad Principal:**
Este complemento está diseñado para optimizar la colaboración y la productividad del equipo mediante la integración de procesos personalizados en varias aplicaciones de Google Workspace.

**Características Clave:**
-   **Integración con Gmail:** Automatiza tareas relacionadas con correos electrónicos, incluyendo la lectura de mensajes, la composición y el manejo de archivos adjuntos.
-   **Integración con Google Drive:** Permite acciones contextuales sobre elementos seleccionados en Google Drive.
-   **Página de Inicio Centralizada:** Proporciona una interfaz de usuario unificada al iniciar el complemento, con enlaces rápidos a Gmail y Drive.
-   **Procesamiento de Documentos:** Utiliza la librería `ProcesadorAIdeDocumentos` para el procesamiento avanzado de documentos, posiblemente a través de Google Document AI.
-   **Integración con Rose Rocket:** Se conecta con la API de Rose Rocket a través de la librería `RoseRocketGASLibrary` para funcionalidades relacionadas con la gestión de órdenes o logística.
-   **Gestión Segura de Propiedades:** Utiliza `PropertiesService` para almacenar de forma segura configuraciones y credenciales sensibles.

**Tecnologías y Dependencias:**
-   **Google Apps Script:** Plataforma de desarrollo principal.
-   **Servicios Avanzados de Google:** Gmail Service, CardService.
-   **Librerías Externas:**
    -   `RoseRocketGASLibrary`
    -   `ProcesadorAIdeDocumentos`
-   **Scopes de OAuth:** Requiere permisos para interactuar con Gmail, Drive, Hojas de Cálculo, y realizar solicitudes externas, entre otros.

**Uso:**
Este es un complemento interno, lo que implica que su despliegue y gestión se realizan a través de la consola de administración de Google Workspace o directamente desde el proyecto de Google Apps Script.