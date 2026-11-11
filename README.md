# Sistema de Evaluación Online

Proyecto de portafolio: plataforma web de evaluaciones académicas desarrollada con **Google Apps Script**, **Google Sheets**, **HTML**, **CSS** y **JavaScript**.

## Descripción

Este proyecto implementa una aplicación web para administrar el proceso de evaluación de estudiantes. La arquitectura separa la interfaz de usuario, la lógica de negocio y el almacenamiento de información.

## Funcionalidades principales

- Validación de estudiantes y cursos antes de iniciar una evaluación.
- Consulta de evaluaciones publicadas.
- Presentación de preguntas y alternativas desde una base de datos en Google Sheets.
- Corrección de respuestas en el servidor.
- Protección de las respuestas correctas para que no se envíen al navegador.
- Registro y control de intentos.
- Prevención de envíos simultáneos.
- Panel de acceso docente y consulta de resultados.
- Análisis de resultados por pregunta.
- Interfaz adaptable a distintos tamaños de pantalla.

## Tecnologías

- Google Apps Script (backend y publicación como aplicación web).
- Google Sheets (almacenamiento de datos).
- HTML5, CSS3 y JavaScript (interfaz).

## Estructura

```text
sistema-evaluacion-online/
├── Code.gs       # Lógica del servidor y comunicación con Google Sheets
├── Index.html    # Interfaz de estudiante y panel docente
├── README.md     # Documentación del proyecto
└── .gitignore    # Exclusión de archivos locales y credenciales
```

## Aviso sobre esta versión pública

Esta es una **versión de portafolio**. Para proteger información institucional, no incluye bancos de preguntas reales, respuestas correctas, nóminas de estudiantes, credenciales, identificadores de hojas de cálculo ni herramientas internas de mantenimiento/diagnóstico del sistema original. La interfaz y el backend muestran la arquitectura y las funciones principales; para ejecutar una instalación propia se debe configurar una hoja de cálculo nueva con datos ficticios y revisar las funciones de integración antes de desplegarla.

No conectes esta copia a la base de datos de un establecimiento ni a datos reales de estudiantes.

## Configuración general

1. Crea un proyecto nuevo en Google Apps Script.
2. Agrega el contenido de `Code.gs` al archivo de servidor.
3. Crea un archivo HTML llamado `Index` y pega el contenido de `Index.html`.
4. Crea una hoja de cálculo de prueba con datos ficticios y las pestañas/columnas que requiere la aplicación.
5. Configura las propiedades de acceso docente desde el entorno de prueba.
6. Revisa permisos y funciones antes de publicar la aplicación web.

Esta documentación no incluye datos reales ni credenciales de acceso.

## Autoría

Desarrollado como proyecto de programación y automatización de procesos educativos.
