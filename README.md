# Fútbol 7 mixtos

Aplicación para armar equipos de fútbol 7 y registrar partidos mixtos.

## Calificación de jugadores

Cada perfil requiere dos selecciones independientes:

- Estado o potencia física: baja, intermedia o alta.
- Habilidad con el balón: baja, intermedia o alta.

Ambas se consultan al crear el perfil, sin valores preseleccionados, y se pueden
modificar al editarlo. La guía usa la misma escala para todos los géneros.

Para equilibrar equipos, cada nivel corresponde a 2, 6 o 10 puntos. El promedio
de ambas variables da la puntuación base, con el mismo peso para cada una.
Se conserva el ajuste existente por resultados de partidos (hasta ±1,5 puntos).
La puntuación calculada se limita al rango 1–10.

Los jugadores anteriores quedan pendientes de evaluación. Mientras se completan
sus dos variables, usan una base provisional de 6 puntos. Las estrellas antiguas
no se convierten automáticamente: no permiten deducir ambas dimensiones.

## Activación

Actualizar Apps Script antes de publicar el frontend:

1. Abrir el proyecto de Google Apps Script vinculado a la hoja utilizada por la aplicación.
2. Guardar una copia de la hoja y del script vigente.
3. Reemplazar el código del proyecto por el archivo Codigo.gs de este repositorio.
4. Ejecutar initSheets() desde el editor y autorizar su ejecución si Google lo solicita.
   Añade estadoFisico y habilidadBalon al final de Jugadores, además de aliases
   si faltaba. No borra ni convierte los valores antiguos, ni los partidos.
   Ejecutar la migración otra vez no duplica columnas.
5. En Implementar → Administrar implementaciones, editar la implementación web
   existente, seleccionar una nueva versión y guardar. Conservar su URL y sus
   permisos de acceso vigentes.
6. Publicar index.html mediante el mecanismo habitual del sitio y recargarlo.
   Si se crea otra implementación con una URL diferente, actualizar API_URL.
7. Crear un jugador con ambos niveles, recargar y comprobar que se conservan.
   Editar un perfil antiguo y completar sus calificaciones.

La lectura de datos también asegura las columnas nuevas. El frontend bloquea el
guardado de jugadores si detecta un backend anterior, con un mensaje de actualización.
Los errores JSON del backend se muestran como errores, incluso con respuesta HTTP 200.

La columna rating se conserva como dato histórico; la nueva versión no la utiliza
para calcular ni actualizar calificaciones. El frontend anterior no debe utilizarse
para crear jugadores después de actualizar Apps Script.

## Validación local

Ejecutar con Node.js, sin instalar dependencias:

    node test.cjs

Las pruebas simulan SpreadsheetApp, las respuestas HTTP y el DOM. Cubren migración
repetible, conservación de perfiles e historial, valores inválidos, guardado y
recarga de ambas variables, creación/edición, nueve combinaciones de niveles,
compatibilidad con perfiles antiguos, errores de API y estructura HTML.
La verificación real en la hoja se realiza con el paso 7 de activación.
