# Buscador CG INF (borrador)

Este directorio contiene el primer esqueleto del complemento Office.js. El manifiesto está configurado para un shared runtime y la app registra un evento `Worksheet.onChanged` en `buscador`, filtrado para `G2`.

## Antes de cargarlo en Excel

1. Publicar los archivos estáticos de este directorio en GitHub Pages.
2. Cargar `manifest.xml` desde Excel para probar el complemento.
3. Abrir el panel y activar la búsqueda automática una vez en cada libro. Office guarda ese comportamiento en el documento; en aperturas posteriores registra el evento al cargar el shared runtime.

El nombre mensual aceptado es uno de los doce meses en español seguido por un año de cuatro dígitos, por ejemplo `NOVIEMBRE 2026`. Este es un borrador y todavía requiere validación dentro de Excel Online, especialmente para comprobar disponibilidad de `SharedRuntime 1.1`, lectura de comentarios modernos y políticas de carga de complementos de la organización.
