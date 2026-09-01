/**
 * Publicar el Panel KPIs con un link · Apps Script
 * ------------------------------------------------
 * Sirve el panel como página web con una URL que puedes enviar al equipo.
 * Sin hosting externo, sin cuentas nuevas, y con el acceso controlado por Google.
 *
 * Montaje (5 minutos):
 *  1. script.google.com › Nuevo proyecto. Llámalo "Panel KPIs 101".
 *  2. Pega este archivo en Código.gs.
 *  3. Botón + junto a "Archivos" › HTML › nómbralo exactamente  panel
 *     (Apps Script le añade el .html solo).
 *  4. Borra el contenido de ejemplo de panel.html y pega dentro TODO el
 *     contenido de panel-kpis-101.html (ábrelo con un editor de texto,
 *     Ctrl+A, Ctrl+C, y pégalo).
 *  5. ANTES de pegar, rellena en panel-kpis-101.html el bloque CONFIG_FIJA
 *     (está al principio del <script>) con las URLs y tokens de tus dos
 *     conectores: así el equipo lo abre ya conectado.
 *  6. Implementar › Nueva implementación › Aplicación web:
 *       · Ejecutar como: yo
 *       · Quién tiene acceso: "Cualquier usuario con una cuenta de Google"
 *  7. La URL /exec resultante es el link para el equipo.
 *
 * Sobre el acceso — aquí es AL REVÉS que en los conectores:
 *  · Los conectores de DATOS necesitan "Cualquier usuario" para que el panel
 *    pueda leerlos por fetch; los protege el token.
 *  · El PANEL lo abre una persona en el navegador, así que aquí
 *    "Cualquier usuario con una cuenta de Google" sí funciona, y hace que
 *    Google pida login antes de mostrar nada. Los tokens incrustados en
 *    CONFIG_FIJA solo los ve quien pasa ese login.
 *
 * Para actualizar el panel más adelante: reemplaza el contenido de panel.html
 * y publica una versión nueva (Gestionar implementaciones › lápiz › Nueva
 * versión). La URL no cambia.
 */

function doGet() {
  return HtmlService.createHtmlOutputFromFile('panel')
    .setTitle('Panel KPIs · Area101 Communities')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
}
