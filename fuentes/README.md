# Fuentes del Panel KPIs 101

Copia de seguridad del código. **Los tokens están sustituidos por marcadores** porque este repo es público: los valores reales viven en los despliegues de Apps Script (y horneados dentro del `index.html` cifrado). Para restaurar un conector: pegar el `.gs`, poner el token real en la constante `TOKEN` (y `OPENAI_KEY` / `HUBSPOT_PAT` si aplica), ejecutar `instalar()` donde exista, e Implementar como aplicación web con acceso "Cualquier usuario".

- `conector-tdc.gs` — sirve ventas, oportunidades ("Lo que se viene") y objetivos (Dashboard 2026) desde el TDC. Caché de servidor 10 min (`?nocache=1` la salta).
- `conector-comunidad.gs` — ingesta diaria de los ficheros Miembros-Completos a un Sheet de caché; sirve comunidad, correos de invitados pendientes, `?insights=1` (OpenAI, servidor) y `?hubspot=1` (negocios abiertos). Incluye `informeSemanal()` por correo.
- `app-101.gs` — alternativa todo-en-uno (HtmlService + google.script.run, sin tokens, login Google).
- `publicar-panel.gs` — utilidad de publicación.

El panel (`index.html` en la raíz) va cifrado con StatiCrypt; el fuente en claro NO se sube aquí porque lleva los tokens horneados y datos de negocio embebidos.
