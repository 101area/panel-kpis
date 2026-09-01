/**
 * Conector TDC → Panel KPIs Area101 Communities
 * ---------------------------------------------
 * Se instala DENTRO del propio TDC (Extensiones › Apps Script), así que el Sheet
 * sigue siendo privado: este script solo publica las columnas que el panel necesita.
 *
 * Instalación
 *  1. Cambia TOKEN por una cadena larga tuya.
 *  2. Implementar › Nueva implementación › Aplicación web.
 *     · Ejecutar como: yo (tu cuenta)
 *     · Quién tiene acceso: cualquier usuario     ← necesario para que el panel pueda leer
 *  3. Copia la URL que acaba en /exec y pégala en el panel, en «Datos y conexión».
 *
 * Nota sobre el acceso: la URL es pública pero inútil sin el token, y solo devuelve
 * cliente, proyecto, unidad, fechas e importes de venta. No expone costes, márgenes,
 * facturación, contactos ni nada de Holded.
 */

const TOKEN = 'PON-AQUI-EL-TOKEN-DEL-TDC';
const OPORTUNIDADES_ID = '1hqJ4abJ4OcOehG3u7NnfY4sCPO-31cAQKR1CxKMZoVs';  // hoja "Lo que se viene"

/** Pestaña donde se guardan los cierres mensuales del panel. Se crea sola. */
const HOJA_ESTADO = 'Panel_KPIs';

/** Unidades de negocio que el panel espera en el cuadro de mando. */
const UNIDADES = ['101', 'Area101', 'Area101 Communities', 'Area101 Innovación',
  'Area101 Talento', 'Solutions101', 'Solutions101 Tech', 'Solutions101 Services', 'Lab101'];

const MESES = ['2025', 'Ene', 'Feb', 'Mar', 'Abr', 'May', 'Jun', 'Jul', 'Ago', 'Sep', 'Oct', 'Nov', 'Dic'];

/* ============================== entrada ============================== */

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (TOKEN && p.token !== TOKEN) return json({ error: 'Token no válido' });
  try {
    const cs = CacheService.getScriptCache();
    if (!p.nocache) {
      const hit = cs.get('tdc_v2');
      if (hit) return ContentService.createTextOutput(hit).setMimeType(ContentService.MimeType.JSON);
    }
    const ss = SpreadsheetApp.getActive();
    return jsonCache(cs, {
      generado: Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd HH:mm'),
      origen: 'TDC en vivo · ' + ss.getName(),
      meses: MESES,
      lineas: leerLineas(ss),
      resumen: leerResumen(ss),
      mensual: leerMensual(ss),
      oportunidades: leerOportunidades(),
      objetivos: leerObjetivos(ss),
      estado: leerEstado(ss)
    });
  } catch (err) {
    return json({ error: String(err) });
  }
}

function doPost(e) {
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (TOKEN && body.token !== TOKEN) return json({ error: 'Token no válido' });
    guardarEstado(SpreadsheetApp.getActive(), body.estado || {});
    return json({ ok: true });
  } catch (err) {
    return json({ error: String(err) });
  }
}

function json(o) {
  return ContentService.createTextOutput(JSON.stringify(o))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ========================= líneas de proyecto ========================= */

/** Alias de cabecera → campo que espera el panel. */
const COLS = {
  cliente: ['cliente'],
  proyecto: ['proyecto'],
  unidad: ['unidad de negocio', 'area'],
  servicio: ['servicio', 'categoria', 'categoría'],
  tipo: ['tipo', 'tipo de venta'],
  origen: ['origen'],
  dir: ['dir'],
  estado: ['avance'],
  inicio: ['inicio'],
  fin: ['fin'],
  ventaAnio: ['venta año en curso', 'ventas año'],
  ventaTotal: ['venta total (s/iva)', 'venta (s/iva)']
};

function norm(s) {
  return String(s == null ? '' : s).trim().toLowerCase();
}

/** Busca en cada pestaña una fila de cabecera con Cliente + Proyecto y lee las filas siguientes. */
function leerLineas(ss) {
  const out = [];
  ss.getSheets().forEach(function (sh) {
    if (sh.getName() === HOJA_ESTADO) return;
    const rng = sh.getDataRange();
    if (rng.getNumRows() < 2) return;
    const vals = rng.getValues();

    let hi = -1;
    for (let i = 0; i < Math.min(vals.length, 30); i++) {
      const row = vals[i].map(norm);
      if (row.indexOf('cliente') >= 0 && row.indexOf('proyecto') >= 0) { hi = i; break; }
    }
    if (hi < 0) return;

    const hdr = vals[hi].map(norm);
    const idx = {};
    Object.keys(COLS).forEach(function (campo) {
      for (let a = 0; a < COLS[campo].length; a++) {
        const j = hdr.indexOf(COLS[campo][a]);
        if (j >= 0) { idx[campo] = j; return; }
      }
    });
    if (idx.cliente == null || idx.proyecto == null) return;

    const esAntigua = idx.unidad != null && hdr[idx.unidad] === 'area';

    for (let i = hi + 1; i < vals.length; i++) {
      const r = vals[i];
      const cliente = String(r[idx.cliente] || '').trim();
      const proyecto = String(r[idx.proyecto] || '').trim();
      if (!cliente || !proyecto) continue;

      let unidad = idx.unidad != null ? String(r[idx.unidad] || '').trim() : '';
      // la pestaña antigua usa "Area101" donde la nueva usa "Area101 Communities"
      if (esAntigua && unidad === 'Area101') unidad = 'Area101 Communities';

      out.push({
        cliente: cliente,
        proyecto: proyecto,
        unidad: unidad || '—',
        servicio: idx.servicio != null ? String(r[idx.servicio] || '').trim() : '',
        tipo: idx.tipo != null ? String(r[idx.tipo] || '').trim() : '',
        origen: idx.origen != null ? String(r[idx.origen] || '').trim() : '',
        dir: idx.dir != null ? String(r[idx.dir] || '').trim() : '',
        estado: idx.estado != null ? String(r[idx.estado] || '').trim() : '',
        inicio: fecha(r[idx.inicio], ss),
        fin: fecha(r[idx.fin], ss),
        ventaAnio: numero(r[idx.ventaAnio]),
        ventaTotal: numero(r[idx.ventaTotal])
      });
    }
  });
  return out;
}

function fecha(v, ss) {
  if (!v) return null;
  if (Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v)) {
    return Utilities.formatDate(v, ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd');
  }
  const m = String(v).match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/);
  if (m) return m[3] + '-' + pad(m[2]) + '-' + pad(m[1]);
  const m2 = String(v).match(/^\d{4}-\d{2}-\d{2}/);
  return m2 ? m2[0] : null;
}

function pad(n) { return ('0' + n).slice(-2); }

function numero(v) {
  if (typeof v === 'number') return v;
  if (v == null || v === '') return 0;
  let s = String(v).replace(/[^\d,.\-]/g, '');
  if (/,\d{1,2}$/.test(s)) s = s.replace(/\./g, '').replace(',', '.');
  else s = s.replace(/,/g, '');
  const n = parseFloat(s);
  return isNaN(n) ? 0 : n;
}

/* ========================== cuadro de mando ========================== */

/** Bloque BP / Reforecast / YTD por unidad de negocio. */
function leerResumen(ss) {
  const out = [];
  const vistos = {};
  ss.getSheets().forEach(function (sh) {
    if (sh.getName() === HOJA_ESTADO) return;
    const vals = sh.getDataRange().getValues();
    for (let i = 0; i < vals.length; i++) {
      const fila = vals[i];
      const u = String(fila[0] || '').trim();
      if (UNIDADES.indexOf(u) < 0 || vistos[u]) continue;
      // fila de resumen: BP, Reforecast, YTD en las tres columnas siguientes
      const bp = numero(fila[1]), ytd = numero(fila[3]);
      if (!bp || String(fila[1]).toUpperCase() === 'BP') continue;
      vistos[u] = true;
      out.push({ unidad: u, bp: bp, reforecast: numero(fila[2]), ytd: ytd });
    }
  });
  return out;
}

/** Series acumuladas de 13 puntos (arrastre + 12 meses) por unidad. */
function leerMensual(ss) {
  const out = {};
  ss.getSheets().forEach(function (sh) {
    if (sh.getName() === HOJA_ESTADO) return;
    const vals = sh.getDataRange().getValues();
    for (let i = 0; i < vals.length; i++) {
      const u = String(vals[i][0] || '').trim();
      if (UNIDADES.indexOf(u) < 0) continue;
      if (String(vals[i][1] || '').trim().toUpperCase() !== 'BP') continue;
      if (out[u]) continue;
      const s = { bp: [], reforecast: [], ytd: [] };
      for (let j = i + 1; j < vals.length && s.bp.length < 13; j++) {
        if (!/^\d+\./.test(String(vals[j][0] || '').trim())) break;
        s.bp.push(numero(vals[j][1]));
        s.reforecast.push(numero(vals[j][4]));
        const y = numero(vals[j][7]);
        s.ytd.push(y ? y : null);
      }
      if (s.bp.length) out[u] = s;
    }
  });
  return out;
}

/* ===================== cierres mensuales del panel ===================== */

function hojaEstado(ss) {
  let sh = ss.getSheetByName(HOJA_ESTADO);
  if (!sh) {
    sh = ss.insertSheet(HOJA_ESTADO);
    sh.getRange('A1').setValue('No editar a mano: el Panel KPIs guarda aquí los cierres mensuales.');
    sh.hideSheet();
  }
  return sh;
}

const TROZO = 40000; // límite práctico por celda

function guardarEstado(ss, estado) {
  const sh = hojaEstado(ss);
  const txt = JSON.stringify(estado);
  const trozos = [];
  for (let i = 0; i < txt.length; i += TROZO) trozos.push([txt.substr(i, TROZO)]);
  if (!trozos.length) trozos.push(['']);
  const ultima = Math.max(sh.getLastRow(), 2);
  sh.getRange(2, 1, ultima - 1, 1).clearContent();
  sh.getRange(2, 1, trozos.length, 1).setValues(trozos);
  sh.getRange('B1').setValue('Última escritura: ' +
    Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd HH:mm'));
}

function leerEstado(ss) {
  const sh = ss.getSheetByName(HOJA_ESTADO);
  if (!sh || sh.getLastRow() < 2) return null;
  const txt = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues()
    .map(function (r) { return r[0] || ''; }).join('');
  if (!txt) return null;
  try { return JSON.parse(txt); } catch (err) { return null; }
}

function jsonCache(cs, o) {
  const s = JSON.stringify(o);
  try { if (s.length < 95000) cs.put('tdc_v2', s, 600); } catch (err) { }
  return ContentService.createTextOutput(s).setMimeType(ContentService.MimeType.JSON);
}

/* ==================== oportunidades y objetivos (v2) ==================== */
function leerOportunidades() {
  try {
    const sh = SpreadsheetApp.openById(OPORTUNIDADES_ID).getSheets()[0];
    const vals = sh.getDataRange().getValues(), out = [];
    vals.forEach(function (r) {
      const c = String(r[3] || '').trim();
      if (!c || c === 'Cliente') return;
      const v = numero(r[7]); if (!v) return;
      const pn = function (x) { const n = numero(x); return n == null ? null : (n <= 1 ? n * 100 : n); };
      out.push({ u: String(r[1] || '').trim() || '—', c: c, p: String(r[4] || '').trim(),
        f: fecha(r[5]), e: String(r[6] || '').trim(), v: v,
        pct: pn(r[8]), prob: pn(r[10]), w: numero(r[13]) || numero(r[11]) || 0 });
    });
    return out;
  } catch (err) { return []; }
}
function leerObjetivos(ss) {
  try {
    const shs = ss.getSheets();
    for (let k = 0; k < shs.length; k++) {
      const n = Math.min(shs[k].getLastRow(), 400); if (n < 14) continue;
      const vals = shs[k].getRange(1, 1, n, 9).getValues();
      for (let i = 0; i < vals.length - 13; i++) {
        if (String(vals[i][1]).trim() === 'BP' && String(vals[i][2]).trim() === 'BP Mes') {
          const bp = [], refor = [], ytd = [];
          for (let m = 1; m <= 12; m++) {
            const r = vals[i + 1 + m] || [];
            bp.push(numero(r[1]) || 0); refor.push(numero(r[4]) || 0);
            const y = numero(r[7]); ytd.push(y ? y : null);
          }
          return { bp: bp, refor: refor, ytd: ytd, bpAnual: bp[11], reforAnual: refor[11],
            arrastre2025: numero((vals[i + 1] || [])[1]) || 0 };
        }
      }
    }
  } catch (err) { /* opcional */ }
  return null;
}

/* ============================ comprobación ============================ */

/** Ejecuta esto desde el editor para ver si el conector lee bien el TDC. */
function probar() {
  const ss = SpreadsheetApp.getActive();
  const lineas = leerLineas(ss);
  const resumen = leerResumen(ss);
  const mensual = leerMensual(ss);
  Logger.log('Líneas leídas: %s', lineas.length);
  Logger.log('Unidades en el resumen: %s', resumen.map(function (r) { return r.unidad; }).join(', '));
  Logger.log('Series mensuales: %s', Object.keys(mensual).join(', '));
  Logger.log('Ejemplo de línea: %s', JSON.stringify(lineas[0]));
  const memb = lineas.filter(function (l) { return /leader|starter|lite/i.test(l.proyecto); });
  Logger.log('Líneas de membresía: %s', memb.length);
}
