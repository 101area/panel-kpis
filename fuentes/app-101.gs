/**
 * Panel KPIs 101 · Aplicación unificada (Apps Script) · v3
 * ---------------------------------------------------------
 * Una sola app que lo hace todo, pensada para Workspaces donde el administrador
 * bloquea el acceso anónimo a Apps Script (los conectores por fetch devuelven 404):
 *
 *   · Sirve el panel como página web (HtmlService) con login de Google delante.
 *   · Lee el TDC directamente (sin conector intermedio ni token).
 *   · Ingiere cada madrugada los ficheros diarios Miembros-Completos de Drive.
 *   · Guarda los cierres mensuales del equipo en un Sheet de caché propio.
 *
 * El panel, cuando corre aquí dentro, NO hace fetch a ninguna URL: habla con
 * este servidor por google.script.run, así que la política del dominio no le
 * afecta y no existen tokens que custodiar. La puerta es la cuenta de Google.
 *
 * Montaje (10 minutos):
 *  1. script.google.com › Nuevo proyecto → "Panel KPIs 101".
 *  2. Pega este archivo en Código.gs.
 *  3. Botón + › HTML › nómbralo exactamente  panel  y pega dentro TODO el
 *     contenido de panel-kpis-101.html.
 *  4. Ejecuta instalar() una vez: crea el Sheet de caché en tu Drive, deja el
 *     disparador diario a las 4:00 y arranca la primera ingesta. Ejecuta
 *     procesarPendientes() varias veces hasta que diga que no quedan días.
 *  5. Ejecuta probarTodo() y revisa el log.
 *  6. Implementar › Aplicación web → Ejecutar como: yo · Acceso:
 *     "Cualquier usuario con una cuenta de Google" (o solo tu dominio, si te
 *     aparece la opción: ambas funcionan aquí).
 *  7. La URL /exec es el link para el equipo.
 *
 * Para actualizar el panel: reemplaza panel.html y publica versión nueva.
 */

const TDC_ID = '1vHic67RLmqYuxaDoLn_ppHl6qsDJ1pVJzJ-e2-yhRSE';      // Sheet del TDC
const FOLDER_ID = '1BAZ15zXXRN7y4XesqceThgvrW3p0OsNI';               // carpeta con 2026/…

/** Sheet de caché de comunidad: se crea solo la primera vez y queda registrado. */
function cache() {
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty('cacheId');
  if (id) { try { return SpreadsheetApp.openById(id); } catch (e) { /* recrear */ } }
  const ss = SpreadsheetApp.create('Panel101 · Cache Comunidad');
  props.setProperty('cacheId', ss.getId());
  return ss;
}

function tdc() { return SpreadsheetApp.openById(TDC_ID); }

/* ============================================================
   Servir el panel y puentes google.script.run
   ============================================================ */

function doGet() {
  return HtmlService.createHtmlOutputFromFile('panel')
    .setTitle('Panel KPIs · Area101 Communities')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
}

/** Datos de negocio del TDC, mismo formato que el conector clásico. */
function obtenerTDC() {
  const ss = tdc();
  return {
    generado: Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd HH:mm'),
    origen: 'TDC en vivo · ' + ss.getName(),
    meses: MESES,
    lineas: leerLineas(ss),
    resumen: leerResumen(ss),
    mensual: leerMensual(ss)
  };
}

/** Datos de comunidad acumulados en la caché, mismo formato que el conector clásico. */
function obtenerComunidad() {
  const ss = cache();
  const meses = 18;
  const desde = new Date(); desde.setMonth(desde.getMonth() - meses);
  const corte = Utilities.formatDate(desde, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const base = construirComunidad_(ss, corte);
  base.version = 2;
  base.generado = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
  base.origen = 'Comunidad en vivo · ' + ss.getName();
  return base;
}

/** Cierres mensuales y ajustes del equipo (compartidos). */
function obtenerEstado() { return leerEstado(cache()); }
function guardarCierres(estado) { guardarEstado(cache(), estado || {}); return { ok: true }; }

function construirComunidad_(ss, corte) {

    const orgs = tabla(ss, T.ORG).map(function (r) {
      let msu = {};
      try { msu = JSON.parse(r[8] || '{}'); } catch (err) { msu = {}; }
      return { id: r[0], nombre: r[1], pais: r[2], sector: r[3], asientos: r[4],
        susc: r[5], fin: iso(r[6]), usos: num(r[7]), msu: msu };
    });
    const cols = {};
    orgs.forEach(function (o) { Object.keys(o.msu).forEach(function (k) { cols[k] = 1; }); });

    return {
      version: 2,
      generado: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm'),
      origen: 'Comunidad en vivo · ' + ss.getName(),
      msuCols: Object.keys(cols),
      orgs: orgs,
      miembros: tabla(ss, T.MEM).map(function (r) {
        return { id: r[0], org: r[1], tier: r[2], acceso: iso(r[3]), registro: iso(r[4]),
          invitacion: iso(r[5]) };
      }),
      eventos: tabla(ss, T.EV)
        .filter(function (r) { return iso(r[0]) >= corte; })
        .map(function (r) {
          return { f: iso(r[0]), t: r[1], id: r[2], org: r[3], est: r[4], tit: r[5] };
        }),
      invitPend: tabla(ss, T.INV).map(function (r) {
        let fechas = [];
        try { fechas = JSON.parse(r[3] || '[]'); } catch (err) { fechas = []; }
        return { org: r[0], pendientes: num(r[1]), caducadas: num(r[2]), fechas: fechas };
      }),
      serie: tabla(ss, T.SERIE).map(function (r) {
        return { f: iso(r[0]), miembros: num(r[1]),
          porTier: { leader: num(r[2]), starter: num(r[3]), lite: num(r[4]), member: num(r[5]) },
          orgsPago: num(r[6]), orgsTotal: num(r[7]), acceso30d: num(r[8]) };
      }),
      dias: tabla(ss, T.LOG).length,
      errores: tabla(ss, T.LOG).filter(function (r) { return r[2] === 'error'; })
        .map(function (r) { return { fecha: iso(r[1]), nota: r[4] }; })
    };

}

/** Unidades de negocio que el panel espera en el cuadro de mando. */
const UNIDADES = ['101', 'Area101', 'Area101 Communities', 'Area101 Innovación',
  'Area101 Talento', 'Solutions101', 'Solutions101 Tech', 'Solutions101 Services', 'Lab101'];

const MESES = ['2025', 'Ene', 'Feb', 'Mar', 'Abr', 'May', 'Jun', 'Jul', 'Ago', 'Sep', 'Oct', 'Nov', 'Dic'];


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


const MAX_SEGUNDOS = 240;
const MIN_MIEMBROS = 50;

const T = { EV: 'eventos', ORG: 'orgs', MEM: 'miembros', SERIE: 'serie',
  INV: 'invitaciones', LOG: '_procesados' };


/** Crea la pestaña si falta y reescribe SIEMPRE la fila de cabecera: así una
 *  versión nueva con columnas nuevas se aplica sin tocar los datos. */
function cab(ss, nombre, cols) {
  let sh = ss.getSheetByName(nombre);
  if (!sh) sh = ss.insertSheet(nombre);
  sh.getRange(1, 1, 1, cols.length).setValues([cols]).setFontWeight('bold');
  sh.setFrozenRows(1);
  return sh;
}

/* ============================== ingesta ============================== */

function procesarPendientes() {
  const t0 = Date.now();
  const ss = cache();
  const log = ss.getSheetByName(T.LOG) || cab(ss, T.LOG, ['fileId', 'fecha', 'estado', 'procesadoEn', 'nota']);
  const hechos = {};
  if (log.getLastRow() > 1) {
    log.getRange(2, 1, log.getLastRow() - 1, 3).getValues()
      .forEach(function (r) { if (r[0]) hechos[r[0]] = r[2]; });
  }

  const ficheros = listarDias().filter(function (f) { return hechos[f.id] !== 'ok'; });
  if (!ficheros.length) { Logger.log('No quedan días por procesar.'); return; }
  Logger.log('Pendientes: %s días.', ficheros.length);

  const evIdx = indiceEventos(ss);
  let n = 0;
  for (let k = 0; k < ficheros.length; k++) {
    if (Date.now() - t0 > MAX_SEGUNDOS * 1000) {
      Logger.log('Corto por tiempo. Procesados %s, quedan %s. Vuelve a ejecutar.', n, ficheros.length - k);
      break;
    }
    const f = ficheros[k];
    try {
      const r = procesarDia(ss, f, evIdx);
      log.appendRow([f.id, f.fecha, 'ok', new Date(), r]);
      n++;
    } catch (err) {
      log.appendRow([f.id, f.fecha, 'error', new Date(), String(err).slice(0, 200)]);
      Logger.log('Error en %s: %s', f.nombre, err);
    }
  }
  Logger.log('Procesados %s días en esta ejecución.', n);
}

function listarDias() {
  const raiz = DriveApp.getFolderById(FOLDER_ID);
  const out = [];
  const anios = raiz.getFolders();
  while (anios.hasNext()) {
    const meses = anios.next().getFolders();
    while (meses.hasNext()) {
      const archivos = meses.next().getFiles();
      while (archivos.hasNext()) {
        const a = archivos.next();
        const m = a.getName().match(/(\d{2})-(\d{2})-(\d{4})/);
        if (!m || a.getMimeType() !== MimeType.GOOGLE_SHEETS) continue;
        out.push({ id: a.getId(), nombre: a.getName(), fecha: m[3] + '-' + m[2] + '-' + m[1] });
      }
    }
  }
  out.sort(function (a, b) { return a.fecha < b.fecha ? -1 : 1; });
  return out;
}

/* Índice de eventos ya guardados: clave → fila, para poder actualizar estados.
   Sesiones y perks: una fila por persona-evento (el estado y la fecha cambian).
   Contenidos y manuales: la fecha forma parte de la clave (repetir es legítimo). */
function claveEv(t, id, tit, f) {
  const base = [t, id, String(tit).slice(0, 60)].join('|');
  return (t === 'sesion' || t === 'perk') ? base : base + '|' + f;
}
function indiceEventos(ss) {
  const sh = ss.getSheetByName(T.EV);
  const idx = {};
  if (sh && sh.getLastRow() > 1) {
    sh.getRange(2, 1, sh.getLastRow() - 1, 6).getValues().forEach(function (r, i) {
      idx[claveEv(r[1], r[2], r[5], iso(r[0]))] = { fila: i + 2, est: String(r[4]), f: iso(r[0]) };
    });
  }
  return idx;
}

function procesarDia(ss, f, evIdx) {
  const src = SpreadsheetApp.openById(f.id);

  const mem = leerHoja(src, 'Miembros');
  if (mem.length < MIN_MIEMBROS) throw new Error('fichero incompleto: ' + mem.length + ' miembros');
  const orgs = leerHoja(src, 'Organizaciones');

  /* nombre → id para emparejar los servicios manuales */
  const porNombre = {};
  mem.forEach(function (m) {
    const k = slugNombre(String(m['Nombre'] || '') + ' ' + String(m['Apellidos'] || ''));
    if (k) porNombre[k] = ent(m['ID']);
  });

  /* --- eventos: alta de nuevos + actualización de estados --- */
  const shEv = ss.getSheetByName(T.EV);
  const nuevos = [], cambios = [];
  function evento(fecha, tipo, id, org, est, tit) {
    const k = claveEv(tipo, id, tit, fecha);
    const ya = evIdx[k];
    if (!ya) {
      evIdx[k] = { fila: null, est: est, f: fecha };
      nuevos.push([fecha, tipo, id, org, est, tit]);
    } else if ((tipo === 'sesion' || tipo === 'perk') && (ya.est !== est || ya.f !== fecha) && ya.fila) {
      cambios.push({ fila: ya.fila, valores: [fecha, tipo, id, org, est, tit] });
      ya.est = est; ya.f = fecha;
    }
  }
  [['Sesiones', 'Evento', 'sesion'], ['Perks', 'Perk', 'perk'], ['Contenidos', 'Título', 'contenido']]
    .forEach(function (cfg) {
      leerHoja(src, cfg[0]).forEach(function (x) {
        const fecha = fISO(x['Fecha']);
        if (!fecha) return;
        evento(fecha, cfg[2], ent(x['ID']), x['Organización'] || '—',
          String(x['Estado'] || x['Tipo'] || ''), String(x[cfg[1]] || '').slice(0, 120));
      });
    });
  leerHoja(src, 'Servicios manuales').forEach(function (x) {
    const fecha = fISO(x['Cuándo']);
    if (!fecha) return;
    const tit = (String(x['Servicio'] || '') + ' · ' + String(x['Detalle'] || '')).slice(0, 120);
    evento(fecha, 'manual', buscarId(porNombre, x['Persona']), x['Empresa'] || '—', 'Servicio', tit);
  });
  cambios.forEach(function (c) { shEv.getRange(c.fila, 1, 1, 6).setValues([c.valores]); });
  if (nuevos.length) {
    const fila0 = shEv.getLastRow() + 1;
    shEv.getRange(fila0, 1, nuevos.length, 6).setValues(nuevos);
    nuevos.forEach(function (v, i) {
      evIdx[claveEv(v[1], v[2], v[5], v[0])].fila = fila0 + i;
    });
  }

  /* --- estado (miembros, orgs, invitaciones): solo si el día es el más reciente --- */
  const props = PropertiesService.getScriptProperties();
  const ultimo = props.getProperty('ultimoEstado') || '';
  if (f.fecha >= ultimo) {
    volcar(ss, T.MEM, mem.map(function (m) {
      return [ent(m['ID']), m['Organización'] || '—', tier(m['Membership']),
        fISO(m['Último Acceso']), fISO(m['Fecha Registro']), fISO(m['Fecha Invitación'])];
    }), 6);

    const FIJAS = { 'ID': 1, 'Nombre': 1, 'Imagen': 1, 'País': 1, 'Sector': 1, '# Asientos': 1,
      'Suscripción': 1, 'Fecha terminación suscripción': 1 };
    volcar(ss, T.ORG, orgs.map(function (o) {
      const msu = {};
      let usos = 0;
      Object.keys(o).forEach(function (k) {
        if (FIJAS[k] || !k) return;
        const v = num(o[k]);
        if (v) { msu[k] = v; usos += v; }
      });
      return [ent(o['ID']), o['Nombre'] || '—', o['País'] || '', o['Sector'] || '',
        ent(o['# Asientos']) || 0, tier(o['Suscripción']), fISO(o['Fecha terminación suscripción']),
        usos, JSON.stringify(msu)];
    }), 9);

    const porOrg = {};
    leerHoja(src, 'Invitados sin registrar').forEach(function (x) {
      const org = x['Organización'] || '—';
      const a = porOrg[org] || (porOrg[org] = { p: 0, c: 0, f: [] });
      a.p++;
      if (/^(s[ií]|yes|true|1)$/i.test(String(x['Invitación caducada'] || '').trim())) a.c++;
      const fi = fISO(x['Fecha invitación']);
      if (fi) a.f.push(fi);
    });
    volcar(ss, T.INV, Object.keys(porOrg).map(function (org) {
      return [org, porOrg[org].p, porOrg[org].c, JSON.stringify(porOrg[org].f.sort())];
    }), 4);

    props.setProperty('ultimoEstado', f.fecha);
  }

  /* --- punto de la serie diaria --- */
  const cnt = { leader: 0, starter: 0, lite: 0, member: 0 };
  let acc30 = 0;
  const hoy = new Date(f.fecha + 'T00:00:00');
  mem.forEach(function (m) {
    const t = tier(m['Membership']);
    if (cnt[t] != null) cnt[t]++;
    const a = fISO(m['Último Acceso']);
    if (a && (hoy - new Date(a + 'T00:00:00')) / 86400000 <= 30) acc30++;
  });
  const pago = orgs.filter(function (o) {
    const t = tier(o['Suscripción']); return t === 'leader' || t === 'starter' || t === 'lite';
  }).length;
  upsertSerie(ss, [f.fecha, mem.length, cnt.leader, cnt.starter, cnt.lite, cnt.member,
    pago, orgs.length, acc30]);

  return mem.length + ' miembros, ' + nuevos.length + ' eventos nuevos, ' + cambios.length + ' estados actualizados';
}

function leerHoja(ss, nombre) {
  const sh = ss.getSheetByName(nombre);
  if (!sh || sh.getLastRow() < 2) return [];
  const vals = sh.getDataRange().getValues();
  const hdr = vals[0].map(function (c) { return String(c == null ? '' : c).trim(); });
  const out = [];
  for (let i = 1; i < vals.length; i++) {
    const r = vals[i];
    if (r.every(function (c) { return c === '' || c == null; })) continue;
    const o = {};
    for (let j = 0; j < hdr.length; j++) if (hdr[j]) o[hdr[j]] = r[j];
    out.push(o);
  }
  return out;
}

function volcar(ss, nombre, filas, cols) {
  const sh = ss.getSheetByName(nombre);
  const filasViejas = sh.getLastRow() - 1, colsViejas = sh.getLastColumn();
  if (filasViejas > 0) sh.getRange(2, 1, filasViejas, Math.max(cols, colsViejas)).clearContent();
  if (filas.length) sh.getRange(2, 1, filas.length, cols).setValues(filas);
}

function upsertSerie(ss, fila) {
  const sh = ss.getSheetByName(T.SERIE);
  const n = sh.getLastRow();
  if (n > 1) {
    const fechas = sh.getRange(2, 1, n - 1, 1).getValues();
    for (let i = 0; i < fechas.length; i++) {
      if (String(fechas[i][0]).slice(0, 10) === fila[0] || iso(fechas[i][0]) === fila[0]) {
        sh.getRange(i + 2, 1, 1, fila.length).setValues([fila]); return;
      }
    }
  }
  sh.appendRow(fila);
}

/* ============================== utilidades ============================== */

function tier(v) {
  const s = String(v == null ? '' : v).trim().toLowerCase();
  if (s.indexOf('leader') >= 0) return 'leader';
  if (s.indexOf('starter') >= 0) return 'starter';
  if (s.indexOf('lite') >= 0) return 'lite';
  if (s.indexOf('trial') >= 0) return 'trial';
  if (s.indexOf('member') >= 0) return 'member';
  return 'otro';
}
function slugNombre(s) {
  return String(s || '').toLowerCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-z0-9 ]/g, ' ').replace(/\s+/g, ' ').trim();
}
function buscarId(porNombre, persona) {
  const s = slugNombre(persona);
  if (!s) return null;
  if (porNombre[s] != null) return porNombre[s];
  const claves = Object.keys(porNombre);
  for (let i = 0; i < claves.length; i++) {
    if (claves[i].indexOf(s) >= 0 || s.indexOf(claves[i]) >= 0) return porNombre[claves[i]];
  }
  return null;
}
function fISO(v) {
  if (!v) return '';
  if (Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v)) {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  const m = String(v).match(/^(\d{1,2})-(\d{1,2})-(\d{4})/);
  if (m) return m[3] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[1]).slice(-2);
  const m2 = String(v).match(/^\d{4}-\d{2}-\d{2}/);
  return m2 ? m2[0] : '';
}
function num(v) {
  if (typeof v === 'number') return v;
  const n = parseFloat(String(v == null ? '' : v).replace(/[^\d.\-]/g, ''));
  return isNaN(n) ? 0 : n;
}
function ent(v) { const n = num(v); return n ? Math.round(n) : null; }
function iso(v) {
  if (!v) return '';
  if (Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v)) {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return String(v).slice(0, 10);
}


function tabla(ss, nombre) {
  const sh = ss.getSheetByName(nombre);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues()
    .filter(function (r) { return r[0] !== '' && r[0] != null; });
}


/* ============================================================
   Comprobación
   ============================================================ */
function probarTodo() {
  const t = tdc();
  Logger.log('TDC "%s": %s líneas · resumen %s unidades · series %s',
    t.getName(), leerLineas(t).length, leerResumen(t).length, Object.keys(leerMensual(t)).join(','));
  const dias = listarDias();
  Logger.log('Drive: %s ficheros diarios (último %s)', dias.length, dias.length ? dias[dias.length - 1].fecha : '—');
  const c = cache();
  [T.EV, T.ORG, T.MEM, T.SERIE, T.INV, T.LOG].forEach(function (x) {
    Logger.log('Cache %s: %s filas', x, tabla(c, x).length);
  });
}
