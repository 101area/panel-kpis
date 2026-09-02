/**
 * Conector Comunidad → Panel KPIs Area101 Communities · v2
 * --------------------------------------------------------
 * Lee los ficheros diarios "Miembros-Completos-DD-MM-AAAA" de Drive, los va
 * acumulando en este mismo Sheet y los sirve al panel como JSON.
 *
 * Novedades v2 (para la nueva versión de los ficheros):
 *  · Miembros: guarda la Fecha Invitación → funnel de invitaciones automático.
 *  · Servicios manuales: los apuntes del equipo cuentan como interacción,
 *    emparejando la persona con su ID de miembro por nombre.
 *  · Invitados sin registrar: se agrega por organización (sin guardar correos).
 *  · Estados de sesión que cambian (Registrado → Asistido) se ACTUALIZAN en
 *    lugar de ignorarse: sin esto, la asistencia se quedaba en cero.
 *
 * Migración desde v1: pega este archivo encima del anterior y ejecuta
 * instalar() una vez. No borra datos; añade la pestaña y columnas nuevas.
 * Después, Implementar › Gestionar implementaciones › lápiz › Nueva versión.
 *
 * Instalación desde cero: igual que v1 (ver puesta-en-marcha.md).
 *
 * Privacidad: el JSON que sirve al panel NO incluye nombres ni correos, solo
 * el ID interno del miembro.
 */

const TOKEN = 'PON-AQUI-EL-TOKEN-DE-COMUNIDAD';

/* Insights IA (opcional): la clave vive AQUI, en servidor; jamás en el panel.
   Recomendado: rota la clave en platform.openai.com si la has compartido por chat. */
const OPENAI_KEY = 'PEGA-AQUI-TU-CLAVE-OPENAI';
const HUBSPOT_PAT = 'PEGA-AQUI-TU-PAT-DE-HUBSPOT-O-DEJALO-ASI';
/* Informe de los lunes 8:00 por correo. Pon destinatarios ('a@x.com,b@x.com')
   y vuelve a ejecutar instalar() para crear el disparador. Vacío = apagado. */
const CORREOS_INFORME = '';
const NPS_SHEET_ID = '1oFApWiAiUSTescI6IFbv4W1Fyo2NFw2uMDYP3Aj7GzY';  // encuesta semestral de miembros
const FOLDER_ID = '1BAZ15zXXRN7y4XesqceThgvrW3p0OsNI';   // carpeta raíz, la que tiene 2026/

const MAX_SEGUNDOS = 240;
const MIN_MIEMBROS = 50;

const T = { EV: 'eventos', ORG: 'orgs', MEM: 'miembros', SERIE: 'serie',
  INV: 'invitaciones', LOG: '_procesados' };

/* ============================== instalación ============================== */

function instalar() {
  if (CORREOS_INFORME && !ScriptApp.getProjectTriggers().some(function (t) { return t.getHandlerFunction() === 'informeSemanal'; }))
    ScriptApp.newTrigger('informeSemanal').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(8).create();
  const ss = SpreadsheetApp.getActive();
  cab(ss, T.EV, ['fecha', 'tipo', 'memberId', 'org', 'estado', 'titulo']);
  cab(ss, T.ORG, ['id', 'nombre', 'pais', 'sector', 'asientos', 'suscripcion', 'fin', 'usos', 'msu']);
  cab(ss, T.MEM, ['id', 'org', 'tier', 'acceso', 'registro', 'invitacion']);
  cab(ss, T.SERIE, ['fecha', 'miembros', 'leader', 'starter', 'lite', 'member', 'orgsPago', 'orgsTotal', 'acceso30d']);
  cab(ss, T.INV, ['org', 'pendientes', 'caducadas', 'fechas', 'correos']);
  cab(ss, T.LOG, ['fileId', 'fecha', 'estado', 'procesadoEn', 'nota']);
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'procesarPendientes') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('procesarPendientes').timeBased().atHour(4).everyDays(1).create();
  Logger.log('Pestañas al día y disparador diario a las 4:00 instalado.');
  procesarPendientes();
}

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
  const ss = SpreadsheetApp.getActive();
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
      const a = porOrg[org] || (porOrg[org] = { p: 0, c: 0, f: [], m: [] });
      a.p++;
      const correo = String(x['Correo'] || x['Email'] || x['Correo electrónico'] || '').trim();
      if (correo) a.m.push(correo);
      if (/^(s[ií]|yes|true|1)$/i.test(String(x['Invitación caducada'] || '').trim())) a.c++;
      const fi = fISO(x['Fecha invitación']);
      if (fi) a.f.push(fi);
    });
    volcar(ss, T.INV, Object.keys(porOrg).map(function (org) {
      return [org, porOrg[org].p, porOrg[org].c, JSON.stringify(porOrg[org].f.sort()),
        JSON.stringify(porOrg[org].m)];
    }), 5);

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

/* ============================== servir ============================== */

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (TOKEN && p.token !== TOKEN) return json({ error: 'Token no válido' });
  if (p.insights) return json(insightsIA());
  if (p.hubspot) return json(hubspotAbiertos());
  try {
    const ss = SpreadsheetApp.getActive();
    const meses = parseInt(p.meses || '18', 10);
    const desde = new Date(); desde.setMonth(desde.getMonth() - meses);
    const corte = Utilities.formatDate(desde, Session.getScriptTimeZone(), 'yyyy-MM-dd');

    const orgs = tabla(ss, T.ORG).map(function (r) {
      let msu = {};
      try { msu = JSON.parse(r[8] || '{}'); } catch (err) { msu = {}; }
      return { id: r[0], nombre: r[1], pais: r[2], sector: r[3], asientos: r[4],
        susc: r[5], fin: iso(r[6]), usos: num(r[7]), msu: msu };
    });
    const cols = {};
    orgs.forEach(function (o) { Object.keys(o.msu).forEach(function (k) { cols[k] = 1; }); });

    return json({
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
        let correos = []; try { correos = JSON.parse(r[4] || '[]'); } catch (err) { correos = []; }
        return { org: r[0], pendientes: num(r[1]), caducadas: num(r[2]), fechas: fechas, correos: correos };
      }),
      serie: tabla(ss, T.SERIE).map(function (r) {
        return { f: iso(r[0]), miembros: num(r[1]),
          porTier: { leader: num(r[2]), starter: num(r[3]), lite: num(r[4]), member: num(r[5]) },
          orgsPago: num(r[6]), orgsTotal: num(r[7]), acceso30d: num(r[8]) };
      }),
      dias: tabla(ss, T.LOG).length,
      errores: tabla(ss, T.LOG).filter(function (r) { return r[2] === 'error'; })
        .map(function (r) { return { fecha: iso(r[1]), nota: r[4] }; }),
      nps: leerNPS(),
      rds: leerRDS(ss)
    });
  } catch (err) {
    return json({ error: String(err) });
  }
}

function tabla(ss, nombre) {
  const sh = ss.getSheetByName(nombre);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues()
    .filter(function (r) { return r[0] !== '' && r[0] != null; });
}
function json(o) {
  return ContentService.createTextOutput(JSON.stringify(o))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ============================== NPS semestral (encuesta) ============================== */
function leerNPS() {
  try {
    const sh = SpreadsheetApp.openById(NPS_SHEET_ID).getSheets()[0];
    const vals = sh.getDataRange().getValues();
    let n = 0, pro = 0, pas = 0, det = 0, fecha = '', porOrg = [], ponentes = [];
    for (let i = 1; i < vals.length; i++) {
      const r = vals[i];
      const score = num(r[3]) != null && r[3] !== '' ? num(r[3]) : num(r[12]);
      if (score == null) continue;
      n++;
      if (score >= 9) pro++; else if (score >= 7) pas++; else det++;
      const f = r[0] instanceof Date ? Utilities.formatDate(r[0], Session.getScriptTimeZone(), 'yyyy-MM') : String(r[0]).slice(6, 10) + '-' + ('0' + String(r[0]).split('/')[1]).slice(-2);
      if (f > fecha) fecha = f;
      porOrg.push({ org: String(r[2] || '').trim(), score: score });
      if (/^s[ií]/i.test(String(r[10] || ''))) ponentes.push({ nombre: String(r[1] || '').trim(), org: String(r[2] || '').trim() });
    }
    if (!n) return [];
    return [{ fecha: fecha, n: n, promotores: pro, pasivos: pas, detractores: det,
      nps: Math.round(100 * (pro - det) / n), porOrg: porOrg, ponentes: ponentes }];
  } catch (err) { return []; }
}

/* ============================== ingesta y servicio de datos RDS ============================== */
function doPost(e) {
  const p = (e && e.parameter) || {};
  if (TOKEN && p.token !== TOKEN) return json({ error: 'Token no válido' });
  try {
    const carga = JSON.parse(e.postData.contents || '{}');
    const ds = carga.datasets || {};
    const ss = SpreadsheetApp.getActive();
    const resumen = {};
    Object.keys(ds).forEach(function (k) {
      const filas = ds[k] || [];
      const nombre = 'rds_' + k.replace(/[^a-z0-9_]/gi, '').slice(0, 40);
      let sh = ss.getSheetByName(nombre) || ss.insertSheet(nombre);
      sh.clearContents();
      if (!filas.length) { resumen[k] = 0; return; }
      const cabs = Object.keys(filas[0]);
      const matriz = [cabs].concat(filas.map(function (f) { return cabs.map(function (c) { return f[c]; }); }));
      sh.getRange(1, 1, matriz.length, cabs.length).setValues(matriz);
      resumen[k] = filas.length;
    });
    return json({ ok: true, filas: resumen, recibido: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm') });
  } catch (err) { return json({ error: 'doPost: ' + err }); }
}

function leerRDS(ss) {
  const out = {};
  ss.getSheets().forEach(function (sh) {
    const nombre = sh.getName();
    if (nombre.indexOf('rds_') !== 0) return;
    const vals = sh.getDataRange().getValues();
    if (vals.length < 2) return;
    const cabs = vals[0];
    out[nombre.slice(4)] = vals.slice(1).map(function (r) {
      const o = {}; cabs.forEach(function (c, i) { o[c] = r[i]; }); return o;
    });
  });
  return out;
}

/* ============================== HubSpot: negocios abiertos ============================== */
function hubspotAbiertos() {
  if (!HUBSPOT_PAT || HUBSPOT_PAT.indexOf('PEGA-AQUI') === 0)
    return { error: 'Falta el PAT de HubSpot en el conector (constante HUBSPOT_PAT)' };
  const cs = CacheService.getScriptCache();
  const hit = cs.get('hs_deals');
  if (hit) return JSON.parse(hit);
  try {
    const r = UrlFetchApp.fetch('https://api.hubapi.com/crm/v3/objects/deals/search', {
      method: 'post', contentType: 'application/json', muteHttpExceptions: true,
      headers: { Authorization: 'Bearer ' + HUBSPOT_PAT },
      payload: JSON.stringify({ limit: 50,
        properties: ['dealname', 'amount', 'dealstage', 'notes_last_updated', 'hs_lastmodifieddate'],
        filterGroups: [{ filters: [{ propertyName: 'hs_is_closed', operator: 'EQ', value: 'false' }] }] })
    });
    if (r.getResponseCode() !== 200) return { error: 'HubSpot HTTP ' + r.getResponseCode() };
    const hoy = Date.now();
    const deals = (JSON.parse(r.getContentText()).results || []).map(function (x) {
      const pr = x.properties || {};
      const ult = pr.notes_last_updated || pr.hs_lastmodifieddate;
      return { n: pr.dealname || '—', v: parseFloat(pr.amount) || null, etapa: pr.dealstage || '',
        dias: ult ? Math.round((hoy - new Date(ult).getTime()) / 86400000) : null };
    });
    const out = { deals: deals };
    try { cs.put('hs_deals', JSON.stringify(out), 1800); } catch (err) { }
    return out;
  } catch (err) { return { error: 'HubSpot: ' + err }; }
}

/* ============================== informe semanal por correo ============================== */
function informeSemanal() {
  if (!CORREOS_INFORME) return;
  const ss = SpreadsheetApp.getActive();
  const orgs = tabla(ss, T.ORG), mems = tabla(ss, T.MEM), evs = tabla(ss, T.EV);
  const hoy = new Date(), mes = Utilities.formatDate(hoy, Session.getScriptTimeZone(), 'yyyy-MM');
  const act = {}; evs.forEach(function (r) { if (String(r[0]).slice(0, 7) === mes) act[r[1]] = 1; });
  const lim = new Date(hoy.getTime() + 120 * 86400000);
  const renov = orgs.filter(function (r) { return r[3] && r[6] && new Date(r[6]) <= lim && new Date(r[6]) >= hoy; })
    .sort(function (a, b) { return String(a[6]).localeCompare(String(b[6])); }).slice(0, 5);
  const lineas = [
    'Panel KPIs 101 - lunes ' + Utilities.formatDate(hoy, Session.getScriptTimeZone(), 'dd/MM'),
    '',
    'Miembros: ' + mems.length + ' | Con actividad en ' + mes + ': ' + Object.keys(act).length,
    'Organizaciones de pago: ' + orgs.filter(function (r) { return r[3]; }).length,
    '',
    'Renovaciones proximas (120 dias):',
  ].concat(renov.map(function (r) { return ' - ' + r[1] + ' -> ' + String(r[6]).slice(0, 10); }))
   .concat(['', 'Panel: https://101area.github.io/panel-kpis/']);
  MailApp.sendEmail(CORREOS_INFORME, 'Panel 101 · resumen del lunes', lineas.join('\n'));
}

/* ============================== insights IA (servidor) ============================== */
function insightsIA() {
  if (!OPENAI_KEY || OPENAI_KEY.indexOf('PEGA-AQUI') === 0)
    return { error: 'Falta la clave de OpenAI en el conector (constante OPENAI_KEY)' };
  const props = PropertiesService.getScriptProperties();
  const hoy = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const cacheado = props.getProperty('insights_' + hoy);
  if (cacheado) return JSON.parse(cacheado);

  const ss = SpreadsheetApp.getActive();
  const orgs = tabla(ss, T.ORG), mems = tabla(ss, T.MEM), evs = tabla(ss, T.EV);
  const mesAct = hoy.slice(0, 7);
  const pago = orgs.filter(function (r) { return r[3]; }).length;
  const actMes = {}; evs.forEach(function (r) {
    if (String(r[0]).slice(0, 7) === mesAct) actMes[r[1]] = 1; });
  let hs = '';
  try {
    if (HUBSPOT_PAT && HUBSPOT_PAT.indexOf('PEGA-AQUI') !== 0) {
      const r = UrlFetchApp.fetch('https://api.hubapi.com/crm/v3/objects/deals/search', {
        method: 'post', contentType: 'application/json', muteHttpExceptions: true,
        headers: { Authorization: 'Bearer ' + HUBSPOT_PAT },
        payload: JSON.stringify({ limit: 50, properties: ['dealname', 'amount', 'dealstage'],
          filterGroups: [{ filters: [{ propertyName: 'hs_is_closed', operator: 'EQ', value: 'false' }] }] })
      });
      if (r.getResponseCode() === 200) {
        const dd = JSON.parse(r.getContentText()).results || [];
        const suma = dd.reduce(function (t, x) { return t + (parseFloat((x.properties || {}).amount) || 0); }, 0);
        hs = 'HubSpot: ' + dd.length + ' negocios abiertos por ' + Math.round(suma) + ' EUR. ';
      }
    }
  } catch (err) { /* HubSpot es opcional */ }

  const resumen = 'Comunidad Area101 Innovacion. Miembros: ' + mems.length +
    '. Organizaciones: ' + orgs.length + ' (' + pago + ' de pago). ' +
    'Miembros con actividad en ' + mesAct + ': ' + Object.keys(actMes).length + '. ' + hs +
    'Contexto: comunidad B2B de innovacion corporativa, tres unidades (Communities, Solutions101, Lab101), objetivo anual 1,11M EUR, vendido 806k, pipeline caliente 90k.';
  try {
    const r = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', {
      method: 'post', contentType: 'application/json', muteHttpExceptions: true,
      headers: { Authorization: 'Bearer ' + OPENAI_KEY },
      payload: JSON.stringify({ model: 'gpt-4o-mini', temperature: 0.4, max_tokens: 420,
        messages: [
          { role: 'system', content: 'Eres analista de negocio de una comunidad B2B. Devuelve 4 o 5 lecturas accionables, concretas y no obvias, en castellano, una por linea, sin markdown.' },
          { role: 'user', content: resumen }
        ] })
    });
    if (r.getResponseCode() !== 200) return { error: 'OpenAI HTTP ' + r.getResponseCode() };
    const texto = (JSON.parse(r.getContentText()).choices[0].message.content || '').trim();
    const out = { generado: hoy, texto: texto };
    props.setProperty('insights_' + hoy, JSON.stringify(out));
    return out;
  } catch (err) { return { error: 'OpenAI: ' + err }; }
}

/* ============================== comprobación ============================== */

function probar() {
  const dias = listarDias();
  Logger.log('Días encontrados en Drive: %s', dias.length);
  if (dias.length) {
    Logger.log('Primero: %s · Último: %s', dias[0].fecha, dias[dias.length - 1].fecha);
    const src = SpreadsheetApp.openById(dias[dias.length - 1].id);
    ['Miembros', 'Organizaciones', 'Sesiones', 'Perks', 'Contenidos', 'Aira',
      'Servicios manuales', 'Invitados sin registrar'].forEach(function (h) {
      Logger.log('  %s: %s filas', h, leerHoja(src, h).length);
    });
  }
  const ss = SpreadsheetApp.getActive();
  [T.EV, T.ORG, T.MEM, T.SERIE, T.INV, T.LOG].forEach(function (t) {
    Logger.log('Cache %s: %s filas', t, tabla(ss, t).length);
  });
}
