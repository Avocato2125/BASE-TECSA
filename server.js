require('dotenv').config();
const express    = require('express');
const { google } = require('googleapis');
const cors       = require('cors');
const path       = require('path');
const PDFDocument = require('pdfkit');
const { Readable } = require('stream');

const app = express();
app.use(express.json({ limit: '50mb' }));
app.use(cors());
app.use(express.static(path.join(__dirname, 'public')));

const SHEET_ID          = process.env.GOOGLE_SHEET_ID;
const PASS_BAJA         = process.env.PASSWORD_BAJA;
const FOLDER_RAIZ       = process.env.FOLDER_RAIZ;
const FOLDER_AUDITORIAS = process.env.FOLDER_AUDITORIAS;
const SHEET_ID_PROVEDORES = process.env.SHEET_ID_PROVEDORES;
const SHEET_ID_ANTIDOPING = process.env.SHEET_ID_ANTIDOPING;
const FOLDER_PROVEDORES   = process.env.FOLDER_PROVEDORES;

// Verificación al arrancar: si falta alguna variable critica, detener el
// servidor con un mensaje claro en vez de seguir con valores por defecto
// inseguros (contraseñas o IDs hardcodeados en el código fuente).
const REQUIRED_VARS = {
  GOOGLE_SHEET_ID: SHEET_ID,
  PASSWORD_BAJA: PASS_BAJA,
  FOLDER_RAIZ: FOLDER_RAIZ,
  FOLDER_AUDITORIAS: FOLDER_AUDITORIAS,
  GOOGLE_SERVICE_ACCOUNT_JSON: process.env.GOOGLE_SERVICE_ACCOUNT_JSON,
  OAUTH_CLIENT_ID: process.env.OAUTH_CLIENT_ID,
  OAUTH_CLIENT_SECRET: process.env.OAUTH_CLIENT_SECRET,
  OAUTH_REFRESH_TOKEN: process.env.OAUTH_REFRESH_TOKEN,
  SHEET_ID_PROVEDORES: process.env.SHEET_ID_PROVEDORES,
  FOLDER_PROVEDORES: process.env.FOLDER_PROVEDORES,
};
const faltantes = Object.entries(REQUIRED_VARS).filter(([, v]) => !v).map(([k]) => k);
if (faltantes.length) {
  console.error('Faltan variables de entorno requeridas:', faltantes.join(', '));
  process.exit(1);
}

// ── Autenticación Google ────────────────────────────────────────
// Sheets → Service Account (ya funciona)
// Drive  → OAuth con cuenta personal (resuelve el error de storage quota)
function getSheetsAuth() {
  const credentials = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON);
  return new google.auth.GoogleAuth({
    credentials,
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
  });
}

function getDriveAuth() {
  const oauth2Client = new google.auth.OAuth2(
    process.env.OAUTH_CLIENT_ID,
    process.env.OAUTH_CLIENT_SECRET,
    'https://developers.google.com/oauthplayground'
  );
  oauth2Client.setCredentials({
    refresh_token: process.env.OAUTH_REFRESH_TOKEN,
  });
  return oauth2Client;
}

async function getClients() {
  const sheetsAuth = await getSheetsAuth().getClient();
  const driveAuth  = getDriveAuth();
  const sheets = google.sheets({ version: 'v4', auth: sheetsAuth });
  const drive  = google.drive({ version: 'v3', auth: driveAuth });
  return { sheets, drive };
}

// ── Zona horaria Monterrey ──────────────────────────────────────
function ahoraMty() {
  const ahora = new Date();
  const fecha = ahora.toLocaleDateString('es-MX',  { day:'2-digit', month:'2-digit', year:'numeric',  timeZone:'America/Monterrey' });
  // hourCycle h23 en vez de hour12:false: con hour12 algunas versiones de
  // Node escriben la medianoche como "24:30" en lugar de "00:30".
  const hora  = ahora.toLocaleTimeString('es-MX',  { hour:'2-digit', minute:'2-digit', hourCycle:'h23',  timeZone:'America/Monterrey' });
  return { fecha, hora };
}

// ── Folio único robusto ─────────────────────────────────────────
// Contadores SEPARADOS: Taller busca solo folios con "T", General solo sin "T".
// El periodo (MM/AA) cambia automáticamente cada mes — si el mes actual
// no tiene folios en el Sheet, el contador reinicia desde 001.
// Con reintentos para manejar race conditions entre instancias de Railway.
async function generarFolio(sheets, esTaller) {
  const MAX_INTENTOS = 5;

  for (let intento = 0; intento < MAX_INTENTOS; intento++) {
    if (intento > 0) {
      await new Promise(r => setTimeout(r, 200 + Math.random() * 300));
    }

    const ahora   = new Date();
    const mes     = String(ahora.toLocaleDateString('es-MX', { month:'2-digit', timeZone:'America/Monterrey' })).padStart(2,'0');
    const anio    = ahora.toLocaleDateString('es-MX', { year:'2-digit',  timeZone:'America/Monterrey' });
    const periodo = `${mes}/${anio}`;   // ej: "06/26", "07/26"
    const prefijo = `${periodo}-`;       // ej: "06/26-"

    const res  = await sheets.spreadsheets.values.get({
      spreadsheetId: SHEET_ID,
      range: 'Entradas!A2:A',
    });
    const filas = (res.data.values || []).flat();

    // Filtrar SOLO los folios del periodo actual Y del tipo correcto
    let max = 0;
    filas.forEach(f => {
      if (typeof f !== 'string') return;
      if (!f.startsWith(prefijo)) return;          // diferente periodo → ignorar
      const resto = f.slice(prefijo.length);       // ej: "T013" o "012"
      if (esTaller && !resto.startsWith('T')) return;   // buscamos T, ignorar generales
      if (!esTaller && resto.startsWith('T')) return;   // buscamos generales, ignorar T
      const num = parseInt(resto.replace('T', ''), 10);
      if (!isNaN(num) && num > max) max = num;
    });

    const siguiente = String(max + 1).padStart(3, '0');
    const folio     = `${prefijo}${esTaller ? 'T' : ''}${siguiente}`;

    // Verificar que el folio no exista ya (protección contra race condition)
    if (filas.includes(folio)) {
      console.warn(`Folio ${folio} ya existe, reintentando (intento ${intento + 1})`);
      continue;
    }

    console.log(`Folio generado: ${folio} (intento ${intento + 1})`);
    return folio;
  }

  // Fallback con timestamp si hay demasiados conflictos
  const ahora   = new Date();
  const mes     = String(ahora.toLocaleDateString('es-MX', { month:'2-digit', timeZone:'America/Monterrey' })).padStart(2,'0');
  const anio    = ahora.toLocaleDateString('es-MX', { year:'2-digit',  timeZone:'America/Monterrey' });
  const fallback = `${mes}/${anio}-${esTaller ? 'T' : ''}ERR${Date.now().toString().slice(-4)}`;
  console.error(`Folio fallback usado: ${fallback}`);
  return fallback;
}

// ── Confirmacion de folio unico ────────────────────────────────
// Calcular "el maximo + 1" no basta: si dos personas registran en el mismo
// segundo, las dos leen la misma lista y sacan el mismo numero. Por eso,
// DESPUES de escribir se relee la columna: si el folio aparece mas de una
// vez, se queda con el el registro que quedo mas arriba y el de abajo toma
// el siguiente numero libre. Es determinista: nunca se quedan los dos.
function filaDeAppend(resp) {
  const m = String(resp?.data?.updates?.updatedRange || '').match(/![A-Z]+(\d+)/);
  return m ? parseInt(m[1], 10) : null;
}
async function confirmarFolio(sheets, { spreadsheetId, hoja, fila, folio, generar }) {
  if (!fila) return folio;
  const rango = hoja ? `'${hoja}'!A2:A` : 'A2:A';
  const celda = n => hoja ? `'${hoja}'!A${n}` : `A${n}`;
  for (let intento = 0; intento < 5; intento++) {
    const r = await sheets.spreadsheets.values.get({ spreadsheetId, range: rango });
    const col = (r.data.values || []).map(x => String((x && x[0]) || '').trim());
    const filas = col.map((f, i) => f === folio ? i + 2 : 0).filter(Boolean);
    if (filas.length <= 1 || filas[0] === fila) return folio;
    console.warn(`[Folio] ${folio} duplicado (filas ${filas.join(', ')}); la fila ${fila} toma otro`);
    await new Promise(res => setTimeout(res, 150 + Math.random() * 350));
    folio = await generar();
    await sheets.spreadsheets.values.update({
      spreadsheetId, range: celda(fila), valueInputOption: 'RAW', requestBody: { values: [[folio]] },
    });
  }
  return folio;
}

// Busca la fila de un folio. Si se da un renglon sugerido y ahi esta el
// folio, lo usa; si no (alguien movio filas en el Sheet), lo busca.
async function filaDeFolio(sheets, { spreadsheetId, hoja, folio, sugerida, colEstado, soloActivo = true }) {
  const rango = hoja ? `'${hoja}'!A2:K` : 'A2:K';
  const r = await sheets.spreadsheets.values.get({ spreadsheetId, range: rango });
  const filas = r.data.values || [];
  const ok = i => filas[i] && String(filas[i][0] || '').trim() === String(folio).trim()
               && (!soloActivo || filas[i][colEstado] === 'ACTIVO');
  const iSug = sugerida ? sugerida - 2 : -1;
  if (iSug >= 0 && ok(iSug)) return sugerida;
  for (let i = 0; i < filas.length; i++) if (ok(i)) return i + 2;
  return null;
}

// ── Drive: obtener o crear carpeta ──────────────────────────────
async function getOCrearCarpeta(drive, padreId, nombre) {
  if (!padreId) throw new Error('getOCrearCarpeta: padreId es undefined para carpeta "' + nombre + '"');
  // Escapar apostrofes y diagonales invertidas: un nombre como "O'Brien"
  // rompia la consulta y la carpeta se duplicaba en cada envio.
  const nombreQ = String(nombre).replace(/\\/g, '\\\\').replace(/'/g, "\\'");
  const q = `'${padreId}' in parents and name='${nombreQ}' and mimeType='application/vnd.google-apps.folder' and trashed=false`;
  const res = await drive.files.list({
    q,
    fields: 'files(id)',
    pageSize: 1,
    supportsAllDrives: true,
    includeItemsFromAllDrives: true,
  });
  if (res.data.files.length > 0) {
    console.log(`Carpeta existente "${nombre}": ${res.data.files[0].id}`);
    return res.data.files[0].id;
  }
  const nuevo = await drive.files.create({
    requestBody: { name: nombre, mimeType: 'application/vnd.google-apps.folder', parents: [padreId] },
    fields: 'id',
    supportsAllDrives: true,
  });
  console.log(`Carpeta creada "${nombre}": ${nuevo.data.id} en padre ${padreId}`);
  return nuevo.data.id;
}

// ── Drive: subir archivo ─────────────────────────────────────────
// Con reintentos: Drive rechaza ráfagas de subidas (rate limit 403/429).
// Espera creciente entre intentos para que la foto siempre llegue.
async function subirArchivo(drive, carpetaId, nombre, base64, mimeType) {
  if (!carpetaId) throw new Error('subirArchivo: carpetaId es undefined para archivo "' + nombre + '"');
  const buffer = Buffer.from(base64, 'base64');
  const MAX_INTENTOS = 4;
  let ultimoError;
  for (let intento = 0; intento < MAX_INTENTOS; intento++) {
    if (intento > 0) {
      const espera = 500 * Math.pow(2, intento - 1) + Math.random() * 300; // 500ms, 1s, 2s
      await new Promise(r => setTimeout(r, espera));
      console.log(`[Drive] Reintento ${intento + 1} para "${nombre}"`);
    }
    try {
      const stream = Readable.from(buffer);
      const res = await drive.files.create({
        requestBody: { name: nombre, parents: [carpetaId] },
        media: { mimeType, body: stream },
        fields: 'id,webViewLink',
        supportsAllDrives: true,
      });
      console.log(`[Drive] Subido "${nombre}" → ${res.data.id}`);
      return res.data;
    } catch(e) {
      ultimoError = e;
      console.error(`[Drive] Error subiendo "${nombre}" (intento ${intento + 1}):`, e.message);
    }
  }
  throw ultimoError;
}

// ── Calcular días en base ───────────────────────────────────────
function calcularDias(fechaStr) {
  if (!fechaStr) return 0;
  const p = fechaStr.trim().split('/');
  if (p.length !== 3) return 0;
  const entrada = new Date(Date.UTC(parseInt(p[2]), parseInt(p[1])-1, parseInt(p[0]), 6, 0));
  const diff = Date.now() - entrada;
  return diff < 0 ? 0 : Math.floor(diff / 86400000);
}
function fechaHoraATimestamp(fechaStr, horaStr) {
  if (!fechaStr || !horaStr) return null;
  const p = fechaStr.trim().split('/');
  if (p.length !== 3) return null;
  const [hh, mm] = horaStr.trim().split(':');
  return Date.UTC(parseInt(p[2]), parseInt(p[1])-1, parseInt(p[0]), parseInt(hh)+6, parseInt(mm));
}

// ═══════════════════════════════════════════════════════════════
// ENDPOINTS EXISTENTES (dashboard)
// ═══════════════════════════════════════════════════════════════

app.get('/api/unidades', async (req, res) => {
  try {
    const { sheets } = await getClients();
    const [resE, resT] = await Promise.all([
      sheets.spreadsheets.values.get({ spreadsheetId: SHEET_ID, range: 'Entradas!A2:I' }),
      sheets.spreadsheets.values.get({ spreadsheetId: SHEET_ID, range: 'Taller!A2:I'  }),
    ]);
    const rowsE = resE.data.values || [];
    const rowsT = resT.data.values || [];

    // Cruce por FOLIO, no por unidad: si la misma unidad tiene dos registros
    // activos, la llave por unidad hacia que el segundo pisara al primero y
    // ambas entradas mostraban el mismo folio de taller.
    const tallerPorFolio  = {};
    const tallerPorUnidad = {};   // respaldo para filas antiguas sin folio
    rowsT.forEach((row, i) => {
      if ((row[8] || '') !== 'ACTIVO') return;
      const folio  = (row[0] || '').toString().trim();
      const unidad = (row[3] || '').toString().toUpperCase().trim();
      const datos  = {
        folioTaller: folio, planta: row[5] || '',
        areaServicio: row[6] || '', reporteFalla: row[7] || '',
        rowIndexTaller: i + 2,
      };
      if (folio)  tallerPorFolio[folio] = datos;
      if (unidad && !tallerPorUnidad[unidad]) tallerPorUnidad[unidad] = datos;
    });

    const activos = rowsE.map((row, i) => {
      const unidad = (row[3] || '').toString().toUpperCase().trim();
      const folio  = (row[0] || '').toString().trim();
      // Primero por folio; solo si no hay coincidencia se usa el respaldo,
      // y ese registro se consume para que no lo tome otra entrada.
      let t = tallerPorFolio[folio] || null;
      if (!t && tallerPorUnidad[unidad]) {
        t = tallerPorUnidad[unidad];
        delete tallerPorUnidad[unidad];
      }
      return {
        rowIndex: i + 2, folio: row[0] || '', fecha: row[1] || '',
        hora: row[2] || '', unidad, operador: row[4] || '',
        motivo: row[5] || '', estado: row[6] || '',
        dias: calcularDias(row[1]),
        timestamp: fechaHoraATimestamp(row[1], row[2]),
        folioTaller:    t ? t.folioTaller    : null,
        planta:         t ? t.planta         : null,
        areaServicio:   t ? t.areaServicio   : null,
        reporteFalla:   t ? t.reporteFalla   : null,
        rowIndexTaller: t ? t.rowIndexTaller : null,
      };
    }).filter(r => r.estado === 'ACTIVO');

    res.json({ ok: true, entradas: activos });
  } catch(e) { console.error(e); res.json({ ok: false, error: e.message }); }
});

app.post('/api/baja', async (req, res) => {
  const { rowIndex, rowIndexTaller, password, folio } = req.body;
  if (password !== PASS_BAJA) return res.json({ ok: false, error: 'Contrasena incorrecta' });
  try {
    const { sheets } = await getClients();
    const { fecha, hora } = ahoraMty();

    // El renglon viene de la lista que cargo el navegador. Si alguien movio
    // filas en el Sheet desde entonces, ese renglon ya es OTRA unidad; por
    // eso se confirma con el folio y, si no coincide, se busca.
    let filaE = parseInt(rowIndex, 10) || null;
    if (folio) {
      filaE = await filaDeFolio(sheets, { spreadsheetId: SHEET_ID, hoja: 'Entradas', folio, sugerida: filaE, colEstado: 6 });
      if (!filaE) return res.json({ ok: false, error: 'Ese folio ya no esta activo. Actualiza la lista.' });
    } else if (!filaE) {
      return res.json({ ok: false, error: 'Falta el folio de la unidad' });
    }
    await sheets.spreadsheets.values.update({
      spreadsheetId: SHEET_ID, range: `Entradas!G${filaE}:I${filaE}`,
      valueInputOption: 'RAW', requestBody: { values: [['BAJA', fecha, hora]] }
    });

    let filaT = parseInt(rowIndexTaller, 10) || null;
    if (folio) filaT = await filaDeFolio(sheets, { spreadsheetId: SHEET_ID, hoja: 'Taller', folio, sugerida: filaT, colEstado: 8 });
    if (filaT) {
      await sheets.spreadsheets.values.update({
        spreadsheetId: SHEET_ID, range: `Taller!I${filaT}`,
        valueInputOption: 'RAW', requestBody: { values: [['BAJA']] }
      });
    }
    res.json({ ok: true });
  } catch(e) { console.error(e); res.json({ ok: false, error: e.message }); }
});

// ═══════════════════════════════════════════════════════════════
// FORMULARIO OPERADOR — Registrar entrada
// ═══════════════════════════════════════════════════════════════
app.post('/api/registrar-entrada', async (req, res) => {
  try {
    const datos = req.body;
    const { sheets } = await getClients();
    const { fecha, hora } = ahoraMty();
    const unidad   = String(datos.unidad || '').toUpperCase().trim();
    const operador = String(datos.operador || '').trim();
    const motivo   = datos.motivo;
    const esTaller = motivo === 'Taller';
    if (!unidad || !operador || !motivo) {
      return res.json({ ok: false, error: 'Faltan datos: unidad, operador y motivo son obligatorios' });
    }

    let motivoTexto = motivo;
    if (motivo === 'Camaras')  motivoTexto = `Camaras/Display — ${datos.planta || ''} — ${datos.reporteFalla || ''}`;
    else if (motivo === 'Inplant') motivoTexto = `Inplant — ${datos.planta || ''} — ${datos.detalle || ''}`;
    else if (motivo === 'Otro' && datos.detalle) motivoTexto = `Otro: ${datos.detalle}`;

    let folio = await generarFolio(sheets, esTaller);

    // Guardar en Entradas y confirmar que nadie tomo el mismo folio a la vez
    const ap = await sheets.spreadsheets.values.append({
      spreadsheetId: SHEET_ID, range: 'Entradas!A:I',
      valueInputOption: 'RAW', insertDataOption: 'INSERT_ROWS',
      requestBody: { values: [[folio, fecha, hora, unidad, operador, motivoTexto, 'ACTIVO', '', '']] }
    });
    folio = await confirmarFolio(sheets, {
      spreadsheetId: SHEET_ID, hoja: 'Entradas', fila: filaDeAppend(ap), folio,
      generar: () => generarFolio(sheets, esTaller),
    });

    // Si es Taller, también guardar en hoja Taller (ya con el folio definitivo)
    if (esTaller) {
      await sheets.spreadsheets.values.append({
        spreadsheetId: SHEET_ID, range: 'Taller!A:I',
        valueInputOption: 'RAW', insertDataOption: 'INSERT_ROWS',
        requestBody: { values: [[folio, fecha, hora, unidad, operador, datos.planta || '', datos.areaServicio || '', datos.reporteFalla || '', 'ACTIVO']] }
      });
    }

    res.json({ ok: true, folio });
  } catch(e) { console.error(e); res.json({ ok: false, error: e.message }); }
});

// ═══════════════════════════════════════════════════════════════
// FORMULARIO TALLER — Guardar reporte + fotos + baja
// ═══════════════════════════════════════════════════════════════
// ═══════════════════════════════════════════════════════════════
// REGISTRO DE EVIDENCIAS
// Cada foto subida se anota con su fileId en la hoja "Evidencias".
// Asi al reimprimir un reporte las fotos se descargan directo por ID,
// sin recorrer carpetas de Drive. Los reportes viejos (sin registro)
// caen al metodo de buscar la carpeta.
// ═══════════════════════════════════════════════════════════════
const HOJA_EVIDENCIAS = 'Evidencias';
const HEADERS_EVID    = ['Folio','Tipo','Area','FileId','Nombre','Fecha'];

async function registrarEvidencias(sheets, filas) {
  if (!filas || !filas.length) return;
  try {
    try {
      await sheets.spreadsheets.values.get({ spreadsheetId: SHEET_ID, range: `${HOJA_EVIDENCIAS}!A1` });
    } catch {
      await sheets.spreadsheets.batchUpdate({
        spreadsheetId: SHEET_ID,
        requestBody: { requests: [{ addSheet: { properties: { title: HOJA_EVIDENCIAS } } }] }
      });
      await sheets.spreadsheets.values.update({
        spreadsheetId: SHEET_ID, range: `${HOJA_EVIDENCIAS}!A1`,
        valueInputOption: 'RAW', requestBody: { values: [HEADERS_EVID] }
      });
    }
    await sheets.spreadsheets.values.append({
      spreadsheetId: SHEET_ID, range: `${HOJA_EVIDENCIAS}!A:A`,
      valueInputOption: 'RAW', insertDataOption: 'INSERT_ROWS',
      requestBody: { values: filas }
    });
    console.log(`[Evidencias] ${filas.length} registradas`);
  } catch(e) {
    console.error('[Evidencias] No se pudieron registrar:', e.message);
  }
}

// Devuelve [{area, fileId, nombre}] de un folio segun la hoja Evidencias
async function evidenciasDeFolio(sheets, folio) {
  try {
    const r = await sheets.spreadsheets.values.get({
      spreadsheetId: SHEET_ID, range: `${HOJA_EVIDENCIAS}!A2:F`,
    });
    return (r.data.values || [])
      .filter(row => String(row[0]).trim() === String(folio).trim())
      .map(row => ({ area: row[2]||'', fileId: row[3]||'', nombre: row[4]||'' }))
      .filter(e => e.fileId);
  } catch(e) { return []; }
}

// Fallback para reportes viejos: recorre Unidad-X/*/fecha y lista imagenes
async function buscarFotosEnDrive(drive, carpetaRaizId, unidad, fechaCarp) {
  const out = [];
  try {
    const qU = `'${carpetaRaizId}' in parents and name='Unidad-${unidad}' and mimeType='application/vnd.google-apps.folder' and trashed=false`;
    const ru = await drive.files.list({ q: qU, fields:'files(id)', pageSize:1, supportsAllDrives:true, includeItemsFromAllDrives:true });
    if (!ru.data.files.length) return out;
    const idUnidad = ru.data.files[0].id;

    const ra = await drive.files.list({
      q: `'${idUnidad}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false`,
      fields:'files(id,name)', pageSize:100, supportsAllDrives:true, includeItemsFromAllDrives:true,
    });
    for (const area of ra.data.files) {
      const rf = await drive.files.list({
        q: `'${area.id}' in parents and name='${fechaCarp}' and mimeType='application/vnd.google-apps.folder' and trashed=false`,
        fields:'files(id)', pageSize:1, supportsAllDrives:true, includeItemsFromAllDrives:true,
      });
      if (!rf.data.files.length) continue;
      const rimg = await drive.files.list({
        q: `'${rf.data.files[0].id}' in parents and trashed=false`,
        fields:'files(id,name)', pageSize:100, supportsAllDrives:true, includeItemsFromAllDrives:true,
      });
      rimg.data.files.forEach(f => out.push({ area: area.name, fileId: f.id, nombre: f.name }));
    }
  } catch(e) { console.error('[Drive] buscarFotos:', e.message); }
  return out;
}

// Descarga los bytes de una imagen de Drive
async function descargarImagen(drive, fileId) {
  const r = await drive.files.get(
    { fileId, alt: 'media', supportsAllDrives: true },
    { responseType: 'arraybuffer' }
  );
  return Buffer.from(r.data);
}

const HOJAS_TALLER = {
  mecanico:  'Reporte Taller Mecanico',
  electrico: 'Reporte Electrico',
  imagen:    'Reporte Imagen',
  llantas:   'Reporte Llantas',
  llenado:   'Llenado',
  suspension:'Suspension',
};
const HEADERS_TALLER = {
  mecanico:  ['Folio','Fecha','Hora','Unidad','Operador','Planta','Area Servicio','Mecanico','Aceite-Km','Aceite-CapTeorica','Aceite-LitAnt','Aceite-NivelBajo','Aceite-LitNuevo','Aceite-Obs','Frenos-Obs','Engrasado','Engrasado-Obs','SrvAceite-Km','SrvAceite-Cap','SrvAceite-LitAnt','SrvAceite-NivelBajo','SrvAceite-LitNuevo','SrvAceite-Obs','Filtro-Aire','FiltroAire-Obs','Filtro-Diesel','FiltroDiesel-Obs','Filtro-Aceite','FiltroAceite-Obs','Filtro-Separador','FiltroSep-Obs','Afinacion-Km','Afinacion-Piezas','Afinacion-Mat','Afinacion-Obs','Piezas-Taller','Ajuste','Obs-Taller'],
  electrico: ['Folio','Fecha','Hora','Unidad','Operador','Planta','Area Servicio','Mecanico','Carga-Bat-VoltAnt','Carga-Bat-VoltNuevo','Carga-Bat-Obs','Cambio-Bat-Motivo','Cambio-Bat-Obs','Piezas-Electrico','Obs-Electrico'],
  imagen:    ['Folio','Fecha','Hora','Unidad','Operador','Planta','Area Servicio','Mecanico','Calcas-Mat','Calcas-Obs','Asiento-Mat','Asiento-Obs','Pintura-Area','Pintura-Mat','Pintura-Obs','Soldadura-Mat','Soldadura-Obs','Piezas-Imagen','Obs-Imagen'],
  llantas:   ['Folio','Fecha','Hora','Unidad','Operador','Planta','Area Servicio','Mecanico','Llanta-Marca','Llanta-Obs','LlantaRep-Vida','LlantaRep-Obs','Obs-Llantas'],
  llenado:   ['Folio','Fecha','Hora','Unidad','Operador','Planta','Area Servicio','Mecanico','Aceite-LitAnt','Aceite-LitPuestos','Aceite-LitDespues','Aceite-Obs','Adblue-LitAnt','Adblue-LitPuestos','Adblue-LitDespues','Adblue-Obs','LiqFrenos-LitAnt','LiqFrenos-LitPuestos','LiqFrenos-LitDespues','LiqFrenos-Obs','Anticong-LitAnt','Anticong-LitPuestos','Anticong-LitDespues','Anticong-Obs','Gasolina-LitAnt','Gasolina-LitPuestos','Gasolina-LitDespues','Gasolina-Obs','Direccion-LitAnt','Direccion-LitPuestos','Direccion-LitDespues','Direccion-Obs','Obs-Llenado'],
  suspension:['Folio','Fecha','Hora','Unidad','Operador','Planta','Area Servicio','Mecanico','Muelles','Muelles-Hojas','Muelles-Obs','Amortiguadores','Amortiguadores-Obs','Piezas-Suspension','Obs-Suspension','Muelles-Posicion'],
};

function v(obj, key) { return (obj && obj[key] != null) ? String(obj[key]) : ''; }
function filtroVal(t, key) { const f = t?.servicio?.filtros?.[key]; return f?.cambiado ? 'Si' : 'No'; }
function filtroObs(t, key)  { const f = t?.servicio?.filtros?.[key]; return f?.obs || ''; }
function piezasTexto(arr) {
  return (arr || []).map((p, i) => {
    const partes = [];
    if (p.nombre)   partes.push(`Nombre de la pz: ${p.nombre}`);
    if (p.material) partes.push(`Material us: ${p.material}`);
    if (p.obs)      partes.push(`Obs: ${p.obs}`);
    return partes.join(' | ');
  }).filter(Boolean).join(' // ');
}

app.post('/api/reporte-taller', async (req, res) => {
  try {
    const datos = req.body;
    const { sheets, drive } = await getClients();
    const { fecha, hora } = ahoraMty();

    // Si el folio ya se cerro (doble toque en Enviar, o dos mecanicos con
    // la misma unidad abierta), no se guarda un segundo reporte: duplicaria
    // litros y refacciones en el reporte diario.
    const folioT = datos.folioTaller || datos.folioEntrada;
    if (!folioT) return res.json({ ok: false, error: 'Falta el folio de la unidad' });
    const sigueActivo = await filaDeFolio(sheets, { spreadsheetId: SHEET_ID, hoja: 'Taller', folio: folioT, colEstado: 8 });
    if (!sigueActivo) {
      return res.json({ ok: false, error: `El folio ${folioT} ya fue cerrado; este reporte ya se habia enviado. Actualiza la pagina.` });
    }
    const t  = datos.taller    || {};
    const el = datos.electrico || {};
    const im = datos.imagen    || {};
    const ll = datos.llantas   || {};

    // ── 1. Subir fotos a Drive ────────────────────────────────
    console.log(`[Taller] Unidad: ${datos.unidad} | FOLDER_RAIZ: ${FOLDER_RAIZ}`);
    let carpRaiz;
    try {
      carpRaiz = await getOCrearCarpeta(drive, FOLDER_RAIZ, `Unidad-${datos.unidad}`);
      console.log(`[Taller] Carpeta unidad: ${carpRaiz}`);
    } catch(errCarp) {
      console.error('[Taller] ERROR creando carpeta unidad:', errCarp.message);
      // Continuar sin fotos en vez de fallar todo el reporte
      carpRaiz = null;
    }
    const fechaCarp  = fecha.replace(/\//g, '-');
    const evidencias = []; // [Folio, Tipo, Area, FileId, Nombre, Fecha]

    async function subirAreaFotos(area, items) {
      if (!carpRaiz || !items || !items.length) return;
      try {
        const carpArea  = await getOCrearCarpeta(drive, carpRaiz, area);
        const carpFecha = await getOCrearCarpeta(drive, carpArea, fechaCarp);
        for (const item of items) {
          if (!item?.fotos) continue;
          const lista = Array.isArray(item.fotos)
            ? item.fotos
            : [...(item.fotos.antes||[]).map(f=>({...f,nombre:'antes-'+f.nombre})),
               ...(item.fotos.despues||[]).map(f=>({...f,nombre:'despues-'+f.nombre}))];
          for (const f of lista) {
            if (!f?.base64) continue;
            try {
              const sub = await subirArchivo(drive, carpFecha, f.nombre||'foto.jpg', f.base64, 'image/jpeg');
              evidencias.push([datos.folioEntrada||'', 'Taller', area, sub.id, f.nombre||'foto.jpg', fecha]);
            } catch(eFoto) {
              console.error(`[Taller] Foto perdida en ${area}: ${f.nombre} — ${eFoto.message}`);
            }
          }
        }
      } catch(errArea) {
        console.error(`[Taller] ERROR subiendo fotos de ${area}:`, errArea.message);
      }
    }

    // Subida SECUENCIAL (no en paralelo): Drive rechaza ráfagas de
    // peticiones simultáneas con rate limit y las fotos se perdían.
    if (t.aceite)    await subirAreaFotos('Aceite',    [t.aceite]);
    if (t.frenos)    await subirAreaFotos('Frenos',    [t.frenos]);
    if (t.servicio?.engrasado) await subirAreaFotos('Engrasado', [t.servicio.engrasado]);
    if (t.servicio?.aceite)    await subirAreaFotos('Aceite-Servicio', [t.servicio.aceite]);
    // Filtros (faltaban en la migración a Node)
    if (t.servicio?.filtros) {
      const filtros = t.servicio.filtros;
      for (const k of Object.keys(filtros)) {
        if (filtros[k]?.cambiado && filtros[k]?.fotos) {
          await subirAreaFotos('Filtro-' + k, [filtros[k]]);
        }
      }
    }
    if (t.afinacion) await subirAreaFotos('Afinacion', [t.afinacion]);
    if (t.ajuste)    await subirAreaFotos('Ajuste',    [t.ajuste]);
    // Piezas de taller (faltaban en la migración a Node)
    if (t.piezas?.length) {
      for (let i = 0; i < t.piezas.length; i++) {
        await subirAreaFotos(`Pieza-Taller-${i+1}-${(t.piezas[i].nombre||'').replace(/[\/\\:*?"<>|]/g,'-')}`, [t.piezas[i]]);
      }
    }
    if (el.cargaBat)  await subirAreaFotos('Carga-Bateria',  [el.cargaBat]);
    if (el.cambioBat) await subirAreaFotos('Cambio-Bateria', [el.cambioBat]);
    if (el.piezas?.length) {
      for (let i = 0; i < el.piezas.length; i++) {
        await subirAreaFotos(`Pieza-Electrico-${i+1}-${(el.piezas[i].nombre||'').replace(/[\/\\:*?"<>|]/g,'-')}`, [el.piezas[i]]);
      }
    }
    if (im.calcas)    await subirAreaFotos('Calcas',    [im.calcas]);
    if (im.asiento)   await subirAreaFotos('Asiento',   [im.asiento]);
    if (im.pintura)   await subirAreaFotos('Pintura',   [im.pintura]);
    if (im.soldadura) await subirAreaFotos('Soldadura', [im.soldadura]);
    if (im.piezas?.length) {
      for (let i = 0; i < im.piezas.length; i++) {
        await subirAreaFotos(`Pieza-Imagen-${i+1}-${(im.piezas[i].nombre||'').replace(/[\/\\:*?"<>|]/g,'-')}`, [im.piezas[i]]);
      }
    }
    if (ll.cambio)     await subirAreaFotos('Llanta-Cambio',     [ll.cambio]);
    if (ll.reparacion) await subirAreaFotos('Llanta-Reparacion', [ll.reparacion]);
    // Rellenados
    const ln = datos.llenado || {};
    if (ln.aceite)         await subirAreaFotos('Rellenado-Aceite',         [ln.aceite]);
    if (ln.adblue)         await subirAreaFotos('Rellenado-Adblue',         [ln.adblue]);
    if (ln.liqFrenos)      await subirAreaFotos('Rellenado-LiqFrenos',      [ln.liqFrenos]);
    if (ln.anticongelante) await subirAreaFotos('Rellenado-Anticongelante', [ln.anticongelante]);
    if (ln.gasolina)       await subirAreaFotos('Rellenado-Gasolina',       [ln.gasolina]);
    if (ln.direccion)      await subirAreaFotos('Rellenado-Direccion',      [ln.direccion]);
    // Suspension
    const sp = datos.suspension || {};
    if (sp.muelles)        await subirAreaFotos('Muelles',        [sp.muelles]);
    if (sp.amortiguadores) await subirAreaFotos('Amortiguadores', [sp.amortiguadores]);
    if (sp.piezas?.length) {
      for (let i = 0; i < sp.piezas.length; i++) {
        await subirAreaFotos(`Pieza-Suspension-${i+1}-${(sp.piezas[i].nombre||'').replace(/[\/\\:*?"<>|]/g,'-')}`, [sp.piezas[i]]);
      }
    }

    // ── 2. Guardar en hojas separadas por área ────────────────
    const comun = [datos.folioEntrada, fecha, hora, datos.unidad, datos.operador, datos.planta, datos.areaServicio, datos.mecanico];

    const rowMec = [...comun,
      v(t.aceite,'km'), v(t.aceite,'capTeórica'), v(t.aceite,'litrosAnt'), v(t.aceite,'nivelBajo'), v(t.aceite,'litrosNuevo'), v(t.aceite,'obs'),
      v(t.frenos,'obs'),
      t.servicio?.engrasado ? 'Si' : 'No', v(t.servicio?.engrasado,'obs'),
      v(t.servicio?.aceite,'km'), v(t.servicio?.aceite,'capTeorica'), v(t.servicio?.aceite,'litrosAnt'), v(t.servicio?.aceite,'nivelBajo'), v(t.servicio?.aceite,'litrosNuevo'), v(t.servicio?.aceite,'obs'),
      filtroVal(t,'fil-aire'), filtroObs(t,'fil-aire'), filtroVal(t,'fil-diesel'), filtroObs(t,'fil-diesel'),
      filtroVal(t,'fil-aceite'), filtroObs(t,'fil-aceite'), filtroVal(t,'fil-separador'), filtroObs(t,'fil-separador'),
      v(t.afinacion,'km'), v(t.afinacion,'piezas'), v(t.afinacion,'materiales'), v(t.afinacion,'obs'),
      piezasTexto(t.piezas),
      t.ajuste ? ('Si — ' + (t.ajuste.pieza||'') + (t.ajuste.obs ? ' [' + t.ajuste.obs + ']' : '')) : 'No',
      t.obsGral||'',
    ];
    const rowElec = [...comun, v(el.cargaBat,'voltAnt'), v(el.cargaBat,'voltNuevo'), v(el.cargaBat,'obs'), v(el.cambioBat,'motivo'), v(el.cambioBat,'obs'), piezasTexto(el.piezas), el.obsGral||''];
    const rowImg  = [...comun, v(im.calcas,'material'), v(im.calcas,'obs'), v(im.asiento,'material'), v(im.asiento,'obs'), v(im.pintura,'area'), v(im.pintura,'material'), v(im.pintura,'obs'), v(im.soldadura,'material'), v(im.soldadura,'obs'), piezasTexto(im.piezas), im.obsGral||''];
    const rowLl   = [...comun, v(ll.cambio,'marca'), v(ll.cambio,'obs'), v(ll.reparacion,'vidaRestante'), v(ll.reparacion,'obs'), ll.obsGral||''];
    const rowLlenado = [...comun,
      v(ln.aceite,'litAnt'),         v(ln.aceite,'litPuestos'),         v(ln.aceite,'litDespues'),         v(ln.aceite,'obs'),
      v(ln.adblue,'litAnt'),         v(ln.adblue,'litPuestos'),         v(ln.adblue,'litDespues'),         v(ln.adblue,'obs'),
      v(ln.liqFrenos,'litAnt'),      v(ln.liqFrenos,'litPuestos'),      v(ln.liqFrenos,'litDespues'),      v(ln.liqFrenos,'obs'),
      v(ln.anticongelante,'litAnt'), v(ln.anticongelante,'litPuestos'), v(ln.anticongelante,'litDespues'), v(ln.anticongelante,'obs'),
      v(ln.gasolina,'litAnt'),       v(ln.gasolina,'litPuestos'),       v(ln.gasolina,'litDespues'),       v(ln.gasolina,'obs'),
      v(ln.direccion,'litAnt'),      v(ln.direccion,'litPuestos'),      v(ln.direccion,'litDespues'),      v(ln.direccion,'obs'),
      ln.obsGral||''
    ];
    const rowSusp = [...comun,
      sp.muelles ? 'Si' : 'No',
      sp.muelles ? ((sp.muelles.hojas||[]).join(', ') + (sp.muelles.completo ? ' (Muelle completo)' : '')) : '',
      v(sp.muelles,'obs'),
      sp.amortiguadores ? 'Si' : 'No',
      v(sp.amortiguadores,'obs'),
      piezasTexto(sp.piezas),
      sp.obsGral||'',
      v(sp.muelles,'posicion'),
    ];

    // "No" es el valor por defecto de Engrasado, filtros, Ajuste, Muelles y
    // Amortiguadores. Si solo hay "No", la seccion no se trabajo y no se
    // escribe fila (antes cada reporte dejaba filas vacias en Mecanico y
    // Suspension aunque solo se hubiera trabajado otra area).
    const tieneContenido = row => row.slice(8).some(c => c != null && String(c).trim() !== '' && String(c).trim() !== 'No');

    async function appendHoja(nombre, headers, row) {
      if (!tieneContenido(row)) return;
      // Verificar si la hoja existe; si no, crearla con headers
      try {
        await sheets.spreadsheets.values.get({ spreadsheetId: SHEET_ID, range: `${nombre}!A1` });
      } catch {
        await sheets.spreadsheets.batchUpdate({
          spreadsheetId: SHEET_ID,
          requestBody: { requests: [{ addSheet: { properties: { title: nombre } } }] }
        });
        await sheets.spreadsheets.values.update({
          spreadsheetId: SHEET_ID, range: `${nombre}!A1`,
          valueInputOption: 'RAW', requestBody: { values: [headers] }
        });
      }
      await sheets.spreadsheets.values.append({
        spreadsheetId: SHEET_ID, range: `${nombre}!A:A`,
        valueInputOption: 'RAW', insertDataOption: 'INSERT_ROWS',
        requestBody: { values: [row] }
      });
    }

    await Promise.all([
      appendHoja(HOJAS_TALLER.mecanico,  HEADERS_TALLER.mecanico,  rowMec),
      appendHoja(HOJAS_TALLER.electrico, HEADERS_TALLER.electrico, rowElec),
      appendHoja(HOJAS_TALLER.imagen,    HEADERS_TALLER.imagen,    rowImg),
      appendHoja(HOJAS_TALLER.llantas,   HEADERS_TALLER.llantas,   rowLl),
      appendHoja(HOJAS_TALLER.llenado,   HEADERS_TALLER.llenado,   rowLlenado),
      appendHoja(HOJAS_TALLER.suspension,HEADERS_TALLER.suspension,rowSusp),
    ]);

    // ── 3. Dar de baja en Taller y Entradas ───────────────────
    const { fecha: fBaja, hora: hBaja } = ahoraMty();
    const bajas = [];
    if (datos.folioTaller) {
      bajas.push(sheets.spreadsheets.values.get({ spreadsheetId: SHEET_ID, range: 'Taller!A2:I' })
        .then(r => {
          const rows = r.data.values || [];
          for (let i = 0; i < rows.length; i++) {
            if (String(rows[i][0]).trim() === String(datos.folioTaller).trim() && rows[i][8] === 'ACTIVO') {
              return sheets.spreadsheets.values.update({ spreadsheetId: SHEET_ID, range: `Taller!I${i+2}`, valueInputOption:'RAW', requestBody:{values:[['BAJA']]} });
            }
          }
        }));
    }
    if (datos.folioEntrada) {
      bajas.push(sheets.spreadsheets.values.get({ spreadsheetId: SHEET_ID, range: 'Entradas!A2:I' })
        .then(r => {
          const rows = r.data.values || [];
          for (let i = 0; i < rows.length; i++) {
            if (String(rows[i][0]).trim() === String(datos.folioEntrada).trim() && rows[i][6] === 'ACTIVO') {
              return sheets.spreadsheets.values.update({ spreadsheetId: SHEET_ID, range: `Entradas!G${i+2}:I${i+2}`, valueInputOption:'RAW', requestBody:{values:[['BAJA', fBaja, hBaja]]} });
            }
          }
        }));
    }
    await Promise.all(bajas);
    await registrarEvidencias(sheets, evidencias);

    res.json({ ok: true });
  } catch(e) { console.error(e); res.json({ ok: false, error: e.message }); }
});

// ═══════════════════════════════════════════════════════════════
// FORMULARIO AUDITORÍA — PDF + fotos + Sheets + baja
// ═══════════════════════════════════════════════════════════════
const SECCIONES_AUD = [
  { nombre: 'Documentos Legales', puntos: ['Licencia Vigente','Tarjeton Vigente','Poliza'] },
  { nombre: 'Seguridad',          puntos: ['Cinturones de Seguridad','Extintor','Botiquin','Camaras','Parabrisas en Buen Estado'] },
  { nombre: 'Limpieza',           puntos: ['Cortinas','Piso','Asientos','Carroceria'] },
  { nombre: 'Personal Operativo', puntos: ['Uniforme Completo','Barba','Zapatos','Sin Gorra','Sin Aretes o Piercing'] },
  { nombre: 'Estado de la Unidad',puntos: ['Luces Largas','Luces Cortas','Faros','Intermitentes','Luces Interiores','Asientos en Buen Estado','Retrovisores Laterales y Central','Neumaticos en Buen Estado','Sin Golpes en Defensa y Laterales','Carroceria sin Dano en la Pintura'] },
];

// Logo TECSA en base64 (sin prefijo data:image/png;base64,)
const LOGO_BASE64 = 'iVBORw0KGgoAAAANSUhEUgAAAh4AAAEGCAYAAAAqrpwhAAAQAElEQVR4AeydW6xtV1nH59yXc3LO7qm1SdNWY62mFBAkQIgWGlADFbkXkJdaQwjYBx8sJl4TH5qAkQcShAdCotHUvqFGDBLusaCNQJCQ+CJIQpUHDzWxlbJPe86+LNc3zx7rjDX3nHPNy7h8Y47fyf72vI3xXX7fXHP+91xr77NV8A8CEIAABCAAAQgEIoDwCASaMBCAAAQgAIHTBPLbg/DIr+dUDAEIQAACEIhGAOERDT2BIQABCECgToDt+RNAeMy/x1QIAQhAAAIQUEMA4aGmFSQCAQhAoE6AbQjMjwDCY349pSIIQAACEICAWgIID7WtITEIQKBOgG0IQCB9AgiP9HtIBRCAAAQgAIFkCCA8kmkViUKgToBtCEAAAukRQHik1zMyhgAEIAABCCRLAOGRbOtIvE6AbQhAAAIQ0E8A4aG/R2QIAQhAAAIQmA0BhMdsWlkvhG0IQAACEICAPgIID309ISMIQAACEIDAbAlkIzxm20EKgwAEIAABCCREAOGRULNIFQIQgAAEIJAogVXaCI8VClYgAAEIQAACEPBNAOHhmzD+IQABCEAAAnUCGW8jPDJuPqVDAAIQgAAEQhNAeIQmTjwIQAACEKgTYDsjAgiPjJpNqRCAAAQgAIHYBBAesTtAfAhAAAJ1AmxDYMYEEB4zbi6lQQACEIAABLQRQHho6wj5QAACdQJsQwACMyKA8JhRMykFAhCAAAQgoJ0AwkN7h8gPAnUCbEMAAhBImADCI+HmkToEIAABCEAgNQIIj9Q6Rr51AmxDAAIQgEBCBBAeCTWLVCEAAQhAAAKpE0B4pN7Bev5sQwACEIAABBQTQHgobg6pQQACEIAABOZGYO7CY279mlzPL73n/Yvr3/kJrIPBZMg4gAAEIACBVgLqhUdZlotl9qdM9tu2tbW12N7ermxnZ2chtru7u6ib7DfjZI6Y7acpVoh9dg6Sn+Q5MO5y+Oavfz144eZBjIAABCAAAQg4IXDaiXrhsViI5jiduOy37fj4uDg6Olqzg4OD4vDw8JSZcTJHzPZzOlKYPXYOJr+BkZtB1Zx84P4X1fawaRP4wcP3lvY26xCAAAQg4JaAeuExplxzE5e5Zt0sZV8KJvn6yPM37/lpbqwtYBEdLWDYDQEIOCOAo6LQLjx6/RRPI1cEevHiBrvitVqByQoFKxCAAAS8EtAuPLwWr915WZZjUuwlPsY4nuscRMdcO0td+gmQYY4EEB6Kuz7h7ZaN4oOb7dXGw+EqB75DAAIQCEUA4RGK9Mg4ZTnqqUevaNpuupeOj1+pLadeIBkEAUcEcAOBHAhoFh4bf2rPoUFS49bWVlGWgwVIL36xb/QS39jhI2/7Z/kbI1KzbzMxfcfBPwQgAAEIrBPYWt9kSyuB7e3t4uzZs8Xe3l6xs7NTuPwnN2GX/mxf+4ujd4j/NjNjRXCImW2WECgKGEAAAnMksDXHouZe04ULF/qW2OuphzgTYSBLFya+jB391dv/psuniA2xrjGuj0lurn3iDwIQgAAE+hFAePTjFHWU/JEz+UNoly9fLvb394snn3xySD69xccQp/c8/0w1XG7idasO1L7tvvvzLxGBUbfaMK+bJk+vQTw6xzUEIACBORDQKjy83CxTbNiE32wZXK7cmLsmyXFjf/sHry9lvWv89rs++0YjNM4d7n+jayzHIAABCEAgDwJahUce9MNV2VvIiZhosz7p7rzrc79gxMbe8TOf7DNn3Jhhs6SmYTMYDQEIQAACPgggPHxQ1emzt/gYm74IjvPHlx4dO9/XPESHL7L4hQAEIDCcAMJjODOVM2ImJYJDLGYObbERHW1k2A8BCEAgDgGNwsP7T+ZxUKcb1X77RARG3bRWhujQ2hnyggAEciagUXg46AcuXBHY+bW/e7XGt0821Yfo2ESI4xCAAATiEEB4xOEeI+qgP31qnmqc3ym/ECPZKTERHVPoMRcCEICAAwIdLhAeHXBmckgEh9jGcozYkOXGwQyAAAQgAAEIjCCA8BgBLZEpIjbEOtMVkWGsc2AiB3nakUijSBMCeRGgWouANuHBB0ut5kxY3Sg4xLcIDlnOxRAdc+kkdUAAAnMmoE14zJl1iNpEcIi1xhKxYax1UIIHEB0JNo2U8yVA5VkTQHik334RGsbaqlnc9twXL0RwtA1IeT+iI+XukTsEIJAbAYRH2h0XwdFagQiNEyueuuuh1nEpH0B0pNw9cj8hwAICWRFAeMyw3Sdig8/LzLC3lAQBCEAgdQKahAc3ymFn06mnHbkJDp52DDthkhlNohCAwKwJaBIeswbtsDgRHGIrl7kJDikc0SEUMAhAAALpEUB4pNEzERrGVhnnKDik+MxEh5SMQQACEJgNAYSH/laK4FjL8nW/85fRf0NFbv62rSXocUNienSPawhAAAIQ8ExAi/Dg8x2nGy2CQ2ztiDzleOx/fnRtX+iNWDf/Km7oYokHAQhAAAJOCWgRHk6L8u2sLMuiLEvfYU75F9FxamfAHXLjF2sKKfvFmo652OfTt4v88AEBCEAAAv0IIDxqnMqyrERFWW5e1qZ62/yjRz7f9taKt5i2Y7npi9n7Qq7HjB2yTmJBAAIQyIGAFuFR+oRdlmWjmNja2ips297eXtuWY7LPtp2dncKYx5xL4/sN7/vE4iNf2DebQZdywxcbElTGiw2Zw1gIQAACEMiHwJaiUuVmO82KonH+YrEom+z4+Li07ejoqKzb4eFhadvBwUF55cqVyoqiOV5RrO0vBv6TGgp5W0Xsn74zcLaj4VPFw9T5pgxXfow/lhCAAAQgEJeAJuERl4SO6CvRoSOdaVlMFQ1T50/LntkQgAAEIDCUQJ/xCI8+lMaP6fvbOiI4xKonHePD6Zsp4kFsaGZj5gyNwXgIQAACEAhPAOERnnk9YiU46jtz3kZ05Nx9aofAnAhQSxMBhEcTlUj75DMdkUKvhfVx4+/rs++4tYTZgAAEIACBZAggPOK2avW043UP/UPft2XiZjwhOqJiAjymQmAGBCgBAkIA4SEUIti58+dXUeVJx2PfPVxtz3mlS3x0HZszE2qDAAQgkBMBhIe/bnc+wXjm0qXqaYeIDn8p6PTcJDCa9unMnqwg4IoAfiCQJwGER5y+V6IjTujuqKEEgMSxrTsrjkIAAhCAwFwIIDzCd3IlOnJ82hEeNxFTIUCeEIBAHgQQHn763Pk2i5+QeIUABCAAAQjoJ4DwCNsjnnaE5Z1wNFKHAAQgME8CCI8IfeUtlgjQCQkBCEAAAioIIDzct4G3WdwzLXAJAQhAAALzIIDwCNfH6m0WzU875LdMwuEgEgQgAAEI5EgA4ZFk10kaAhCAAAQgkCYBhEeYvlVPO8KEIgoEvBCQtxB9mJdkE3DqimUCpZIiBNYJzEJ4rJekd0vz2ywOqbm6oEbxc/PNNy/uvPPOxe23376Q9RtvvHFx/fXXL/b29hbnzp1bnD17dnHmzJnF7u7uYmdnp7Lt7e1qabbluEOesV2ZPvjKw/iXpa8YWvxKjcZc5WT8LcqyrMyV48T9rLgs6wi9vgzJVxcBhEcXnYyO8fmOq83+/ve/Xzz++OPFE088UTz11FPF008/XVy6dKl49tlni8uXLxcHBweVHR4eFkdHR5UdHx9XS3v7qrckv9cv0iGLMLFDxvQRy9RRX/qItfK5WCwKseUOBMgSAl9BCQwKhvAYhKt78HXX7T3RMIK3WRqgaN515cqVSmSIkBBRISYX9D4mdcl4WSZm5iapIW1NufThYfI1yz5zvI6Rc3UZQPJZLrL7yrXuZBqN8HDYqh/+cP+mNncZvc3ShiCp/SI6xOQCXreuQszYrjHKjslFWkxZWlU6WvOqkjv5pj1Hya96AiJvCZ7kPOeF1Bu/PjLoJIDw6MSTx0HeZmnus4gIeXohS7HmUUnulYuzMe0FaM5TctPOr8pPzl85l0V8yOeUqp18g0AkAggPd+CTuQi5KxlPTQSWF3nNb6+lep5K3mJNyEPvkzzEQsedFG95XhYiPuRzSoE/AD0pbybPjwDCw0FPb7nlltd1ucnkbZYuBByLR8DcIGUpFi8TN5Fj1iCxxdxUEsGLiA8xeRtRnnzccMMNSddTQzinWmqlzWsT4eGgnxcvXvx0ixvNP/m2pDx6Ny/60ei8T5xbb9Kux3u7NweQJx/ym1nPPPOMDBaeYrKeqqWef6rcR+WN8BiFbT6T+HyH815qEZtyIRZzXqAShyFrk1hiSkqfloY88SjLq6epCBDL22xqtGpiVSEBhIe/plx9Zfvzj2cItBFI6QbSVkOf/bnU2YfFoDEiPsTk8x61iSky1ZaztnxqLY6/ifBw0wNONDcc8TKNgJyHYtO8MNsmMFuetacdqdY82/7YDZnbOsLDc0cz+WApL35X59F4P7n2wGfdPn2P77SDmfK0Y4Ob2da+oW4OByCA8AgAmRAQ8EyAm4R7wDB1zxSPEKgIIDwqDHxTSiC1tGJ8rif3G2QM5qmdl3PNN/dzP9m+Ijymty7Zk5/faJnefDxEJeBLdCT7mnbYDV9sXaVIj1yRjOAH4eERuvPPd3jMFdfJEZALr1hyiTtIWG6KYg5c4aKBAGwboLDLHQGEhzuWtqecXri53vzsfode18RczvUuc81GYrn2GdOf1NPHvOV45swZ27fkYm9rXHd9/kvNYhprnWVOKQsPDQ1x/QLQUJO2HD6yTOjC0uTCoN2WaXr/0nLOmV5sKrjvuFB+NsUJdXwoFzO+vhyVb1mWxdbWVrG7u1ucP3++OPknvk9W1S60nP+bAKWS56Y6vBxHeHjBilNHBF6y9PPg0n64NL50EJCbk9jQbGSO2JB5ZrxZDpmreazLesSX2OB6y7Is5GnHhQui64tRPgYH1Tch17oddWKcm61x05jVQSCJEzmBD5bKf7z3zQ7OOR6K+VOUnNdiU7mLD7FNfswYs9w03tXxmIyn1DCIkzzt2NnZKfb398vvfe97g+ZOSVLZ3Fzrjt4GhIenFvDB0klgv7Sc/Zml8XWNQMwboo8LdJfPrmPXiLBWJyDcxOr717bLsqzeZrl8+fLGscV8/3mpfb643FaG8HDLMzdvvm6G780N5IZ6fXHeELY67PMCXfct22JV4Jl+C1FfZ4zt7e3i4OCgc4xC9vIaEFOYGikNJYDwGEqse3z1Ytb+tCOBt1l4i6X7PAt1tDqfPQeTGMY8h8rKfRvT8vDwUI4FgEEICDQTQHg0c2EvBLQQiPFTntyYxLQwII/xBKSPYuLBLGU9JYvxGkiJT3K5IjzGtyz3F0Pu9Y8/c3TPTPXmpJtqURQnCcZ63dDXkwYsF7BYQoj5hfDwQP8Vt8W6tngoJrxL+WBp+Kg6I4Y+kbgghzkPQvc1TFVEgUBPAgiPnqBqw45q22ubh8/yZyfWgAzbeGjY8NmODn1zCiw6Ztu3voWF7m/fvLSNg5O2jjjIB+ExDmInt689Uf1BnnGePc9K4IOlU6FoZAAAEABJREFUj3pGgPvTBBAd15iEZCE3VbFr0VmzCaTOJvX87V44Xe+8gTqNNH9nIS9Y86eZWYWUmzUBblBh2i/XaLEw0YjSSgDh0Ypm3AHNv0rr8GkHF8pxp0ffWaH4ykVYrG9ejPNHIFTP/VXg1jM83PJU5Q3hMbwdvCCGMxs4g+EQyJKAXFvEsizeKhoGFow5riI85thVakqZABfd+N2L/RSIcyD+OUAGHgkgPNzA9XqhcpHi8+9+kws3+JgPAfXn7HxQj6okV/GRa92jTpJUJyE8HHZO8+c7vvrAtqsbje8Lg/gXk878onzDIJApAXkd2JYpBidlu7r+OUkmdycJCg89LXvu8+7Uk8z8MpEL7j8uy5JlaFuGjfIldYYIzEU4BGX3MeT8EHPvWYdHX7Vxvuvo7yoLhMcKRa+VtT8c9q1//7ZM4qQWClgqBDhf+3VKMye5QYv1qySNUXOrJw3qY7OcOA/hMQxgKy/Nb7MMK5HREIDACQHN4kNS5GYtFLDkCLTeSJOrhIRDEOBCF4IyMSDQn4C8JsX6z9A30mf+IcSjPqLKM0J4KG8Q6UHAIQEuwsNhpsJMbt5iwyuc7wwNvaMnDecXwqMBSsuutc93tIxRuftXX7DzRZWJkZRNgAuUTUPXutzAjOnK7HQ28c+j0zl17Ukt365aONaTAMKjJ6jlsCZWcjFaHtL99Re/98bX6M6Q7CAAARcEyrIsdnd3uZkXRRLX5iLTf00300xRjC+bD5aOZ8dMCMycQNDyyrIsFotFsbe3t7j11lu1CxBt+SFWAp2tCI9AoGOF4T+Gi0VeXVwuqm5aoprj8fFxcXh4WOzv7xcXL150U7EfLz5Fh+oe+cGZlleER1r9GpRtWZZfGjSBwRDwTWAe/uXGJqa6GnnysUzQ5w1+6X70ly9+vvyOLpSJpwkgPE4zGbzn7tt1fu705bdvPTS4GCZw4eIc6EsglXNFm/iQfMT6cmbczAhszaweX+U0vUhWF53HHt/2FXeS38889KZHJzlgsm8C+E+fwOo6oLyUpmtYjJS15BGy9hxr7uSL8OjEw8ETArxwTkCwgEADAREfYg2HVO3idayqHfkmg/CY2Ht+o2UiQHs66xBIm4CIDzHNVcQUH75ja2ev+bwImhvCIyjucMHuuvlSuGDzicSFaz69jFmJnEdiMXMgNgTUEkB4qG1NMSmzG39kb9L8wJPlIq3BApdNuJkTMOf0zMvsVZ7Ppx1w7tUCPYMQHpt70fSCkRN988yII2647ryr6E31u/KNHwjkQECuF2Jaag39mg4dTwtn8mghkI7waCkgwu7VBUTz5zs+9uA9qzwjMBoSMpU8h9TEWAg0EeBcb6LCvuwIIDyyazkFQwACEQmI+BCLmEIVWp5CiFUbHr+FiOExfVzbBFytIzxckcQPBCAAgf4ERHwY6z+LkRCYAQGExwyamHAJcuFNOH1Sh4ATAvI6EHPiTJGTEE87InIbRDoEi0EJxRyM8BhJ/5f/+Ms5nEg51DjyDEhuGr3U37JUbqJ9SHK+9aGU6RiER3fjW188X/n2/3bPjHT0JTcfFQ7/R9pIVRAWAtkSSFZ8ROgYrCJAdxES4eGCoiIf58/wWlTUDlKBwBgC8iIWGzNXw5zWH9g0JEcO8QkgPIb1QP3F4NPvf5v6HIchZ7RDAtwQ2mEKm7q1j954xMkAXsvtGGHTzkb9EYSH+hZFS1AuwtGCE9gbAfp6Gm0bk7b9pz2wxxCAmSHBspUAwqMVDQc8E+AnlnXAs+axXqqqrU03yk3HVRWTSTK8VhJvNMKjvYGtFxytf7H07h9/sr0ajkDgGoHWc/vakCzW+nLoOy41aNzAw3ZsrufRYIoIj8HI9E64bod26u2OnZmK9dAXQYlnTAMAyWVIHjJebMic3MaG4INYmsFZxZ2quYkhXkDNkSfsPX/uugmz16YmWf9aBWz0IRCizxJDzM6nvm0f870uscXGxpkyd2jMkLGG5lYfn1Ku9dzZDkwA4dEfuHqlvXd+5739y9EzkkxWBNSfY6tMN6/IjUisbWTXsbY5U/e7iil+xKbm0zZffIu1Hc91/5xeH7n2sKob4VFh6P9N6+c7pIKPPviWD8sSg8AAAq5vcOJPbEAKQYb6yMm1T/EnFgSIwyAp5uywfFwNJYDwOE0s8ovodEJ99rzyxy79dp9xjIFAAwE55401HPa2S2J6c37iWGKInWw6X4hvY1Oci48p8+c+l6cdM+owwmNAM8/e//d3DBgedOin/uS+P3UUMNQFUOJoN0dIB7mJfYGt92RT8kPH1/3J/Pq+VLelFmN9axg6vq/fUOMk/1CxiDMTAuqFhxLO1c1ga7F4WEk+pAGBUATkxtJlLvIQ/y781H348luP07QtsftY09xU9kl9IXKtrr8hAhEjDAGER1/O7/j4Tbs7xSv6DmccBCYQkAut2AQXyU0NdRNLDoynhHM7vzxhHOw2xfN8cJGbJiA8NhE6OX7+7Jk/P1llAQEIQCBlAi5ER6gbqItcU+7VLHNHeKy3tfXFtLNdvHl9KFsQ8E4gt4tu6+tvJGnhJzZy+iynwWNqW5k/mQDCYzPCnF6ori/8m+kyYhOBnM6/TSw4roNAqOsE576OfjvPAuHhHCkOIQCBCQR83NS4gV1tiAsO9f5c9cx3CAwggPAYAEvr0B88fK+LC4rW8mLlpYmpplxi9WNq3NwZuqg/pOhwke/Uc4b5ngggPK6BDfmiuhaVNQj0I8CFuB+nrlHCUKxrzLRjOme7qtmVH52UyCoYAYRHMNTqAyG81Leo4MLvpkdwHMcx1DWC/ozrTzKzEB49WqX5/2fhbZYeDRw+RPOFT3ITG15V+BlDI0pdYkPnjRkfKs6Y3FzOcVUnosNlVzL3hfC4egK0vahcvWivRuE7BNwRkHNTzJ3HOJ6kBmOhM4gVN1SdUt/UWG3Xxql+c56fPVOER86nP7WHJUA0m4DcFMXsfbHWteQxqv6ff9ld9jypxZi9f+y6+Mr+RjkWHvOaCSA8mrnIXnnBFZrfZpEkHRkXFkcgI7ipztNlXLNcrqr/0pir5CSmHl49wa9+/SuyS3IXk3WXFvLa4CN/lyzw5YgAwsMRSNwMJsAEdwTsC7asi7nz7s6T5CXmzqN7T5Kfbe4juPVocnXrFW8Q8EgA4eERLq4hEJiA3IQkZMifUiVeHzO59RmraYzkLaYpp1C5aDyPQtVOHI8EEB5FoePFNaLJ/EbLCGh5TJEbZd1CVl6PLdsh4/uIJTUY8+G/j08T3172mZfCGKkphTzJ0QEBhIcDiDNwIS96rKj+ToZwKGb4T+qybWqJtq/6+lTf2ufX65VtVzmLrzZzFaOvn7Y8fOzvm5PvcT5qa/Lpuw7V/rUKD9XQNCTH0w4NXUg6h6aL4ZB9SRfvIfkh7LrGekgNlxDQRQDhoasfZAMBCEAAAhBQQMBfCgiPZrbyE0kuv0rbTIC9EIAABCAAAQ8EEB4eoOISAhCAAATmRYBq3BFAeLhjiScIQAACEIAABDYQQHhsAMRhCEAAAhCoE2AbAuMJ5Cw85O93iDXSy+RPpTfWPmWncBtiU2KZuX3imbEs3RJoY+82ih5vF975iUWT6cmQTCCgn0DOwqOtO9UHS9sOsr+dgNyE2o/6OdI3Zt9xfrKcp9cupnJMzGXl4q/LXMayfdlCw95vr9t52ftdrdv+29ZdxGrzvWl/iNguYhgfbfWY4yz9EkB4+OXrxftc/oZH6DrkYuOlIZZTiSFm7fKyKjGM9QkwZGxff33GhRxzx2/8desTzDF5GMHRZ64ENibz+sxxPUZ67NrnHP11ceo6NkcWsWpCeMQiP8O4IiRsq5doH5P1+vGp2+LT2FRfKc2PfbF8+XNuLQz3q8t7S7P0wdH4tpcS54kru85+Bb5JPDz98L1lm0l825rm28enrpva636mngvGb31p4tT3m21z3NXS+LWXrnzbfmz/Zt0+zrofAlt+3CbrlbdZkm3deuIxLiL2Rd9eX8/Mz1boeHYV//If/21velnfVJ/020vgE6ciOE5WGxebjjdOGrnTrlXWxUa6YhoEohBAeETBPj4oF5nx7ELPNDfLUHFDxzN13f2zZ83qLJZ1jn1FRX2c76ceNuyXPfdWe5N1CKgmgPC41h6edlxj4Xxtb/fYuU9NDus3q1i5xcjjsX+7XLz1Q4/KRxxile00rl1IXUxsCjR0/CZ/fY9//Vv+nzr1zSWlcTFeLynx8ZVrrsLDvracYqv1ZEz5acf+gTnVTuGe1Q7pkZgpKsa5FCKmXaPU+sVvPuXs8xXib4hJvWJD5sxp7Bxrl5ps89kvE8dnDHyvE8jjbrBeM1szJWAuIPbSLrV+s7SPTVmXeFPmT5krNRmz/YTISeLaMWU9RFyJYZvENdaUkzk2l6Vdu6zPpa7YdQhLsdh55BAf4VHrcugTrxa+dTOHC2pr8Q4OhOBnx7DXQ51T737NT6+RChFX6nzrzzwneNy1gGysCEg/VhusNBIQRsaaBoR43TTFzWkfwiOnbmdaq1xkfJWu6SL1oV9/UXnfa29bK1XyE1vb6Xjj4d9/QfXrs7ZbiSlm7/O97rPPvnN34X9O9UsttrngM8RH6HN3SG5zGKtMeARB2vn5jiAZEMQLgdAXqiEXpyFjp8D52H0vLR948+1TXEye+9Kf8vs5bbvP9vrkxBNwYNcr6ybln7vtFrPKcgABYWhswDSGTiSQo/BoQub3StkUkX1eCcjFxAQIddO/4z2PLCSWMRM/9PKDb39x9QTiV9YffnhPQ5iLfeO717S9sPAe2GOAkL8SO6WMr/3XxSnTmbskIOfucsFXAAIIjwCQCRGHwP2vunEV2PUNsNFf2a1fG+esMnS/8vH33Vu+/NZz7h3P3GP9V2L7ig/pb9+xrhDaN0uJ78ovfiDgk0BuwuNpnzDxrYvAR9/9qjUl4OrCbPuRC7+x7/zZ/dWTBrNtljGpfPYDry1f/BMXnKZw5299sveTnTf85DNOY4dyNlR8yDlx7TnP1SzrPq7udf9dzjPjVfIw6yyHEWhhN8wJo3sRyE14XNdFReuJZ19YuvLn2GkCD9x909rO2D2OEf/L73/1mgBbAzJi4+L/Ha3Nkppssw9+6j/TfeJSFw7yNKPNYokOm/Wc1u3zyay7qM/4alra/rnm2jTcr+cmPNwTxKNqAh984O7qKUTMJDVcxCSHtz5vLxgGiWfMRdC7bjpTuXnLCw+qZahvdfHRJ+6YOX38do0R1l3HnR1L3FEfTn3GJI4hevoIj+gtIIEpBOQiYazLj4zpOj7kmPgy1neeGS/LvnO6xokfsa4x9WMP/+E9Tp58SFzb6nFkW36ilKUr+9wHX18JyEd+9x1OahiSlwiJITbEd9+xhnfX+D5juubHPmbyb1u6yq/Nv9nvKg5+2gnkJDw+3IIh+IWsJY/Z7TYvZLOMXaCWPGJzcB3fcK0vXcdR6C+blExvsymYQr0SyBsV2JAAAASVSURBVEl4PLgkKSKjbsvder/e9doX6k2OzCAAAQhAAAIDCeQkPDaiMape0/LD990hQmlj7gyAQHQCJAABCECgBwGERw9IDIEABCAAAQhAwA0BhIcbjniBQJ0A2xCAAAQg0EAA4dEAhV0QgAAEIAABCPghgPDwwxWvdQJsQwACEIAABJYEEB5LCHxBAAIQgAAEIBCGAMIjDOd6FLYhAAEIQAACWRJAeGTZdoqGAAQgAAEIxCGgQ3jEqZ2oEIAABCAAAQgEJoDwCAyccBCAAAQgAAFtBELmg/AISZtYEIAABCAAgcwJIDwyPwEoHwIQgAAE6gTY9kkA4eGTLr4hAAEIQAACEFgjgPBYw8EGBCAAAQjUCbANAZcEEB4uaeILAhCAAAQgAIFOAgiPTjwchAAEIFAnwDYEIDCFAMJjCj3mQgACEIAABCAwiADCYxAuBkMAAnUCbEMAAhAYQgDhMYQWYyEAAQhAAAIQmEQA4TEJH5MhUCfANgQgAAEIdBFAeHTR4RgEIAABCEAAAk4JIDyc4sRZnQDbEIAABCAAAZsAwsOmwToEIAABCEAAAl4JIDy84q07ZxsCEIAABCCQNwGER979p3oIQAACEIBAUAJRhUfQSgkGAQhAAAIQgEB0AgiP6C0gAQhAAAIQgEAUAlGCIjyiYCcoBCAAAQhAIE8CCI88+07VEIAABCBQJ8B2EAIIjyCYCQIBCEAAAhCAgBBAeAgFDAIQgAAE6gTYhoAXAggPL1hxCgEIQAACEIBAEwGERxMV9kEAAhCoE2AbAhBwQgDh4QQjTiAAAQhAAAIQ6EMA4dGHEmMgAIE6AbYhAAEIjCKA8BiFjUkQgAAEIAABCIwhgPAYQ405EKgTYBsCEIAABHoRQHj0wsQgCEAAAhCAAARcEEB4uKCIjzoBtiEAAQhAAAKNBBAejVjYCQEIQAACEICADwIIDx9U6z7ZhgAEIAABCECgIoDwqDDwDQIQgAAEIACBEARiCI8QdREDAhCAAAQgAAGFBBAeCptCShCAAAQgAAF/BOJ6RnjE5U90CEAAAhCAQFYEEB5ZtZtiIQABCECgToDtsAQQHmF5Ew0CEIAABCCQNQGER9btp3gIQAACdQJsQ8AvAYSHX754hwAEIAABCEDAIoDwsGCwCgEIQKBOgG0IQMAtAYSHW554gwAEIAABCECggwDCowMOhyAAgToBtiEAAQhMI4DwmMaP2RCAAAQgAAEIDCCA8BgAi6EQqBNgGwIQgAAEhhFAeAzjxWgIQAACEIAABCYQQHhMgMfUOgG2IQABCEAAAt0EEB7dfDgKAQhAAAIQgIBDAggPhzDrrtiGAAQgAAEIQGCdAMJjnQdbEIAABCAAAQh4JBBQeHisAtcQgAAEIAABCCRBAOGRRJtIEgIQgAAEIDCRgJLpCA8ljSANCEAAAhCAQA4EEB45dJkaIQABCECgToDtSAQQHpHAExYCEIAABCCQIwGER45dp2YIQAACdQJsQyAQAYRHINCEgQAEIAABCECgKBAenAUQgAAEThNgDwQg4IkAwsMTWNxCAAIQgAAEIHCawP8DAAD//2w6LToAAAAGSURBVAMAPHeuoyJ6MZcAAAAASUVORK5CYII=';

async function generarPDFAuditoria(datos, fecha) {
  return new Promise((resolve, reject) => {
    const doc = new PDFDocument({ size: 'LETTER', margin: 36 });
    const chunks = [];
    doc.on('data', c => chunks.push(c));
    doc.on('end',  () => resolve(Buffer.concat(chunks)));
    doc.on('error', reject);

    const AZUL_OSC = '#16213e';
    const W = doc.page.width - 72; // ancho util

    // ── Título centrado arriba ──
    doc.fontSize(16).font('Helvetica-Bold').text('AUDITORIA GENERAL', 0, 30, { align: 'center' });

    // ── Logo a la izquierda, debajo del título ──
    const LOGO_Y = 56;
    try {
      const logoBuf = Buffer.from(LOGO_BASE64, 'base64');
      doc.image(logoBuf, 36, LOGO_Y, { width: 90, height: 36 });
    } catch(e) { console.error('Error insertando logo:', e.message); }

    // ── Header: datos generales, a la derecha del logo ──
    // Posicionado explícitamente para que el texto haga wrap correctamente
    // y nunca se solape con el logo ni se salga del margen derecho.
    const HEADER_X = 135;                  // después del logo
    const HEADER_W = W - (HEADER_X - 36);  // ancho disponible hasta el margen derecho
    doc.fontSize(9).font('Helvetica-Bold').text('Planta: ', HEADER_X, LOGO_Y, { continued: true, width: HEADER_W })
       .font('Helvetica').text(datos.planta || '', { continued: true })
       .font('Helvetica-Bold').text('    Fecha: ', { continued: true })
       .font('Helvetica').text(fecha);
    doc.font('Helvetica-Bold').text('Unidad: ', HEADER_X, doc.y, { continued: true, width: HEADER_W })
       .font('Helvetica').text(String(datos.unidad||''), { continued: true })
       .font('Helvetica-Bold').text('    Nombre del operador: ', { continued: true })
       .font('Helvetica').text(datos.operadorNombreAud || datos.operador || '');
    doc.font('Helvetica-Bold').text('Kilometraje: ', HEADER_X, doc.y, { continued: true, width: HEADER_W })
       .font('Helvetica').text(String(datos.kilometraje||''), { continued: true })
       .font('Helvetica-Bold').text('    Auditor: ', { continued: true })
       .font('Helvetica').text(datos.auditor || '');

    // Asegurar que el cursor quede debajo de lo más alto entre logo y header,
    // con un pequeño margen antes de empezar la tabla
    doc.x = 36;
    doc.y = Math.max(doc.y, LOGO_Y + 36) + 12;


    // ── Tabla ──
    const COL = [W*0.45, W*0.12, W*0.12, W*0.31];
    const X   = [36, 36+COL[0], 36+COL[0]+COL[1], 36+COL[0]+COL[1]+COL[2]];
    const ROW_H = 16;

    // Dibuja una fila con fondo, borde exterior Y líneas verticales entre columnas
    function dibujarFila(cols, y, bg) {
      if (bg) doc.rect(36, y, W, ROW_H).fill(bg);
      doc.rect(36, y, W, ROW_H).stroke('#999999'); // borde exterior de la fila
      // Líneas verticales entre cada columna
      for (let i = 1; i < X.length; i++) {
        doc.moveTo(X[i], y).lineTo(X[i], y + ROW_H).stroke('#999999');
      }
      cols.forEach((txt, i) => {
        doc.fillColor(bg === AZUL_OSC ? '#ffffff' : '#000000')
           .fontSize(8).font(bg === AZUL_OSC ? 'Helvetica-Bold' : 'Helvetica')
           .text(txt, X[i]+3, y+4, { width: COL[i]-6, lineBreak: false });
      });
    }

    let y = doc.y;
    dibujarFila(['Puntos de Revision','Cumple','No Cumple','Observaciones'], y, AZUL_OSC);
    y += ROW_H;

    const puntoMap = {};
    (datos.puntos||[]).forEach(p => { puntoMap[p.nombre] = p; });

    SECCIONES_AUD.forEach(sec => {
      if (y + ROW_H > doc.page.height - 80) { doc.addPage(); y = 36; }
      dibujarFila([sec.nombre,'','',''], y, AZUL_OSC);
      y += ROW_H;
      sec.puntos.forEach(nombre => {
        const p = puntoMap[nombre] || {};
        if (y + ROW_H > doc.page.height - 80) { doc.addPage(); y = 36; }
        const bg = p.valor === 'nocumple' ? '#fff3f3' : '#ffffff';
        dibujarFila([
          nombre,
          p.valor === 'cumple'   ? 'X' : '',
          p.valor === 'nocumple' ? 'X' : '',
          p.obs || '',
        ], y, bg);
        y += ROW_H;
      });
    });

    // ── Firmas ──
    doc.moveDown(1.5);
    const yFirma     = doc.y;
    const mitadIzqX  = 36;
    const mitadDerX  = 36 + W/2;
    const FIRMA_W    = 160;
    const FIRMA_H    = 64;

    // Firma auditor — imagen centrada dentro de su mitad, texto centrado en la misma mitad
    if (datos.firmaAuditor) {
      try {
        const bufA = Buffer.from(datos.firmaAuditor, 'base64');
        const xImgA = mitadIzqX + (W/2 - FIRMA_W) / 2; // centra la imagen en su mitad
        doc.image(bufA, xImgA, yFirma, { width: FIRMA_W, height: FIRMA_H });
      } catch(e) {}
    }
    doc.fontSize(9).font('Helvetica-Bold').text('Nombre y Firma del Auditor', mitadIzqX, yFirma + FIRMA_H + 6, { width: W/2, align: 'center' });
    doc.font('Helvetica').text(datos.auditor || '', mitadIzqX, yFirma + FIRMA_H + 18, { width: W/2, align: 'center' });

    // Firma operador — imagen centrada dentro de su mitad, texto centrado en la misma mitad
    if (datos.firmaOperador) {
      try {
        const bufO = Buffer.from(datos.firmaOperador, 'base64');
        const xImgO = mitadDerX + (W/2 - FIRMA_W) / 2; // centra la imagen en su mitad
        doc.image(bufO, xImgO, yFirma, { width: FIRMA_W, height: FIRMA_H });
      } catch(e) {}
    }
    doc.font('Helvetica-Bold').text('Nombre y Firma del Operador', mitadDerX, yFirma + FIRMA_H + 6, { width: W/2, align: 'center' });
    doc.font('Helvetica').text(datos.operadorNombreAud || datos.operador || '', mitadDerX, yFirma + FIRMA_H + 18, { width: W/2, align: 'center' });

    doc.end();
  });
}

app.post('/api/auditoria', async (req, res) => {
  try {
    const datos = req.body;
    const { sheets, drive } = await getClients();
    const { fecha, hora } = ahoraMty();
    const fechaCarp = fecha.replace(/\//g, '-');

    // Evita una segunda auditoria del mismo folio (doble envio)
    if (!datos.folio) return res.json({ ok: false, error: 'Falta el folio de la unidad' });
    const activa = await filaDeFolio(sheets, { spreadsheetId: SHEET_ID, hoja: 'Entradas', folio: datos.folio, colEstado: 6 });
    if (!activa) {
      return res.json({ ok: false, error: `El folio ${datos.folio} ya fue auditado. Actualiza la pagina.` });
    }

    // ── 1. Carpeta Unidad dentro de FOLDER_AUDITORIAS ────────
    const carpUnidad = await getOCrearCarpeta(drive, FOLDER_AUDITORIAS, `Unidad-${datos.unidad}`);
    const carpFecha  = await getOCrearCarpeta(drive, carpUnidad, fechaCarp);

    // ── 2. Subir fotos de cada punto (secuencial, evita rate limit) ─
    const evidenciasAud = [];
    for (const p of (datos.puntos || [])) {
      if (!p.fotos || !p.fotos.length) continue;
      const nombreBase = (p.nombre||'punto').replace(/[\/\\:*?"<>|]/g, '-');
      for (let i = 0; i < p.fotos.length; i++) {
        if (!p.fotos[i]?.base64) continue;
        const nombre = i === 0 ? `${nombreBase}.jpg` : `${nombreBase}-${i+1}.jpg`;
        try {
          const sub = await subirArchivo(drive, carpFecha, nombre, p.fotos[i].base64, 'image/jpeg');
          evidenciasAud.push([datos.folio||'', 'Auditoria', p.nombre||'', sub.id, nombre, fecha]);
        } catch(eFoto) {
          console.error(`[Auditoria] Foto perdida "${nombre}": ${eFoto.message}`);
        }
      }
    }
    await registrarEvidencias(sheets, evidenciasAud);

    // ── 3. Generar PDF con PDFKit ────────────────────────────
    const pdfBuffer = await generarPDFAuditoria(datos, fecha);
    const pdfNombre = `Auditoria_${datos.unidad}_${fechaCarp}.pdf`;
    const pdfFile   = await subirArchivo(drive, carpUnidad, pdfNombre, pdfBuffer.toString('base64'), 'application/pdf');
    // Hacer el PDF accesible con el link
    await drive.permissions.create({ fileId: pdfFile.id, requestBody: { role:'reader', type:'anyone' } });
    const pdfUrl = `https://drive.google.com/file/d/${pdfFile.id}/view`;

    // ── 4. Registrar en hoja Auditorias ──────────────────────
    const puntoMap = {};
    (datos.puntos||[]).forEach(p => { puntoMap[p.nombre] = p; });

    // Verificar/crear hoja Auditorias
    try { await sheets.spreadsheets.values.get({ spreadsheetId: SHEET_ID, range: 'Auditorias!A1' }); }
    catch {
      await sheets.spreadsheets.batchUpdate({ spreadsheetId: SHEET_ID, requestBody: { requests: [{ addSheet: { properties: { title: 'Auditorias' } } }] } });
      const headers = ['Folio','Fecha','Hora','Unidad','Operador','Planta','Auditor','Kilometraje',
        ...SECCIONES_AUD.flatMap(s => s.puntos), 'PDF'];
      await sheets.spreadsheets.values.update({ spreadsheetId: SHEET_ID, range: 'Auditorias!A1', valueInputOption:'RAW', requestBody:{values:[headers]} });
    }

    const rowBase = [datos.folio, fecha, hora, datos.unidad, datos.operador, datos.planta, datos.auditor, datos.kilometraje||''];
    const rowPuntos = SECCIONES_AUD.flatMap(s => s.puntos.map(nombre => {
      const p = puntoMap[nombre];
      if (!p) return '';
      if (p.valor === 'cumple')   return 'Cumple';
      if (p.valor === 'nocumple') return `No Cumple${p.obs ? ': ' + p.obs : ''}`;
      return '';
    }));

    // Los datos se escriben en RAW (texto tal cual). Antes toda la fila iba
    // en USER_ENTERED y Sheets reinterpretaba la fecha segun la region del
    // archivo. Solo la celda del PDF necesita USER_ENTERED para la formula.
    const apAud = await sheets.spreadsheets.values.append({
      spreadsheetId: SHEET_ID, range: 'Auditorias!A:A',
      valueInputOption: 'RAW', insertDataOption: 'INSERT_ROWS',
      requestBody: { values: [[...rowBase, ...rowPuntos, 'PDF']] }
    });
    const filaAud = filaDeAppend(apAud);
    if (filaAud) {
      const colPdf = rowBase.length + rowPuntos.length + 1;
      let letra = '', n = colPdf;
      while (n > 0) { const r = (n - 1) % 26; letra = String.fromCharCode(65 + r) + letra; n = Math.floor((n - 1) / 26); }
      await sheets.spreadsheets.values.update({
        spreadsheetId: SHEET_ID, range: `Auditorias!${letra}${filaAud}`,
        valueInputOption: 'USER_ENTERED', requestBody: { values: [[`=HYPERLINK("${pdfUrl}","PDF")`]] }
      });
    }

    // ── 5. Dar de baja en Entradas ───────────────────────────
    const { fecha: fBaja, hora: hBaja } = ahoraMty();
    const resE = await sheets.spreadsheets.values.get({ spreadsheetId: SHEET_ID, range: 'Entradas!A2:I' });
    const rowsE = resE.data.values || [];
    for (let i = 0; i < rowsE.length; i++) {
      if (String(rowsE[i][0]).trim() === String(datos.folio).trim() && rowsE[i][6] === 'ACTIVO') {
        await sheets.spreadsheets.values.update({
          spreadsheetId: SHEET_ID, range: `Entradas!G${i+2}:I${i+2}`,
          valueInputOption:'RAW', requestBody:{values:[['BAJA', fBaja, hBaja]]}
        });
        break;
      }
    }

    res.json({ ok: true, pdfUrl });
  } catch(e) { console.error(e); res.json({ ok: false, error: e.message }); }
});

// ═══════════════════════════════════════════════════════════════
// PROVEEDORES EXTERNOS — registro, listado y salida con evidencias
// Spreadsheet separado (SHEET_ID_PROVEDORES), primera hoja.
// Drive: FOLDER_PROVEDORES / NombreProveedor / dd-MM-yyyy / fotos
// ═══════════════════════════════════════════════════════════════
const HEADERS_PROV = ['Folio','Fecha Entrada','Hora Entrada','Proveedor','Piezas que Ingresa','Trabajo a Realizar','Factura','Estado','Fecha Salida','Hora Salida','Trabajo Realizado'];

// Asegura que la primera hoja tenga encabezados (solo la primera vez)
async function asegurarHeadersProv(sheets) {
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SHEET_ID_PROVEDORES, range: 'A1:K1',
  });
  const fila1 = (res.data.values || [])[0] || [];
  if (!fila1.length || !fila1[0]) {
    await sheets.spreadsheets.values.update({
      spreadsheetId: SHEET_ID_PROVEDORES, range: 'A1',
      valueInputOption: 'RAW', requestBody: { values: [HEADERS_PROV] }
    });
  }
}

// Folio de proveedor: MM/AA-P001, contador propio en su spreadsheet,
// con la misma protección de reintentos que el folio general.
async function generarFolioProv(sheets) {
  const MAX_INTENTOS = 5;
  for (let intento = 0; intento < MAX_INTENTOS; intento++) {
    if (intento > 0) await new Promise(r => setTimeout(r, 200 + Math.random() * 300));
    const ahora   = new Date();
    const mes     = String(ahora.toLocaleDateString('es-MX', { month:'2-digit', timeZone:'America/Monterrey' })).padStart(2,'0');
    const anio    = ahora.toLocaleDateString('es-MX', { year:'2-digit', timeZone:'America/Monterrey' });
    const prefijo = `${mes}/${anio}-P`;
    const res  = await sheets.spreadsheets.values.get({ spreadsheetId: SHEET_ID_PROVEDORES, range: 'A2:A' });
    const filas = (res.data.values || []).flat();
    let max = 0;
    filas.forEach(f => {
      if (typeof f === 'string' && f.startsWith(prefijo)) {
        const num = parseInt(f.slice(prefijo.length), 10);
        if (!isNaN(num) && num > max) max = num;
      }
    });
    const folio = `${prefijo}${String(max + 1).padStart(3, '0')}`;
    if (filas.includes(folio)) { console.warn(`Folio prov ${folio} duplicado, reintento`); continue; }
    return folio;
  }
  const ahora = new Date();
  const mes   = String(ahora.toLocaleDateString('es-MX', { month:'2-digit', timeZone:'America/Monterrey' })).padStart(2,'0');
  const anio  = ahora.toLocaleDateString('es-MX', { year:'2-digit', timeZone:'America/Monterrey' });
  return `${mes}/${anio}-PERR${Date.now().toString().slice(-4)}`;
}

// Sube fotos a Proveedores/Nombre/dd-MM-yyyy/ con nombres únicos
// (timestamp en el nombre) para que NUNCA se sobreescriban.
async function subirFotosProveedor(drive, nombreProveedor, fecha, fotos, prefijoNombre) {
  const subidas = [];
  if (!fotos || !fotos.length) return subidas;
  const nombreLimpio = String(nombreProveedor).replace(/[\/\\:*?"<>|]/g, '-').trim();
  const carpProv  = await getOCrearCarpeta(drive, FOLDER_PROVEDORES, nombreLimpio);
  const carpFecha = await getOCrearCarpeta(drive, carpProv, fecha.replace(/\//g, '-'));
  for (let i = 0; i < fotos.length; i++) {
    if (!fotos[i]?.base64) continue;
    const nombre = `${prefijoNombre}-${Date.now()}-${i + 1}.jpg`;
    try {
      const sub = await subirArchivo(drive, carpFecha, nombre, fotos[i].base64, 'image/jpeg');
      subidas.push({ fileId: sub.id, nombre });
    } catch(e) { console.error('Foto prov error:', e.message); }
  }
  return subidas;
}

// ── Registrar entrada de proveedor ─────────────────────────────
app.post('/api/registrar-proveedor', async (req, res) => {
  try {
    const datos = req.body;
    if (!datos.proveedor || !String(datos.proveedor).trim()) {
      return res.json({ ok: false, error: 'Falta el nombre del proveedor' });
    }
    const { sheets, drive } = await getClients();
    const { fecha, hora } = ahoraMty();

    await asegurarHeadersProv(sheets);
    let folio = await generarFolioProv(sheets);

    // Foto de factura (si aplica)
    if (datos.factura === 'Si' && datos.facturaFotos?.length) {
      const subs = await subirFotosProveedor(drive, datos.proveedor, fecha, datos.facturaFotos, 'Factura');
      await registrarEvidencias(sheets, subs.map(s => [folio, 'Proveedor', 'Factura', s.fileId, s.nombre, fecha]));
    }

    const apProv = await sheets.spreadsheets.values.append({
      spreadsheetId: SHEET_ID_PROVEDORES, range: 'A:A',
      valueInputOption: 'RAW', insertDataOption: 'INSERT_ROWS',
      requestBody: { values: [[
        folio, fecha, hora,
        String(datos.proveedor).trim(),
        datos.piezas  || '',
        datos.trabajo || '',
        datos.factura === 'Si' ? 'Si' : 'No',
        'ACTIVO', '', '', ''
      ]] }
    });
    folio = await confirmarFolio(sheets, {
      spreadsheetId: SHEET_ID_PROVEDORES, hoja: null, fila: filaDeAppend(apProv), folio,
      generar: () => generarFolioProv(sheets),
    });

    res.json({ ok: true, folio });
  } catch(e) { console.error(e); res.json({ ok: false, error: e.message }); }
});

// ── Listar proveedores activos ──────────────────────────────────
app.get('/api/proveedores-activos', async (req, res) => {
  try {
    const { sheets } = await getClients();
    const r = await sheets.spreadsheets.values.get({ spreadsheetId: SHEET_ID_PROVEDORES, range: 'A2:K' });
    const rows = r.data.values || [];
    const activos = rows.map((row, i) => ({
      rowIndex: i + 2,
      folio:     row[0] || '',
      fecha:     row[1] || '',
      hora:      row[2] || '',
      proveedor: row[3] || '',
      piezas:    row[4] || '',
      trabajo:   row[5] || '',
      estado:    row[7] || '',
    })).filter(p => p.estado === 'ACTIVO');
    res.json({ ok: true, proveedores: activos });
  } catch(e) { console.error(e); res.json({ ok: false, error: e.message }); }
});

// ── Salida de proveedor: password + trabajo realizado + fotos ──
app.post('/api/salida-proveedor', async (req, res) => {
  try {
    const { password, trabajoRealizado, fotos, folio } = req.body;
    let rowIndex = parseInt(req.body.rowIndex, 10) || null;
    if (password !== PASS_BAJA) return res.json({ ok: false, error: 'Contrasena incorrecta' });
    if (!rowIndex && !folio) return res.json({ ok: false, error: 'Falta el folio del proveedor' });
    if (!trabajoRealizado || !String(trabajoRealizado).trim()) {
      return res.json({ ok: false, error: 'Describe el trabajo realizado' });
    }
    if (!fotos || !fotos.length) {
      return res.json({ ok: false, error: 'Sube al menos una foto de evidencia' });
    }

    const { sheets, drive } = await getClients();
    const { fecha, hora } = ahoraMty();

    // Confirmar que el renglon sigue siendo este proveedor (si alguien movio
    // filas en el Sheet, se busca por folio en vez de cerrar a otro)
    if (folio) {
      rowIndex = await filaDeFolio(sheets, { spreadsheetId: SHEET_ID_PROVEDORES, hoja: null, folio, sugerida: rowIndex, colEstado: 7 });
      if (!rowIndex) return res.json({ ok: false, error: 'Este proveedor ya fue dado de baja. Actualiza la lista.' });
    }

    // Leer la fila para obtener el nombre del proveedor y verificar estado
    const r = await sheets.spreadsheets.values.get({
      spreadsheetId: SHEET_ID_PROVEDORES, range: `A${rowIndex}:K${rowIndex}`,
    });
    const fila = (r.data.values || [])[0];
    if (!fila) return res.json({ ok: false, error: 'Registro no encontrado' });
    if (fila[7] !== 'ACTIVO') return res.json({ ok: false, error: 'Este proveedor ya fue dado de baja' });
    const nombreProveedor = fila[3] || 'SinNombre';

    // Subir evidencias (carpeta del día de salida, nombres únicos)
    const subsSal = await subirFotosProveedor(drive, nombreProveedor, fecha, fotos, 'Evidencia');
    await registrarEvidencias(sheets, subsSal.map(s => [fila[0]||'', 'Proveedor', 'Evidencia de salida', s.fileId, s.nombre, fecha]));

    // Actualizar: Estado, Fecha/Hora Salida, Trabajo Realizado (H:K)
    await sheets.spreadsheets.values.update({
      spreadsheetId: SHEET_ID_PROVEDORES, range: `H${rowIndex}:K${rowIndex}`,
      valueInputOption: 'RAW',
      requestBody: { values: [['BAJA', fecha, hora, String(trabajoRealizado).trim()]] }
    });

    res.json({ ok: true });
  } catch(e) { console.error(e); res.json({ ok: false, error: e.message }); }
});

// ═══════════════════════════════════════════════════════════════
// ANTIDOPING — folio consecutivo y registro en hoja mensual
// Spreadsheet propio (SHEET_ID_ANTIDOPING).
// Hoja por mes: "DOPING MES AÑO". Datos desde la fila 4.
// ═══════════════════════════════════════════════════════════════
const MESES_AD = ['ENERO','FEBRERO','MARZO','ABRIL','MAYO','JUNIO',
                  'JULIO','AGOSTO','SEPTIEMBRE','OCTUBRE','NOVIEMBRE','DICIEMBRE'];
const HEADERS_AD = ['#','FECHA','OPERADOR','PLANTA','UNIDAD',
  'RAZON DEL EXAMEN','MEDICAMENTO EL ULTIMO MES','ESPECIMEN',
  'AMP','MET','TCH','COC','OPI','MPD','BZD'];
const DATA_START_ROW_AD = 4;

function tituloHojaAD(fecha) {
  return `DOPING ${MESES_AD[fecha.getMonth()]} ${fecha.getFullYear()}`;
}

// Folio global: el mayor numero de la columna A en TODAS las hojas DOPING.
// Reemplaza al PropertiesService de Apps Script (que no existe en Node) y
// ademas sobrevive a un redeploy, porque la fuente de verdad es la hoja.
async function ultimoFolioAD(sheets) {
  const meta = await sheets.spreadsheets.get({ spreadsheetId: SHEET_ID_ANTIDOPING });
  const hojas = (meta.data.sheets || [])
    .map(s => s.properties.title)
    .filter(t => t.toUpperCase().trim().startsWith('DOPING '));
  if (!hojas.length) return 0;
  const rangos = hojas.map(h => `'${h}'!A${DATA_START_ROW_AD}:A`);
  const res = await sheets.spreadsheets.values.batchGet({
    spreadsheetId: SHEET_ID_ANTIDOPING, ranges: rangos,
  });
  let max = 0;
  (res.data.valueRanges || []).forEach(vr => {
    (vr.values || []).flat().forEach(v => {
      const n = parseInt(v, 10);
      if (!isNaN(n) && n > max) max = n;
    });
  });
  return max;
}

// Crea la hoja del mes con titulo, encabezados y formato si aun no existe
async function getOCrearHojaAD(sheets, fecha) {
  const titulo = tituloHojaAD(fecha);
  const meta = await sheets.spreadsheets.get({ spreadsheetId: SHEET_ID_ANTIDOPING });
  const existente = (meta.data.sheets || [])
    .find(s => s.properties.title.toUpperCase().trim() === titulo);
  if (existente) return titulo;

  const add = await sheets.spreadsheets.batchUpdate({
    spreadsheetId: SHEET_ID_ANTIDOPING,
    requestBody: { requests: [{ addSheet: { properties: { title: titulo } } }] }
  });
  const sheetId = add.data.replies[0].addSheet.properties.sheetId;

  // Titulo en A1:O1 y encabezados en la fila 3
  await sheets.spreadsheets.values.batchUpdate({
    spreadsheetId: SHEET_ID_ANTIDOPING,
    requestBody: { valueInputOption: 'RAW', data: [
      { range: `'${titulo}'!A1`, values: [[titulo]] },
      { range: `'${titulo}'!A3`, values: [HEADERS_AD] },
    ]}
  });

  const anchos = [40,90,200,130,70,160,160,90,80,80,80,80,80,80,80];
  await sheets.spreadsheets.batchUpdate({
    spreadsheetId: SHEET_ID_ANTIDOPING,
    requestBody: { requests: [
      { mergeCells: { range: { sheetId, startRowIndex:0, endRowIndex:1, startColumnIndex:0, endColumnIndex:15 }, mergeType:'MERGE_ALL' } },
      { repeatCell: {
          range: { sheetId, startRowIndex:0, endRowIndex:1, startColumnIndex:0, endColumnIndex:15 },
          cell: { userEnteredFormat: {
            backgroundColor:{ red:0.965, green:0.722, blue:0 },
            horizontalAlignment:'CENTER', verticalAlignment:'MIDDLE',
            textFormat:{ bold:true, fontSize:18, foregroundColor:{ red:1, green:1, blue:1 } }
          }},
          fields:'userEnteredFormat' } },
      { repeatCell: {
          range: { sheetId, startRowIndex:2, endRowIndex:3, startColumnIndex:0, endColumnIndex:15 },
          cell: { userEnteredFormat: {
            backgroundColor:{ red:0.965, green:0.722, blue:0 },
            horizontalAlignment:'CENTER', wrapStrategy:'WRAP',
            textFormat:{ bold:true, foregroundColor:{ red:1, green:1, blue:1 } }
          }},
          fields:'userEnteredFormat' } },
      { updateDimensionProperties: { range:{ sheetId, dimension:'ROWS', startIndex:0, endIndex:1 }, properties:{ pixelSize:50 }, fields:'pixelSize' } },
      { updateDimensionProperties: { range:{ sheetId, dimension:'ROWS', startIndex:1, endIndex:2 }, properties:{ pixelSize:8  }, fields:'pixelSize' } },
      { updateDimensionProperties: { range:{ sheetId, dimension:'ROWS', startIndex:2, endIndex:3 }, properties:{ pixelSize:36 }, fields:'pixelSize' } },
      ...anchos.map((w,i) => ({ updateDimensionProperties: {
        range:{ sheetId, dimension:'COLUMNS', startIndex:i, endIndex:i+1 },
        properties:{ pixelSize:w }, fields:'pixelSize' } })),
    ]}
  });
  return titulo;
}

// Primera fila vacia desde la fila 4 (respeta contenido previo de la hoja)
async function primeraFilaVaciaAD(sheets, titulo) {
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SHEET_ID_ANTIDOPING, range: `'${titulo}'!A${DATA_START_ROW_AD}:A`,
  });
  const filas = res.data.values || [];
  for (let i = 0; i < filas.length; i++) {
    const c = filas[i][0];
    if (c === '' || c === null || c === undefined) return DATA_START_ROW_AD + i;
  }
  return DATA_START_ROW_AD + filas.length;
}

app.get('/api/antidoping-folio', async (req, res) => {
  try {
    if (!SHEET_ID_ANTIDOPING) return res.json({ ok:false, error:'Falta SHEET_ID_ANTIDOPING' });
    const { sheets } = await getClients();
    const ultimo = await ultimoFolioAD(sheets);
    res.json({ ok:true, folio: ultimo + 1 });
  } catch(e) { console.error('[Antidoping] folio:', e.message); res.json({ ok:false, error:e.message }); }
});

app.post('/api/antidoping', async (req, res) => {
  try {
    if (!SHEET_ID_ANTIDOPING) return res.json({ result:'error', error:'Falta SHEET_ID_ANTIDOPING' });
    const d = req.body;
    const { sheets } = await getClients();

    const fecha  = d.fecha ? new Date(d.fecha + 'T12:00:00') : new Date();
    const titulo = await getOCrearHojaAD(sheets, fecha);

    const folioNum = (await ultimoFolioAD(sheets)) + 1;
    const fila = [
      folioNum, d.fecha||'', d.operador||'', d.planta||'', d.unidad||'',
      d.razon||'', d.medicamento||'NO', d.especimen||'',
      d.AMP||'', d.MET||'', d.TCH||'', d.COC||'', d.OPI||'', d.MPD||'', d.BZD||'',
    ];

    const filaDestino = await primeraFilaVaciaAD(sheets, titulo);
    await sheets.spreadsheets.values.update({
      spreadsheetId: SHEET_ID_ANTIDOPING,
      range: `'${titulo}'!A${filaDestino}:O${filaDestino}`,
      valueInputOption: 'RAW', requestBody: { values: [fila] }
    });

    console.log(`[Antidoping] Folio ${folioNum} → ${titulo} fila ${filaDestino}`);
    res.json({
      result: 'success',
      folio: String(folioNum).padStart(5,'0'),
      hoja: titulo,
      nextFolio: folioNum + 1,
    });
  } catch(e) { console.error('[Antidoping] guardar:', e.message); res.json({ result:'error', error:e.message }); }
});

// ═══════════════════════════════════════════════════════════════
// CONSULTA E IMPRESION DE REPORTES FINALIZADOS
// Lee lo que ya esta guardado en las hojas y arma un PDF:
// la informacion distribuida en la primera hoja y las fotos
// en hojas aparte con su titulo.
// ═══════════════════════════════════════════════════════════════

// Convierte una fila + sus encabezados en pares [etiqueta, valor],
// omitiendo las 8 columnas comunes y los campos vacios.
function paresDeFila(headers, fila) {
  const out = [];
  for (let i = 8; i < headers.length; i++) {
    const val = fila[i];
    if (val === undefined || val === null || String(val).trim() === '') continue;
    if (String(val).trim() === 'No') continue;
    out.push([headers[i], String(val)]);
  }
  return out;
}

// Solo las hojas de area que ya existen. Si alguna aun no se crea
// (nadie ha enviado, por ejemplo, un reporte de Suspension), pedirla
// hacia fallar la lectura COMPLETA y el listado salia vacio.
async function hojasTallerExistentes(sheets) {
  const meta = await sheets.spreadsheets.get({ spreadsheetId: SHEET_ID, fields: 'sheets.properties.title' });
  const titulos = new Set((meta.data.sheets || []).map(x => x.properties.title));
  return Object.keys(HOJAS_TALLER).filter(k => titulos.has(HOJAS_TALLER[k]));
}
// Una fila con solo "No" (valores por defecto) no es trabajo realizado
const filaConTrabajo = fila => fila.slice(8).some(c => c != null && String(c).trim() !== '' && String(c).trim() !== 'No');

// Orden cronologico de folios MM/AA-[T|P]NNN (el orden de texto ponia
// enero del año siguiente debajo de diciembre)
function claveFolio(f) {
  const m = String(f).match(/^(\d{2})\/(\d{2})-[A-Z]*?(\d+)/);
  return m ? (+m[2]) * 1e7 + (+m[1]) * 1e5 + (+m[3]) : 0;
}

// ── Listado de reportes finalizados ────────────────────────────
app.get('/api/reportes', async (req, res) => {
  try {
    const tipo = req.query.tipo || 'taller';
    const { sheets } = await getClients();

    if (tipo === 'taller') {
      const claves = await hojasTallerExistentes(sheets);
      if (!claves.length) return res.json({ ok:true, reportes: [] });
      const rangos = claves.map(k => `'${HOJAS_TALLER[k]}'!A2:BZ`);
      const r = await sheets.spreadsheets.values.batchGet({ spreadsheetId: SHEET_ID, ranges: rangos });
      const porFolio = {};
      (r.data.valueRanges || []).forEach((vr, idx) => {
        const clave = claves[idx];
        (vr.values || []).forEach(fila => {
          const folio = String(fila[0] || '').trim();
          if (!folio || !filaConTrabajo(fila)) return;
          if (!porFolio[folio]) porFolio[folio] = {
            folio, fecha: fila[1]||'', hora: fila[2]||'', unidad: fila[3]||'',
            operador: fila[4]||'', planta: fila[5]||'', areaServicio: fila[6]||'',
            mecanico: fila[7]||'', areas: [],
          };
          porFolio[folio].areas.push(HOJAS_TALLER[clave]);
        });
      });
      const lista = Object.values(porFolio).sort((a,b) => claveFolio(b.folio) - claveFolio(a.folio));
      return res.json({ ok:true, reportes: lista });
    }

    if (tipo === 'auditoria') {
      // Dos lecturas: la formateada da fecha/hora legibles y la de formula
      // permite sacar la URL de dentro del =HYPERLINK(...). El rango llega
      // hasta BZ porque la columna del PDF es la 36.
      const [rFmt, rForm] = await Promise.all([
        sheets.spreadsheets.values.get({
          spreadsheetId: SHEET_ID, range: 'Auditorias!A2:BZ',
        }).catch(() => ({ data: { values: [] } })),
        sheets.spreadsheets.values.get({
          spreadsheetId: SHEET_ID, range: 'Auditorias!A2:BZ', valueRenderOption: 'FORMULA',
        }).catch(() => ({ data: { values: [] } })),
      ]);

      const filasFmt  = rFmt.data.values  || [];
      const filasForm = rForm.data.values || [];

      const lista = filasFmt.map((f, i) => {
        // El PDF es la ultima celda de la fila que contenga un enlace
        const crudo = filasForm[i] || [];
        let pdfUrl = '';
        for (let c = crudo.length - 1; c >= 8; c--) {
          const celda = String(crudo[c] || '');
          const m = celda.match(/HYPERLINK\("([^"]+)"/i);
          if (m) { pdfUrl = m[1]; break; }
          if (/^https?:\/\//i.test(celda)) { pdfUrl = celda; break; }
        }
        return {
          folio: f[0]||'', fecha: f[1]||'', hora: f[2]||'', unidad: f[3]||'',
          operador: f[4]||'', planta: f[5]||'', auditor: f[6]||'', kilometraje: f[7]||'',
          pdfUrl,
        };
      }).filter(x => x.folio).reverse();

      return res.json({ ok:true, reportes: lista });
    }

    if (tipo === 'proveedor') {
      const r = await sheets.spreadsheets.values.get({ spreadsheetId: SHEET_ID_PROVEDORES, range: 'A2:K' })
        .catch(() => ({ data: { values: [] } }));
      const lista = (r.data.values || []).map((f, i) => ({
        rowIndex: i+2, folio: f[0]||'', fecha: f[1]||'', hora: f[2]||'',
        proveedor: f[3]||'', piezas: f[4]||'', trabajo: f[5]||'', factura: f[6]||'',
        estado: f[7]||'', fechaSalida: f[8]||'', horaSalida: f[9]||'', trabajoRealizado: f[10]||'',
      })).filter(x => x.folio && x.estado === 'BAJA').reverse();
      return res.json({ ok:true, reportes: lista });
    }

    res.json({ ok:false, error:'Tipo no valido' });
  } catch(e) { console.error('[Reportes]', e.message); res.json({ ok:false, error:e.message }); }
});

// ── PDF de un reporte ──────────────────────────────────────────
function armarPDFReporte(info, fotos) {
  return new Promise((resolve, reject) => {
    // bufferPages permite volver atras para numerar las paginas al final
    const doc = new PDFDocument({ size:'LETTER', margin:36, bufferPages:true });
    const chunks = [];
    doc.on('data', c => chunks.push(c));
    doc.on('end',  () => resolve(Buffer.concat(chunks)));
    doc.on('error', reject);

    const AZUL = '#16213e';
    const W = doc.page.width - 72;

    function encabezado(titulo, subtitulo) {
      doc.fontSize(15).font('Helvetica-Bold').fillColor('#000000')
         .text(titulo, 0, 30, { align:'center' });
      try {
        doc.image(Buffer.from(LOGO_BASE64,'base64'), 36, 54, { width:80, height:32 });
      } catch(e) {}
      if (subtitulo) {
        doc.fontSize(9).font('Helvetica').fillColor('#555555')
           .text(subtitulo, 130, 58, { width: W-100 });
      }
      doc.y = 94;
    }

    encabezado(info.titulo, info.subtitulo);

    // Datos generales en dos columnas
    const COL_W = W/2;
    let yG = doc.y;
    info.generales.forEach((par, i) => {
      const x = 36 + (i % 2) * COL_W;
      if (i % 2 === 0 && i > 0) yG += 22;
      doc.fontSize(7).font('Helvetica').fillColor('#888888')
         .text(String(par[0]).toUpperCase(), x, yG, { width: COL_W-10, lineBreak:false });
      doc.fontSize(9.5).font('Helvetica-Bold').fillColor('#000000')
         .text(String(par[1]||'—'), x, yG+7, { width: COL_W-10, lineBreak:false });
    });
    doc.y = yG + 30;

    // Secciones con sus campos
    info.secciones.forEach(sec => {
      if (!sec.campos.length) return;
      if (doc.y + 40 > doc.page.height - 60) doc.addPage();
      doc.rect(36, doc.y, W, 15).fill(AZUL);
      doc.fontSize(8.5).font('Helvetica-Bold').fillColor('#ffffff')
         .text(sec.nombre.toUpperCase(), 40, doc.y + 4.5, { lineBreak:false });
      doc.y += 20;
      doc.fillColor('#000000');

      sec.campos.forEach(par => {
        const etiqueta = String(par[0]);
        const valor    = String(par[1]);
        const alto = doc.heightOfString(valor, { width: W-150, align:'left' });
        if (doc.y + alto + 6 > doc.page.height - 50) doc.addPage();
        const y0 = doc.y;
        doc.fontSize(8).font('Helvetica').fillColor('#666666')
           .text(etiqueta, 40, y0, { width: 140 });
        doc.fontSize(9).font('Helvetica-Bold').fillColor('#000000')
           .text(valor, 185, y0, { width: W-155 });
        doc.y = y0 + Math.max(alto, 11) + 4;
        doc.moveTo(40, doc.y-2).lineTo(36+W, doc.y-2).lineWidth(0.3).stroke('#e8e8e8');
      });
      doc.y += 8;
    });

    // ── Hojas de fotos ───────────────────────────────────────
    if (fotos && fotos.length) {
      const CW = (W - 12) / 2, CH = 76;
      let col = 0, yF = 0;

      function hojaFotos() {
        doc.addPage();
        doc.fontSize(13).font('Helvetica-Bold').fillColor('#000000')
           .text('EVIDENCIA FOTOGRAFICA', 0, 34, { align:'center' });
        doc.fontSize(8.5).font('Helvetica').fillColor('#666666')
           .text(info.subtitulo || '', 0, 52, { align:'center' });
        doc.moveTo(36, 66).lineTo(36+W, 66).lineWidth(0.5).stroke('#cccccc');
        yF = 74; col = 0;
      }

      hojaFotos();
      fotos.forEach(f => {
        if (col === 0 && yF + CH + 14 > doc.page.height - 40) hojaFotos();
        const x = 36 + col * (CW + 12);
        try {
          doc.image(f.buffer, x, yF, { fit:[CW, CH], align:'center', valign:'center' });
          doc.rect(x, yF, CW, CH).lineWidth(0.4).stroke('#dddddd');
        } catch(e) {
          doc.fontSize(7).fillColor('#999999').text('(imagen no disponible)', x, yF+CH/2);
        }
        doc.fontSize(7).font('Helvetica').fillColor('#666666')
           .text(f.area || '', x, yF + CH + 3, { width: CW, lineBreak:false });
        col++;
        if (col === 2) { col = 0; yF += CH + 14; }
      });
    }

    // Pie de pagina
    const total = doc.bufferedPageRange().count;
    for (let i = 0; i < total; i++) {
      doc.switchToPage(i);
      doc.fontSize(7).font('Helvetica').fillColor('#999999')
         .text(`${info.folio}  —  TECSA Transportes  —  Pag. ${i+1} de ${total}`,
               36, doc.page.height - 46,
               { width: W, align:'center', lineBreak:false });
    }

    doc.end();
  });
}

// Corta una promesa si excede el tiempo dado. Evita que una consulta
// lenta a Drive deje el reporte generandose para siempre.
function conTiempoLimite(promesa, ms) {
  return Promise.race([
    promesa,
    new Promise((_, rej) => setTimeout(() => rej(new Error('Tiempo de espera agotado')), ms)),
  ]);
}

app.get('/api/reporte-pdf', async (req, res) => {
  try {
    const tipo  = req.query.tipo  || 'taller';
    const folio = req.query.folio;
    if (!folio) return res.status(400).send('Falta folio');

    const { sheets, drive } = await getClients();
    let info = null;
    let unidad = '', fechaRep = '';

    if (tipo === 'taller') {
      const claves = await hojasTallerExistentes(sheets);
      if (!claves.length) return res.status(404).send('Reporte no encontrado');
      const r = await sheets.spreadsheets.values.batchGet({
        spreadsheetId: SHEET_ID, ranges: claves.map(k => `'${HOJAS_TALLER[k]}'!A1:BZ`),
      });
      const secciones = [];
      let gen = null;
      (r.data.valueRanges || []).forEach((vr, idx) => {
        const filas = vr.values || [];
        if (filas.length < 2) return;
        const headers = filas[0];
        const fila = filas.slice(1).find(f => String(f[0]||'').trim() === String(folio).trim());
        if (!fila) return;
        if (!gen) gen = fila;
        const campos = paresDeFila(headers, fila);
        if (campos.length) secciones.push({ nombre: HOJAS_TALLER[claves[idx]], campos });
      });
      if (!gen) return res.status(404).send('Reporte no encontrado');
      unidad = gen[3] || ''; fechaRep = gen[1] || '';
      info = {
        folio, titulo: 'REPORTE DE TALLER',
        subtitulo: `Unidad ${unidad}  |  Folio ${folio}  |  ${fechaRep}`,
        generales: [
          ['Folio', folio], ['Fecha', gen[1]], ['Hora', gen[2]], ['Unidad', gen[3]],
          ['Operador', gen[4]], ['Planta', gen[5]], ['Area de servicio', gen[6]], ['Mecanico', gen[7]],
        ],
        secciones,
      };
    }

    else if (tipo === 'auditoria') {
      const r = await sheets.spreadsheets.values.get({ spreadsheetId: SHEET_ID, range: 'Auditorias!A1:AZ' });
      const filas = r.data.values || [];
      const headers = filas[0] || [];
      const fila = filas.slice(1).find(f => String(f[0]||'').trim() === String(folio).trim());
      if (!fila) return res.status(404).send('Auditoria no encontrada');
      unidad = fila[3] || ''; fechaRep = fila[1] || '';
      const puntos = [];
      for (let i = 8; i < headers.length; i++) {
        if (String(headers[i]).toUpperCase() === 'PDF') continue;
        const val = fila[i];
        if (val === undefined || String(val).trim() === '') continue;
        puntos.push([headers[i], String(val)]);
      }
      info = {
        folio, titulo: 'AUDITORIA GENERAL',
        subtitulo: `Unidad ${unidad}  |  Folio ${folio}  |  ${fechaRep}`,
        generales: [
          ['Folio', folio], ['Fecha', fila[1]], ['Hora', fila[2]], ['Unidad', fila[3]],
          ['Operador', fila[4]], ['Planta', fila[5]], ['Auditor', fila[6]], ['Kilometraje', fila[7]],
        ],
        secciones: [{ nombre: 'Puntos de revision', campos: puntos }],
      };
    }

    else if (tipo === 'proveedor') {
      const r = await sheets.spreadsheets.values.get({ spreadsheetId: SHEET_ID_PROVEDORES, range: 'A2:K' });
      const fila = (r.data.values || []).find(f => String(f[0]||'').trim() === String(folio).trim());
      if (!fila) return res.status(404).send('Proveedor no encontrado');
      fechaRep = fila[1] || '';
      info = {
        folio, titulo: 'REPORTE DE PROVEEDOR EXTERNO',
        subtitulo: `${fila[3]||''}  |  Folio ${folio}  |  ${fechaRep}`,
        generales: [
          ['Folio', folio], ['Proveedor', fila[3]], ['Fecha entrada', fila[1]], ['Hora entrada', fila[2]],
          ['Fecha salida', fila[8]], ['Hora salida', fila[9]], ['Factura', fila[6]], ['Estado', fila[7]],
        ],
        secciones: [{ nombre: 'Detalle del trabajo', campos: [
          ['Piezas que ingresa', fila[4]||'—'],
          ['Trabajo a realizar', fila[5]||'—'],
          ['Trabajo realizado',  fila[10]||'—'],
        ]}],
      };
    }

    else return res.status(400).send('Tipo no valido');

    // ── Fotos: primero el registro, si no hay se busca en Drive ──
    // Si el usuario ya acepto generar sin fotos, nos saltamos todo esto.
    const sinFotos = req.query.sinFotos === '1';
    const fotos = [];

    if (!sinFotos) {
      let refs = [];
      try {
        // Tope de tiempo para localizar las fotos: si Drive tarda o la
        // carpeta no existe, no dejamos el reporte colgado indefinidamente.
        refs = await conTiempoLimite((async () => {
          let r = await evidenciasDeFolio(sheets, folio);
          if (!r.length && tipo === 'taller' && unidad && fechaRep) {
            r = await buscarFotosEnDrive(drive, FOLDER_RAIZ, unidad, fechaRep.replace(/\//g,'-'));
          }
          if (!r.length && tipo === 'auditoria' && unidad && fechaRep) {
            r = await buscarFotosEnDrive(drive, FOLDER_AUDITORIAS, unidad, fechaRep.replace(/\//g,'-'));
          }
          return r;
        })(), 15000);
      } catch(e) {
        console.error('[ReportePDF] busqueda de fotos:', e.message);
        refs = [];
      }

      for (const ref of refs.slice(0, 40)) { // tope para no tardar demasiado
        try {
          const buffer = await conTiempoLimite(descargarImagen(drive, ref.fileId), 12000);
          fotos.push({ area: ref.area, buffer });
        } catch(e) { console.error('[ReportePDF] foto:', ref.fileId, e.message); }
      }

      // Sin ninguna foto recuperada: en vez de devolver un PDF incompleto
      // sin avisar, el front pregunta si se genera igual.
      if (!fotos.length) {
        return res.status(409).json({
          ok: false, sinFotos: true,
          error: 'No se pudieron obtener las fotografias de este reporte.'
        });
      }
    }

    const pdf = await armarPDFReporte(info, fotos);
    const nombre = `${tipo}_${String(folio).replace(/\//g,'-')}.pdf`;
    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', `inline; filename="${nombre}"`);
    res.send(pdf);
  } catch(e) {
    console.error('[ReportePDF]', e.message);
    res.status(500).send('Error generando PDF: ' + e.message);
  }
});

// ── Reporte diario de taller (pestaña en Sheets + PDF + automatico) ──
require('./reporteDiario')(app, {
  getClients, PDFDocument, SHEET_ID, HOJAS_TALLER, HEADERS_TALLER,
  HOJA_EVIDENCIAS, LOGO_BASE64,
});

// ── Catch-all ──────────────────────────────────────────────────
app.get('*', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`TECSA en puerto ${PORT}`);
  console.log(`FOLDER_RAIZ: ${FOLDER_RAIZ}`);
  console.log(`FOLDER_AUDITORIAS: ${FOLDER_AUDITORIAS}`);
  console.log(`SHEET_ID: ${SHEET_ID}`);
});
