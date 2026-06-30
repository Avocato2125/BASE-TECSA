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
  const hora  = ahora.toLocaleTimeString('es-MX',  { hour:'2-digit', minute:'2-digit', hour12:false,  timeZone:'America/Monterrey' });
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

// ── Drive: obtener o crear carpeta ──────────────────────────────
async function getOCrearCarpeta(drive, padreId, nombre) {
  if (!padreId) throw new Error('getOCrearCarpeta: padreId es undefined para carpeta "' + nombre + '"');
  const q = `'${padreId}' in parents and name='${nombre}' and mimeType='application/vnd.google-apps.folder' and trashed=false`;
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
async function subirArchivo(drive, carpetaId, nombre, base64, mimeType) {
  if (!carpetaId) throw new Error('subirArchivo: carpetaId es undefined para archivo "' + nombre + '"');
  console.log(`Subiendo "${nombre}" a carpeta ${carpetaId}`);
  const buffer = Buffer.from(base64, 'base64');
  const stream = Readable.from(buffer);
  const res = await drive.files.create({
    requestBody: { name: nombre, parents: [carpetaId] },
    media: { mimeType, body: stream },
    fields: 'id,webViewLink',
    supportsAllDrives: true,
  });
  console.log(`Archivo subido: ${res.data.id}`);
  return res.data;
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

    const tallerMap = {};
    rowsT.forEach((row, i) => {
      const unidad = (row[3] || '').toString().toUpperCase().trim();
      if ((row[8] || '') === 'ACTIVO' && unidad) {
        tallerMap[unidad] = {
          folioTaller: row[0] || '', planta: row[5] || '',
          areaServicio: row[6] || '', reporteFalla: row[7] || '',
          rowIndexTaller: i + 2,
        };
      }
    });

    const activos = rowsE.map((row, i) => {
      const unidad = (row[3] || '').toString().toUpperCase().trim();
      const t = tallerMap[unidad] || null;
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
  const { rowIndex, rowIndexTaller, password } = req.body;
  if (password !== PASS_BAJA) return res.json({ ok: false, error: 'Contrasena incorrecta' });
  try {
    const { sheets } = await getClients();
    const { fecha, hora } = ahoraMty();
    await sheets.spreadsheets.values.update({
      spreadsheetId: SHEET_ID, range: `Entradas!G${rowIndex}:I${rowIndex}`,
      valueInputOption: 'RAW', requestBody: { values: [['BAJA', fecha, hora]] }
    });
    if (rowIndexTaller) {
      await sheets.spreadsheets.values.update({
        spreadsheetId: SHEET_ID, range: `Taller!I${rowIndexTaller}`,
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
    const unidad   = String(datos.unidad).toUpperCase().trim();
    const operador = String(datos.operador).trim();
    const motivo   = datos.motivo;
    const esTaller = motivo === 'Taller';

    let motivoTexto = motivo;
    if (motivo === 'Camaras')  motivoTexto = `Camaras/Display — ${datos.planta || ''} — ${datos.reporteFalla || ''}`;
    else if (motivo === 'Inplant') motivoTexto = `Inplant — ${datos.planta || ''} — ${datos.detalle || ''}`;
    else if (motivo === 'Otro' && datos.detalle) motivoTexto = `Otro: ${datos.detalle}`;

    const folio = await generarFolio(sheets, esTaller);

    // Guardar en Entradas
    await sheets.spreadsheets.values.append({
      spreadsheetId: SHEET_ID, range: 'Entradas!A:I',
      valueInputOption: 'RAW', insertDataOption: 'INSERT_ROWS',
      requestBody: { values: [[folio, fecha, hora, unidad, operador, motivoTexto, 'ACTIVO', '', '']] }
    });

    // Si es Taller, también guardar en hoja Taller
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
const HOJAS_TALLER = {
  mecanico:  'Reporte Taller Mecanico',
  electrico: 'Reporte Electrico',
  imagen:    'Reporte Imagen',
  llantas:   'Reporte Llantas',
};
const HEADERS_TALLER = {
  mecanico:  ['Folio','Fecha','Hora','Unidad','Operador','Planta','Area Servicio','Mecanico','Aceite-Km','Aceite-CapTeorica','Aceite-LitAnt','Aceite-NivelBajo','Aceite-LitNuevo','Aceite-Obs','Frenos-Obs','Engrasado','Engrasado-Obs','SrvAceite-Km','SrvAceite-Cap','SrvAceite-LitAnt','SrvAceite-NivelBajo','SrvAceite-LitNuevo','SrvAceite-Obs','Filtro-Aire','FiltroAire-Obs','Filtro-Diesel','FiltroDiesel-Obs','Filtro-Aceite','FiltroAceite-Obs','Filtro-Separador','FiltroSep-Obs','Afinacion-Km','Afinacion-Piezas','Afinacion-Mat','Afinacion-Obs','Piezas-Taller','Obs-Taller'],
  electrico: ['Folio','Fecha','Hora','Unidad','Operador','Planta','Area Servicio','Mecanico','Carga-Bat-VoltAnt','Carga-Bat-VoltNuevo','Carga-Bat-Obs','Cambio-Bat-Motivo','Cambio-Bat-Obs','Piezas-Electrico','Obs-Electrico'],
  imagen:    ['Folio','Fecha','Hora','Unidad','Operador','Planta','Area Servicio','Mecanico','Calcas-Mat','Calcas-Obs','Asiento-Mat','Asiento-Obs','Pintura-Area','Pintura-Mat','Pintura-Obs','Soldadura-Mat','Soldadura-Obs','Piezas-Imagen','Obs-Imagen'],
  llantas:   ['Folio','Fecha','Hora','Unidad','Operador','Planta','Area Servicio','Mecanico','Llanta-Marca','Llanta-Obs','LlantaRep-Vida','LlantaRep-Obs','Obs-Llantas'],
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
    const t  = datos.taller    || {};
    const el = datos.electrico || {};
    const im = datos.imagen    || {};
    const ll = datos.llantas   || {};

    // ── 1. Subir fotos a Drive ────────────────────────────────
    const carpRaiz   = await getOCrearCarpeta(drive, FOLDER_RAIZ, `Unidad-${datos.unidad}`);
    const fechaCarp  = fecha.replace(/\//g, '-');

    async function subirAreaFotos(area, items) {
      if (!items || !items.length) return;
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
          await subirArchivo(drive, carpFecha, f.nombre||'foto.jpg', f.base64, 'image/jpeg').catch(()=>{});
        }
      }
    }

    await Promise.allSettled([
      t.aceite    && subirAreaFotos('Aceite',           [t.aceite]),
      t.frenos    && subirAreaFotos('Frenos',           [t.frenos]),
      t.servicio?.engrasado && subirAreaFotos('Engrasado', [t.servicio.engrasado]),
      t.servicio?.aceite    && subirAreaFotos('Aceite-Servicio', [t.servicio.aceite]),
      t.afinacion && subirAreaFotos('Afinacion',        [t.afinacion]),
      el.cargaBat  && subirAreaFotos('Carga-Bateria',   [el.cargaBat]),
      el.cambioBat && subirAreaFotos('Cambio-Bateria',  [el.cambioBat]),
      im.calcas    && subirAreaFotos('Calcas',          [im.calcas]),
      im.asiento   && subirAreaFotos('Asiento',         [im.asiento]),
      im.pintura   && subirAreaFotos('Pintura',         [im.pintura]),
      im.soldadura && subirAreaFotos('Soldadura',       [im.soldadura]),
      ll.cambio    && subirAreaFotos('Llanta-Cambio',   [ll.cambio]),
      ll.reparacion && subirAreaFotos('Llanta-Reparacion', [ll.reparacion]),
    ]);

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
      piezasTexto(t.piezas), t.obsGral||'',
    ];
    const rowElec = [...comun, v(el.cargaBat,'voltAnt'), v(el.cargaBat,'voltNuevo'), v(el.cargaBat,'obs'), v(el.cambioBat,'motivo'), v(el.cambioBat,'obs'), piezasTexto(el.piezas), el.obsGral||''];
    const rowImg  = [...comun, v(im.calcas,'material'), v(im.calcas,'obs'), v(im.asiento,'material'), v(im.asiento,'obs'), v(im.pintura,'area'), v(im.pintura,'material'), v(im.pintura,'obs'), v(im.soldadura,'material'), v(im.soldadura,'obs'), piezasTexto(im.piezas), im.obsGral||''];
    const rowLl   = [...comun, v(ll.cambio,'marca'), v(ll.cambio,'obs'), v(ll.reparacion,'vidaRestante'), v(ll.reparacion,'obs'), ll.obsGral||''];

    const tieneContenido = row => row.slice(8).some(c => c !== '' && c != null);

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
const LOGO_BASE64 = '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';

async function generarPDFAuditoria(datos, fecha) {
  return new Promise((resolve, reject) => {
    const doc = new PDFDocument({ size: 'LETTER', margin: 36 });
    const chunks = [];
    doc.on('data', c => chunks.push(c));
    doc.on('end',  () => resolve(Buffer.concat(chunks)));
    doc.on('error', reject);

    const AZUL_OSC = '#16213e';
    const W = doc.page.width - 72; // ancho util

    // ── Logo + título alineados ──
    try {
      const logoBuf = Buffer.from(LOGO_BASE64, 'base64');
      doc.image(logoBuf, 36, 30, { width: 90, height: 36 });
    } catch(e) { console.error('Error insertando logo:', e.message); }
    doc.fontSize(16).font('Helvetica-Bold').text('AUDITORIA GENERAL', 0, 36, { align: 'center' });
    doc.y = 75;
    doc.moveDown(0.3);

    // ── Header: datos generales ──
    doc.fontSize(9).font('Helvetica-Bold').text('Planta: ', { continued: true })
       .font('Helvetica').text(datos.planta || '', { continued: true })
       .font('Helvetica-Bold').text('    Fecha: ', { continued: true })
       .font('Helvetica').text(fecha);
    doc.font('Helvetica-Bold').text('Unidad: ', { continued: true })
       .font('Helvetica').text(String(datos.unidad||''), { continued: true })
       .font('Helvetica-Bold').text('    Nombre del operador: ', { continued: true })
       .font('Helvetica').text(datos.operadorNombreAud || datos.operador || '');
    doc.font('Helvetica-Bold').text('Kilometraje: ', { continued: true })
       .font('Helvetica').text(String(datos.kilometraje||''), { continued: true })
       .font('Helvetica-Bold').text('    Auditor: ', { continued: true })
       .font('Helvetica').text(datos.auditor || '');
    doc.moveDown(0.5);

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
        doc.fillColor(bg && bg !== '#ffffff' ? '#ffffff' : '#000000')
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
    const yFirma = doc.y;
    const mitad  = 36 + W/2;

    // Firma auditor
    if (datos.firmaAuditor) {
      try {
        const bufA = Buffer.from(datos.firmaAuditor, 'base64');
        doc.image(bufA, 36, yFirma, { width: 180, height: 72 });
      } catch(e) {}
    }
    doc.fontSize(9).font('Helvetica-Bold').text('Nombre y Firma del Auditor', 36, yFirma + 78, { width: W/2, align: 'center' });
    doc.font('Helvetica').text(datos.auditor || '', 36, yFirma + 90, { width: W/2, align: 'center' });

    // Firma operador
    if (datos.firmaOperador) {
      try {
        const bufO = Buffer.from(datos.firmaOperador, 'base64');
        doc.image(bufO, mitad, yFirma, { width: 180, height: 72 });
      } catch(e) {}
    }
    doc.font('Helvetica-Bold').text('Nombre y Firma del Operador', mitad, yFirma + 78, { width: W/2, align: 'center' });
    doc.font('Helvetica').text(datos.operadorNombreAud || datos.operador || '', mitad, yFirma + 90, { width: W/2, align: 'center' });

    doc.end();
  });
}

app.post('/api/auditoria', async (req, res) => {
  try {
    const datos = req.body;
    const { sheets, drive } = await getClients();
    const { fecha, hora } = ahoraMty();
    const fechaCarp = fecha.replace(/\//g, '-');

    // ── 1. Carpeta Unidad dentro de FOLDER_AUDITORIAS ────────
    const carpUnidad = await getOCrearCarpeta(drive, FOLDER_AUDITORIAS, `Unidad-${datos.unidad}`);
    const carpFecha  = await getOCrearCarpeta(drive, carpUnidad, fechaCarp);

    // ── 2. Subir fotos de cada punto ─────────────────────────
    const subirFotos = (datos.puntos || []).map(async p => {
      if (!p.fotos || !p.fotos.length) return;
      const nombreBase = (p.nombre||'punto').replace(/[\/\\:*?"<>|]/g, '-');
      for (let i = 0; i < p.fotos.length; i++) {
        if (!p.fotos[i]?.base64) continue;
        const nombre = i === 0 ? `${nombreBase}.jpg` : `${nombreBase}-${i+1}.jpg`;
        await subirArchivo(drive, carpFecha, nombre, p.fotos[i].base64, 'image/jpeg').catch(()=>{});
      }
    });
    await Promise.allSettled(subirFotos);

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

    // Agregar fila con fórmula HYPERLINK para PDF
    const lastRow = (await sheets.spreadsheets.values.get({ spreadsheetId: SHEET_ID, range: 'Auditorias!A:A' })).data.values?.length || 1;
    await sheets.spreadsheets.values.append({
      spreadsheetId: SHEET_ID, range: 'Auditorias!A:A',
      valueInputOption: 'USER_ENTERED', insertDataOption: 'INSERT_ROWS',
      requestBody: { values: [[...rowBase, ...rowPuntos, `=HYPERLINK("${pdfUrl}","PDF")`]] }
    });

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
