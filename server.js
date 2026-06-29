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

const SHEET_ID         = process.env.GOOGLE_SHEET_ID;
const PASS_BAJA        = process.env.PASSWORD_BAJA   || 'Tecsa2125';
const FOLDER_RAIZ      = process.env.FOLDER_RAIZ     || '14jjhGkt9Zq6T-RG4w3A2xMPqxlG3JhNY';
const FOLDER_AUDITORIAS = process.env.FOLDER_AUDITORIAS || '174I1SxpnD8dGbBjDRRRt1WqLE5FyVbcw';

// ── Autenticación Google (Sheets + Drive) ──────────────────────
function getAuth() {
  const credentials = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON);
  return new google.auth.GoogleAuth({
    credentials,
    scopes: [
      'https://www.googleapis.com/auth/spreadsheets',
      'https://www.googleapis.com/auth/drive',
    ],
  });
}
async function getClients() {
  const auth   = await getAuth().getClient();
  const sheets = google.sheets({ version: 'v4', auth });
  const drive  = google.drive({ version: 'v3', auth });
  return { sheets, drive };
}

// ── Zona horaria Monterrey ──────────────────────────────────────
function ahoraMty() {
  const ahora = new Date();
  const fecha = ahora.toLocaleDateString('es-MX',  { day:'2-digit', month:'2-digit', year:'numeric',  timeZone:'America/Monterrey' });
  const hora  = ahora.toLocaleTimeString('es-MX',  { hour:'2-digit', minute:'2-digit', hour12:false,  timeZone:'America/Monterrey' });
  return { fecha, hora };
}

// ── Folio único ─────────────────────────────────────────────────
// Usa la hoja Entradas para calcular el último folio del mes actual,
// igual que hacía Apps Script pero ahora en Node.
async function generarFolio(sheets, esTaller) {
  const ahora = new Date();
  const mes   = String(ahora.toLocaleDateString('es-MX', { month:'2-digit', timeZone:'America/Monterrey' })).padStart(2,'0');
  const anio  = ahora.toLocaleDateString('es-MX', { year:'2-digit', timeZone:'America/Monterrey' });
  const periodo = `${mes}/${anio}`;

  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SHEET_ID,
    range: 'Entradas!A2:A',
  });
  const rows = (res.data.values || []).flat();
  let max = 0;
  rows.forEach(f => {
    if (typeof f === 'string' && f.startsWith(periodo + '-')) {
      const num = parseInt(f.replace(periodo + '-', '').replace('T', ''), 10);
      if (!isNaN(num) && num > max) max = num;
    }
  });
  const siguiente = String(max + 1).padStart(3, '0');
  return `${periodo}-${esTaller ? 'T' : ''}${siguiente}`;
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
function piezasTexto(arr)   { return (arr||[]).map((p,i) => `${i+1}. ${p.nombre||''} — ${p.material||''}${p.obs?' ['+p.obs+']':''}`).join(' | '); }

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

async function generarPDFAuditoria(datos, fecha) {
  return new Promise((resolve, reject) => {
    const doc = new PDFDocument({ size: 'LETTER', margin: 36 });
    const chunks = [];
    doc.on('data', c => chunks.push(c));
    doc.on('end',  () => resolve(Buffer.concat(chunks)));
    doc.on('error', reject);

    const AZUL_OSC = '#16213e';
    const W = doc.page.width - 72; // ancho util

    // ── Título ──
    doc.fontSize(16).font('Helvetica-Bold').text('AUDITORIA GENERAL', { align: 'center' });
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

    function dibujarFila(cols, y, bg) {
      if (bg) { doc.rect(36, y, W, ROW_H).fill(bg).stroke(bg); }
      doc.rect(36, y, W, ROW_H).stroke('#aaaaaa');
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
