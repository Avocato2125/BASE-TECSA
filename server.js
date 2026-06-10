require('dotenv').config();
const express = require('express');
const { google } = require('googleapis');
const cors = require('cors');
const path = require('path');

const app = express();
app.use(express.json());
app.use(cors());
app.use(express.static(path.join(__dirname, 'public')));

const SHEET_ID  = process.env.GOOGLE_SHEET_ID;
const PASS_BAJA = process.env.PASSWORD_BAJA || 'tecsa2024';

function getAuth() {
  const credentials = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON);
  return new google.auth.GoogleAuth({
    credentials,
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
  });
}

async function getSheetsClient() {
  const auth = await getAuth().getClient();
  return google.sheets({ version: 'v4', auth });
}

function calcularDias(fechaStr, horaStr) {
  if (!fechaStr) return 0;
  const partes = fechaStr.trim().split('/');
  if (partes.length !== 3) return 0;
  const dia  = parseInt(partes[0], 10);
  const mes  = parseInt(partes[1], 10) - 1;
  const anio = parseInt(partes[2], 10);
  if (isNaN(dia) || isNaN(mes) || isNaN(anio)) return 0;
  const [hh, mm] = (horaStr || '00:00').split(':').map(Number);
  // Fecha y hora exacta de entrada en zona Monterrey (UTC-6)
  const entrada = new Date(Date.UTC(anio, mes, dia, hh + 6, mm));
  const ahora   = new Date();
  const diffMs  = ahora - entrada;
  if (diffMs < 0) return 0;
  return Math.floor(diffMs / (1000 * 60 * 60 * 24));
}

// Convierte "dd/MM/yyyy HH:mm" a timestamp UTC para el cronómetro
function fechaHoraATimestamp(fechaStr, horaStr) {
  if (!fechaStr || !horaStr) return null;
  const p = fechaStr.trim().split('/');
  if (p.length !== 3) return null;
  const [hh, mm] = horaStr.trim().split(':');
  // Monterrey es UTC-6 (sin horario de verano simplificado)
  const utc = Date.UTC(parseInt(p[2]), parseInt(p[1])-1, parseInt(p[0]), parseInt(hh)+6, parseInt(mm));
  return utc;
}

// ── API: Unidades activas ──────────────────────────────────────
app.get('/api/unidades', async (req, res) => {
  try {
    const sheets = await getSheetsClient();

    const [resEntradas, resTaller] = await Promise.all([
      sheets.spreadsheets.values.get({ spreadsheetId: SHEET_ID, range: 'Entradas!A2:I' }),
      sheets.spreadsheets.values.get({ spreadsheetId: SHEET_ID, range: 'Taller!A2:I'  }),
    ]);

    const rowsEntradas = resEntradas.data.values || [];
    const rowsTaller   = resTaller.data.values   || [];

    // Mapa unidad -> datos de taller activos (para cruzar)
    const tallerPorUnidad = {};
    rowsTaller.forEach((row, i) => {
      const estado = row[8] || '';
      const unidad = (row[3] || '').toString().toUpperCase().trim();
      if (estado === 'ACTIVO' && unidad) {
        tallerPorUnidad[unidad] = {
          folioTaller:  row[0] || '',
          planta:       row[5] || '',
          areaServicio: row[6] || '',
          reporteFalla: row[7] || '',
          rowIndexTaller: i + 2,
        };
      }
    });

    const activos = rowsEntradas
      .map((row, i) => {
        const unidad = (row[3] || '').toString().toUpperCase().trim();
        const taller = tallerPorUnidad[unidad] || null;
        return {
          rowIndex:     i + 2,
          folio:        row[0] || '',
          fecha:        row[1] || '',
          hora:         row[2] || '',
          unidad,
          operador:     row[4] || '',
          motivo:       row[5] || '',
          estado:       row[6] || '',
          dias:         calcularDias(row[1]),
          timestamp:    fechaHoraATimestamp(row[1], row[2]),
          sheet:        'Entradas',
          // Datos de taller cruzados (null si no aplica)
          folioTaller:  taller ? taller.folioTaller  : null,
          planta:       taller ? taller.planta        : null,
          areaServicio: taller ? taller.areaServicio  : null,
          reporteFalla: taller ? taller.reporteFalla  : null,
          rowIndexTaller: taller ? taller.rowIndexTaller : null,
        };
      })
      .filter(r => r.estado === 'ACTIVO');

    res.json({ ok: true, entradas: activos });
  } catch (e) {
    console.error(e);
    res.json({ ok: false, error: e.message });
  }
});

// ── API: Dar de baja ───────────────────────────────────────────
app.post('/api/baja', async (req, res) => {
  const { rowIndex, rowIndexTaller, password } = req.body;

  if (password !== PASS_BAJA) {
    return res.json({ ok: false, error: 'Contrasena incorrecta' });
  }

  try {
    const sheets = await getSheetsClient();
    const ahora = new Date();
    const fechaSalida = ahora.toLocaleDateString('es-MX', { day:'2-digit', month:'2-digit', year:'numeric', timeZone:'America/Monterrey' });
    const horaSalida  = ahora.toLocaleTimeString('es-MX', { hour:'2-digit', minute:'2-digit', hour12:false, timeZone:'America/Monterrey' });

    // Baja en Entradas
    await sheets.spreadsheets.values.update({
      spreadsheetId: SHEET_ID,
      range: `Entradas!G${rowIndex}:I${rowIndex}`,
      valueInputOption: 'RAW',
      requestBody: { values: [['BAJA', fechaSalida, horaSalida]] }
    });

    // Baja en Taller si existe registro cruzado
    if (rowIndexTaller) {
      await sheets.spreadsheets.values.update({
        spreadsheetId: SHEET_ID,
        range: `Taller!I${rowIndexTaller}`,
        valueInputOption: 'RAW',
        requestBody: { values: [['BAJA']] }
      });
    }

    res.json({ ok: true });
  } catch (e) {
    console.error(e);
    res.json({ ok: false, error: e.message });
  }
});

app.get('*', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`TECSA Dashboard en puerto ${PORT}`));
