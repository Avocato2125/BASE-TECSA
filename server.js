require('dotenv').config();
const express = require('express');
const { google } = require('googleapis');
const cors = require('cors');
const path = require('path');

const app = express();
app.use(express.json());
app.use(cors());
app.use(express.static(path.join(__dirname, 'public')));

// ── Configuración Google Sheets ──────────────────────────────
const SHEET_ID = process.env.GOOGLE_SHEET_ID;
const PASS_BAJA = process.env.PASSWORD_BAJA || 'tecsa2024';

// Autenticación con Service Account
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

// ── Helpers ──────────────────────────────────────────────────
function calcularDias(fechaStr) {
  if (!fechaStr) return 0;
  // Formato esperado: dd/MM/yyyy
  const partes = fechaStr.trim().split('/');
  if (partes.length !== 3) return 0;
  const dia  = parseInt(partes[0], 10);
  const mes  = parseInt(partes[1], 10) - 1; // 0-indexed
  const anio = parseInt(partes[2], 10);
  if (isNaN(dia) || isNaN(mes) || isNaN(anio)) return 0;
  const fecha = new Date(anio, mes, dia);
  const hoy = new Date();
  hoy.setHours(0, 0, 0, 0);
  fecha.setHours(0, 0, 0, 0);
  const diff = Math.floor((hoy - fecha) / (1000 * 60 * 60 * 24));
  return Math.max(0, diff);
}

// ── API: Obtener unidades activas ────────────────────────────
app.get('/api/unidades', async (req, res) => {
  try {
    const sheets = await getSheetsClient();

    // Leer hoja Entradas
    const entradas = await sheets.spreadsheets.values.get({
      spreadsheetId: SHEET_ID,
      range: 'Entradas!A2:I',
    });

    // Leer hoja Taller
    const taller = await sheets.spreadsheets.values.get({
      spreadsheetId: SHEET_ID,
      range: 'Taller!A2:I',
    });

    const rowsEntradas = entradas.data.values || [];
    const rowsTaller = taller.data.values || [];

    // Filtrar solo ACTIVOS de Entradas
    const activos = rowsEntradas
      .map((row, i) => ({
        rowIndex: i + 2, // 1-indexed + header
        folio: row[0] || '',
        fecha: row[1] || '',
        hora: row[2] || '',
        unidad: row[3] || '',
        operador: row[4] || '',
        motivo: row[5] || '',
        estado: row[6] || '',
        fechaSalida: row[7] || '',
        horaSalida: row[8] || '',
        dias: calcularDias(row[1]),
        sheet: 'Entradas'
      }))
      .filter(r => r.estado === 'ACTIVO');

    // Filtrar ACTIVOS de Taller
    const activosTaller = rowsTaller
      .map((row, i) => ({
        rowIndex: i + 2,
        folio: row[0] || '',
        fecha: row[1] || '',
        hora: row[2] || '',
        unidad: row[3] || '',
        operador: row[4] || '',
        planta: row[5] || '',
        areaServicio: row[6] || '',
        reporteFalla: row[7] || '',
        estado: row[8] || '',
        dias: calcularDias(row[1]),
        sheet: 'Taller'
      }))
      .filter(r => r.estado === 'ACTIVO');

    res.json({ ok: true, entradas: activos, taller: activosTaller });
  } catch (e) {
    console.error(e);
    res.json({ ok: false, error: e.message });
  }
});

// ── API: Dar de baja una unidad ──────────────────────────────
app.post('/api/baja', async (req, res) => {
  const { folio, sheet, rowIndex, password } = req.body;

  if (password !== PASS_BAJA) {
    return res.json({ ok: false, error: 'Contraseña incorrecta' });
  }

  try {
    const sheets = await getSheetsClient();
    const ahora = new Date();
    const fechaSalida = ahora.toLocaleDateString('es-MX', {
      day: '2-digit', month: '2-digit', year: 'numeric',
      timeZone: 'America/Monterrey'
    });
    const horaSalida = ahora.toLocaleTimeString('es-MX', {
      hour: '2-digit', minute: '2-digit', hour12: false,
      timeZone: 'America/Monterrey'
    });

    if (sheet === 'Entradas') {
      // Actualizar columnas G (estado), H (fechaSalida), I (horaSalida)
      await sheets.spreadsheets.values.update({
        spreadsheetId: SHEET_ID,
        range: `Entradas!G${rowIndex}:I${rowIndex}`,
        valueInputOption: 'RAW',
        requestBody: { values: [['BAJA', fechaSalida, horaSalida]] }
      });
    } else if (sheet === 'Taller') {
      // Actualizar columna I (estado)
      await sheets.spreadsheets.values.update({
        spreadsheetId: SHEET_ID,
        range: `Taller!I${rowIndex}`,
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

// ── Servir frontend ──────────────────────────────────────────
app.get('*', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`TECSA Dashboard corriendo en puerto ${PORT}`));
