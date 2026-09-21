// ═══════════════════════════════════════════════════════════════════
// REPORTE DIARIO DE TALLER — TECSA
//
// Arma el resumen de un dia a partir de las hojas que ya existen:
//   · Reportes finalizados ese dia (con tiempo de reparacion y servicios)
//   · Unidades que siguen en taller al cierre del dia
//   · Impacto por planta
//   · Consumo del dia: litros y piezas
//
// Salidas:
//   GET  /api/reporte-diario?fecha=YYYY-MM-DD         → JSON (vista previa)
//   GET  /api/reporte-diario-pdf?fecha=YYYY-MM-DD     → PDF descargable
//   POST /api/reporte-diario/sheets?fecha=YYYY-MM-DD  → escribe la pestaña
//   + generacion automatica diaria de la pestaña (hora configurable)
//
// Todo se reconstruye desde las hojas, asi que se puede regenerar
// cualquier dia pasado y siempre da el mismo resultado.
// ═══════════════════════════════════════════════════════════════════

module.exports = function registrarReporteDiario(app, deps) {
  const { getClients, PDFDocument, SHEET_ID, HOJAS_TALLER, HEADERS_TALLER,
          HOJA_EVIDENCIAS, LOGO_BASE64 } = deps;

  // Pestañas en un spreadsheet aparte si se configura (recomendado),
  // para no llenar el Sheet principal con una pestaña por dia.
  const SHEET_ID_DIARIO = process.env.SHEET_ID_DIARIO || SHEET_ID;
  const HORA_AUTO       = process.env.REPORTE_DIARIO_HORA || '19:00';

  // ── Fechas y horas (hora de Monterrey) ─────────────────────────
  function partesMty(d = new Date()) {
    const p = {};
    new Intl.DateTimeFormat('en-GB', {
      timeZone: 'America/Monterrey', year:'numeric', month:'2-digit', day:'2-digit',
      hour:'2-digit', minute:'2-digit', hour12:false,
    }).formatToParts(d).forEach(x => { p[x.type] = x.value; });
    return { d:+p.day, m:+p.month, y:+p.year, hh:+p.hour % 24, mm:+p.minute };
  }
  const pad = n => String(n).padStart(2, '0');
  const fechaStr = ({ d, m, y }) => `${pad(d)}/${pad(m)}/${y}`;

  // "dd/MM/yyyy" (tolera d/m/yyyy) → {d,m,y} o null
  function parseFecha(s) {
    const t = String(s || '').trim().split('/');
    if (t.length !== 3) return null;
    const r = { d: parseInt(t[0], 10), m: parseInt(t[1], 10), y: parseInt(t[2], 10) };
    return (isNaN(r.d) || isNaN(r.m) || isNaN(r.y)) ? null : r;
  }
  const claveDia = f => f ? f.y * 10000 + f.m * 100 + f.d : 0;

  // Milisegundos "naive" (misma base para todo, sin zona) para restar tiempos
  function aMs(fecha, hora) {
    const f = parseFecha(fecha);
    if (!f) return null;
    const [hh, mm] = String(hora || '00:00').split(':').map(n => parseInt(n, 10) || 0);
    return Date.UTC(f.y, f.m - 1, f.d, hh, mm);
  }
  function duracion(ms) {
    if (ms == null || ms < 0) return '';
    const min = Math.round(ms / 60000);
    const d = Math.floor(min / 1440), h = Math.floor((min % 1440) / 60), m = min % 60;
    if (d) return `${d} d ${h} h`;
    if (h) return `${h} h ${m} min`;
    return `${m} min`;
  }

  // "YYYY-MM-DD" del query → "dd/MM/yyyy"; sin parametro = hoy
  function fechaDeQuery(q) {
    const m = String(q || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (m) return `${m[3]}/${m[2]}/${m[1]}`;
    return fechaStr(partesMty());
  }

  // ── Utilidades de filas ────────────────────────────────────────
  const idx = (hoja, nombre) => HEADERS_TALLER[hoja].indexOf(nombre);
  const celda = (fila, hoja, nombre) => {
    const i = idx(hoja, nombre);
    return i < 0 ? '' : String(fila[i] == null ? '' : fila[i]).trim();
  };
  const esSi = v => /^si\b/i.test(String(v || '').trim());
  const num  = v => { const n = parseFloat(String(v || '').replace(',', '.')); return isNaN(n) ? 0 : n; };

  // Texto de piezas → [{pieza, material}]
  // Formato actual: "Nombre de la pz: X | Material us: Y | Obs: Z // ..."
  // Formato viejo:  "1. X — Y [obs] | 2. ..."
  function parsePiezas(txt) {
    const s = String(txt || '').trim();
    if (!s) return [];
    if (/Nombre de la pz:/i.test(s)) {
      return s.split('//').map(bloque => {
        const nom = (bloque.match(/Nombre de la pz:\s*([^|]*)/i) || [])[1] || '';
        const mat = (bloque.match(/Material us:\s*([^|]*)/i) || [])[1] || '';
        return { pieza: nom.trim(), material: mat.trim() };
      }).filter(p => p.pieza || p.material);
    }
    return s.split('|').map(x => {
      const limpio = x.replace(/^\s*\d+\.\s*/, '').trim();
      const [pieza, material] = limpio.split('—').map(t => (t || '').replace(/\[.*\]/, '').trim());
      return { pieza: pieza || limpio, material: material || '' };
    }).filter(p => p.pieza);
  }

  // ── Lectura de todas las hojas en una sola llamada ─────────────
  async function leerHojas(sheets) {
    const meta = await sheets.spreadsheets.get({
      spreadsheetId: SHEET_ID, fields: 'sheets.properties.title',
    });
    const existentes = new Set((meta.data.sheets || []).map(s => s.properties.title));

    const pedidos = [
      ['entradas', 'Entradas'],
      ['taller',   'Taller'],
      ...Object.entries(HOJAS_TALLER),
      ['evidencias', HOJA_EVIDENCIAS],
    ].filter(([, titulo]) => existentes.has(titulo));

    const res = await sheets.spreadsheets.values.batchGet({
      spreadsheetId: SHEET_ID,
      ranges: pedidos.map(([, titulo]) => `'${titulo}'!A2:BZ`),
    });
    const datos = {};
    pedidos.forEach(([clave], i) => { datos[clave] = res.data.valueRanges[i].values || []; });
    return datos;
  }

  // ── Construccion del reporte de un dia ─────────────────────────
  async function construirReporte(sheets, fechaDia) {
    const H = await leerHojas(sheets);
    const dia = parseFecha(fechaDia);
    const kDia = claveDia(dia);

    // Si el dia es hoy, el "cierre" es ahora; si es pasado, las 23:59
    const hoy = partesMty();
    const esHoy = kDia === claveDia(hoy);
    const msCierre = esHoy
      ? Date.UTC(hoy.y, hoy.m - 1, hoy.d, hoy.hh, hoy.mm)
      : Date.UTC(dia.y, dia.m - 1, dia.d, 23, 59);

    // Entradas por folio (para la fecha/hora de salida)
    const entradas = {};
    (H.entradas || []).forEach(f => {
      const folio = String(f[0] || '').trim();
      if (folio) entradas[folio] = { estado: f[6] || '', fSal: f[7] || '', hSal: f[8] || '' };
    });

    // Taller por folio (entrada, planta, area, reporte de falla)
    const taller = {};
    (H.taller || []).forEach(f => {
      const folio = String(f[0] || '').trim();
      if (!folio) return;
      taller[folio] = {
        folio, fEnt: f[1] || '', hEnt: f[2] || '', unidad: f[3] || '', operador: f[4] || '',
        planta: f[5] || '', area: f[6] || '', falla: f[7] || '', estado: f[8] || '',
      };
    });

    // Evidencias con area "Frenos" (el apartado de frenos solo tiene
    // observaciones; la foto obligatoria es la mejor señal de que se hizo)
    const frenosPorFolio = new Set();
    (H.evidencias || []).forEach(f => {
      if (String(f[2] || '').trim() === 'Frenos') frenosPorFolio.add(String(f[0] || '').trim());
    });

    // ── Recorrer las 6 hojas de reportes del dia ──
    const fin = {};         // folio → datos del finalizado
    const liquidos = {};    // concepto → { litros, unidades:Set }
    const piezas = [];      // { folio, unidad, area, pieza, material }

    const sumarLitros = (concepto, litros, unidad) => {
      if (!litros) return;
      liquidos[concepto] = liquidos[concepto] || { litros: 0, unidades: new Set() };
      liquidos[concepto].litros += litros;
      if (unidad) liquidos[concepto].unidades.add(unidad);
    };

    Object.keys(HOJAS_TALLER).forEach(hoja => {
      (H[hoja] || []).forEach(fila => {
        const fechaRep = celda(fila, hoja, 'Fecha');
        if (claveDia(parseFecha(fechaRep)) !== kDia) return;

        const folio  = celda(fila, hoja, 'Folio');
        const unidad = celda(fila, hoja, 'Unidad');
        if (!folio) return;

        const r = fin[folio] = fin[folio] || {
          folio, fecha: fechaRep, hora: celda(fila, hoja, 'Hora'),
          unidad, operador: celda(fila, hoja, 'Operador'),
          planta: celda(fila, hoja, 'Planta'), area: celda(fila, hoja, 'Area Servicio'),
          mecanicos: new Set(), servicios: [],
        };
        const mec = celda(fila, hoja, 'Mecanico');
        if (mec) r.mecanicos.add(mec);

        const serv = [];
        const agregarPiezas = (col, area) => parsePiezas(celda(fila, hoja, col))
          .forEach(p => piezas.push({ folio, unidad, area, ...p }));

        if (hoja === 'mecanico') {
          if (celda(fila, hoja, 'Aceite-Km') || celda(fila, hoja, 'Aceite-LitNuevo')) {
            serv.push('Cambio de aceite');
            sumarLitros('Aceite (cambio de aceite)', num(celda(fila, hoja, 'Aceite-LitNuevo')), unidad);
          }
          if (celda(fila, hoja, 'Frenos-Obs') || frenosPorFolio.has(folio)) serv.push('Frenos');
          if (esSi(celda(fila, hoja, 'Engrasado'))) serv.push('Engrasado');
          if (celda(fila, hoja, 'SrvAceite-Km') || celda(fila, hoja, 'SrvAceite-LitNuevo')) {
            serv.push('Servicio de aceite');
            sumarLitros('Aceite (servicio)', num(celda(fila, hoja, 'SrvAceite-LitNuevo')), unidad);
          }
          [['Filtro-Aire','Filtro de aire'], ['Filtro-Diesel','Filtro de diesel'],
           ['Filtro-Aceite','Filtro de aceite'], ['Filtro-Separador','Filtro separador de agua']]
            .forEach(([col, nombre]) => {
              if (esSi(celda(fila, hoja, col))) {
                serv.push(nombre);
                piezas.push({ folio, unidad, area: 'Mecanico', pieza: nombre, material: '' });
              }
            });
          if (celda(fila, hoja, 'Afinacion-Km')) serv.push('Afinacion');
          if (celda(fila, hoja, 'Piezas-Taller')) { serv.push('Cambio de piezas'); agregarPiezas('Piezas-Taller', 'Mecanico'); }
          const aj = celda(fila, hoja, 'Ajuste');
          if (aj === 'Si' || aj.startsWith('Si —')) serv.push('Ajuste');
        }
        else if (hoja === 'electrico') {
          if (celda(fila, hoja, 'Carga-Bat-VoltAnt') || celda(fila, hoja, 'Carga-Bat-VoltNuevo')) serv.push('Carga de bateria');
          if (celda(fila, hoja, 'Cambio-Bat-Motivo')) {
            serv.push('Cambio de bateria');
            piezas.push({ folio, unidad, area: 'Electrico', pieza: 'Bateria', material: celda(fila, hoja, 'Cambio-Bat-Motivo') });
          }
          if (celda(fila, hoja, 'Piezas-Electrico')) { serv.push('Cambio de piezas (electrico)'); agregarPiezas('Piezas-Electrico', 'Electrico'); }
        }
        else if (hoja === 'imagen') {
          if (celda(fila, hoja, 'Calcas-Mat'))    serv.push('Calcas');
          if (celda(fila, hoja, 'Asiento-Mat'))   serv.push('Asiento');
          if (celda(fila, hoja, 'Pintura-Area') || celda(fila, hoja, 'Pintura-Mat')) serv.push('Pintura');
          if (celda(fila, hoja, 'Soldadura-Mat')) serv.push('Soldadura');
          if (celda(fila, hoja, 'Piezas-Imagen')) { serv.push('Cambio de piezas (imagen)'); agregarPiezas('Piezas-Imagen', 'Imagen'); }
        }
        else if (hoja === 'llantas') {
          if (celda(fila, hoja, 'Llanta-Marca')) {
            serv.push('Cambio de llanta');
            piezas.push({ folio, unidad, area: 'Llantas', pieza: 'Llanta', material: celda(fila, hoja, 'Llanta-Marca') });
          }
          if (celda(fila, hoja, 'LlantaRep-Vida')) serv.push('Reparacion de llanta');
        }
        else if (hoja === 'llenado') {
          [['Aceite','Aceite (rellenado)'], ['Adblue','AdBlue'], ['LiqFrenos','Liquido de frenos'],
           ['Anticong','Anticongelante'], ['Gasolina','Gasolina'], ['Direccion','Liquido de direccion hidraulica']]
            .forEach(([pref, nombre]) => {
              const litros = num(celda(fila, hoja, `${pref}-LitPuestos`));
              if (litros) { serv.push(`Rellenado: ${nombre}`); sumarLitros(nombre, litros, unidad); }
            });
        }
        else if (hoja === 'suspension') {
          if (esSi(celda(fila, hoja, 'Muelles'))) {
            const pos   = celda(fila, hoja, 'Muelles-Posicion');
            const hojas = celda(fila, hoja, 'Muelles-Hojas');
            serv.push('Cambio de muelles' + (pos ? ` (${pos})` : ''));
            piezas.push({ folio, unidad, area: 'Suspension',
              pieza: 'Muelle' + (pos ? ` ${pos.toLowerCase()}` : ''), material: hojas ? `Hojas: ${hojas}` : '' });
          }
          if (esSi(celda(fila, hoja, 'Amortiguadores'))) {
            serv.push('Amortiguadores');
            piezas.push({ folio, unidad, area: 'Suspension', pieza: 'Amortiguadores', material: '' });
          }
          if (celda(fila, hoja, 'Piezas-Suspension')) { serv.push('Cambio de piezas (suspension)'); agregarPiezas('Piezas-Suspension', 'Suspension'); }
        }

        // Si no se detecto nada concreto (solo observaciones), al menos el area
        const nombres = { mecanico:'Mecanico', electrico:'Electrico', imagen:'Imagen',
                          llantas:'Llantas', llenado:'Rellenados', suspension:'Suspension' };
        r.servicios.push(...(serv.length ? serv : [`${nombres[hoja]} (ver observaciones)`]));
      });
    });

    // Salidas del dia sin reporte de taller (baja directa desde el dashboard)
    Object.values(taller).forEach(t => {
      const e = entradas[t.folio];
      if (!e || fin[t.folio]) return;
      if (claveDia(parseFecha(e.fSal)) !== kDia) return;
      fin[t.folio] = {
        folio: t.folio, fecha: e.fSal, hora: e.hSal, unidad: t.unidad, operador: t.operador,
        planta: t.planta, area: t.area, mecanicos: new Set(),
        servicios: ['Salida registrada sin reporte de taller'],
      };
    });

    // ── Finalizados: completar entrada, salida y tiempo ──
    const finalizados = Object.values(fin).map(r => {
      const t = taller[r.folio] || {};
      const e = entradas[r.folio] || {};
      const fSal = e.fSal || r.fecha, hSal = e.hSal || r.hora;
      const msEnt = aMs(t.fEnt, t.hEnt), msSal = aMs(fSal, hSal);
      const ms = (msEnt != null && msSal != null) ? msSal - msEnt : null;
      return {
        folio: r.folio, fecha: r.fecha, hora: r.hora,
        unidad: r.unidad || t.unidad || '', operador: r.operador || t.operador || '',
        planta: r.planta || t.planta || '', areaServicio: r.area || t.area || '',
        fechaEntrada: t.fEnt ? `${t.fEnt} ${t.hEnt || ''}`.trim() : '',
        fechaSalida:  fSal ? `${fSal} ${hSal || ''}`.trim() : '',
        tiempo: duracion(ms), horas: (ms != null && ms >= 0) ? ms / 3600000 : null,
        mecanico: [...r.mecanicos].join(', '),
        reporteFalla: t.falla || '',
        servicios: [...new Set(r.servicios)].join(', '),
      };
    }).sort((a, b) => (aMs(a.fecha, a.hora) || 0) - (aMs(b.fecha, b.hora) || 0));

    // ── Activos al cierre: entraron a mas tardar ese dia y no habian salido ──
    const activos = Object.values(taller).filter(t => {
      const kEnt = claveDia(parseFecha(t.fEnt));
      if (!kEnt || kEnt > kDia) return false;
      const e = entradas[t.folio];
      if (e) {
        const kSal = claveDia(parseFecha(e.fSal));
        return !kSal || kSal > kDia;
      }
      return t.estado === 'ACTIVO';
    }).map(t => {
      const msEnt = aMs(t.fEnt, t.hEnt);
      const ms = msEnt != null ? msCierre - msEnt : null;
      return {
        folio: t.folio, fechaEntrada: t.fEnt, horaEntrada: t.hEnt,
        unidad: t.unidad, operador: t.operador, planta: t.planta, areaServicio: t.area,
        tiempo: duracion(ms), horas: (ms != null && ms >= 0) ? ms / 3600000 : null,
        reporteFalla: t.falla,
      };
    }).sort((a, b) => (b.horas || 0) - (a.horas || 0));

    // ── Impacto por planta ──
    const porPlanta = {};
    const tocarPlanta = p => (porPlanta[p || 'Sin planta'] = porPlanta[p || 'Sin planta'] || { activas: 0, finalizadas: 0 });
    activos.forEach(a => tocarPlanta(a.planta).activas++);
    finalizados.forEach(f => tocarPlanta(f.planta).finalizadas++);
    const plantas = Object.entries(porPlanta)
      .map(([planta, c]) => ({ planta, ...c }))
      .sort((a, b) => b.activas - a.activas || b.finalizadas - a.finalizadas);

    // ── Resumen ──
    const entraron = Object.values(taller).filter(t => claveDia(parseFecha(t.fEnt)) === kDia).length;
    const conHoras = finalizados.filter(f => f.horas != null);
    const promedio = conHoras.length
      ? duracion(conHoras.reduce((s, f) => s + f.horas, 0) / conHoras.length * 3600000) : '';

    return {
      fecha: fechaDia,
      generado: `${fechaStr(hoy)} ${pad(hoy.hh)}:${pad(hoy.mm)}`,
      resumen: { entraron, finalizados: finalizados.length, activos: activos.length, promedio },
      finalizados, activos, plantas,
      liquidos: Object.entries(liquidos).map(([concepto, x]) => ({
        concepto, litros: Math.round(x.litros * 10) / 10, unidades: x.unidades.size,
      })).sort((a, b) => b.litros - a.litros),
      piezas,
    };
  }

  // ═══════════════════════════════════════════════════════════════
  // PESTAÑA EN SHEETS
  // ═══════════════════════════════════════════════════════════════
  async function escribirPestana(sheets, rep) {
    const titulo = `Diario ${rep.fecha.replace(/\//g, '-')}`;
    const filas = [];
    const fmt = { titulo: [], seccion: [], encabezado: [] };
    const push = (fila, tipo) => { if (tipo) fmt[tipo].push(filas.length); filas.push(fila); };
    const vacio = () => filas.push([]);

    push([`REPORTE DIARIO DE TALLER — ${rep.fecha}`], 'titulo');
    push([`Generado: ${rep.generado}`]);
    vacio();

    push(['RESUMEN'], 'seccion');
    push(['Entraron a taller', 'Finalizados', 'Activos al cierre', 'Tiempo promedio de reparacion'], 'encabezado');
    push([rep.resumen.entraron, rep.resumen.finalizados, rep.resumen.activos, rep.resumen.promedio || '—']);
    vacio();

    push(['REPORTES FINALIZADOS'], 'seccion');
    push(['Folio','Fecha','Hora','Unidad','Operador','Planta','Area Servicio','Fecha Entrada',
          'Fecha Salida','Tiempo Reparacion','Mecanico','Reporte de Falla','Servicios Realizados'], 'encabezado');
    if (rep.finalizados.length) {
      rep.finalizados.forEach(f => push([f.folio, f.fecha, f.hora, f.unidad, f.operador, f.planta,
        f.areaServicio, f.fechaEntrada, f.fechaSalida, f.tiempo, f.mecanico, f.reporteFalla, f.servicios]));
    } else push(['Sin reportes finalizados este dia']);
    vacio();

    push(['SIGUEN ACTIVOS AL CIERRE'], 'seccion');
    push(['Folio','Fecha Entrada','Hora Entrada','Unidad','Operador','Planta','Area Servicio',
          'Tiempo en Taller','Reporte de Falla'], 'encabezado');
    if (rep.activos.length) {
      rep.activos.forEach(a => push([a.folio, a.fechaEntrada, a.horaEntrada, a.unidad, a.operador,
        a.planta, a.areaServicio, a.tiempo, a.reporteFalla]));
    } else push(['Sin unidades activas al cierre']);
    vacio();

    push(['IMPACTO POR PLANTA'], 'seccion');
    push(['Planta', 'Activas al cierre', 'Finalizadas en el dia'], 'encabezado');
    if (rep.plantas.length) rep.plantas.forEach(p => push([p.planta, p.activas, p.finalizadas]));
    else push(['Sin movimientos']);
    vacio();

    push(['CONSUMO DEL DIA — LIQUIDOS'], 'seccion');
    push(['Concepto', 'Litros', 'Unidades'], 'encabezado');
    if (rep.liquidos.length) rep.liquidos.forEach(l => push([l.concepto, l.litros, l.unidades]));
    else push(['Sin consumo de liquidos registrado']);
    vacio();

    push(['CONSUMO DEL DIA — PIEZAS'], 'seccion');
    push(['Folio', 'Unidad', 'Area', 'Pieza', 'Material / Detalle'], 'encabezado');
    if (rep.piezas.length) rep.piezas.forEach(p => push([p.folio, p.unidad, p.area, p.pieza, p.material]));
    else push(['Sin piezas cambiadas registradas']);

    // Borrar y recrear la pestaña: regenerar un dia siempre deja el mismo
    // resultado (y si dos instancias corren a la vez, no se duplica nada).
    const meta = await sheets.spreadsheets.get({
      spreadsheetId: SHEET_ID_DIARIO, fields: 'sheets.properties(sheetId,title)',
    });
    const previa = (meta.data.sheets || []).find(s => s.properties.title === titulo);
    const requests = [];
    if (previa) requests.push({ deleteSheet: { sheetId: previa.properties.sheetId } });
    requests.push({ addSheet: { properties: { title: titulo, index: 0, gridProperties: { rowCount: Math.max(filas.length + 5, 50), columnCount: 13 } } } });
    const r = await sheets.spreadsheets.batchUpdate({ spreadsheetId: SHEET_ID_DIARIO, requestBody: { requests } });
    const sheetId = r.data.replies[r.data.replies.length - 1].addSheet.properties.sheetId;

    await sheets.spreadsheets.values.update({
      spreadsheetId: SHEET_ID_DIARIO, range: `'${titulo}'!A1`,
      valueInputOption: 'RAW', requestBody: { values: filas },
    });

    const rango = (fila, c0 = 0, c1 = 13) => ({ sheetId, startRowIndex: fila, endRowIndex: fila + 1, startColumnIndex: c0, endColumnIndex: c1 });
    const color = hex => ({ red: parseInt(hex.slice(1,3),16)/255, green: parseInt(hex.slice(3,5),16)/255, blue: parseInt(hex.slice(5,7),16)/255 });
    const anchos = [90, 85, 60, 60, 150, 110, 130, 115, 115, 105, 130, 240, 300];

    const formato = [
      ...fmt.titulo.map(f => ({ repeatCell: { range: rango(f),
        cell: { userEnteredFormat: { textFormat: { bold: true, fontSize: 14 } } }, fields: 'userEnteredFormat.textFormat' } })),
      ...fmt.seccion.map(f => ({ repeatCell: { range: rango(f),
        cell: { userEnteredFormat: { backgroundColor: color('#e8b84b'), textFormat: { bold: true, foregroundColor: color('#16213e') } } },
        fields: 'userEnteredFormat(backgroundColor,textFormat)' } })),
      ...fmt.encabezado.map(f => ({ repeatCell: { range: rango(f),
        cell: { userEnteredFormat: { backgroundColor: color('#16213e'), textFormat: { bold: true, foregroundColor: color('#ffffff') } } },
        fields: 'userEnteredFormat(backgroundColor,textFormat)' } })),
      { repeatCell: { range: { sheetId, startRowIndex: 0, endRowIndex: filas.length, startColumnIndex: 0, endColumnIndex: 13 },
        cell: { userEnteredFormat: { wrapStrategy: 'WRAP', verticalAlignment: 'TOP' } },
        fields: 'userEnteredFormat(wrapStrategy,verticalAlignment)' } },
      ...anchos.map((w, i) => ({ updateDimensionProperties: {
        range: { sheetId, dimension: 'COLUMNS', startIndex: i, endIndex: i + 1 },
        properties: { pixelSize: w }, fields: 'pixelSize' } })),
    ];
    await sheets.spreadsheets.batchUpdate({ spreadsheetId: SHEET_ID_DIARIO, requestBody: { requests: formato } });

    console.log(`[ReporteDiario] Pestaña "${titulo}" escrita (${filas.length} filas)`);
    return titulo;
  }

  // ═══════════════════════════════════════════════════════════════
  // PDF (horizontal, tablas con alto de fila automatico)
  // ═══════════════════════════════════════════════════════════════
  function generarPDF(rep) {
    return new Promise((resolve, reject) => {
      const doc = new PDFDocument({ size: 'LETTER', layout: 'landscape', margin: 30, bufferPages: true });
      const partes = [];
      doc.on('data', c => partes.push(c));
      doc.on('end', () => resolve(Buffer.concat(partes)));
      doc.on('error', reject);

      const AZUL = '#16213e', DORADO = '#e8b84b', GRIS = '#f3f5f8', LINEA = '#c9d0da';
      const X0 = 30, W = doc.page.width - 60, Y_MAX = doc.page.height - 40;

      // Encabezado de la primera hoja
      try { doc.image(Buffer.from(LOGO_BASE64, 'base64'), X0, 24, { width: 90, height: 36 }); } catch (e) {}
      doc.fillColor(AZUL).font('Helvetica-Bold').fontSize(15)
         .text('REPORTE DIARIO DE TALLER', X0, 30, { width: W, align: 'center' });
      doc.font('Helvetica').fontSize(9).fillColor('#444')
         .text(`Fecha: ${rep.fecha}`, X0, 30, { width: W, align: 'right' })
         .text(`Generado: ${rep.generado}`, X0, 42, { width: W, align: 'right' });
      let y = 72;

      // Resumen en 4 recuadros
      const cajas = [
        ['Entraron a taller', rep.resumen.entraron],
        ['Finalizados', rep.resumen.finalizados],
        ['Activos al cierre', rep.resumen.activos],
        ['Tiempo promedio de reparacion', rep.resumen.promedio || '—'],
      ];
      const cw = (W - 30) / 4;
      cajas.forEach(([t, v], i) => {
        const x = X0 + i * (cw + 10);
        doc.roundedRect(x, y, cw, 42, 4).fillAndStroke(GRIS, LINEA);
        doc.fillColor('#555').font('Helvetica').fontSize(7.5).text(t.toUpperCase(), x + 8, y + 7, { width: cw - 16 });
        doc.fillColor(AZUL).font('Helvetica-Bold').fontSize(15).text(String(v), x + 8, y + 19, { width: cw - 16 });
      });
      y += 56;

      function seccion(titulo) {
        if (y + 40 > Y_MAX) { doc.addPage(); y = 30; }
        doc.rect(X0, y, 4, 13).fill(DORADO);
        doc.fillColor(AZUL).font('Helvetica-Bold').fontSize(10).text(titulo, X0 + 9, y + 1.5);
        y += 19;
      }

      // Tabla generica: cols = [{h, w, k}], w en proporcion del ancho total
      function tabla(cols, filas, vacioTxt) {
        const total = cols.reduce((s, c) => s + c.w, 0);
        const anchos = cols.map(c => c.w / total * W);
        const FS = 7, PAD = 3.5;

        const altoFila = (valores, fuente) => {
          doc.font(fuente).fontSize(FS);
          return Math.max(...valores.map((t, i) =>
            doc.heightOfString(String(t == null ? '' : t), { width: anchos[i] - PAD * 2 }))) + PAD * 2;
        };
        const dibujar = (valores, esEnc, zebra) => {
          const fuente = esEnc ? 'Helvetica-Bold' : 'Helvetica';
          const h = altoFila(valores, fuente);
          if (y + h > Y_MAX) { doc.addPage(); y = 30; if (!esEnc) dibujar(cols.map(c => c.h), true); }
          doc.rect(X0, y, W, h).fill(esEnc ? AZUL : (zebra ? GRIS : '#ffffff'));
          let x = X0;
          valores.forEach((t, i) => {
            doc.fillColor(esEnc ? '#ffffff' : '#222').font(fuente).fontSize(FS)
               .text(String(t == null ? '' : t), x + PAD, y + PAD, { width: anchos[i] - PAD * 2 });
            x += anchos[i];
          });
          doc.lineWidth(0.4).strokeColor(LINEA).rect(X0, y, W, h).stroke();
          let xl = X0;
          anchos.slice(0, -1).forEach(a => { xl += a; doc.moveTo(xl, y).lineTo(xl, y + h).stroke(); });
          y += h;
        };

        dibujar(cols.map(c => c.h), true);
        if (!filas.length) dibujar([vacioTxt, ...cols.slice(1).map(() => '')], false, false);
        else filas.forEach((f, i) => dibujar(cols.map(c => f[c.k]), false, i % 2 === 1));
        y += 14;
      }

      seccion('REPORTES FINALIZADOS');
      tabla([
        { h: 'Folio', w: 52, k: 'folio' }, { h: 'Unidad', w: 34, k: 'unidad' },
        { h: 'Operador', w: 70, k: 'operador' }, { h: 'Planta', w: 58, k: 'planta' },
        { h: 'Area servicio', w: 64, k: 'areaServicio' }, { h: 'Entrada', w: 56, k: 'fechaEntrada' },
        { h: 'Salida', w: 56, k: 'fechaSalida' }, { h: 'Tiempo', w: 42, k: 'tiempo' },
        { h: 'Mecanico', w: 62, k: 'mecanico' }, { h: 'Reporte de falla', w: 110, k: 'reporteFalla' },
        { h: 'Servicios realizados', w: 124, k: 'servicios' },
      ], rep.finalizados, 'Sin reportes finalizados este dia');

      seccion('SIGUEN ACTIVOS AL CIERRE');
      tabla([
        { h: 'Folio', w: 52, k: 'folio' }, { h: 'Unidad', w: 34, k: 'unidad' },
        { h: 'Operador', w: 80, k: 'operador' }, { h: 'Planta', w: 62, k: 'planta' },
        { h: 'Area servicio', w: 70, k: 'areaServicio' }, { h: 'Fecha entrada', w: 56, k: 'fechaEntrada' },
        { h: 'Hora', w: 32, k: 'horaEntrada' }, { h: 'Tiempo en taller', w: 52, k: 'tiempo' },
        { h: 'Reporte de falla', w: 190, k: 'reporteFalla' },
      ], rep.activos, 'Sin unidades activas al cierre');

      seccion('IMPACTO POR PLANTA');
      tabla([
        { h: 'Planta', w: 2, k: 'planta' }, { h: 'Activas al cierre', w: 1, k: 'activas' },
        { h: 'Finalizadas en el dia', w: 1, k: 'finalizadas' },
      ], rep.plantas, 'Sin movimientos');

      seccion('CONSUMO DEL DIA — LIQUIDOS');
      tabla([
        { h: 'Concepto', w: 2, k: 'concepto' }, { h: 'Litros', w: 1, k: 'litros' },
        { h: 'Unidades', w: 1, k: 'unidades' },
      ], rep.liquidos, 'Sin consumo de liquidos registrado');

      seccion('CONSUMO DEL DIA — PIEZAS');
      tabla([
        { h: 'Folio', w: 1, k: 'folio' }, { h: 'Unidad', w: 0.7, k: 'unidad' },
        { h: 'Area', w: 1, k: 'area' }, { h: 'Pieza', w: 2.2, k: 'pieza' },
        { h: 'Material / detalle', w: 3, k: 'material' },
      ], rep.piezas, 'Sin piezas cambiadas registradas');

      // Pie con numero de pagina. Se quita el margen inferior mientras se
      // escribe: si no, PDFKit crea una hoja en blanco por cada pie.
      const rango = doc.bufferedPageRange();
      for (let i = 0; i < rango.count; i++) {
        doc.switchToPage(rango.start + i);
        const mb = doc.page.margins.bottom;
        doc.page.margins.bottom = 0;
        doc.fillColor('#888').font('Helvetica').fontSize(7)
           .text(`TECSA — Reporte diario ${rep.fecha}   ·   Pag. ${i + 1} de ${rango.count}`,
                 X0, doc.page.height - 26, { width: W, align: 'center', lineBreak: false });
        doc.page.margins.bottom = mb;
      }
      doc.end();
    });
  }

  // ═══════════════════════════════════════════════════════════════
  // ENDPOINTS
  // ═══════════════════════════════════════════════════════════════
  app.get('/api/reporte-diario', async (req, res) => {
    try {
      const { sheets } = await getClients();
      res.json({ ok: true, reporte: await construirReporte(sheets, fechaDeQuery(req.query.fecha)) });
    } catch (e) { console.error('[ReporteDiario]', e.message); res.json({ ok: false, error: e.message }); }
  });

  app.get('/api/reporte-diario-pdf', async (req, res) => {
    try {
      const { sheets } = await getClients();
      const rep = await construirReporte(sheets, fechaDeQuery(req.query.fecha));
      const pdf = await generarPDF(rep);
      res.setHeader('Content-Type', 'application/pdf');
      res.setHeader('Content-Disposition', `attachment; filename="Reporte_Diario_${rep.fecha.replace(/\//g, '-')}.pdf"`);
      res.send(pdf);
    } catch (e) { console.error('[ReporteDiario PDF]', e.message); res.status(500).send('Error generando PDF: ' + e.message); }
  });

  app.post('/api/reporte-diario/sheets', async (req, res) => {
    try {
      const { sheets } = await getClients();
      const rep = await construirReporte(sheets, fechaDeQuery(req.query.fecha));
      res.json({ ok: true, hoja: await escribirPestana(sheets, rep) });
    } catch (e) { console.error('[ReporteDiario Sheets]', e.message); res.json({ ok: false, error: e.message }); }
  });

  // ═══════════════════════════════════════════════════════════════
  // GENERACION AUTOMATICA
  // Revisa cada minuto; a partir de la hora configurada escribe la
  // pestaña del dia una sola vez. Si el servidor se reinicia despues
  // de esa hora, la vuelve a escribir (es idempotente, no duplica).
  // ═══════════════════════════════════════════════════════════════
  let ultimoDia = null;
  const [hAuto, mAuto] = HORA_AUTO.split(':').map(n => parseInt(n, 10) || 0);
  setInterval(async () => {
    const p = partesMty();
    const hoyTxt = fechaStr(p);
    if (ultimoDia === hoyTxt) return;
    if (p.hh * 60 + p.mm < hAuto * 60 + mAuto) return;
    ultimoDia = hoyTxt; // se marca antes para no encimar ejecuciones
    try {
      const { sheets } = await getClients();
      await escribirPestana(sheets, await construirReporte(sheets, hoyTxt));
    } catch (e) {
      console.error('[ReporteDiario auto]', e.message);
      ultimoDia = null; // se reintenta en el siguiente minuto
    }
  }, 60 * 1000);

  console.log(`[ReporteDiario] Automatico a las ${HORA_AUTO} (hora de Monterrey) → ${SHEET_ID_DIARIO === SHEET_ID ? 'Sheet principal' : 'SHEET_ID_DIARIO'}`);
};
