// ═══════════════════════════════════════════════════════════════════
// REPORTE DE TALLER POR PERIODO — TECSA
//
// Un dia o un rango de fechas (semana, mes...), opcionalmente de una
// sola planta. Se reconstruye desde las hojas, asi que cualquier
// periodo pasado se puede regenerar y siempre da lo mismo.
//
//   GET  /api/reporte-diario?desde=YYYY-MM-DD&hasta=YYYY-MM-DD&planta=X  → JSON
//   GET  /api/reporte-diario-pdf?...mismos parametros...                  → PDF
//   POST /api/reporte-diario/sheets?...mismos parametros...               → pestaña
//   (se acepta ?fecha=YYYY-MM-DD como atajo de un solo dia)
//
// Ademas escribe sola la pestaña del dia (todas las plantas) a la hora
// configurada en REPORTE_DIARIO_HORA.
// ═══════════════════════════════════════════════════════════════════

module.exports = function registrarReporteDiario(app, deps) {
  const { getClients, PDFDocument, SHEET_ID, HOJAS_TALLER, HEADERS_TALLER,
          HOJA_EVIDENCIAS, LOGO_BASE64 } = deps;

  const SHEET_ID_DIARIO = process.env.SHEET_ID_DIARIO || SHEET_ID;
  const HORA_AUTO       = process.env.REPORTE_DIARIO_HORA || '19:00';

  // ── Fechas (hora de Monterrey) ─────────────────────────────────
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

  function parseFecha(s) {
    const t = String(s || '').trim().split('/');
    if (t.length !== 3) return null;
    const r = { d: parseInt(t[0], 10), m: parseInt(t[1], 10), y: parseInt(t[2], 10) };
    return (isNaN(r.d) || isNaN(r.m) || isNaN(r.y)) ? null : r;
  }
  const claveDia = f => f ? f.y * 10000 + f.m * 100 + f.d : 0;

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

  const isoAFecha = s => {
    const m = String(s || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
    return m ? `${m[3]}/${m[2]}/${m[1]}` : null;
  };

  // Parametros del query → { desde, hasta, planta } en dd/MM/yyyy
  function periodoDeQuery(q) {
    const hoy = fechaStr(partesMty());
    let desde = isoAFecha(q.desde) || isoAFecha(q.fecha) || hoy;
    let hasta = isoAFecha(q.hasta) || desde;
    if (claveDia(parseFecha(hasta)) < claveDia(parseFecha(desde))) [desde, hasta] = [hasta, desde];
    const planta = String(q.planta || '').trim().toUpperCase();
    return { desde, hasta, planta };
  }

  // ── Utilidades de filas ────────────────────────────────────────
  const idx = (hoja, nombre) => HEADERS_TALLER[hoja].indexOf(nombre);
  const celda = (fila, hoja, nombre) => {
    const i = idx(hoja, nombre);
    return i < 0 ? '' : String(fila[i] == null ? '' : fila[i]).trim();
  };
  const esSi = v => /^si\b/i.test(String(v || '').trim());
  const num  = v => { const n = parseFloat(String(v || '').replace(',', '.')); return isNaN(n) ? 0 : n; };
  const normPlanta = p => String(p || '').trim().toUpperCase() || 'SIN PLANTA';

  function parsePiezas(txt) {
    const s = String(txt || '').trim();
    if (!s) return [];
    if (/Nombre de la pz:/i.test(s)) {
      return s.split('//').map(b => {
        const nom = (b.match(/Nombre de la pz:\s*([^|]*)/i) || [])[1] || '';
        const mat = (b.match(/Material us:\s*([^|]*)/i) || [])[1] || '';
        return { pieza: nom.trim(), material: mat.trim() };
      }).filter(p => p.pieza || p.material);
    }
    return s.split('|').map(x => {
      const limpio = x.replace(/^\s*\d+\.\s*/, '').trim();
      const [pieza, material] = limpio.split('—').map(t => (t || '').replace(/\[.*\]/, '').trim());
      return { pieza: pieza || limpio, material: material || '' };
    }).filter(p => p.pieza);
  }

  async function leerHojas(sheets) {
    const meta = await sheets.spreadsheets.get({ spreadsheetId: SHEET_ID, fields: 'sheets.properties.title' });
    const existentes = new Set((meta.data.sheets || []).map(s => s.properties.title));
    const pedidos = [
      ['entradas', 'Entradas'], ['taller', 'Taller'],
      ...Object.entries(HOJAS_TALLER), ['evidencias', HOJA_EVIDENCIAS],
    ].filter(([, t]) => existentes.has(t));
    const res = await sheets.spreadsheets.values.batchGet({
      spreadsheetId: SHEET_ID, ranges: pedidos.map(([, t]) => `'${t}'!A2:BZ`),
    });
    const datos = {};
    pedidos.forEach(([k], i) => { datos[k] = res.data.valueRanges[i].values || []; });
    return datos;
  }

  // ═══════════════════════════════════════════════════════════════
  // CONSTRUCCION DEL REPORTE
  // ═══════════════════════════════════════════════════════════════
  async function construirReporte(sheets, { desde, hasta, planta }) {
    const H = await leerHojas(sheets);
    const kDesde = claveDia(parseFecha(desde));
    const kHasta = claveDia(parseFecha(hasta));
    const enRango = k => k >= kDesde && k <= kHasta;
    const esDia = kDesde === kHasta;

    // Cierre: si el periodo termina hoy es "ahora"; si no, 23:59 del ultimo dia
    const hoy = partesMty();
    const fH = parseFecha(hasta);
    const msCierre = kHasta === claveDia(hoy)
      ? Date.UTC(hoy.y, hoy.m - 1, hoy.d, hoy.hh, hoy.mm)
      : Date.UTC(fH.y, fH.m - 1, fH.d, 23, 59);

    const entradas = {};
    (H.entradas || []).forEach(f => {
      const folio = String(f[0] || '').trim();
      if (folio) entradas[folio] = { fSal: f[7] || '', hSal: f[8] || '' };
    });

    const taller = {};
    const catalogo = new Set();
    (H.taller || []).forEach(f => {
      const folio = String(f[0] || '').trim();
      if (!folio) return;
      taller[folio] = {
        folio, fEnt: f[1] || '', hEnt: f[2] || '', unidad: f[3] || '', operador: f[4] || '',
        planta: f[5] || '', area: f[6] || '', falla: f[7] || '', estado: f[8] || '',
      };
      if (f[5]) catalogo.add(normPlanta(f[5]));
    });

    // Filtro de planta: se decide por la planta de la entrada a taller
    const pasaPlanta = p => !planta || normPlanta(p) === planta;

    const frenosPorFolio = new Set();
    (H.evidencias || []).forEach(f => {
      if (String(f[2] || '').trim() === 'Frenos') frenosPorFolio.add(String(f[0] || '').trim());
    });

    const fin = {};
    const liquidos = {};
    const piezas = [];

    const sumarLitros = (concepto, litros, folio, unidad) => {
      if (!litros) return;
      const l = liquidos[concepto] = liquidos[concepto] || { litros: 0, unidades: new Set(), detalle: [] };
      l.litros += litros;
      if (unidad) l.unidades.add(unidad);
      l.detalle.push({ folio, unidad, litros: Math.round(litros * 10) / 10 });
    };

    const nombresHoja = { mecanico:'Mecanico', electrico:'Electrico', imagen:'Imagen',
                          llantas:'Llantas', llenado:'Rellenados', suspension:'Suspension' };

    Object.keys(HOJAS_TALLER).forEach(hoja => {
      (H[hoja] || []).forEach(fila => {
        const fechaRep = celda(fila, hoja, 'Fecha');
        if (!enRango(claveDia(parseFecha(fechaRep)))) return;
        const folio = celda(fila, hoja, 'Folio');
        if (!folio) return;
        const t = taller[folio] || {};
        const plantaFila = t.planta || celda(fila, hoja, 'Planta');
        if (!pasaPlanta(plantaFila)) return;

        const unidad = celda(fila, hoja, 'Unidad') || t.unidad || '';
        const r = fin[folio] = fin[folio] || {
          folio, fecha: fechaRep, hora: celda(fila, hoja, 'Hora'), unidad,
          operador: celda(fila, hoja, 'Operador'), planta: plantaFila,
          area: celda(fila, hoja, 'Area Servicio'), mecanicos: new Set(), servicios: [],
        };
        const mec = celda(fila, hoja, 'Mecanico');
        if (mec) r.mecanicos.add(mec);

        const serv = [];
        const pz = (area, pieza, material) => piezas.push({ folio, unidad, planta: plantaFila, area, pieza, material: material || '' });
        const agregarPiezas = (col, area) => parsePiezas(celda(fila, hoja, col)).forEach(p => pz(area, p.pieza, p.material));

        if (hoja === 'mecanico') {
          if (celda(fila, hoja, 'Aceite-Km') || celda(fila, hoja, 'Aceite-LitNuevo')) {
            serv.push('Cambio de aceite');
            sumarLitros('Aceite (cambio de aceite)', num(celda(fila, hoja, 'Aceite-LitNuevo')), folio, unidad);
          }
          if (celda(fila, hoja, 'Frenos-Obs') || frenosPorFolio.has(folio)) serv.push('Frenos');
          if (esSi(celda(fila, hoja, 'Engrasado'))) serv.push('Engrasado');
          if (celda(fila, hoja, 'SrvAceite-Km') || celda(fila, hoja, 'SrvAceite-LitNuevo')) {
            serv.push('Servicio de aceite');
            sumarLitros('Aceite (servicio)', num(celda(fila, hoja, 'SrvAceite-LitNuevo')), folio, unidad);
          }
          [['Filtro-Aire','Filtro de aire'], ['Filtro-Diesel','Filtro de diesel'],
           ['Filtro-Aceite','Filtro de aceite'], ['Filtro-Separador','Filtro separador de agua']]
            .forEach(([col, nombre]) => {
              if (esSi(celda(fila, hoja, col))) { serv.push(nombre); pz('Mecanico', nombre, ''); }
            });
          if (celda(fila, hoja, 'Afinacion-Km')) serv.push('Afinacion');
          if (celda(fila, hoja, 'Piezas-Taller')) { serv.push('Cambio de piezas'); agregarPiezas('Piezas-Taller', 'Mecanico'); }
          const aj = celda(fila, hoja, 'Ajuste');
          if (aj === 'Si' || aj.startsWith('Si —')) serv.push('Ajuste');
        }
        else if (hoja === 'electrico') {
          if (celda(fila, hoja, 'Carga-Bat-VoltAnt') || celda(fila, hoja, 'Carga-Bat-VoltNuevo')) serv.push('Carga de bateria');
          if (celda(fila, hoja, 'Cambio-Bat-Motivo')) { serv.push('Cambio de bateria'); pz('Electrico', 'Bateria', celda(fila, hoja, 'Cambio-Bat-Motivo')); }
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
          if (celda(fila, hoja, 'Llanta-Marca')) { serv.push('Cambio de llanta'); pz('Llantas', 'Llanta', celda(fila, hoja, 'Llanta-Marca')); }
          if (celda(fila, hoja, 'LlantaRep-Vida')) serv.push('Reparacion de llanta');
        }
        else if (hoja === 'llenado') {
          [['Aceite','Aceite (rellenado)'], ['Adblue','AdBlue'], ['LiqFrenos','Liquido de frenos'],
           ['Anticong','Anticongelante'], ['Gasolina','Gasolina'], ['Direccion','Liquido de direccion hidraulica']]
            .forEach(([pref, nombre]) => {
              const litros = num(celda(fila, hoja, `${pref}-LitPuestos`));
              if (litros) { serv.push(`Rellenado: ${nombre}`); sumarLitros(nombre, litros, folio, unidad); }
            });
        }
        else if (hoja === 'suspension') {
          if (esSi(celda(fila, hoja, 'Muelles'))) {
            const pos = celda(fila, hoja, 'Muelles-Posicion'), hj = celda(fila, hoja, 'Muelles-Hojas');
            serv.push('Cambio de muelles' + (pos ? ` (${pos})` : ''));
            pz('Suspension', 'Muelle' + (pos ? ` ${pos.toLowerCase()}` : ''), hj ? `Hojas: ${hj}` : '');
          }
          if (esSi(celda(fila, hoja, 'Amortiguadores'))) { serv.push('Amortiguadores'); pz('Suspension', 'Amortiguadores', ''); }
          if (celda(fila, hoja, 'Piezas-Suspension')) { serv.push('Cambio de piezas (suspension)'); agregarPiezas('Piezas-Suspension', 'Suspension'); }
        }
        r.servicios.push(...(serv.length ? serv : [`${nombresHoja[hoja]} (ver observaciones)`]));
      });
    });

    // Salidas del periodo sin reporte de taller (baja directa en el dashboard)
    Object.values(taller).forEach(t => {
      const e = entradas[t.folio];
      if (!e || fin[t.folio] || !pasaPlanta(t.planta)) return;
      if (!enRango(claveDia(parseFecha(e.fSal)))) return;
      fin[t.folio] = {
        folio: t.folio, fecha: e.fSal, hora: e.hSal, unidad: t.unidad, operador: t.operador,
        planta: t.planta, area: t.area, mecanicos: new Set(),
        servicios: ['Salida registrada sin reporte de taller'],
      };
    });

    // ── Finalizados, marcando si entraron dentro del periodo o antes ──
    const finalizados = Object.values(fin).map(r => {
      const t = taller[r.folio] || {};
      const e = entradas[r.folio] || {};
      const fSal = e.fSal || r.fecha, hSal = e.hSal || r.hora;
      const msEnt = aMs(t.fEnt, t.hEnt), msSal = aMs(fSal, hSal);
      const ms = (msEnt != null && msSal != null) ? msSal - msEnt : null;
      const kEnt = claveDia(parseFecha(t.fEnt));
      return {
        folio: r.folio, unidad: r.unidad || t.unidad || '', operador: r.operador || t.operador || '',
        planta: r.planta || t.planta || '', areaServicio: r.area || t.area || '',
        fechaEntrada: t.fEnt ? `${t.fEnt} ${t.hEnt || ''}`.trim() : '',
        fechaSalida:  fSal ? `${fSal} ${hSal || ''}`.trim() : '',
        tiempo: duracion(ms), horas: (ms != null && ms >= 0) ? ms / 3600000 : null,
        mecanico: [...r.mecanicos].join(', '),
        reporteFalla: t.falla || '',
        servicios: [...new Set(r.servicios)].join(', '),
        origen: (kEnt && kEnt < kDesde) ? 'anterior' : 'periodo',
        _orden: msSal || 0,
      };
    }).sort((a, b) => a._orden - b._orden).map(({ _orden, ...x }) => x);

    // ── Activos al cierre, separados por cuando entraron ──
    const activos = Object.values(taller).filter(t => {
      if (!pasaPlanta(t.planta)) return false;
      const kEnt = claveDia(parseFecha(t.fEnt));
      if (!kEnt || kEnt > kHasta) return false;
      const e = entradas[t.folio];
      if (e) { const kSal = claveDia(parseFecha(e.fSal)); return !kSal || kSal > kHasta; }
      return t.estado === 'ACTIVO';
    }).map(t => {
      const msEnt = aMs(t.fEnt, t.hEnt);
      const ms = msEnt != null ? msCierre - msEnt : null;
      return {
        folio: t.folio, fechaEntrada: t.fEnt, horaEntrada: t.hEnt,
        unidad: t.unidad, operador: t.operador, planta: t.planta, areaServicio: t.area,
        tiempo: duracion(ms), horas: (ms != null && ms >= 0) ? ms / 3600000 : null,
        reporteFalla: t.falla,
        origen: claveDia(parseFecha(t.fEnt)) < kDesde ? 'anterior' : 'periodo',
      };
    }).sort((a, b) => (b.horas || 0) - (a.horas || 0));

    // ── Impacto por planta, con el detalle de cada folio ──
    const porPlanta = {};
    const tocar = p => (porPlanta[normPlanta(p)] = porPlanta[normPlanta(p)] || { activas: [], finalizadas: [] });
    activos.forEach(a => tocar(a.planta).activas.push({
      folio: a.folio, unidad: a.unidad, area: a.areaServicio, motivo: a.reporteFalla, tiempo: a.tiempo }));
    finalizados.forEach(f => tocar(f.planta).finalizadas.push({
      folio: f.folio, unidad: f.unidad, area: f.areaServicio, motivo: f.reporteFalla,
      tiempo: f.tiempo, servicios: f.servicios }));
    const plantas = Object.entries(porPlanta).map(([p, x]) => ({
      planta: p, activas: x.activas.length, finalizadas: x.finalizadas.length,
      detalleActivas: x.activas, detalleFinalizadas: x.finalizadas,
    })).sort((a, b) => b.activas - a.activas || b.finalizadas - a.finalizadas);

    // ── Resumen ──
    const entraron = Object.values(taller)
      .filter(t => pasaPlanta(t.planta) && enRango(claveDia(parseFecha(t.fEnt)))).length;
    const conHoras = finalizados.filter(f => f.horas != null);
    const promedio = conHoras.length
      ? duracion(conHoras.reduce((s, f) => s + f.horas, 0) / conHoras.length * 3600000) : '';

    return {
      desde, hasta, esDia, planta: planta || '',
      generado: `${fechaStr(hoy)} ${pad(hoy.hh)}:${pad(hoy.mm)}`,
      resumen: {
        entraron,
        finPeriodo:   finalizados.filter(f => f.origen === 'periodo').length,
        finAnteriores: finalizados.filter(f => f.origen === 'anterior').length,
        activos: activos.length,
        promedio, promedioN: conHoras.length,
      },
      finalizados,
      activosPeriodo:   activos.filter(a => a.origen === 'periodo'),
      activosAnteriores: activos.filter(a => a.origen === 'anterior'),
      plantas,
      liquidos: Object.entries(liquidos).map(([concepto, x]) => ({
        concepto, litros: Math.round(x.litros * 10) / 10, unidades: x.unidades.size, detalle: x.detalle,
      })).sort((a, b) => b.litros - a.litros),
      piezas,
      catalogoPlantas: [...catalogo].sort(),
    };
  }

  // Textos que cambian si es un dia o un periodo
  function etiquetas(rep) {
    return rep.esDia ? {
      periodo: rep.desde,
      finPeriodo: 'Finalizados del dia', finAnteriores: 'Finalizados de dias anteriores',
      actPeriodo: 'Activos — entraron este dia', actAnteriores: 'Activos — de dias anteriores',
    } : {
      periodo: `${rep.desde} al ${rep.hasta}`,
      finPeriodo: 'Finalizados que entraron en el periodo', finAnteriores: 'Finalizados que entraron antes',
      actPeriodo: 'Activos — entraron en el periodo', actAnteriores: 'Activos — de antes del periodo',
    };
  }

  // Detalle en texto (para Sheets y PDF, donde no hay clic)
  const lineaActiva = d => `${d.folio} · U${d.unidad} · ${d.area}${d.motivo ? ' — ' + d.motivo : ''}${d.tiempo ? ' (' + d.tiempo + ')' : ''}`;
  const lineaFinal  = d => `${d.folio} · U${d.unidad} · ${d.area}${d.motivo ? ' — ' + d.motivo : ''}`;
  const lineaLitro  = d => `${d.folio} · U${d.unidad}: ${d.litros} L`;

  // ═══════════════════════════════════════════════════════════════
  // PESTAÑA EN SHEETS
  // Cuadricula de 11 columnas; cada tabla ocupa todo el ancho y los
  // campos largos se combinan en varias columnas para que se lean bien.
  // ═══════════════════════════════════════════════════════════════
  async function escribirPestana(sheets, rep) {
    const E = etiquetas(rep);
    const N = 11;
    const base = rep.esDia ? `Diario ${rep.desde.replace(/\//g, '-')}`
                           : `Periodo ${rep.desde.replace(/\//g, '-')} a ${rep.hasta.replace(/\//g, '-')}`;
    const titulo = (rep.planta ? `${base} · ${rep.planta}` : base).slice(0, 99);

    const filas = [], merges = [], estilos = [], bordes = [];
    const alturas = {};
    const nueva = () => { filas.push(new Array(N).fill('')); return filas.length - 1; };
    const combinar = (r, c0, c1) => { if (c1 - c0 > 1) merges.push({ r, c0, c1 }); };

    // Titulo y subtitulo
    let r = nueva(); filas[r][0] = `REPORTE DE TALLER — ${E.periodo}`; combinar(r, 0, N); estilos.push({ r, c0: 0, c1: N, t: 'titulo' }); alturas[r] = 38;
    r = nueva(); filas[r][0] = `Planta: ${rep.planta || 'Todas'}     ·     Generado: ${rep.generado}`; combinar(r, 0, N); estilos.push({ r, c0: 0, c1: N, t: 'sub' });
    nueva();

    // Tarjetas de resumen: etiqueta arriba, valor grande abajo
    const tarjetas = [
      [0, 2, 'Entraron a taller', rep.resumen.entraron],
      [2, 4, E.finPeriodo, rep.resumen.finPeriodo],
      [4, 6, E.finAnteriores, rep.resumen.finAnteriores],
      [6, 8, 'Activos al cierre', rep.resumen.activos],
      [8, 11, 'Tiempo promedio de reparacion', rep.resumen.promedio ? `${rep.resumen.promedio}  (${rep.resumen.promedioN} reportes)` : '—'],
    ];
    const rLbl = nueva(), rVal = nueva();
    alturas[rVal] = 34;
    tarjetas.forEach(([c0, c1, lbl, val]) => {
      filas[rLbl][c0] = lbl; filas[rVal][c0] = val;
      combinar(rLbl, c0, c1); combinar(rVal, c0, c1);
      estilos.push({ r: rLbl, c0, c1, t: 'cardLbl' }, { r: rVal, c0, c1, t: 'cardVal' });
      bordes.push({ r0: rLbl, r1: rVal + 1, c0, c1, caja: true });
    });
    nueva();

    // Tabla generica: cols = [{h, k, span, centro}]; la suma de span es 11
    function tabla(tituloSec, cols, datos, vacioTxt) {
      const rs = nueva(); filas[rs][0] = tituloSec; combinar(rs, 0, N); estilos.push({ r: rs, c0: 0, c1: N, t: 'seccion' });
      const rh = nueva();
      let c = 0;
      cols.forEach(col => { filas[rh][c] = col.h; combinar(rh, c, c + col.span); c += col.span; });
      estilos.push({ r: rh, c0: 0, c1: N, t: 'enc' });
      const r0 = rh;
      if (!datos.length) {
        const rv = nueva(); filas[rv][0] = vacioTxt; combinar(rv, 0, N); estilos.push({ r: rv, c0: 0, c1: N, t: 'vacio' });
      }
      datos.forEach((d, i) => {
        const rd = nueva();
        let cc = 0;
        cols.forEach(col => {
          const v = typeof col.k === 'function' ? col.k(d) : d[col.k];
          filas[rd][cc] = v == null ? '' : v;
          combinar(rd, cc, cc + col.span);
          if (col.centro) estilos.push({ r: rd, c0: cc, c1: cc + col.span, t: 'centro' });
          cc += col.span;
        });
        if (i % 2 === 1) estilos.push({ r: rd, c0: 0, c1: N, t: 'zebra' });
      });
      bordes.push({ r0, r1: filas.length, c0: 0, c1: N });
      nueva();
    }

    const colsFin = [
      { h: 'Folio', k: 'folio', span: 1 }, { h: 'Unidad', k: 'unidad', span: 1, centro: true },
      { h: 'Operador', k: 'operador', span: 1 }, { h: 'Planta', k: 'planta', span: 1 },
      { h: 'Area servicio', k: 'areaServicio', span: 1 }, { h: 'Entrada', k: 'fechaEntrada', span: 1, centro: true },
      { h: 'Salida', k: 'fechaSalida', span: 1, centro: true }, { h: 'Tiempo', k: 'tiempo', span: 1, centro: true },
      { h: 'Mecanico', k: 'mecanico', span: 1 }, { h: 'Reporte de falla', k: 'reporteFalla', span: 1 },
      { h: 'Servicios realizados', k: 'servicios', span: 1 },
    ];
    const colsAct = [
      { h: 'Folio', k: 'folio', span: 1 }, { h: 'Unidad', k: 'unidad', span: 1, centro: true },
      { h: 'Operador', k: 'operador', span: 1 }, { h: 'Planta', k: 'planta', span: 1 },
      { h: 'Area servicio', k: 'areaServicio', span: 1 },
      { h: 'Entrada', k: d => `${d.fechaEntrada} ${d.horaEntrada || ''}`.trim(), span: 1, centro: true },
      { h: 'Tiempo en taller', k: 'tiempo', span: 1, centro: true },
      { h: 'Reporte de falla', k: 'reporteFalla', span: 4 },
    ];

    tabla('REPORTES FINALIZADOS', colsFin, rep.finalizados, 'Sin reportes finalizados en el periodo');
    tabla(`SIGUEN ACTIVOS AL CIERRE — ${E.actPeriodo.replace('Activos — ', '').toUpperCase()}`, colsAct, rep.activosPeriodo, 'Ninguna');
    tabla(`SIGUEN ACTIVOS AL CIERRE — ${E.actAnteriores.replace('Activos — ', '').toUpperCase()}`, colsAct, rep.activosAnteriores, 'Ninguna');
    tabla('IMPACTO POR PLANTA', [
      { h: 'Planta', k: 'planta', span: 2 },
      { h: 'Activas', k: 'activas', span: 1, centro: true },
      { h: 'Finalizadas', k: 'finalizadas', span: 1, centro: true },
      { h: 'Detalle activas (folio · unidad · area — motivo)', k: d => d.detalleActivas.map(lineaActiva).join('\n'), span: 3 },
      { h: 'Detalle finalizadas (folio · unidad · area — motivo)', k: d => d.detalleFinalizadas.map(lineaFinal).join('\n'), span: 4 },
    ], rep.plantas, 'Sin movimientos');
    tabla('CONSUMO DE LIQUIDOS', [
      { h: 'Concepto', k: 'concepto', span: 2 },
      { h: 'Litros', k: 'litros', span: 1, centro: true },
      { h: 'Unidades', k: 'unidades', span: 1, centro: true },
      { h: 'Detalle (folio · unidad: litros)', k: d => d.detalle.map(lineaLitro).join('\n'), span: 7 },
    ], rep.liquidos, 'Sin consumo de liquidos registrado');
    tabla('CONSUMO DE REFACCIONES', [
      { h: 'Folio', k: 'folio', span: 1 }, { h: 'Unidad', k: 'unidad', span: 1, centro: true },
      { h: 'Area', k: 'area', span: 1 }, { h: 'Pieza', k: 'pieza', span: 2 },
      { h: 'Material / detalle', k: 'material', span: 6 },
    ], rep.piezas, 'Sin piezas cambiadas registradas');

    // ── Crear la pestaña (reemplaza la anterior del mismo periodo) ──
    const meta = await sheets.spreadsheets.get({ spreadsheetId: SHEET_ID_DIARIO, fields: 'sheets.properties(sheetId,title)' });
    const previa = (meta.data.sheets || []).find(s => s.properties.title === titulo);
    const reqs = [];
    if (previa) reqs.push({ deleteSheet: { sheetId: previa.properties.sheetId } });
    reqs.push({ addSheet: { properties: { title: titulo, index: 0,
      gridProperties: { rowCount: filas.length + 5, columnCount: N, hideGridlines: true } } } });
    const resp = await sheets.spreadsheets.batchUpdate({ spreadsheetId: SHEET_ID_DIARIO, requestBody: { requests: reqs } });
    const sheetId = resp.data.replies[resp.data.replies.length - 1].addSheet.properties.sheetId;

    await sheets.spreadsheets.values.update({
      spreadsheetId: SHEET_ID_DIARIO, range: `'${titulo}'!A1`,
      valueInputOption: 'RAW', requestBody: { values: filas },
    });

    const col = hex => ({ red: parseInt(hex.slice(1,3),16)/255, green: parseInt(hex.slice(3,5),16)/255, blue: parseInt(hex.slice(5,7),16)/255 });
    const rng = (r0, r1, c0, c1) => ({ sheetId, startRowIndex: r0, endRowIndex: r1, startColumnIndex: c0, endColumnIndex: c1 });
    const FMT = {
      titulo:  { backgroundColor: col('#16213e'), horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
                 textFormat: { bold: true, fontSize: 16, foregroundColor: col('#ffffff') } },
      sub:     { horizontalAlignment: 'CENTER', textFormat: { italic: true, fontSize: 9, foregroundColor: col('#5a6b85') } },
      cardLbl: { backgroundColor: col('#eef2f7'), horizontalAlignment: 'CENTER', verticalAlignment: 'BOTTOM',
                 textFormat: { fontSize: 9, foregroundColor: col('#5a6b85') } },
      cardVal: { backgroundColor: col('#eef2f7'), horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
                 textFormat: { bold: true, fontSize: 18, foregroundColor: col('#16213e') } },
      seccion: { backgroundColor: col('#e8b84b'), verticalAlignment: 'MIDDLE',
                 textFormat: { bold: true, fontSize: 11, foregroundColor: col('#16213e') } },
      enc:     { backgroundColor: col('#16213e'), horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE', wrapStrategy: 'WRAP',
                 textFormat: { bold: true, fontSize: 9, foregroundColor: col('#ffffff') } },
      vacio:   { horizontalAlignment: 'CENTER', textFormat: { italic: true, foregroundColor: col('#8a94a6') } },
      zebra:   { backgroundColor: col('#f5f7fa') },
      centro:  { horizontalAlignment: 'CENTER' },
    };
    const CAMPOS = {
      titulo: 'userEnteredFormat(backgroundColor,horizontalAlignment,verticalAlignment,textFormat)',
      sub: 'userEnteredFormat(horizontalAlignment,textFormat)',
      cardLbl: 'userEnteredFormat(backgroundColor,horizontalAlignment,verticalAlignment,textFormat)',
      cardVal: 'userEnteredFormat(backgroundColor,horizontalAlignment,verticalAlignment,textFormat)',
      seccion: 'userEnteredFormat(backgroundColor,verticalAlignment,textFormat)',
      enc: 'userEnteredFormat(backgroundColor,horizontalAlignment,verticalAlignment,wrapStrategy,textFormat)',
      vacio: 'userEnteredFormat(horizontalAlignment,textFormat)',
      zebra: 'userEnteredFormat.backgroundColor',
      centro: 'userEnteredFormat.horizontalAlignment',
    };
    const lineaGris = { style: 'SOLID', width: 1, color: col('#c9d0da') };
    const anchos = [95, 60, 150, 115, 140, 120, 120, 95, 130, 250, 300];

    const formato = [
      // Base: todo con ajuste de texto y alineado arriba
      { repeatCell: { range: rng(0, filas.length, 0, N),
        cell: { userEnteredFormat: { wrapStrategy: 'WRAP', verticalAlignment: 'TOP', textFormat: { fontSize: 9 } } },
        fields: 'userEnteredFormat(wrapStrategy,verticalAlignment,textFormat)' } },
      ...estilos.map(e => ({ repeatCell: { range: rng(e.r, e.r + 1, e.c0, e.c1),
        cell: { userEnteredFormat: FMT[e.t] }, fields: CAMPOS[e.t] } })),
      ...merges.map(m => ({ mergeCells: { range: rng(m.r, m.r + 1, m.c0, m.c1), mergeType: 'MERGE_ALL' } })),
      ...bordes.map(b => ({ updateBorders: { range: rng(b.r0, b.r1, b.c0, b.c1),
        top: lineaGris, bottom: lineaGris, left: lineaGris, right: lineaGris,
        ...(b.caja ? {} : { innerHorizontal: lineaGris, innerVertical: lineaGris }) } })),
      ...anchos.map((w, i) => ({ updateDimensionProperties: {
        range: { sheetId, dimension: 'COLUMNS', startIndex: i, endIndex: i + 1 },
        properties: { pixelSize: w }, fields: 'pixelSize' } })),
      ...Object.entries(alturas).map(([row, px]) => ({ updateDimensionProperties: {
        range: { sheetId, dimension: 'ROWS', startIndex: +row, endIndex: +row + 1 },
        properties: { pixelSize: px }, fields: 'pixelSize' } })),
    ];
    await sheets.spreadsheets.batchUpdate({ spreadsheetId: SHEET_ID_DIARIO, requestBody: { requests: formato } });

    console.log(`[ReporteDiario] Pestaña "${titulo}" escrita (${filas.length} filas)`);
    return titulo;
  }

  // ═══════════════════════════════════════════════════════════════
  // PDF (horizontal)
  // ═══════════════════════════════════════════════════════════════
  function generarPDF(rep) {
    const E = etiquetas(rep);
    return new Promise((resolve, reject) => {
      const doc = new PDFDocument({ size: 'LETTER', layout: 'landscape', margin: 30, bufferPages: true });
      const partes = [];
      doc.on('data', c => partes.push(c));
      doc.on('end', () => resolve(Buffer.concat(partes)));
      doc.on('error', reject);

      const AZUL = '#16213e', DORADO = '#e8b84b', GRIS = '#f3f5f8', LINEA = '#c9d0da';
      const X0 = 30, W = doc.page.width - 60, Y_MAX = doc.page.height - 40;

      try { doc.image(Buffer.from(LOGO_BASE64, 'base64'), X0, 24, { width: 90, height: 36 }); } catch (e) {}
      doc.fillColor(AZUL).font('Helvetica-Bold').fontSize(15)
         .text('REPORTE DE TALLER', X0, 28, { width: W, align: 'center' });
      doc.font('Helvetica').fontSize(9).fillColor('#444')
         .text(E.periodo, X0, 47, { width: W, align: 'center' });
      doc.fontSize(8.5)
         .text(`Planta: ${rep.planta || 'Todas'}`, X0, 30, { width: W, align: 'right' })
         .text(`Generado: ${rep.generado}`, X0, 42, { width: W, align: 'right' });
      let y = 72;

      const cajas = [
        ['Entraron a taller', rep.resumen.entraron],
        [E.finPeriodo, rep.resumen.finPeriodo],
        [E.finAnteriores, rep.resumen.finAnteriores],
        ['Activos al cierre', rep.resumen.activos],
        ['Tiempo promedio de reparacion', rep.resumen.promedio ? rep.resumen.promedio : '—'],
      ];
      const cw = (W - 40) / 5;
      cajas.forEach(([t, v], i) => {
        const x = X0 + i * (cw + 10);
        doc.roundedRect(x, y, cw, 44, 4).fillAndStroke(GRIS, LINEA);
        doc.fillColor('#555').font('Helvetica').fontSize(7).text(t.toUpperCase(), x + 7, y + 6, { width: cw - 14 });
        doc.fillColor(AZUL).font('Helvetica-Bold').fontSize(14).text(String(v), x + 7, y + 23, { width: cw - 14 });
      });
      y += 58;

      function seccion(titulo) {
        if (y + 40 > Y_MAX) { doc.addPage(); y = 30; }
        doc.rect(X0, y, 4, 13).fill(DORADO);
        doc.fillColor(AZUL).font('Helvetica-Bold').fontSize(10).text(titulo, X0 + 9, y + 1.5);
        y += 19;
      }

      function tabla(cols, filas, vacioTxt) {
        const total = cols.reduce((s, c) => s + c.w, 0);
        const anchos = cols.map(c => c.w / total * W);
        const FS = 7, PAD = 3.5;
        const valor = (f, c) => { const v = typeof c.k === 'function' ? c.k(f) : f[c.k]; return v == null ? '' : String(v); };
        const alto = (vals, fuente) => {
          doc.font(fuente).fontSize(FS);
          return Math.max(...vals.map((t, i) => doc.heightOfString(t, { width: anchos[i] - PAD * 2 }))) + PAD * 2;
        };
        const dibujar = (vals, esEnc, zebra) => {
          const fuente = esEnc ? 'Helvetica-Bold' : 'Helvetica';
          const h = alto(vals, fuente);
          if (y + h > Y_MAX) { doc.addPage(); y = 30; if (!esEnc) dibujar(cols.map(c => c.h), true); }
          doc.rect(X0, y, W, h).fill(esEnc ? AZUL : (zebra ? GRIS : '#ffffff'));
          let x = X0;
          vals.forEach((t, i) => {
            doc.fillColor(esEnc ? '#ffffff' : '#222').font(fuente).fontSize(FS)
               .text(t, x + PAD, y + PAD, { width: anchos[i] - PAD * 2 });
            x += anchos[i];
          });
          doc.lineWidth(0.4).strokeColor(LINEA).rect(X0, y, W, h).stroke();
          let xl = X0;
          anchos.slice(0, -1).forEach(a => { xl += a; doc.moveTo(xl, y).lineTo(xl, y + h).stroke(); });
          y += h;
        };
        dibujar(cols.map(c => c.h), true);
        if (!filas.length) dibujar([vacioTxt, ...cols.slice(1).map(() => '')], false, false);
        else filas.forEach((f, i) => dibujar(cols.map(c => valor(f, c)), false, i % 2 === 1));
        y += 14;
      }

      const colsAct = [
        { h: 'Folio', w: 52, k: 'folio' }, { h: 'Unidad', w: 34, k: 'unidad' },
        { h: 'Operador', w: 80, k: 'operador' }, { h: 'Planta', w: 62, k: 'planta' },
        { h: 'Area servicio', w: 70, k: 'areaServicio' },
        { h: 'Entrada', w: 62, k: f => `${f.fechaEntrada} ${f.horaEntrada || ''}`.trim() },
        { h: 'Tiempo en taller', w: 52, k: 'tiempo' }, { h: 'Reporte de falla', w: 216, k: 'reporteFalla' },
      ];

      seccion('REPORTES FINALIZADOS');
      tabla([
        { h: 'Folio', w: 52, k: 'folio' }, { h: 'Unidad', w: 34, k: 'unidad' },
        { h: 'Operador', w: 70, k: 'operador' }, { h: 'Planta', w: 58, k: 'planta' },
        { h: 'Area servicio', w: 64, k: 'areaServicio' }, { h: 'Entrada', w: 56, k: 'fechaEntrada' },
        { h: 'Salida', w: 56, k: 'fechaSalida' }, { h: 'Tiempo', w: 42, k: 'tiempo' },
        { h: 'Mecanico', w: 62, k: 'mecanico' }, { h: 'Reporte de falla', w: 110, k: 'reporteFalla' },
        { h: 'Servicios realizados', w: 124, k: 'servicios' },
      ], rep.finalizados, 'Sin reportes finalizados en el periodo');

      seccion(`SIGUEN ACTIVOS AL CIERRE — ${E.actPeriodo.replace('Activos — ', '').toUpperCase()}`);
      tabla(colsAct, rep.activosPeriodo, 'Ninguna');
      seccion(`SIGUEN ACTIVOS AL CIERRE — ${E.actAnteriores.replace('Activos — ', '').toUpperCase()}`);
      tabla(colsAct, rep.activosAnteriores, 'Ninguna');

      seccion('IMPACTO POR PLANTA');
      tabla([
        { h: 'Planta', w: 1.2, k: 'planta' }, { h: 'Activas', w: 0.6, k: 'activas' },
        { h: 'Finalizadas', w: 0.7, k: 'finalizadas' },
        { h: 'Detalle activas', w: 3.2, k: f => f.detalleActivas.map(lineaActiva).join('\n') },
        { h: 'Detalle finalizadas', w: 3.2, k: f => f.detalleFinalizadas.map(lineaFinal).join('\n') },
      ], rep.plantas, 'Sin movimientos');

      seccion('CONSUMO DE LIQUIDOS');
      tabla([
        { h: 'Concepto', w: 1.6, k: 'concepto' }, { h: 'Litros', w: 0.6, k: 'litros' },
        { h: 'Unidades', w: 0.6, k: 'unidades' },
        { h: 'Detalle (folio · unidad: litros)', w: 4.5, k: f => f.detalle.map(lineaLitro).join('\n') },
      ], rep.liquidos, 'Sin consumo de liquidos registrado');

      seccion('CONSUMO DE REFACCIONES');
      tabla([
        { h: 'Folio', w: 1, k: 'folio' }, { h: 'Unidad', w: 0.7, k: 'unidad' },
        { h: 'Area', w: 1, k: 'area' }, { h: 'Pieza', w: 2.2, k: 'pieza' },
        { h: 'Material / detalle', w: 3, k: 'material' },
      ], rep.piezas, 'Sin piezas cambiadas registradas');

      const rango = doc.bufferedPageRange();
      for (let i = 0; i < rango.count; i++) {
        doc.switchToPage(rango.start + i);
        const mb = doc.page.margins.bottom;
        doc.page.margins.bottom = 0;
        doc.fillColor('#888').font('Helvetica').fontSize(7)
           .text(`TECSA — Reporte de taller ${E.periodo}${rep.planta ? ' — ' + rep.planta : ''}   ·   Pag. ${i + 1} de ${rango.count}`,
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
      res.json({ ok: true, reporte: await construirReporte(sheets, periodoDeQuery(req.query)) });
    } catch (e) { console.error('[ReporteDiario]', e.message); res.json({ ok: false, error: e.message }); }
  });

  app.get('/api/reporte-diario-pdf', async (req, res) => {
    try {
      const { sheets } = await getClients();
      const rep = await construirReporte(sheets, periodoDeQuery(req.query));
      const pdf = await generarPDF(rep);
      const f = s => s.replace(/\//g, '-');
      const nombre = `Reporte_Taller_${f(rep.desde)}${rep.esDia ? '' : '_a_' + f(rep.hasta)}${rep.planta ? '_' + rep.planta.replace(/[^A-Z0-9]+/g, '-') : ''}.pdf`;
      res.setHeader('Content-Type', 'application/pdf');
      res.setHeader('Content-Disposition', `attachment; filename="${nombre}"`);
      res.send(pdf);
    } catch (e) { console.error('[ReporteDiario PDF]', e.message); res.status(500).send('Error generando PDF: ' + e.message); }
  });

  app.post('/api/reporte-diario/sheets', async (req, res) => {
    try {
      const { sheets } = await getClients();
      const rep = await construirReporte(sheets, periodoDeQuery(req.query));
      res.json({ ok: true, hoja: await escribirPestana(sheets, rep) });
    } catch (e) { console.error('[ReporteDiario Sheets]', e.message); res.json({ ok: false, error: e.message }); }
  });

  // ── Generacion automatica del dia (todas las plantas) ──
  let ultimoDia = null;
  const [hAuto, mAuto] = HORA_AUTO.split(':').map(n => parseInt(n, 10) || 0);
  setInterval(async () => {
    const p = partesMty();
    const hoyTxt = fechaStr(p);
    if (ultimoDia === hoyTxt) return;
    if (p.hh * 60 + p.mm < hAuto * 60 + mAuto) return;
    ultimoDia = hoyTxt;
    try {
      const { sheets } = await getClients();
      await escribirPestana(sheets, await construirReporte(sheets, { desde: hoyTxt, hasta: hoyTxt, planta: '' }));
    } catch (e) {
      console.error('[ReporteDiario auto]', e.message);
      ultimoDia = null;
    }
  }, 60 * 1000);

  console.log(`[ReporteDiario] Automatico a las ${HORA_AUTO} (hora de Monterrey) → ${SHEET_ID_DIARIO === SHEET_ID ? 'Sheet principal' : 'SHEET_ID_DIARIO'}`);
};
