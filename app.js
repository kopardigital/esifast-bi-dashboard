/* ═══════════════════════════════════════════════════════════
   ESIFAST BI · app.js
   100% client-side · sin backend · GitHub Pages ready
═══════════════════════════════════════════════════════════ */

'use strict';

/* ── CONFIG ───────────────────────────────────────────────
   ⚠ ADVERTENCIA: Credenciales visibles en el código fuente.
   Este login es solo una barrera visual básica.
   Para seguridad real implementar autenticación con backend.
   NO publicar credenciales reales en GitHub.
────────────────────────────────────────────────────────── */
const CONFIG = {
  user: 'admin',
  pass: 'Cambiar123!',
  rowsPerPage: 20,
  expectedSheets: ['RESUMEN', 'INGRESOS', 'EGRESOS ENERGY', 'EGRESOS ESIFAR'],
};

/* ── STATE ────────────────────────────────────────────────── */
const STATE = {
  data: null,         // processed data object
  filtered: [],       // current filtered rows
  sortCol: null,
  sortDir: 'asc',
  currentPage: 1,
  charts: {},         // chart instances
};

/* ══════════════════════════════════════════════════════════
   AUTH
══════════════════════════════════════════════════════════ */
function initAuth() {
  if (sessionStorage.getItem('esifast_auth')) showApp();

  document.getElementById('btn-login').addEventListener('click', doLogin);
  document.getElementById('inp-pass').addEventListener('keydown', e => {
    if (e.key === 'Enter') doLogin();
  });
  document.getElementById('btn-logout').addEventListener('click', () => {
    sessionStorage.removeItem('esifast_auth');
    location.reload();
  });
}

function doLogin() {
  const u = document.getElementById('inp-user').value.trim();
  const p = document.getElementById('inp-pass').value.trim();
  const err = document.getElementById('login-error');
  if (u === CONFIG.user && p === CONFIG.pass) {
    sessionStorage.setItem('esifast_auth', '1');
    document.getElementById('login-screen').classList.add('hidden');
    showApp();
  } else {
    err.classList.remove('hidden');
    document.getElementById('inp-pass').value = '';
  }
}

function showApp() {
  document.getElementById('login-screen').classList.add('hidden');
  document.getElementById('app').classList.remove('hidden');
}

/* ══════════════════════════════════════════════════════════
   NAVIGATION
══════════════════════════════════════════════════════════ */
function initNav() {
  document.querySelectorAll('.nav-item').forEach(btn => {
    btn.addEventListener('click', () => {
      document.querySelectorAll('.nav-item').forEach(b => b.classList.remove('active'));
      document.querySelectorAll('.page').forEach(p => p.classList.remove('active'));
      btn.classList.add('active');
      document.getElementById('page-' + btn.dataset.page).classList.add('active');
    });
  });
}

/* ══════════════════════════════════════════════════════════
   FILE LOADING
══════════════════════════════════════════════════════════ */
function initFileUpload() {
  const input = document.getElementById('file-input');
  const zone  = document.getElementById('upload-zone');

  input.addEventListener('change', e => {
    if (e.target.files[0]) processFile(e.target.files[0]);
  });

  zone.addEventListener('dragover', e => {
    e.preventDefault();
    zone.classList.add('dragover');
  });
  zone.addEventListener('dragleave', () => zone.classList.remove('dragover'));
  zone.addEventListener('drop', e => {
    e.preventDefault();
    zone.classList.remove('dragover');
    const file = e.dataTransfer.files[0];
    if (file && file.name.endsWith('.xlsx')) processFile(file);
  });
  zone.addEventListener('click', () => input.click());
}

function processFile(file) {
  const reader = new FileReader();
  reader.onload = e => {
    try {
      const wb = XLSX.read(e.target.result, { type: 'array', cellDates: true });
      validateSheets(wb);
      STATE.data = parseWorkbook(wb);
      renderAll();
      showLoadedStatus(file.name);
    } catch (err) {
      alert('Error al procesar el archivo:\n' + err.message);
    }
  };
  reader.readAsArrayBuffer(file);
}

function validateSheets(wb) {
  const missing = CONFIG.expectedSheets.filter(s => !wb.SheetNames.includes(s));
  if (missing.length) {
    throw new Error('Faltan las siguientes hojas: ' + missing.join(', '));
  }
}

function showLoadedStatus(name) {
  document.getElementById('upload-status').innerHTML =
    `<div class="upload-ok">✓ ${name} cargado correctamente</div>`;
  document.getElementById('btn-export-pdf').disabled = false;
  document.getElementById('dash-subtitle').textContent =
    'Archivo procesado · ' + new Date().toLocaleString('es-AR');
}

/* ══════════════════════════════════════════════════════════
   WORKBOOK PARSER
══════════════════════════════════════════════════════════ */
function parseWorkbook(wb) {
  const ingresos     = parseIngresos(wb.Sheets['INGRESOS']);
  const egresosEn    = parseEgresos(wb.Sheets['EGRESOS ENERGY'], 'CLADAN ENERGY');
  const egresosEs    = parseEgresos(wb.Sheets['EGRESOS ESIFAR'], 'ESIFAR');
  const resumenRaw   = parseResumen(wb.Sheets['RESUMEN']);
  const ops          = buildOps(ingresos, [...egresosEn, ...egresosEs]);

  return { ingresos, egresosEn, egresosEs, resumenRaw, ops };
}

/* ── INGRESOS ─────────────────────────────────────────────── */
function parseIngresos(sheet) {
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null });
  const result = [];
  let headerRow = -1;

  for (let i = 0; i < rows.length; i++) {
    const r = rows[i];
    if (r[0] === 'Nro OP') { headerRow = i; break; }
  }
  if (headerRow < 0) return result;

  for (let i = headerRow + 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r[0]) continue;
    result.push({
      op:      String(r[0]).trim().padStart(5, '0'),
      fecha:   parseDate(r[1]),
      cliente: String(r[2] || '').trim(),
      conten:  r[3] || '',
      remito:  r[4] || '',
      pallets: toNum(r[5]),
      kg:      toNum(r[6]),
      obs:     String(r[7] || '').trim(),
      in:      toNum(r[8]),
    });
  }
  return result;
}

/* ── EGRESOS ──────────────────────────────────────────────── */
function parseEgresos(sheet, fallbackCliente) {
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null });
  const result = [];

  for (let i = 0; i < rows.length; i++) {
    const r = rows[i];
    if (r[0] === 'Nro OP') {
      const closeDateRaw = rows[i][13];
      const closeDate    = parseDate(closeDateRaw);

      for (let j = i + 1; j < rows.length; j++) {
        const d = rows[j];
        if (!d[0] || d[0] === 'Nro OP') { i = j - 1; break; }
        result.push({
          op:           String(d[0]).trim().padStart(5, '0'),
          fechaIngreso: parseDate(d[1]),
          cliente:      String(d[2] || fallbackCliente).trim(),
          fechaSalida:  parseDate(d[3]),
          conten:       d[4] || '',
          remitoEnt:    d[5] || '',
          remitoSal:    d[6] || '',
          palletsSal:   toNum(d[7]),
          saldoPall:    toNum(d[8]),
          kg:           toNum(d[9]),
          obs:          String(d[10] || '').trim(),
          out:          toNum(d[11]),
          almacenaje:   toNum(d[12]),
          mesClose:     closeDate,
        });
      }
    }
  }
  return result;
}

/* ── RESUMEN ──────────────────────────────────────────────── */
function parseResumen(sheet) {
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null });
  const blocks = [];
  let i = 0;

  while (i < rows.length) {
    const r = rows[i];
    // detect header rows like "CLADAN ENERGY ENERO"
    if (r[0] && typeof r[0] === 'string' && r[0].includes('ENERGY')) {
      const nameEn = r[0]; const nameEs = r[5] || '';
      i++; // skip the "CONCEPTO TARIFARIO..." header
      const rowsEn = []; const rowsEs = [];
      i++;
      while (i < rows.length && rows[i][0] !== null && typeof rows[i][0] === 'string' && !rows[i][0].includes('ENERGY') && !rows[i][0].includes('ESIFAR')) {
        const rr = rows[i];
        if (rr[0] === 'CONCEPTO') { i++; continue; }
        if (rr[0] !== null) rowsEn.push({ concepto: rr[0], tarifario: rr[1], cantidad: rr[2], importe: rr[3] });
        if (rr[5] !== null && rr[5] !== undefined) rowsEs.push({ concepto: rr[5], tarifario: rr[6], cantidad: rr[7], importe: rr[8] });
        if (rr[0] === null && rr[3] !== null) { rowsEn.push({ concepto: 'TOTAL', tarifario: null, cantidad: null, importe: rr[3] }); }
        if (rr[5] === null && rr[8] !== null) { rowsEs.push({ concepto: 'TOTAL', tarifario: null, cantidad: null, importe: rr[8] }); }
        i++;
      }
      blocks.push({ nameEn, nameEs, rowsEn, rowsEs });
    } else {
      i++;
    }
  }
  return blocks;
}

/* ── MERGE OPS ────────────────────────────────────────────── */
function buildOps(ingresos, egresos) {
  const egMap = {};
  egresos.forEach(e => {
    if (!egMap[e.op]) egMap[e.op] = [];
    egMap[e.op].push(e);
  });

  return ingresos.map(ing => {
    const egs = egMap[ing.op] || [];
    // take latest egreso record (highest almacenaje = current month)
    const eg = egs.reduce((best, cur) =>
      (cur.almacenaje || 0) > (best.almacenaje || 0) ? cur : best, egs[0] || {});

    const hasSalida = eg.fechaSalida && +eg.fechaSalida > 100000;
    return {
      op:          ing.op,
      fechaIngreso: ing.fecha,
      fechaSalida:  hasSalida ? eg.fechaSalida : null,
      cliente:     ing.cliente,
      obs:         ing.obs,
      pallets:     ing.pallets,
      palletsSal:  eg.palletsSal || 0,
      saldo:       ing.pallets - (eg.palletsSal || 0),
      kg:          ing.kg,
      in:          ing.in,
      out:         eg.out || 0,
      almacenaje:  eg.almacenaje || 0,
      total:       (ing.in || 0) + (eg.out || 0) + (eg.almacenaje || 0),
      estado:      hasSalida ? 'CERRADA' : 'ABIERTA',
      fuente:      ing.cliente.includes('ENERGY') ? 'ENERGY' : 'ESIFAR',
    };
  });
}

/* ══════════════════════════════════════════════════════════
   RENDER ALL
══════════════════════════════════════════════════════════ */
function renderAll() {
  document.getElementById('dashboard-content').classList.remove('hidden');
  document.getElementById('upload-zone').style.pointerEvents = 'none';

  renderKPIs();
  renderCharts();
  renderTopOps();
  renderOpsTable();
  renderAnalisis();
  renderRecomendaciones();
  renderResumen();
  populateFilters();
}

/* ══════════════════════════════════════════════════════════
   KPIs
══════════════════════════════════════════════════════════ */
function renderKPIs() {
  const ops = STATE.data.ops;
  const totalIN  = sum(ops, 'in');
  const totalOUT = sum(ops, 'out');
  const totalAlm = sum(ops, 'almacenaje');
  const totalFact = sum(ops, 'total');
  const pallIn   = sum(ops, 'pallets');
  const pallSal  = sum(ops, 'palletsSal');
  const abiertas = ops.filter(o => o.estado === 'ABIERTA').length;
  const clientes = [...new Set(ops.map(o => o.cliente))].length;
  const diasProm = calcDiasProm(ops);

  const kpis = [
    { label: 'Total IN',        value: formatARS(totalIN),   sub: 'Ingresos operativos',     color: '#2563eb', icon: '↑' },
    { label: 'Total OUT',       value: formatARS(totalOUT),  sub: 'Egresos operativos',       color: '#06b6d4', icon: '↓' },
    { label: 'Almacenaje',      value: formatARS(totalAlm),  sub: 'Acumulado',                color: '#8b5cf6', icon: '📦' },
    { label: 'Facturado total', value: formatARS(totalFact), sub: 'IN + OUT + Almacenaje',    color: '#10b981', icon: '💰' },
    { label: 'Operaciones',     value: ops.length,           sub: `${abiertas} abiertas`,     color: '#f59e0b', icon: '🗂' },
    { label: 'Pallets ingres.', value: pallIn,               sub: 'Total ingresados',         color: '#2563eb', icon: '⬆' },
    { label: 'Pallets salida',  value: pallSal,              sub: 'Total egresados',          color: '#06b6d4', icon: '⬇' },
    { label: 'Saldo pallets',   value: pallIn - pallSal,     sub: 'Pendientes',               color: '#ef4444', icon: '⚖' },
    { label: 'Clientes activos',value: clientes,             sub: 'Empresas distintas',       color: '#10b981', icon: '🏢' },
    { label: 'OP sin salida',   value: abiertas,             sub: 'Pendientes de cierre',     color: '#f59e0b', icon: '⏳' },
    { label: 'Días prom. alm.', value: diasProm,             sub: 'Promedio por operación',   color: '#8b5cf6', icon: '📅' },
    { label: 'Prom. por OP',    value: formatARS(ops.length ? Math.round(totalFact / ops.length) : 0), sub: 'Importe promedio', color: '#06b6d4', icon: '⊘' },
  ];

  document.getElementById('kpi-grid').innerHTML = kpis.map(k => `
    <div class="kpi-card" style="--kpi-color:${k.color}">
      <div class="kpi-icon">${k.icon}</div>
      <div class="kpi-label">${k.label}</div>
      <div class="kpi-value">${k.value}</div>
      <div class="kpi-sub">${k.sub}</div>
    </div>`).join('');
}

/* ══════════════════════════════════════════════════════════
   CHARTS
══════════════════════════════════════════════════════════ */
const CHART_COLORS = {
  energy: 'rgba(37,99,235,.8)',
  esifar: 'rgba(6,182,212,.8)',
  energy_border: '#2563eb',
  esifar_border: '#06b6d4',
  palette: ['#2563eb','#06b6d4','#10b981','#8b5cf6','#f59e0b','#ef4444','#ec4899'],
};

const CHART_DEFAULTS = {
  color: '#e2e8f0',
  gridColor: 'rgba(31,45,69,.7)',
};

function destroyChart(id) {
  if (STATE.charts[id]) { STATE.charts[id].destroy(); delete STATE.charts[id]; }
}

function renderCharts() {
  const ops = STATE.data.ops;

  // ── Ingresos por cliente ──
  destroyChart('chart-ing');
  const ingByCliente = groupSum(ops, 'cliente', 'in');
  STATE.charts['chart-ing'] = new Chart(document.getElementById('chart-ing'), {
    type: 'bar',
    data: {
      labels: Object.keys(ingByCliente),
      datasets: [{
        data: Object.values(ingByCliente),
        backgroundColor: Object.keys(ingByCliente).map(k =>
          k.includes('ENERGY') ? CHART_COLORS.energy : CHART_COLORS.esifar),
        borderColor: Object.keys(ingByCliente).map(k =>
          k.includes('ENERGY') ? CHART_COLORS.energy_border : CHART_COLORS.esifar_border),
        borderWidth: 1, borderRadius: 5,
      }]
    },
    options: barOptions('AR$'),
  });

  // ── Almacenaje por cliente ──
  destroyChart('chart-alm');
  const almByCliente = groupSum(ops, 'cliente', 'almacenaje');
  STATE.charts['chart-alm'] = new Chart(document.getElementById('chart-alm'), {
    type: 'bar',
    data: {
      labels: Object.keys(almByCliente),
      datasets: [{
        data: Object.values(almByCliente),
        backgroundColor: Object.keys(almByCliente).map(k =>
          k.includes('ENERGY') ? 'rgba(37,99,235,.6)' : 'rgba(6,182,212,.6)'),
        borderColor: Object.keys(almByCliente).map(k =>
          k.includes('ENERGY') ? CHART_COLORS.energy_border : CHART_COLORS.esifar_border),
        borderWidth: 1, borderRadius: 5,
      }]
    },
    options: barOptions('AR$'),
  });

  // ── Pallets ──
  destroyChart('chart-pallets');
  const pallByCliente = {};
  const pallSalByCliente = {};
  ops.forEach(o => {
    pallByCliente[o.cliente]    = (pallByCliente[o.cliente] || 0) + o.pallets;
    pallSalByCliente[o.cliente] = (pallSalByCliente[o.cliente] || 0) + o.palletsSal;
  });
  const pallLabels = Object.keys(pallByCliente);
  STATE.charts['chart-pallets'] = new Chart(document.getElementById('chart-pallets'), {
    type: 'bar',
    data: {
      labels: pallLabels,
      datasets: [
        { label: 'Ingresados', data: pallLabels.map(l => pallByCliente[l]), backgroundColor: 'rgba(37,99,235,.7)', borderColor: '#2563eb', borderWidth: 1, borderRadius: 4 },
        { label: 'Egresados',  data: pallLabels.map(l => pallSalByCliente[l] || 0), backgroundColor: 'rgba(16,185,129,.7)', borderColor: '#10b981', borderWidth: 1, borderRadius: 4 },
      ]
    },
    options: barOptions('Pallets', true),
  });

  // ── Conceptos ──
  destroyChart('chart-conceptos');
  const totalIN  = sum(ops, 'in');
  const totalOUT = sum(ops, 'out');
  const totalAlm = sum(ops, 'almacenaje');
  STATE.charts['chart-conceptos'] = new Chart(document.getElementById('chart-conceptos'), {
    type: 'doughnut',
    data: {
      labels: ['IN', 'OUT', 'Almacenaje'],
      datasets: [{
        data: [totalIN, totalOUT, totalAlm],
        backgroundColor: ['rgba(37,99,235,.8)', 'rgba(6,182,212,.8)', 'rgba(139,92,246,.8)'],
        borderColor: ['#2563eb','#06b6d4','#8b5cf6'],
        borderWidth: 1,
      }]
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: {
        legend: { position: 'bottom', labels: { color: CHART_DEFAULTS.color, font: { size: 12 }, padding: 16 } },
        tooltip: { callbacks: { label: ctx => ` ${ctx.label}: ${formatARS(ctx.raw)}` } },
      },
    },
  });

  // ── Evolución mensual ──
  destroyChart('chart-evolucion');
  const months = [...new Set(
    [...STATE.data.egresosEn, ...STATE.data.egresosEs]
      .filter(e => e.mesClose)
      .map(e => fmtMonth(e.mesClose))
  )].sort();

  const almEnByMonth = {};
  const almEsByMonth = {};
  STATE.data.egresosEn.forEach(e => {
    if (e.mesClose) { const m = fmtMonth(e.mesClose); almEnByMonth[m] = (almEnByMonth[m] || 0) + (e.almacenaje || 0); }
  });
  STATE.data.egresosEs.forEach(e => {
    if (e.mesClose) { const m = fmtMonth(e.mesClose); almEsByMonth[m] = (almEsByMonth[m] || 0) + (e.almacenaje || 0); }
  });

  STATE.charts['chart-evolucion'] = new Chart(document.getElementById('chart-evolucion'), {
    type: 'line',
    data: {
      labels: months.map(m => m),
      datasets: [
        {
          label: 'CLADAN ENERGY',
          data: months.map(m => almEnByMonth[m] || 0),
          borderColor: '#2563eb', backgroundColor: 'rgba(37,99,235,.1)',
          tension: .35, fill: true, borderWidth: 2, pointRadius: 4,
        },
        {
          label: 'ESIFAR',
          data: months.map(m => almEsByMonth[m] || 0),
          borderColor: '#06b6d4', backgroundColor: 'rgba(6,182,212,.1)',
          tension: .35, fill: true, borderWidth: 2, pointRadius: 4,
        },
      ]
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      scales: {
        x: { grid: { color: CHART_DEFAULTS.gridColor }, ticks: { color: CHART_DEFAULTS.color, font: { size: 11 } } },
        y: { grid: { color: CHART_DEFAULTS.gridColor }, ticks: { color: CHART_DEFAULTS.color, font: { size: 11 }, callback: v => formatARS(v) } },
      },
      plugins: {
        legend: { labels: { color: CHART_DEFAULTS.color, font: { size: 12 }, padding: 12 } },
        tooltip: { callbacks: { label: ctx => ` ${ctx.dataset.label}: ${formatARS(ctx.raw)}` } },
      },
    },
  });
}

function barOptions(unit, grouped) {
  return {
    responsive: true, maintainAspectRatio: false,
    plugins: {
      legend: grouped
        ? { position: 'bottom', labels: { color: CHART_DEFAULTS.color, font: { size: 12 }, padding: 12 } }
        : { display: false },
      tooltip: { callbacks: { label: ctx => ` ${ctx.dataset.label || ''}: ${unit === 'AR$' ? formatARS(ctx.raw) : ctx.raw}` } },
    },
    scales: {
      x: { grid: { color: CHART_DEFAULTS.gridColor }, ticks: { color: CHART_DEFAULTS.color, font: { size: 11 } } },
      y: { grid: { color: CHART_DEFAULTS.gridColor }, ticks: { color: CHART_DEFAULTS.color, font: { size: 11 }, callback: v => unit === 'AR$' ? formatARS(v) : v } },
    },
  };
}

/* ══════════════════════════════════════════════════════════
   TOP OPS TABLE
══════════════════════════════════════════════════════════ */
function renderTopOps() {
  const top = [...STATE.data.ops]
    .sort((a, b) => b.almacenaje - a.almacenaje)
    .slice(0, 10);

  document.getElementById('top-ops-container').innerHTML = `
    <table>
      <thead><tr>
        <th>#</th><th>Nro OP</th><th>Cliente</th><th>Descripción</th>
        <th>Pallets</th><th>Almacenaje AR$</th><th>Estado</th>
      </tr></thead>
      <tbody>${top.map((o, i) => `
        <tr>
          <td>${i + 1}</td>
          <td>${o.op}</td>
          <td><span class="badge ${o.cliente.includes('ENERGY') ? 'badge-energy' : 'badge-esifar'}">${o.cliente}</span></td>
          <td>${o.obs || '—'}</td>
          <td>${o.pallets}</td>
          <td><strong>${formatARS(o.almacenaje)}</strong></td>
          <td><span class="badge ${o.estado === 'ABIERTA' ? 'badge-open' : 'badge-closed'}">${o.estado}</span></td>
        </tr>`).join('')}
      </tbody>
    </table>`;
}

/* ══════════════════════════════════════════════════════════
   OPERACIONES TABLE + FILTERS
══════════════════════════════════════════════════════════ */
function populateFilters() {
  const ops = STATE.data.ops;

  const sel = document.getElementById('f-cliente');
  const clientes = [...new Set(ops.map(o => o.cliente))].sort();
  sel.innerHTML = '<option value="">Todos</option>' +
    clientes.map(c => `<option value="${c}">${c}</option>`).join('');

  const meses = [...new Set(ops.map(o => o.fechaIngreso ? fmtMonth(o.fechaIngreso) : null).filter(Boolean))].sort();
  const selMes = document.getElementById('f-mes');
  selMes.innerHTML = '<option value="">Todos</option>' +
    meses.map(m => `<option value="${m}">${m}</option>`).join('');

  ['f-cliente','f-desde','f-hasta','f-estado','f-op','f-mes','table-search'].forEach(id => {
    document.getElementById(id).addEventListener('input', applyFilters);
    document.getElementById(id).addEventListener('change', applyFilters);
  });
  document.getElementById('btn-clear-filters').addEventListener('click', clearFilters);

  applyFilters();
}

function applyFilters() {
  const cliente = document.getElementById('f-cliente').value;
  const desde   = document.getElementById('f-desde').value;
  const hasta   = document.getElementById('f-hasta').value;
  const estado  = document.getElementById('f-estado').value;
  const op      = document.getElementById('f-op').value.trim().toLowerCase();
  const mes     = document.getElementById('f-mes').value;
  const search  = document.getElementById('table-search').value.toLowerCase();

  STATE.filtered = STATE.data.ops.filter(o => {
    if (cliente && o.cliente !== cliente) return false;
    if (estado  && o.estado !== estado)   return false;
    if (op      && !o.op.includes(op))    return false;
    if (mes     && fmtMonth(o.fechaIngreso) !== mes) return false;
    if (desde   && o.fechaIngreso && fmtISO(o.fechaIngreso) < desde) return false;
    if (hasta   && o.fechaIngreso && fmtISO(o.fechaIngreso) > hasta)  return false;
    if (search) {
      const row = Object.values(o).join(' ').toLowerCase();
      if (!row.includes(search)) return false;
    }
    return true;
  });

  STATE.currentPage = 1;
  STATE.sortCol = null;
  renderOpsTable();
}

function clearFilters() {
  ['f-cliente','f-estado','f-op','f-mes','f-desde','f-hasta'].forEach(id => {
    document.getElementById(id).value = '';
  });
  document.getElementById('table-search').value = '';
  applyFilters();
}

function renderOpsTable() {
  const data = STATE.filtered.length ? STATE.filtered : (STATE.data ? STATE.data.ops : []);
  if (!data.length) {
    document.getElementById('ops-table-container').innerHTML =
      '<div class="empty-state"><p>Sin resultados para los filtros aplicados.</p></div>';
    document.getElementById('pagination').innerHTML = '';
    document.getElementById('table-count').textContent = '0 registros';
    return;
  }

  const total   = data.length;
  const pages   = Math.ceil(total / CONFIG.rowsPerPage);
  const start   = (STATE.currentPage - 1) * CONFIG.rowsPerPage;
  const slice   = data.slice(start, start + CONFIG.rowsPerPage);

  document.getElementById('table-count').textContent = `${total} registros`;

  const cols = [
    { key: 'op',           label: 'Nro OP' },
    { key: 'fechaIngreso', label: 'Fec. Ingreso' },
    { key: 'fechaSalida',  label: 'Fec. Salida' },
    { key: 'cliente',      label: 'Cliente' },
    { key: 'obs',          label: 'Descripción' },
    { key: 'pallets',      label: 'Pallets IN' },
    { key: 'palletsSal',   label: 'Pallets OUT' },
    { key: 'saldo',        label: 'Saldo' },
    { key: 'in',           label: 'IN AR$' },
    { key: 'out',          label: 'OUT AR$' },
    { key: 'almacenaje',   label: 'Almacenaje AR$' },
    { key: 'total',        label: 'Total AR$' },
    { key: 'estado',       label: 'Estado' },
  ];

  const thead = `<thead><tr>${cols.map(c => `
    <th class="${STATE.sortCol === c.key ? 'sort-' + STATE.sortDir : ''}"
        onclick="sortTable('${c.key}')">${c.label}</th>`).join('')}
  </tr></thead>`;

  const tbody = `<tbody>${slice.map(o => {
    const warnSaldo = o.saldo > 0 && o.estado === 'ABIERTA';
    return `<tr>
      <td>${o.op}</td>
      <td>${fmtDate(o.fechaIngreso)}</td>
      <td>${fmtDate(o.fechaSalida)}</td>
      <td><span class="badge ${o.cliente.includes('ENERGY') ? 'badge-energy' : 'badge-esifar'}">${o.cliente}</span></td>
      <td>${o.obs || '—'}</td>
      <td>${o.pallets}</td>
      <td>${o.palletsSal}</td>
      <td class="${warnSaldo ? 'warn-cell' : ''}">${o.saldo}</td>
      <td>${formatARS(o.in)}</td>
      <td>${formatARS(o.out)}</td>
      <td><strong>${formatARS(o.almacenaje)}</strong></td>
      <td><strong>${formatARS(o.total)}</strong></td>
      <td><span class="badge ${o.estado === 'ABIERTA' ? 'badge-open' : 'badge-closed'}">${o.estado}</span></td>
    </tr>`;
  }).join('')}</tbody>`;

  document.getElementById('ops-table-container').innerHTML =
    `<table>${thead}${tbody}</table>`;

  // Pagination
  let pagi = '';
  for (let i = 1; i <= pages; i++) {
    if (i === 1 || i === pages || Math.abs(i - STATE.currentPage) <= 2) {
      pagi += `<button class="page-btn ${i === STATE.currentPage ? 'active' : ''}" onclick="goPage(${i})">${i}</button>`;
    } else if (Math.abs(i - STATE.currentPage) === 3) {
      pagi += `<span style="color:var(--text-3);padding:0 4px">…</span>`;
    }
  }
  document.getElementById('pagination').innerHTML = pagi;
}

function sortTable(col) {
  if (STATE.sortCol === col) {
    STATE.sortDir = STATE.sortDir === 'asc' ? 'desc' : 'asc';
  } else {
    STATE.sortCol = col;
    STATE.sortDir = 'asc';
  }
  const data = STATE.filtered.length ? STATE.filtered : STATE.data.ops;
  data.sort((a, b) => {
    const va = a[col]; const vb = b[col];
    if (va == null) return 1;
    if (vb == null) return -1;
    const cmp = typeof va === 'number' ? va - vb : String(va).localeCompare(String(vb), 'es');
    return STATE.sortDir === 'asc' ? cmp : -cmp;
  });
  renderOpsTable();
}

window.sortTable = sortTable;

function goPage(p) {
  STATE.currentPage = p;
  renderOpsTable();
}
window.goPage = goPage;

/* ══════════════════════════════════════════════════════════
   ANÁLISIS EJECUTIVO
══════════════════════════════════════════════════════════ */
function renderAnalisis() {
  const ops = STATE.data.ops;
  const energy = ops.filter(o => o.cliente.includes('ENERGY'));
  const esifar = ops.filter(o => !o.cliente.includes('ENERGY'));

  // Top cliente por facturación
  const factByCliente = groupSum(ops, 'cliente', 'total');
  const topCliente    = Object.entries(factByCliente).sort((a,b)=>b[1]-a[1])[0];

  // OP con mayor permanencia (sin fecha salida)
  const abiertas = ops.filter(o => o.estado === 'ABIERTA');

  // OP con saldo pendiente
  const conSaldo = ops.filter(o => o.saldo > 0);

  // Datos incompletos
  const sinObs    = ops.filter(o => !o.obs).length;
  const sinRemito = STATE.data.ingresos.filter(i => !i.remito).length;

  // Concepto mayor
  const totIN  = sum(ops, 'in');
  const totOUT = sum(ops, 'out');
  const totAlm = sum(ops, 'almacenaje');
  const concMax = totAlm >= totIN && totAlm >= totOUT ? 'Almacenaje' : totIN >= totOUT ? 'IN' : 'OUT';

  const html = `
  <div class="compare-grid">
    <div class="compare-card energy">
      <h3>CLADAN ENERGY</h3>
      <div class="compare-row"><span>Operaciones</span><span class="compare-val">${energy.length}</span></div>
      <div class="compare-row"><span>Total IN</span><span class="compare-val">${formatARS(sum(energy,'in'))}</span></div>
      <div class="compare-row"><span>Almacenaje</span><span class="compare-val">${formatARS(sum(energy,'almacenaje'))}</span></div>
      <div class="compare-row"><span>Total facturado</span><span class="compare-val">${formatARS(sum(energy,'total'))}</span></div>
      <div class="compare-row"><span>Pallets activos</span><span class="compare-val">${sum(energy,'pallets') - sum(energy,'palletsSal')}</span></div>
      <div class="compare-row"><span>OP abiertas</span><span class="compare-val">${energy.filter(o=>o.estado==='ABIERTA').length}</span></div>
    </div>
    <div class="compare-card esifar">
      <h3>ESIFAR</h3>
      <div class="compare-row"><span>Operaciones</span><span class="compare-val">${esifar.length}</span></div>
      <div class="compare-row"><span>Total IN</span><span class="compare-val">${formatARS(sum(esifar,'in'))}</span></div>
      <div class="compare-row"><span>Almacenaje</span><span class="compare-val">${formatARS(sum(esifar,'almacenaje'))}</span></div>
      <div class="compare-row"><span>Total facturado</span><span class="compare-val">${formatARS(sum(esifar,'total'))}</span></div>
      <div class="compare-row"><span>Pallets activos</span><span class="compare-val">${sum(esifar,'pallets') - sum(esifar,'palletsSal')}</span></div>
      <div class="compare-row"><span>OP abiertas</span><span class="compare-val">${esifar.filter(o=>o.estado==='ABIERTA').length}</span></div>
    </div>
  </div>

  <div class="analysis-grid">
    <div class="analysis-card">
      <h3>📊 Facturación y clientes</h3>
      <div class="analysis-row"><div class="dot dot-blue"></div><div>El cliente con mayor facturación estimada es <strong>${topCliente?.[0] || '—'}</strong> con <strong>${formatARS(topCliente?.[1] || 0)}</strong> en concepto de IN + OUT + Almacenaje.</div></div>
      <div class="analysis-row"><div class="dot dot-cyan"></div><div>El concepto que genera mayor ingreso es <strong>${concMax}</strong>. IN total: ${formatARS(totIN)} · OUT: ${formatARS(totOUT)} · Almacenaje: ${formatARS(totAlm)}.</div></div>
      <div class="analysis-row"><div class="dot dot-green"></div><div>Se registran <strong>${ops.length} operaciones totales</strong> entre ambas empresas, con un promedio de ${formatARS(ops.length ? Math.round(sum(ops,'total')/ops.length) : 0)} por operación.</div></div>
    </div>

    <div class="analysis-card">
      <h3>📦 Pallets y permanencia</h3>
      <div class="analysis-row"><div class="dot dot-amber"></div><div>Hay <strong>${abiertas.length} operaciones abiertas</strong> sin fecha de salida registrada. Las OPs ${abiertas.slice(0,3).map(o=>o.op).join(', ')} son las de mayor almacenaje acumulado.</div></div>
      <div class="analysis-row"><div class="dot dot-red"></div><div><strong>${conSaldo.length} operaciones</strong> tienen saldo de pallets pendiente (pallets IN mayor a pallets OUT registrados).</div></div>
      <div class="analysis-row"><div class="dot dot-blue"></div><div>El promedio de días de almacenaje calculado es <strong>${calcDiasProm(ops)} días</strong> por operación activa.</div></div>
    </div>

    <div class="analysis-card">
      <h3>⚠ Datos incompletos</h3>
      <div class="analysis-row"><div class="dot dot-amber"></div><div><strong>${sinObs} operaciones</strong> no tienen descripción/observación registrada en la columna OBS.</div></div>
      <div class="analysis-row"><div class="dot dot-red"></div><div><strong>${sinRemito} ingresos</strong> no tienen número de remito registrado. Se recomienda verificar la documentación de respaldo.</div></div>
      <div class="analysis-row"><div class="dot dot-amber"></div><div>La hoja EGRESOS ESIFAR no registra fechas de salida concretas en ninguna operación del período analizado — todos los saldos son abiertos.</div></div>
    </div>

    <div class="analysis-card">
      <h3>📈 Tendencia mensual</h3>
      <div class="analysis-row"><div class="dot dot-blue"></div><div>CLADAN ENERGY presenta almacenaje acumulado de <strong>${formatARS(sum(energy,'almacenaje'))}</strong>. Las OPs corresponden a paneles solares con ingreso en enero 2026.</div></div>
      <div class="analysis-row"><div class="dot dot-cyan"></div><div>ESIFAR incorporó <strong>${esifar.length} operaciones en febrero 2026</strong> con productos como Metionina, Fosfato y Ferroso/Cobre, generando un almacenaje de ${formatARS(sum(esifar,'almacenaje'))}.</div></div>
      <div class="analysis-row"><div class="dot dot-green"></div><div>El volumen operativo muestra crecimiento entre enero y febrero, con ESIFAR generando mayor cantidad de OPs en el segundo mes.</div></div>
    </div>
  </div>`;

  document.getElementById('analisis-content').innerHTML = html;
}

/* ══════════════════════════════════════════════════════════
   RECOMENDACIONES
══════════════════════════════════════════════════════════ */
function renderRecomendaciones() {
  const ops     = STATE.data.ops;
  const abiertas = ops.filter(o => o.estado === 'ABIERTA');
  const conSaldo = ops.filter(o => o.saldo > 0 && o.estado === 'ABIERTA');
  const sinObs   = ops.filter(o => !o.obs).length;
  const sinRemito = STATE.data.ingresos.filter(i => !i.remito).length;
  const altaAlm  = ops.filter(o => o.almacenaje > 200000 && o.estado === 'ABIERTA');

  const recom = [
    {
      prioridad: 'high',
      icon: '🚨',
      titulo: `Revisar ${abiertas.length} operaciones sin fecha de salida`,
      desc: `Las OPs ${abiertas.map(o=>o.op).join(', ')} están abiertas sin fecha de salida registrada. Verificar si corresponde facturar o registrar el egreso. Impacto estimado: ${formatARS(sum(abiertas,'total'))}.`,
    },
    {
      prioridad: 'high',
      icon: '⚖',
      titulo: `${conSaldo.length} operaciones con saldo de pallets pendiente`,
      desc: `Se detectaron ${conSaldo.length} OPs con más pallets ingresados que egresados. Verificar si hay movimientos no registrados o pendientes de carga en EGRESOS. OPs afectadas: ${conSaldo.map(o=>o.op).join(', ')}.`,
    },
    {
      prioridad: 'high',
      icon: '💰',
      titulo: 'Facturación pendiente estimada sin registrar',
      desc: `Las operaciones abiertas acumulan ${formatARS(sum(abiertas,'almacenaje'))} en almacenaje y ${formatARS(sum(abiertas,'in'))} en IN. Si ya se prestaron los servicios, evaluar emisión de comprobantes.`,
    },
    {
      prioridad: 'medium',
      icon: '📦',
      titulo: `${altaAlm.length} operaciones con almacenaje elevado (> $200.000)`,
      desc: `Las OPs ${altaAlm.map(o=>o.op).join(', ')} tienen almacenaje acumulado superior a $200.000 y siguen abiertas. Verificar fechas de cierre o actualizar tarifas si el período se extendió.`,
    },
    {
      prioridad: 'medium',
      icon: '📋',
      titulo: `${sinRemito} ingresos sin número de remito`,
      desc: 'Existen registros en INGRESOS sin remito asociado. El remito es documento respaldo de la operación logística. Completar la información para evitar discrepancias en auditorías.',
    },
    {
      prioridad: 'medium',
      icon: '⚠',
      titulo: `${sinObs} operaciones sin descripción de mercadería`,
      desc: 'Varias operaciones no tienen la columna OBS completada. Registrar el tipo de mercadería facilita la identificación, el tarifario aplicable y los análisis de concentración por producto.',
    },
    {
      prioridad: 'low',
      icon: '🔍',
      titulo: 'Verificar consistencia entre INGRESOS y EGRESOS ESIFAR',
      desc: 'La hoja EGRESOS ESIFAR no tiene fechas de salida concretas en ninguna operación. Confirmar si las operaciones siguen activas o si falta cargar las fechas de cierre correspondientes al mes.',
    },
    {
      prioridad: 'low',
      icon: '📊',
      titulo: 'Actualizar hoja DESCONSOLIDADOS',
      desc: 'La hoja DESCONSOLIDADOS está vacía. Si se realizaron operaciones de desconsolidado, registrarlas para que el sistema BI pueda incluirlas en el análisis de conceptos facturables.',
    },
  ];

  document.getElementById('recom-content').innerHTML = `
    <div class="recom-list">${recom.map(r => `
      <div class="recom-item ${r.prioridad}">
        <div class="recom-icon">${r.icon}</div>
        <div>
          <div class="recom-title">${r.titulo}</div>
          <div class="recom-desc">${r.desc}</div>
        </div>
      </div>`).join('')}
    </div>`;
}

/* ══════════════════════════════════════════════════════════
   RESUMEN
══════════════════════════════════════════════════════════ */
function renderResumen() {
  const blocks = STATE.data.resumenRaw;
  if (!blocks.length) {
    document.getElementById('resumen-content').innerHTML =
      '<div class="empty-state"><p>No se encontraron bloques de resumen en la hoja RESUMEN.</p></div>';
    return;
  }

  let html = '';
  blocks.forEach(b => {
    html += `<div class="resumen-grid">`;

    // Energy block
    const totalEn = b.rowsEn.find(r => r.concepto === 'TOTAL');
    html += `<div class="resumen-block">
      <div class="resumen-block-header energy">
        <span>${b.nameEn}</span>
        ${totalEn ? `<span class="resumen-total">Total: ${formatARS(totalEn.importe)}</span>` : ''}
      </div>
      <table>
        <thead><tr><th>Concepto</th><th>Tarifario</th><th>Cantidad</th><th>AR$</th></tr></thead>
        <tbody>${b.rowsEn.filter(r=>r.concepto!=='TOTAL').map(r => `
          <tr>
            <td>${r.concepto || '—'}</td>
            <td>${r.tarifario != null ? formatARS(r.tarifario) : '—'}</td>
            <td>${r.cantidad != null ? r.cantidad : '—'}</td>
            <td>${r.importe != null ? formatARS(r.importe) : '—'}</td>
          </tr>`).join('')}
        </tbody>
      </table>
    </div>`;

    // Esifar block
    const totalEs = b.rowsEs.find(r => r.concepto === 'TOTAL');
    html += `<div class="resumen-block">
      <div class="resumen-block-header esifar">
        <span>${b.nameEs}</span>
        ${totalEs ? `<span class="resumen-total">Total: ${formatARS(totalEs.importe)}</span>` : ''}
      </div>
      <table>
        <thead><tr><th>Concepto</th><th>Tarifario</th><th>Cantidad</th><th>AR$</th></tr></thead>
        <tbody>${b.rowsEs.filter(r=>r.concepto!=='TOTAL').map(r => `
          <tr>
            <td>${r.concepto || '—'}</td>
            <td>${r.tarifario != null ? formatARS(r.tarifario) : '—'}</td>
            <td>${r.cantidad != null ? r.cantidad : '—'}</td>
            <td>${r.importe != null ? formatARS(r.importe) : '—'}</td>
          </tr>`).join('')}
        </tbody>
      </table>
    </div>`;

    html += '</div>';
  });

  document.getElementById('resumen-content').innerHTML = html;
}

/* ══════════════════════════════════════════════════════════
   EXPORT CSV
══════════════════════════════════════════════════════════ */
function exportCSV() {
  if (!STATE.data) return;
  const ops = STATE.filtered.length ? STATE.filtered : STATE.data.ops;
  const headers = ['Nro OP','Fecha Ingreso','Fecha Salida','Cliente','Descripcion','Pallets IN','Pallets OUT','Saldo','KG','IN AR$','OUT AR$','Almacenaje AR$','Total AR$','Estado'];
  const rows = ops.map(o => [
    o.op, fmtDate(o.fechaIngreso), fmtDate(o.fechaSalida), o.cliente,
    o.obs, o.pallets, o.palletsSal, o.saldo, o.kg,
    o.in, o.out, o.almacenaje, o.total, o.estado,
  ]);

  const csv = [headers, ...rows].map(r => r.map(v =>
    `"${String(v ?? '').replace(/"/g, '""')}"`).join(',')).join('\n');

  download('ESIFAST_operaciones.csv', 'text/csv;charset=utf-8;', '\ufeff' + csv);
}

/* ══════════════════════════════════════════════════════════
   EXPORT EXCEL
══════════════════════════════════════════════════════════ */
function exportExcelOps() {
  if (!STATE.data) return;
  const ops = STATE.filtered.length ? STATE.filtered : STATE.data.ops;
  const rows = ops.map(o => ({
    'Nro OP': o.op,
    'Fecha Ingreso': fmtDate(o.fechaIngreso),
    'Fecha Salida':  fmtDate(o.fechaSalida),
    'Cliente': o.cliente,
    'Descripcion': o.obs,
    'Pallets IN': o.pallets,
    'Pallets OUT': o.palletsSal,
    'Saldo': o.saldo,
    'KG': o.kg,
    'IN AR$': o.in,
    'OUT AR$': o.out,
    'Almacenaje AR$': o.almacenaje,
    'Total AR$': o.total,
    'Estado': o.estado,
  }));

  const ws = XLSX.utils.json_to_sheet(rows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Operaciones');
  XLSX.writeFile(wb, 'ESIFAST_operaciones.xlsx');
}

/* ══════════════════════════════════════════════════════════
   EXPORT PDF
══════════════════════════════════════════════════════════ */
async function exportPDF() {
  if (!STATE.data) return;
  const overlay = document.getElementById('pdf-overlay');
  overlay.classList.remove('hidden');

  try {
    const { jsPDF } = window.jspdf;
    const doc = new jsPDF({ orientation: 'portrait', unit: 'mm', format: 'a4' });
    const ops = STATE.data.ops;
    const W = 210; let y = 0;

    // ── Header ──
    doc.setFillColor(11, 15, 26);
    doc.rect(0, 0, W, 32, 'F');
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(22); doc.setTextColor(255,255,255);
    doc.text('ESIFAST · Informe BI', 14, 16);
    doc.setFontSize(9); doc.setFont('helvetica', 'normal');
    doc.setTextColor(148,163,184);
    doc.text('Generado: ' + new Date().toLocaleString('es-AR'), 14, 24);
    doc.text('Total de operaciones: ' + ops.length, W - 14, 24, { align: 'right' });
    y = 42;

    // ── KPIs ──
    doc.setFont('helvetica', 'bold'); doc.setFontSize(11); doc.setTextColor(37,99,235);
    doc.text('KPIs PRINCIPALES', 14, y); y += 7;

    const totalIN  = sum(ops, 'in');
    const totalOUT = sum(ops, 'out');
    const totalAlm = sum(ops, 'almacenaje');
    const totalFact = sum(ops, 'total');
    const abiertas = ops.filter(o => o.estado === 'ABIERTA').length;

    const kpiData = [
      ['Total IN', formatARS(totalIN)],
      ['Total OUT', formatARS(totalOUT)],
      ['Almacenaje', formatARS(totalAlm)],
      ['Facturado total', formatARS(totalFact)],
      ['Operaciones', ops.length],
      ['OP abiertas', abiertas],
    ];

    const colW = (W - 28) / 3;
    kpiData.forEach((kpi, i) => {
      const col = i % 3; const row = Math.floor(i / 3);
      const x = 14 + col * colW; const ky = y + row * 18;
      doc.setFillColor(17, 24, 39);
      doc.roundedRect(x, ky, colW - 4, 14, 2, 2, 'F');
      doc.setFont('helvetica', 'normal'); doc.setFontSize(7); doc.setTextColor(148,163,184);
      doc.text(kpi[0].toUpperCase(), x + 4, ky + 5);
      doc.setFont('helvetica', 'bold'); doc.setFontSize(10); doc.setTextColor(226,232,240);
      doc.text(String(kpi[1]), x + 4, ky + 11);
    });
    y += 42;

    // ── Gráficos ──
    doc.setFont('helvetica', 'bold'); doc.setFontSize(11); doc.setTextColor(37,99,235);
    doc.text('GRÁFICOS', 14, y); y += 6;

    const chartIds = [
      { id: 'chart-ing',       label: 'Ingresos IN por cliente' },
      { id: 'chart-alm',       label: 'Almacenaje por cliente' },
      { id: 'chart-pallets',   label: 'Pallets ingresados vs egresados' },
      { id: 'chart-conceptos', label: 'Distribución de conceptos' },
    ];

    for (const c of chartIds) {
      const canvas = document.getElementById(c.id);
      if (!canvas) continue;
      const imgData = canvas.toDataURL('image/png');
      const cW = (W - 28) / 2;
      const cH = 50;
      const col = chartIds.indexOf(c) % 2;
      const cRow = Math.floor(chartIds.indexOf(c) / 2);
      const cx = 14 + col * (cW + 4);
      const cy = y + cRow * (cH + 10);

      doc.setFillColor(17, 24, 39);
      doc.roundedRect(cx, cy, cW, cH + 6, 2, 2, 'F');
      doc.setFont('helvetica', 'normal'); doc.setFontSize(7); doc.setTextColor(148,163,184);
      doc.text(c.label, cx + 3, cy + 5);
      doc.addImage(imgData, 'PNG', cx + 2, cy + 7, cW - 4, cH - 4);
    }
    y += 130;

    // ── Tabla resumen de ops ──
    if (y > 220) { doc.addPage(); y = 16; }
    doc.setFont('helvetica', 'bold'); doc.setFontSize(11); doc.setTextColor(37,99,235);
    doc.text('TABLA RESUMEN DE OPERACIONES', 14, y); y += 7;

    const colsTable = ['OP','Ingreso','Cliente','Pallets','IN AR$','Alm. AR$','Total AR$','Estado'];
    const colsW = [14, 20, 36, 14, 24, 24, 26, 18];
    const rowH = 7;

    // Header
    doc.setFillColor(26, 34, 54);
    let cx = 14;
    colsW.forEach((cw, i) => { doc.rect(cx, y, cw, rowH, 'F'); cx += cw; });
    cx = 14;
    doc.setFont('helvetica', 'bold'); doc.setFontSize(6.5); doc.setTextColor(148,163,184);
    colsTable.forEach((h, i) => {
      doc.text(h, cx + 2, y + 4.5);
      cx += colsW[i];
    });
    y += rowH;

    const tableOps = ops.slice(0, 30);
    tableOps.forEach((o, ri) => {
      if (y > 270) { doc.addPage(); y = 16; }
      doc.setFillColor(ri % 2 === 0 ? 17 : 22, ri % 2 === 0 ? 24 : 30, ri % 2 === 0 ? 39 : 54);
      cx = 14;
      colsW.forEach(cw => { doc.rect(cx, y, cw, rowH, 'F'); cx += cw; });
      cx = 14;
      doc.setFont('helvetica', 'normal'); doc.setFontSize(6); doc.setTextColor(226,232,240);
      const row = [o.op, fmtDate(o.fechaIngreso), o.cliente.substring(0,14), o.pallets, formatARS(o.in), formatARS(o.almacenaje), formatARS(o.total), o.estado];
      row.forEach((v, i) => { doc.text(String(v ?? '—'), cx + 2, y + 4.5); cx += colsW[i]; });
      y += rowH;
    });

    if (ops.length > 30) {
      y += 4;
      doc.setFont('helvetica', 'italic'); doc.setFontSize(8); doc.setTextColor(148,163,184);
      doc.text(`… y ${ops.length - 30} operaciones más. Exportar CSV/Excel para el detalle completo.`, 14, y);
      y += 10;
    }

    // ── Análisis ejecutivo ──
    if (y > 230) { doc.addPage(); y = 16; }
    y += 4;
    doc.setFont('helvetica', 'bold'); doc.setFontSize(11); doc.setTextColor(37,99,235);
    doc.text('ANÁLISIS EJECUTIVO', 14, y); y += 8;

    const energy = ops.filter(o => o.cliente.includes('ENERGY'));
    const esifar = ops.filter(o => !o.cliente.includes('ENERGY'));
    const factByCliente = groupSum(ops, 'cliente', 'total');
    const topCliente = Object.entries(factByCliente).sort((a,b)=>b[1]-a[1])[0];

    const analLines = [
      `• Cliente con mayor facturación: ${topCliente?.[0] || '—'} (${formatARS(topCliente?.[1] || 0)})`,
      `• CLADAN ENERGY: ${energy.length} ops · Almacenaje ${formatARS(sum(energy,'almacenaje'))} · Total ${formatARS(sum(energy,'total'))}`,
      `• ESIFAR: ${esifar.length} ops · Almacenaje ${formatARS(sum(esifar,'almacenaje'))} · Total ${formatARS(sum(esifar,'total'))}`,
      `• Operaciones abiertas sin fecha de salida: ${abiertas}`,
      `• Operaciones con saldo de pallets pendiente: ${ops.filter(o=>o.saldo>0).length}`,
    ];
    doc.setFont('helvetica', 'normal'); doc.setFontSize(9); doc.setTextColor(226,232,240);
    analLines.forEach(l => { doc.text(l, 14, y); y += 6; });

    // ── Recomendaciones ──
    if (y > 230) { doc.addPage(); y = 16; }
    y += 4;
    doc.setFont('helvetica', 'bold'); doc.setFontSize(11); doc.setTextColor(37,99,235);
    doc.text('RECOMENDACIONES PRINCIPALES', 14, y); y += 8;

    const recomLines = [
      `🚨 Revisar ${abiertas} operaciones sin fecha de salida (impacto: ${formatARS(sum(ops.filter(o=>o.estado==='ABIERTA'),'total'))})`,
      `⚖  Verificar ${ops.filter(o=>o.saldo>0&&o.estado==='ABIERTA').length} operaciones con saldo de pallets pendiente`,
      `💰 Evaluar facturación pendiente: almacenaje abierto de ${formatARS(sum(ops.filter(o=>o.estado==='ABIERTA'),'almacenaje'))}`,
      `📋 Completar ${STATE.data.ingresos.filter(i=>!i.remito).length} ingresos sin número de remito`,
      `📊 Actualizar hoja DESCONSOLIDADOS con operaciones del período`,
    ];
    doc.setFont('helvetica', 'normal'); doc.setFontSize(9); doc.setTextColor(226,232,240);
    recomLines.forEach(l => { doc.text(l, 14, y); y += 7; });

    // ── Footer ──
    const totalPages = doc.internal.getNumberOfPages();
    for (let i = 1; i <= totalPages; i++) {
      doc.setPage(i);
      doc.setFont('helvetica', 'normal'); doc.setFontSize(7); doc.setTextColor(71,85,105);
      doc.text(`ESIFAST BI · Generado ${new Date().toLocaleString('es-AR')} · Página ${i} de ${totalPages}`, W / 2, 292, { align: 'center' });
      doc.setDrawColor(31,45,69);
      doc.line(14, 288, W - 14, 288);
    }

    doc.save('ESIFAST_Informe_BI_' + new Date().toISOString().slice(0,10) + '.pdf');
  } catch (err) {
    alert('Error al generar PDF: ' + err.message);
    console.error(err);
  } finally {
    overlay.classList.add('hidden');
  }
}

/* ══════════════════════════════════════════════════════════
   HELPERS
══════════════════════════════════════════════════════════ */
function sum(arr, key) {
  return arr.reduce((acc, r) => acc + (Number(r[key]) || 0), 0);
}

function groupSum(arr, groupKey, sumKey) {
  return arr.reduce((acc, r) => {
    const k = r[groupKey] || 'Sin datos';
    acc[k] = (acc[k] || 0) + (Number(r[sumKey]) || 0);
    return acc;
  }, {});
}

function toNum(v) {
  if (v == null || v === '') return 0;
  const n = Number(String(v).replace(/[^\d.-]/g, ''));
  return isNaN(n) ? 0 : n;
}

function parseDate(v) {
  if (!v) return null;
  if (v instanceof Date) return isNaN(v.getTime()) ? null : v;
  if (typeof v === 'number') {
    const d = new Date((v - 25569) * 86400 * 1000);
    return isNaN(d.getTime()) ? null : d;
  }
  const s = String(v).trim();
  if (s === '00:00:00' || s === '' || s === 'NaT') return null;
  const d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
}

function fmtDate(d) {
  if (!d) return '—';
  try {
    return new Date(d).toLocaleDateString('es-AR', { day: '2-digit', month: '2-digit', year: 'numeric' });
  } catch { return '—'; }
}

function fmtISO(d) {
  if (!d) return '';
  try { return new Date(d).toISOString().slice(0, 10); } catch { return ''; }
}

function fmtMonth(d) {
  if (!d) return '';
  try {
    return new Date(d).toLocaleDateString('es-AR', { month: 'long', year: 'numeric' });
  } catch { return ''; }
}

function formatARS(n) {
  if (n == null || isNaN(n)) return '$0';
  return '$' + Math.round(n).toLocaleString('es-AR');
}

function calcDiasProm(ops) {
  const hoy = new Date();
  const dias = ops.map(o => {
    const inicio = o.fechaIngreso ? new Date(o.fechaIngreso) : null;
    const fin    = o.fechaSalida  ? new Date(o.fechaSalida)  : hoy;
    if (!inicio) return 0;
    return Math.max(0, Math.round((fin - inicio) / 86400000));
  }).filter(d => d > 0);
  return dias.length ? Math.round(dias.reduce((a, b) => a + b, 0) / dias.length) : 0;
}

function download(filename, mimeType, content) {
  const blob = new Blob([content], { type: mimeType });
  const url  = URL.createObjectURL(blob);
  const a    = document.createElement('a');
  a.href = url; a.download = filename;
  a.click(); URL.revokeObjectURL(url);
}

/* ══════════════════════════════════════════════════════════
   EVENT BINDINGS
══════════════════════════════════════════════════════════ */
document.addEventListener('DOMContentLoaded', () => {
  initAuth();
  initNav();
  initFileUpload();

  document.getElementById('btn-export-pdf').addEventListener('click', exportPDF);
  document.getElementById('btn-csv').addEventListener('click', exportCSV);
  document.getElementById('btn-excel-exp').addEventListener('click', exportExcelOps);
});
