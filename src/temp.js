'use strict';
/* SAGD Temp Viewer — thermocouple temperature / subcool along the wellbore.
   Reads the same workbook as the production viewer (one sheet per well) and uses
   the Temp_Point_n / Subcool_Point_n columns as positions along the lateral. */

const DAY = 86400000;
const STORE_KEY = 'sagdProfileViewer.v1';
const MD_KEY = 'sagdProfileViewer.md.v1';
const PALETTE = ['#2a78d6', '#eb6834', '#1baf7a', '#eda100', '#e87ba4', '#008300', '#4a3aa7', '#e34948',
                 '#164f94', '#a3401a', '#0c6e4a', '#8f6100', '#a83a66', '#4c4c4c', '#8c7fe0', '#8f2423'];
const VARS = {
  Temp:    { label: 'Temperature', colorscale: [[0, '#fff5d6'], [0.25, '#fdc96b'], [0.5, '#f2853a'], [0.75, '#c9372c'], [1, '#5e0d1c']] },
  // low subcool (steam near the producer) red, high subcool (liquid pooling) blue
  Subcool: { label: 'Subcool', colorscale: [[0, '#b2182b'], [0.25, '#ef8a62'], [0.5, '#f7f7f7'], [0.75, '#67a9cf'], [1, '#2166ac']] },
};
const CONTEXT_COLS = [
  ['Prod_Oil_rate_bbld', 'Oil', 'bbl/d'], ['Prod_Water_rate_bbld', 'Water', 'bbl/d'], ['Inj_Steam_rate_bbld', 'Steam', 'bbl/d'],
  ['Temp_Max', 'Temp max', '°C'], ['Subcool_Min', 'Subcool min', '°C'], ['Production_Hours', 'Prod. hours', 'h'],
];
const TSAT_MAX = 374; // °C, critical point of water
const AXIS = { gridcolor: '#eceef1', linecolor: '#c5cbd4', zeroline: false, ticks: 'outside', tickcolor: '#c5cbd4', automargin: true, showline: true };

const P = {
  fileName: null,
  sheets: [],          // parsed sheets as read, handed to the Production Viewer
  wells: [],
  data: {},            // well -> { t: Float64Array, cols: {name: Float64Array}, n }
  points: [],          // sorted point numbers found in Temp_/Subcool_Point_n columns
  tmin: 0, tmax: 0,
  settings: { well: '', v: 'Temp', avg: 1, axis: 'idx', reverse: false, clean: true, mode: 'dates',
              cur: 0, dates: [], wells: [], vmin: '', vmax: '', step: 7 },
  md: {},              // well -> [md per point index] (stored in this browser)
  cache: new Map(),
  exportRows: null,
  playing: null,
};

/* ================= small helpers ================= */
function $(sel, root = document) { return root.querySelector(sel); }
function h(tag, attrs = {}, ...kids) {
  const el = document.createElement(tag);
  for (const [k, v] of Object.entries(attrs || {})) {
    if (v === undefined || v === null || v === false) continue;
    if (k === 'class') el.className = v;
    else if (k.startsWith('on')) el.addEventListener(k.slice(2), v);
    else el.setAttribute(k, v === true ? '' : v);
  }
  for (const c of kids.flat()) if (c !== null && c !== undefined) el.append(c);
  return el;
}
function isoDate(ms) { return new Date(ms).toISOString().slice(0, 10); }
function parseIso(s) { const m = /^(\d{4})-(\d{2})-(\d{2})/.exec(s || ''); return m ? Date.UTC(+m[1], +m[2] - 1, +m[3]) : NaN; }
function escapeHtml(s) { return String(s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c])); }
function fmt(v) { return Number.isFinite(v) ? (Math.abs(v) >= 100 ? v.toFixed(0) : v.toFixed(1)) : '–'; }
function num(x) { return x === '' || x === null || x === undefined || !Number.isFinite(+x) ? null : +x; }
function addMonths(ms, k) { const d = new Date(ms); return Date.UTC(d.getUTCFullYear(), d.getUTCMonth() + k, d.getUTCDate()); }

let toastTimer;
function toast(msg, err) {
  const t = $('#toast');
  t.textContent = msg;
  t.className = 'toast on' + (err ? ' err' : '');
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => { t.className = 'toast' + (err ? ' err' : ''); }, err ? 6000 : 2800);
}
function showLoading(msg) { $('#loadingMsg').textContent = msg; $('#loading').classList.add('on'); }
function hideLoading() { $('#loading').classList.remove('on'); }
function download(name, blob) {
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = name;
  document.body.append(a);
  a.click();
  setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 1500);
}

function persist() {
  try {
    localStorage.setItem(STORE_KEY, JSON.stringify(P.settings));
    localStorage.setItem(MD_KEY, JSON.stringify(P.md));
  } catch (e) { /* storage unavailable */ }
}
function restore(key) { try { return JSON.parse(localStorage.getItem(key) || 'null'); } catch (e) { return null; } }

/* ================= loading ================= */
async function loadExcel(file) {
  showLoading(`Reading ${file.name} (${(file.size / 1048576).toFixed(1)} MB)…`);
  const t0 = performance.now();
  try {
    const sheets = await readWorkbook(file, msg => { $('#loadingMsg').textContent = msg; });
    if (!sheets.length) throw new Error('No sheet with a date column and data rows was found.');
    ingest(file.name, sheets);
    toast(`Loaded ${P.wells.length} wells in ${((performance.now() - t0) / 1000).toFixed(1)} s`);
  } catch (e) {
    console.error(e);
    toast('Failed to read file: ' + (e.message || e), true);
  } finally {
    hideLoading();
  }
}

/** Opened from the Production Viewer link: take the workbook that viewer already has loaded. */
async function loadFromOpener() {
  if (!window.opener) return;
  showLoading('Receiving data from the Production Viewer…');
  try {
    const r = await receiveHandoff();
    if (r) {
      ingest(r.fileName, r.sheets);
      const f = applyFocus(r.focus);
      toast(`Loaded ${P.wells.length} wells from the Production Viewer` + (f ? ` · ${f}` : ''));
    }
  } catch (e) {
    console.error(e);
    toast('Could not use the data from the Production Viewer: ' + (e.message || e), true);
  } finally {
    hideLoading();
  }
}

/** Hosted demo: `?data=…` opens with that workbook loaded. */
async function loadFromQuery() {
  try {
    const data = await fetchQueryFile('data');
    if (data) await loadExcel(data);
  } catch (e) {
    console.error(e);
    toast('Could not load the demo data: ' + (e.message || e), true);
  }
}

/** Show the well (and date, when given) the Production Viewer handed over. Returns a label, or '' when nothing applied. */
function applyFocus(f) {
  const st = P.settings;
  if (!f || !P.data[f.well]) return '';
  st.well = f.well;
  const ms = parseIso(f.date);
  if (Number.isFinite(ms)) st.cur = Math.min(P.tmax, Math.max(P.tmin, ms));
  syncToolbar();
  renderAll();
  persist();
  return Number.isFinite(ms) ? `${st.well} @ ${isoDate(st.cur)}` : st.well;
}

function ingest(fileName, sheets) {
  const pts = new Set();
  for (const sh of sheets) for (const c of sh.header) {
    const m = /^(Temp|Subcool)_Point_(\d+)$/i.exec(c);
    if (m) pts.add(+m[2]);
  }
  if (!pts.size) throw new Error('No Temp_Point_n / Subcool_Point_n columns found.');
  P.fileName = fileName;
  P.sheets = sheets;
  P.points = [...pts].sort((a, b) => a - b);
  P.wells = sheets.map(s => s.name);
  P.data = {};
  P.tmin = Infinity; P.tmax = -Infinity;
  for (const sh of sheets) {
    P.data[sh.name] = { t: sh.t, cols: sh.cols, n: sh.t.length };
    P.tmin = Math.min(P.tmin, sh.t[0]); P.tmax = Math.max(P.tmax, sh.t[sh.t.length - 1]);
  }
  P.cache.clear();

  // settings: keep what still applies from the last session
  const st = P.settings, saved = restore(STORE_KEY);
  if (saved) Object.assign(st, saved);
  P.md = restore(MD_KEY) || {};
  if (!P.data[st.well]) st.well = P.wells.find(w => hasAnyData(w, 'Temp')) || P.wells[0];
  st.wells = (st.wells || []).filter(w => P.data[w]);
  if (!st.wells.length) st.wells = P.wells.slice();
  if (!(st.cur >= P.tmin && st.cur <= P.tmax)) st.cur = P.tmax;
  st.dates = (st.dates || []).filter(d => d >= P.tmin && d <= P.tmax);
  if (!saved || !saved.dates || !saved.dates.length) st.dates = snapshots(12, 5);

  $('#fileInfo').textContent = `${fileName}  ·  ${P.wells.length} wells  ·  ${P.points.length} points  ·  ${isoDate(P.tmin)} → ${isoDate(P.tmax)}`;
  $('#fileInfo').title = $('#fileInfo').textContent;
  $('#toolbar').setAttribute('aria-disabled', 'false');
  $('#emptyState').classList.add('hidden');
  $('#main').classList.remove('hidden');
  const slider = $('#slider');
  slider.max = String(Math.round((P.tmax - P.tmin) / DAY));
  $('#dateInput').min = isoDate(P.tmin);
  $('#dateInput').max = isoDate(P.tmax);
  $('#wellSel').replaceChildren(...P.wells.map(w => h('option', { value: w }, w)));
  syncToolbar();
  renderAll();
}

/* ================= data ================= */
function colName(v, k) { return `${v}_Point_${k}`; }

/** Cleaned, trailing-averaged daily values for every point: { t, y: [Float64Array per point] }. */
function profileSeries(well, v) {
  const st = P.settings, key = `${well}|${v}|${st.avg}|${st.clean}`;
  if (P.cache.has(key)) return P.cache.get(key);
  const W = P.data[well];
  const y = P.points.map(k => {
    const src = W.cols[colName(v, k)];
    const out = new Float64Array(W.n).fill(NaN);
    if (!src) return out;
    const temp = W.cols[colName('Temp', k)];
    for (let i = 0; i < W.n; i++) {
      let x = src[i];
      if (st.clean && temp) {
        // a dead thermocouple logs 0 °C, which also turns its subcool into Tsat − 0
        if (!(temp[i] > 0)) x = NaN;
        // Tsat = T + subcool cannot exceed water's critical temperature
        else if (v === 'Subcool' && temp[i] + x > TSAT_MAX) x = NaN;
      }
      out[i] = x;
    }
    if (st.avg > 1) {
      const sm = new Float64Array(W.n);
      let sum = 0, n = 0;
      for (let i = 0; i < W.n; i++) {
        if (Number.isFinite(out[i])) { sum += out[i]; n++; }
        if (i >= st.avg) { const o = out[i - st.avg]; if (Number.isFinite(o)) { sum -= o; n--; } }
        sm[i] = n ? sum / n : NaN;
      }
      return sm;
    }
    return out;
  });
  const res = { t: W.t, y };
  P.cache.set(key, res);
  return res;
}

function hasAnyData(well, v) {
  const s = profileSeries(well, v);
  return s.y.some(a => a.some(Number.isFinite));
}

/** Row index of `ms` in a well's daily series, or -1 when the well has no row for that day. */
function rowAt(well, ms) {
  const t = P.data[well].t;
  let lo = 0, hi = t.length - 1;
  if (ms < t[0] - DAY / 2 || ms > t[hi] + DAY / 2) return -1;
  while (lo < hi) { const mid = (lo + hi + 1) >> 1; if (t[mid] <= ms + DAY / 2) lo = mid; else hi = mid - 1; }
  return Math.abs(t[lo] - ms) < DAY ? lo : -1;
}

function profileAt(well, v, ms) {
  const i = rowAt(well, ms);
  if (i < 0) return null;
  const y = profileSeries(well, v).y.map(a => a[i]);
  return y.some(Number.isFinite) ? y : null;
}

/** Position of each point: MD when every point of `well` has one, else the point number. */
function positionsFor(well) {
  const md = P.md[well];
  if (P.settings.axis === 'md' && md && P.points.every((_, j) => Number.isFinite(md[j]))) return { x: P.points.map((_, j) => md[j]), md: true };
  return { x: P.points.slice(), md: false };
}

/** 2nd–98th percentile of every value of the given wells (used for color and axis bounds). */
function valueRange(wells, v) {
  const all = [];
  for (const w of wells) for (const a of profileSeries(w, v).y) for (const x of a) if (Number.isFinite(x)) all.push(x);
  if (!all.length) return null;
  all.sort((a, b) => a - b);
  const q = p => all[Math.min(all.length - 1, Math.max(0, Math.round(p * (all.length - 1))))];
  return { lo: q(0.02), hi: q(0.98), min: all[0], max: all[all.length - 1] };
}

function userRange(auto) {
  const lo = num(P.settings.vmin), hi = num(P.settings.vmax);
  if (lo === null && hi === null) return auto;
  return [lo ?? auto[0], hi ?? auto[1]];
}

/** Kept dates: `count` steps of `months` going back from the current date, within the well's data. */
function snapshots(months, count) {
  const st = P.settings, out = [];
  for (let k = 1; k <= count; k++) {
    const d = addMonths(st.cur, -months * k);
    if (d < P.tmin) break;
    if (P.data[st.well] && !profileAt(st.well, st.v, d)) continue;
    out.push(d);
  }
  return out.sort((a, b) => a - b);
}

/* ================= rendering ================= */
function renderAll() {
  renderProfile();
  renderHeat();
  renderCfg();
  syncTime();
}

function unitLabel() { return `${VARS[P.settings.v].label} [°C]`; }

function renderProfile() {
  const st = P.settings, v = st.v, traces = [], exp = [];
  let wells, lines, missing = [];
  if (st.mode === 'wells') {
    wells = st.wells.filter(w => P.data[w]);
    lines = wells.map(w => ({ well: w, ms: st.cur, name: w, color: PALETTE[P.wells.indexOf(w) % PALETTE.length], width: 2 }));
  } else {
    wells = [st.well];
    lines = st.dates.filter(d => d !== st.cur).map((d, i) => ({ well: st.well, ms: d, name: isoDate(d), color: PALETTE[i % PALETTE.length], width: 1.6, dash: 'dot' }));
    lines.push({ well: st.well, ms: st.cur, name: `${isoDate(st.cur)} (current)`, color: '#1b1f24', width: 3 });
  }
  // MD only when every plotted well has it, so all lines share one position axis
  const useMD = st.axis === 'md' && wells.length && wells.every(w => positionsFor(w).md);
  for (const ln of lines) {
    const y = profileAt(ln.well, v, ln.ms);
    if (!y) { missing.push(ln.name); continue; }
    const x = useMD ? positionsFor(ln.well).x : P.points.slice();
    traces.push({
      type: 'scatter', mode: 'lines+markers', x, y: y.map(z => Number.isFinite(z) ? z : null), name: ln.name,
      customdata: P.points, connectgaps: false,
      line: { color: ln.color, width: ln.width, dash: ln.dash || 'solid' }, marker: { color: ln.color, size: ln.width > 2 ? 8 : 6 },
      hovertemplate: `Point %{customdata}${useMD ? ' · MD %{x}' : ''}<br>%{y:.1f} °C<extra>${escapeHtml(ln.name)}</extra>`,
    });
    exp.push({ name: ln.name, y });
  }
  const rng = valueRange(wells, v);
  const L = {
    height: 440, margin: { l: 64, r: 24, t: 18, b: 48 },
    paper_bgcolor: '#ffffff', plot_bgcolor: '#ffffff',
    font: { family: 'Segoe UI, system-ui, sans-serif', size: 12, color: '#1b1f24' },
    legend: { orientation: 'h', x: 0, y: -0.16, yanchor: 'top', font: { size: 11 } },
    hovermode: 'closest',
    xaxis: { ...AXIS, title: { text: useMD ? 'MD [m]' : 'Thermocouple point' }, autorange: st.reverse ? 'reversed' : true,
             ...(useMD ? {} : { tickmode: 'array', tickvals: P.points, ticktext: P.points.map(String) }) },
    yaxis: { ...AXIS, title: { text: unitLabel() } },
    uirevision: `${st.mode}|${st.well}|${v}|${useMD}|${st.reverse}`,
  };
  // fixed value axis so the lines don't jump while stepping through time
  if (rng) {
    const pad = 0.04 * (rng.max - rng.min || 1);
    L.yaxis.range = userRange([rng.min - pad, rng.max + pad]);
    L.yaxis.autorange = false;
  }
  if (st.axis === 'md' && !useMD) {
    L.annotations = [{ xref: 'paper', yref: 'paper', x: 1, y: 1, xanchor: 'right', yanchor: 'bottom', showarrow: false,
      text: 'MD not set for every plotted well — showing point numbers', font: { size: 11, color: '#c63b3a' } }];
    L.margin.t = 28;
  }
  $('#profTitle').textContent = st.mode === 'wells'
    ? `${VARS[v].label} profile · ${isoDate(st.cur)} · ${wells.length} wells`
    : `${VARS[v].label} profile · ${st.well}`;
  const msg = traces.length ? '' : st.mode === 'wells' && !wells.length ? 'Select at least one well.'
    : `No ${VARS[v].label.toLowerCase()} data on ${isoDate(st.cur)}.`;
  $('#profMsg').textContent = msg;
  $('#profMsg').classList.toggle('on', !!msg);
  P.missing = missing;
  P.exportRows = { useMD, exp, x: useMD ? positionsFor(wells[0]).x : null };
  Plotly.react('profChart', traces, L, { responsive: true, displaylogo: false, modeBarButtonsToRemove: ['lasso2d', 'select2d', 'autoScale2d'],
    toImageButtonOptions: { filename: `${st.mode === 'wells' ? 'wells_' + isoDate(st.cur) : st.well}_${v}_profile`, scale: 2 } });
}

function renderHeat() {
  const st = P.settings, v = st.v, well = st.well;
  const s = profileSeries(well, v), pos = positionsFor(well);
  const x = Array.from(s.t, isoDate);
  const z = s.y.map(a => Array.from(a, q => Number.isFinite(q) ? q : null));
  const rng = valueRange([well], v);
  const [zmin, zmax] = rng ? userRange([rng.lo, rng.hi]) : [undefined, undefined];
  const trace = {
    type: 'heatmap', x, y: pos.x, z, zmin, zmax, colorscale: VARS[v].colorscale, zsmooth: false, hoverongaps: false,
    colorbar: { title: { text: '°C', side: 'right' }, thickness: 12, len: 0.9 },
    hovertemplate: `%{x}<br>${pos.md ? 'MD' : 'Point'} %{y}<br>%{z:.1f} °C<extra></extra>`,
  };
  const L = {
    height: 360, margin: { l: 64, r: 24, t: 12, b: 40 },
    paper_bgcolor: '#ffffff', plot_bgcolor: '#f4f5f7',
    font: { family: 'Segoe UI, system-ui, sans-serif', size: 12, color: '#1b1f24' },
    xaxis: { ...AXIS, type: 'date' },
    yaxis: { ...AXIS, title: { text: pos.md ? 'MD [m]' : 'Thermocouple point' }, autorange: st.reverse ? 'reversed' : true,
             ...(pos.md ? {} : { tickmode: 'array', tickvals: P.points, ticktext: P.points.map(String) }) },
    shapes: timeShapes(),
    uirevision: `${well}|${pos.md}|${st.reverse}`,
  };
  $('#heatTitle').textContent = `${VARS[v].label} · time × position · ${well}`;
  const empty = !z.some(r => r.some(q => q !== null));
  $('#heatMsg').textContent = empty ? `No ${VARS[v].label.toLowerCase()} data for ${well}.` : '';
  $('#heatMsg').classList.toggle('on', empty);
  const el = $('#heatChart');
  Plotly.react(el, [trace], L, { responsive: true, displaylogo: false, modeBarButtonsToRemove: ['lasso2d', 'select2d', 'autoScale2d'],
    toImageButtonOptions: { filename: `${well}_${v}_map`, scale: 2 } });
  if (!el._clickBound) {
    el._clickBound = true;
    el.on('plotly_click', ev => { const pt = ev.points && ev.points[0]; if (pt) setCur(parseIso(String(pt.x))); });
  }
}

function timeShapes() {
  const st = P.settings, line = (ms, dash, color, width) => ({ type: 'line', xref: 'x', yref: 'paper', x0: isoDate(ms), x1: isoDate(ms), y0: 0, y1: 1,
    line: { color, width, dash } });
  const out = st.mode === 'dates' ? st.dates.filter(d => d !== st.cur).map(d => line(d, 'dash', '#4c4c4c', 1)) : [];
  out.push(line(st.cur, 'solid', '#1b1f24', 2));
  return out;
}

/** Left-hand panel of the profile card: kept dates or well picker, plus a readout at the current date. */
function renderCfg() {
  const st = P.settings, box = $('#profCfg');
  const parts = [];
  if (st.mode === 'dates') {
    const kept = st.dates.filter(d => d !== st.cur);
    parts.push(h('div', {},
      h('h4', {}, 'Kept dates', h('span', { class: 'links' },
        h('a', { onclick: () => { st.dates = snapshots(12, 5); afterDates(); } }, 'Yearly'),
        h('a', { onclick: () => { st.dates = snapshots(6, 6); afterDates(); } }, '6 months'),
        h('a', { onclick: () => { st.dates = []; afterDates(); } }, 'Clear'))),
      h('div', { class: 'chips' },
        kept.map((d, i) => h('span', { class: 'datechip' },
          h('span', { class: 'dot', style: `background:${PALETTE[i % PALETTE.length]}` }),
          h('a', { title: 'Make this the current date', style: 'cursor:pointer', onclick: () => setCur(d) }, isoDate(d)),
          h('button', { class: 'xbtn', title: 'Remove', onclick: () => { st.dates = st.dates.filter(x => x !== d); afterDates(); } }, '×'))),
        h('span', { class: 'datechip cur' }, h('span', { class: 'dot', style: 'background:#1b1f24' }), isoDate(st.cur))),
      kept.length ? null : h('div', { class: 'note' }, 'Use “+ Keep date” below the chart, or the Yearly / 6 months shortcuts, to overlay more dates.')));
  } else {
    parts.push(h('div', {},
      h('h4', {}, 'Wells', h('span', { class: 'links' },
        h('a', { onclick: () => { st.wells = P.wells.slice(); afterWells(); } }, 'All'),
        h('a', { onclick: () => { st.wells = []; afterWells(); } }, 'None'))),
      h('div', { class: 'chips' }, P.wells.map(w => {
        const on = st.wells.includes(w);
        return h('span', { class: 'chip' + (on ? ' on' : ''), onclick: () => {
          st.wells = on ? st.wells.filter(x => x !== w) : P.wells.filter(x => x === w || st.wells.includes(x));
          afterWells();
        } }, h('span', { class: 'dot', style: `background:${PALETTE[P.wells.indexOf(w) % PALETTE.length]}` }), w);
      }))));
  }
  if (P.missing && P.missing.length) parts.push(h('div', { class: 'note' }, `No data: ${P.missing.join(', ')}`));

  // context at the current date for the well shown in the map
  const W = P.data[st.well], i = rowAt(st.well, st.cur);
  const rows = CONTEXT_COLS.filter(([c]) => W.cols[c]).map(([c, label, unit]) =>
    h('tr', {}, h('td', {}, label), h('td', { class: 'num' }, i < 0 ? '–' : fmt(W.cols[c][i])), h('td', { class: 'unit' }, unit)));
  if (rows.length) parts.push(h('div', {}, h('h4', {}, `${st.well} · ${isoDate(st.cur)}`), h('table', { class: 'ctx' }, rows)));
  box.replaceChildren(...parts);
}

function afterDates() { P.settings.dates.sort((a, b) => a - b); persist(); renderProfile(); renderCfg(); updateShapes(); }
function afterWells() { persist(); renderProfile(); renderCfg(); }

function updateShapes() { Plotly.relayout('heatChart', { shapes: timeShapes() }); }

/* ================= time control ================= */
function syncTime() {
  const st = P.settings;
  $('#slider').value = String(Math.round((st.cur - P.tmin) / DAY));
  $('#dateInput').value = isoDate(st.cur);
}

function setCur(ms) {
  if (!Number.isFinite(ms)) return;
  const st = P.settings;
  st.cur = Math.min(P.tmax, Math.max(P.tmin, Math.round(ms / DAY) * DAY));
  syncTime();
  renderProfile();
  renderCfg();
  updateShapes();
  persist();
}

function step(dir) { setCur(P.settings.cur + dir * P.settings.step * DAY); }

function togglePlay() {
  const btn = $('#playBtn');
  if (P.playing) { clearInterval(P.playing); P.playing = null; btn.textContent = '▶ Play'; return; }
  if (P.settings.cur >= P.tmax) setCur(P.tmin);
  btn.textContent = '❚❚ Pause';
  P.playing = setInterval(() => { if (P.settings.cur >= P.tmax) togglePlay(); else step(1); }, 120);
}

/* ================= MD editor ================= */
function renderMD() {
  const st = P.settings, md = P.md[st.well] || [];
  $('#mdWell').textContent = st.well;
  $('#mdGrid').replaceChildren(...P.points.map((k, j) => h('label', {}, `P${k}`,
    h('input', { type: 'number', step: 'any', value: Number.isFinite(md[j]) ? md[j] : '', onchange: e => {
      const arr = P.md[st.well] || (P.md[st.well] = []);
      const x = num(e.target.value);
      arr[j] = x === null ? null : x;
      persist();
      renderAll();
    } }))));
}

/* ================= export ================= */
function exportCSV() {
  const ex = P.exportRows;
  if (!ex || !ex.exp.length) { toast('Nothing to export.', true); return; }
  const st = P.settings;
  const header = ['Point', ...(ex.useMD ? ['MD_m'] : []), ...ex.exp.map(e => e.name)];
  const rows = P.points.map((k, j) => [k, ...(ex.useMD ? [ex.x[j]] : []), ...ex.exp.map(e => Number.isFinite(e.y[j]) ? e.y[j] : '')]);
  const meta = `# ${VARS[st.v].label} profile | file: ${P.fileName} | ${st.mode === 'wells' ? 'date: ' + isoDate(st.cur) : 'well: ' + st.well} | average: ${st.avg} d | clean: ${st.clean}`;
  const esc = v => { const s = String(v ?? ''); return /[",\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s; };
  const lines = [meta, header.map(esc).join(','), ...rows.map(r => r.map(esc).join(','))];
  const name = `${st.mode === 'wells' ? 'wells_' + isoDate(st.cur) : st.well}_${st.v}_profile.csv`;
  download(name, new Blob(['﻿' + lines.join('\r\n')], { type: 'text/csv;charset=utf-8' }));
}

/* ================= toolbar & events ================= */
function syncToolbar() {
  const st = P.settings;
  const seg = (id, val) => document.querySelectorAll(`#${id} button`).forEach(b => b.classList.toggle('on', b.dataset.v === String(val)));
  seg('varSeg', st.v); seg('avgSeg', st.avg); seg('axisSeg', st.axis); seg('modeSeg', st.mode);
  $('#wellSel').value = st.well;
  $('#revChk').checked = !!st.reverse;
  $('#cleanChk').checked = !!st.clean;
  $('#vmin').value = st.vmin; $('#vmax').value = st.vmax;
  $('#stepSel').value = String(st.step);
}

function setSetting(k, v) {
  P.settings[k] = v;
  if (k === 'avg' || k === 'clean') P.cache.clear();
  syncToolbar();
  persist();
  if (k === 'well' && !$('#mdPanel').hidden) renderMD();
  renderAll();
}

function handleFiles(files) {
  for (const f of files) {
    if (/\.(xlsx|xlsm|xls|xlsb|csv)$/i.test(f.name)) loadExcel(f);
    else toast(`Unsupported file: ${f.name}`, true);
  }
}

function init() {
  const saved = restore(STORE_KEY);
  if (saved) Object.assign(P.settings, saved);
  syncToolbar();
  $('#fileInput').addEventListener('change', e => { handleFiles(e.target.files); e.target.value = ''; });
  $('#fileInput2').addEventListener('change', e => { handleFiles(e.target.files); e.target.value = ''; });
  linkHandoff($('#prodLink'), () => P.wells.length
    ? { fileName: P.fileName, sheets: P.sheets, focus: { well: P.settings.well, date: isoDate(P.settings.cur) } } : null);
  $('#wellSel').addEventListener('change', e => setSetting('well', e.target.value));
  const seg = (id, key, conv = x => x) => document.querySelectorAll(`#${id} button`).forEach(b => b.addEventListener('click', () => setSetting(key, conv(b.dataset.v))));
  seg('varSeg', 'v'); seg('avgSeg', 'avg', Number); seg('axisSeg', 'axis'); seg('modeSeg', 'mode');
  $('#revChk').addEventListener('change', e => setSetting('reverse', e.target.checked));
  $('#cleanChk').addEventListener('change', e => setSetting('clean', e.target.checked));
  $('#vmin').addEventListener('change', e => setSetting('vmin', e.target.value));
  $('#vmax').addEventListener('change', e => setSetting('vmax', e.target.value));
  $('#stepSel').addEventListener('change', e => { P.settings.step = +e.target.value; persist(); });
  $('#slider').addEventListener('input', e => setCur(P.tmin + (+e.target.value) * DAY));
  $('#dateInput').addEventListener('change', e => setCur(parseIso(e.target.value)));
  $('#prevBtn').addEventListener('click', () => step(-1));
  $('#nextBtn').addEventListener('click', () => step(1));
  $('#playBtn').addEventListener('click', togglePlay);
  $('#pinBtn').addEventListener('click', () => {
    const st = P.settings;
    if (st.mode !== 'dates') setSetting('mode', 'dates');
    if (!st.dates.includes(st.cur)) st.dates.push(st.cur);
    afterDates();
  });
  $('#csvBtn').addEventListener('click', exportCSV);
  $('#mdBtn').addEventListener('click', () => { const p = $('#mdPanel'); p.hidden = !p.hidden; if (!p.hidden) renderMD(); });
  $('#mdClose').addEventListener('click', () => { $('#mdPanel').hidden = true; });
  $('#mdClear').addEventListener('click', () => { delete P.md[P.settings.well]; persist(); renderMD(); renderAll(); });
  $('#mdCopy').addEventListener('click', () => {
    const src = P.md[P.settings.well];
    if (!src) { toast('Enter MD values for this well first.', true); return; }
    for (const w of P.wells) P.md[w] = src.slice();
    persist();
    toast(`MD copied to ${P.wells.length} wells`);
    renderAll();
  });

  // drag & drop anywhere
  let depth = 0;
  const ov = $('#dragOverlay');
  window.addEventListener('dragenter', e => { if (e.dataTransfer && [...e.dataTransfer.types].includes('Files')) { depth++; ov.classList.add('on'); e.preventDefault(); } });
  window.addEventListener('dragleave', () => { depth = Math.max(0, depth - 1); if (!depth) ov.classList.remove('on'); });
  window.addEventListener('dragover', e => e.preventDefault());
  window.addEventListener('drop', e => {
    e.preventDefault(); depth = 0; ov.classList.remove('on');
    if (e.dataTransfer && e.dataTransfer.files.length) handleFiles(e.dataTransfer.files);
  });
}

init();
loadFromOpener().then(() => { if (!P.wells.length) loadFromQuery(); });
