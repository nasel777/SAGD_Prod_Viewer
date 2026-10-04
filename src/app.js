'use strict';
/* SAGD Production Viewer — client-side Excel → Plotly viewer.
   Each sheet = one well. First "Date" column = daily timestamp. */

const DAY = 86400000;
const STORE_KEY = 'bgProdViewer.layout.v1';
// Categorical palette: 8 validated hues + 8 darker companions for wells 9-16.
const PALETTE = ['#2a78d6', '#eb6834', '#1baf7a', '#eda100', '#e87ba4', '#008300', '#4a3aa7', '#e34948',
                 '#164f94', '#a3401a', '#0c6e4a', '#8f6100', '#a83a66', '#4c4c4c', '#8c7fe0', '#8f2423'];
const DASHES = ['solid', 'dash', 'dot', 'dashdot', 'longdash', 'longdashdot'];
const SYMBOLS = ['circle', 'square', 'diamond', 'triangle-up', 'triangle-down', 'cross', 'x', 'star',
                 'hexagon', 'pentagon', 'circle-open', 'square-open', 'diamond-open', 'triangle-up-open', 'star-open', 'hexagon-open'];
const HEIGHTS = { S: 300, M: 420, L: 560, XL: 720 };
// before / after windows around the redrill date in Statistics plots
const EVENT_WINDOWS = { normal: { label: 'Last normal before', color: '#b8bfc8' }, before: { label: 'Just before', color: '#8fb3e3' },
                        after: { label: 'After redrill', color: '#2a78d6' } };
const NORMAL_UPTIME_H = 20; // "normal" = trailing 90-day mean Production_Hours at least this
const FOCUS_HALF_SPAN = 182 * DAY; // zoom around the date handed over by the Profile Viewer
const TYPE_LABEL = { ts: 'Time series', xp: 'Crossplot', st: 'Statistics' };

const S = {
  fileName: null,
  sheets: [],         // parsed sheets as read, handed to the Profile Viewer
  wells: [],          // well (sheet) names in workbook order
  data: {},           // name -> { t: Float64Array(ms UTC), cols: {name: Float64Array}, start: ms }
  rawCols: [],        // union of numeric columns (excluding date)
  flagCols: new Set(),
  allVars: [],        // rawCols + derived
  derived: {},        // active derived definitions
  settings: { res: 'D', ma: 0, excludeShutin: false, zeroGap: false, sync: true, cols: 1 },
  plots: [],
  nextId: 1,
  exports: {},
  syncing: false,
  pendingLayout: null,
  eventCol: null,     // 0/1 column marking wells after a redrill (W.event = its first day)
  focus: null,        // {well, date, range} handed over by the Profile Viewer (this session only)
};

/* ================= derived variables ================= */
const DERIVED_DEFS = {
  // daily sum of columns; behaves like a raw value (masks, averaging, smoothing all apply)
  'Prod_Emulsion_rate_bbld': { kind: 'value', sum: ['Prod_Oil_rate_bbld', 'Prod_Water_rate_bbld'], group: 'Production' },
  'SOR':           { kind: 'ratio', num: ['Inj_Steam_rate_bbld'], den: ['Prod_Oil_rate_bbld'], scale: 1, minDen: 1 },
  'Water_Cut_%':   { kind: 'ratio', num: ['Prod_Water_rate_bbld'], den: ['Prod_Oil_rate_bbld', 'Prod_Water_rate_bbld'], scale: 100 },
  'Cum_Oil_bbl':   { kind: 'cum', src: 'Prod_Oil_rate_bbld' },
  'Cum_Water_bbl': { kind: 'cum', src: 'Prod_Water_rate_bbld' },
  'Cum_Steam_bbl': { kind: 'cum', src: 'Inj_Steam_rate_bbld' },
  'CSOR':          { kind: 'cum', ratioOf: ['Cum_Steam_bbl', 'Cum_Oil_bbl'] },
};

/* ================= loading ================= */
async function loadExcel(file) {
  showLoading(`Reading ${file.name} (${(file.size / 1048576).toFixed(1)} MB)…`);
  const t0 = performance.now();
  try {
    const sheets = await readWorkbook(file, msg => { $('#loadingMsg').textContent = msg; });
    if (!sheets.length) throw new Error('No sheet with a date column and data rows was found.');
    ingest(file.name, sheets);
    toast(`Loaded ${S.wells.length} wells in ${((performance.now() - t0) / 1000).toFixed(1)} s`);
  } catch (e) {
    console.error(e);
    toast('Failed to read file: ' + (e.message || e), true);
  } finally {
    hideLoading();
  }
}

/** Opened from the Profile Viewer link: take the workbook that viewer already has loaded. */
async function loadFromOpener() {
  if (!window.opener) return;
  showLoading('Receiving data from the Profile Viewer…');
  try {
    const r = await receiveHandoff();
    if (r) {
      ingest(r.fileName, r.sheets);
      applyFocus(r.focus);
      toast(`Loaded ${S.wells.length} wells from the Profile Viewer` + (S.focus ? ` · ${S.focus.well} @ ${S.focus.date}` : ''));
    }
  } catch (e) {
    console.error(e);
    toast('Could not use the data from the Profile Viewer: ' + (e.message || e), true);
  } finally {
    hideLoading();
  }
}

function ingest(fileName, sheets) {
  S.fileName = fileName;
  S.sheets = sheets;
  S.focus = null;
  S.wells = [];
  S.data = {};
  const colOrder = [];
  const seen = new Set();
  for (const sh of sheets) {
    S.wells.push(sh.name);
    S.data[sh.name] = { t: sh.t, cols: sh.cols, n: sh.t.length };
    for (const h of sh.header) if (!seen.has(h)) { seen.add(h); colOrder.push(h); }
  }
  S.rawCols = colOrder;

  // 0/1 flag columns
  S.flagCols = new Set();
  for (const c of colOrder) {
    let ok = true;
    for (const w of S.wells) {
      const a = S.data[w].cols[c];
      if (!a) continue;
      for (let i = 0; i < a.length; i++) { const v = a[i]; if (!isNaN(v) && v !== 0 && v !== 1) { ok = false; break; } }
      if (!ok) break;
    }
    if (ok) S.flagCols.add(c);
  }

  // derived columns
  S.derived = {};
  for (const [name, d] of Object.entries(DERIVED_DEFS)) {
    if (d.sum && d.sum.every(c => seen.has(c))) S.derived[name] = d;
    if (d.kind === 'ratio' && [...d.num, ...d.den].every(c => seen.has(c))) S.derived[name] = d;
    if (d.kind === 'cum' && d.src && seen.has(d.src)) S.derived[name] = d;
    if (d.kind === 'cum' && d.ratioOf && d.ratioOf.every(c => S.derived[c])) S.derived[name] = d;
  }
  for (const w of S.wells) {
    const W = S.data[w];
    for (const [name, d] of Object.entries(S.derived)) {
      if (!d.sum && d.kind !== 'cum') continue;
      const out = new Float64Array(W.n);
      if (d.sum) {
        const srcs = d.sum.map(c => W.cols[c]).filter(Boolean);
        for (let i = 0; i < W.n; i++) {
          let s = 0, any = false;
          for (const a of srcs) if (Number.isFinite(a[i])) { s += a[i]; any = true; }
          out[i] = any ? s : NaN;
        }
      } else if (d.src) {
        const a = W.cols[d.src];
        let s = 0;
        for (let i = 0; i < W.n; i++) { const v = a ? a[i] : NaN; if (Number.isFinite(v)) s += v; out[i] = s; }
      } else {
        const a = W.cols[d.ratioOf[0]], b = W.cols[d.ratioOf[1]];
        for (let i = 0; i < W.n; i++) out[i] = b[i] > 0 ? a[i] / b[i] : NaN;
      }
      W.cols[name] = out;
    }
    // first-oil date for "days since start"
    const oil = W.cols['Prod_Oil_rate_bbld'];
    W.start = W.t[0];
    if (oil) for (let i = 0; i < W.n; i++) if (oil[i] > 0) { W.start = W.t[i]; break; }
  }
  // redrill date: first day the redrill flag is 1
  S.eventCol = S.rawCols.find(c => S.flagCols.has(c) && /redrill/i.test(c)) || null;
  for (const w of S.wells) {
    const W = S.data[w], col = S.eventCol && W.cols[S.eventCol];
    W.event = null;
    if (col) for (let i = 0; i < W.n; i++) if (col[i] === 1) { W.event = W.t[i]; break; }
  }
  S.allVars = [...S.rawCols, ...Object.keys(S.derived)];

  // file info
  let tmin = Infinity, tmax = -Infinity;
  for (const w of S.wells) { const t = S.data[w].t; tmin = Math.min(tmin, t[0]); tmax = Math.max(tmax, t[t.length - 1]); }
  $('#fileInfo').textContent = `${fileName}  ·  ${S.wells.length} wells  ·  ${S.rawCols.length} columns  ·  ${isoDate(tmin)} → ${isoDate(tmax)}`;
  $('#fileInfo').title = $('#fileInfo').textContent;
  $('#toolbar').setAttribute('aria-disabled', 'false');
  $('#saveCfg').disabled = false;
  $('#emptyState').classList.add('hidden');

  // plots: pending layout file > existing plots > last session > defaults
  if (S.pendingLayout) { applyLayout(S.pendingLayout); S.pendingLayout = null; return; }
  if (S.plots.length) { S.plots.forEach(sanitize); rebuildAll(); return; }
  let restored = null;
  try { restored = JSON.parse(localStorage.getItem(STORE_KEY) || 'null'); } catch (e) { /* storage unavailable */ }
  if (restored && restored.plots && restored.plots.length) { applyLayout(restored, true); return; }
  defaultPlots();
}

/* ================= series pipeline ================= */
function varKind(name) {
  const d = S.derived[name];
  if (d) return d.kind;
  if (S.flagCols.has(name)) return 'flag';
  return 'value';
}

function isoDate(ms) { return new Date(ms).toISOString().slice(0, 10); }

function binKey(ms, res) {
  if (res === 'W') { const d = Math.floor(ms / DAY); return (d - ((d + 3) % 7)) * DAY; } // Monday start
  if (res === 'M') { const dt = new Date(ms); return Date.UTC(dt.getUTCFullYear(), dt.getUTCMonth(), 1); }
  return ms;
}

// daily values with shut-in / zero masks applied
function maskedDaily(well, name, allowZeroGap) {
  const W = S.data[well];
  const src = W && W.cols[name];
  if (!src) return null;
  const kind = varKind(name);
  if (kind === 'flag' || kind === 'cum') return src;
  const out = Float64Array.from(src);
  const hrs = W.cols['Production_Hours'];
  const ex = S.settings.excludeShutin && hrs, zg = allowZeroGap && S.settings.zeroGap;
  if (ex || zg) {
    for (let i = 0; i < out.length; i++) {
      if (ex && hrs[i] === 0) out[i] = NaN;
      else if (zg && out[i] === 0) out[i] = NaN;
    }
  }
  return out;
}

function aggregate(t, y, res, mode) {
  if (res === 'D') return { t: Array.from(t), y: Array.from(y) };
  const tt = [], yy = [];
  let cur = null, sum = 0, n = 0, last = NaN;
  const flush = () => { tt.push(cur); yy.push(mode === 'last' ? last : (n ? sum / n : NaN)); };
  for (let i = 0; i < t.length; i++) {
    const k = binKey(t[i], res);
    if (k !== cur) { if (cur !== null) flush(); cur = k; sum = 0; n = 0; last = NaN; }
    const v = y[i];
    if (Number.isFinite(v)) { sum += v; n++; last = v; }
  }
  if (cur !== null) flush();
  return { t: tt, y: yy };
}

function movingAvg(y, N) {
  if (!N || N < 2) return y;
  const out = new Array(y.length);
  let sum = 0, n = 0;
  for (let i = 0; i < y.length; i++) {
    const v = y[i];
    if (Number.isFinite(v)) { sum += v; n++; }
    if (i >= N) { const o = y[i - N]; if (Number.isFinite(o)) { sum -= o; n--; } }
    out[i] = Number.isFinite(v) && n ? sum / n : NaN;
  }
  return out;
}

function summedComponents(well, names, res, ma) {
  let acc = null;
  for (const nm of names) {
    const d = maskedDaily(well, nm, false);
    if (!d) return null;
    const a = aggregate(S.data[well].t, d, res, 'mean');
    a.y = movingAvg(a.y, ma);
    if (!acc) acc = a;
    else acc.y = acc.y.map((v, i) => v + a.y[i]);
  }
  return acc;
}

/** Returns {t:[ms], y:[number|NaN]} for a well/variable under current settings. */
function getSeries(well, name, opts = {}) {
  const W = S.data[well];
  if (!W) return null;
  const res = opts.res || S.settings.res;
  const ma = opts.noMA ? 0 : S.settings.ma;
  const d = S.derived[name];
  if (d && d.kind === 'ratio') {
    const num = summedComponents(well, d.num, res, ma), den = summedComponents(well, d.den, res, ma);
    if (!num || !den) return null;
    let y = num.y.map((v, i) => (den.y[i] > (d.minDen || 0) && Number.isFinite(v)) ? v / den.y[i] * d.scale : NaN);
    if (S.settings.zeroGap) y = y.map(v => v === 0 ? NaN : v);
    return { t: num.t, y };
  }
  const daily = maskedDaily(well, name, true);
  if (!daily) return null;
  const kind = varKind(name);
  const a = aggregate(W.t, daily, res, kind === 'cum' ? 'last' : 'mean');
  if (kind === 'value') a.y = movingAvg(a.y, ma);
  return a;
}

/* ================= styling helpers ================= */
function wellIndex(w) { return Math.max(0, S.wells.indexOf(w)); }
function wellColor(w) { return PALETTE[wellIndex(w) % PALETTE.length]; }
function wellDash(w) { return DASHES[Math.floor(wellIndex(w) / PALETTE.length) % DASHES.length]; }
function varUnit(name) {
  const m = name.match(/_(t\/hr|Kg\/hr|kPa|Hz|Am3\/h|%|N\/m|Amp|bbld|bbl)$/i);
  if (/Temp|Subcool/i.test(name)) return '°C';
  if (name === 'SOR' || name === 'CSOR') return 'bbl/bbl';
  if (m) return m[1] === 'bbld' ? 'bbl/d' : m[1];
  return '';
}
function axisTitle(names) {
  if (!names.length) return '';
  const units = [...new Set(names.map(varUnit).filter(Boolean))];
  const label = names.length <= 2 ? names.join(', ') : `${names.length} variables`;
  return units.length === 1 ? `${label} [${units[0]}]` : label;
}

function baseLayout(p) {
  return {
    height: HEIGHTS[p.height] || HEIGHTS.M,
    margin: { l: 64, r: 64, t: 18, b: 40 },
    paper_bgcolor: '#ffffff', plot_bgcolor: '#ffffff',
    font: { family: 'Segoe UI, system-ui, sans-serif', size: 12, color: '#1b1f24' },
    legend: { orientation: 'h', x: 0, y: -0.14, yanchor: 'top', font: { size: 11 } },
    hoverlabel: { font: { size: 12 } },
    uirevision: `${p.id}`,
  };
}
const AXIS = { gridcolor: '#eceef1', linecolor: '#c5cbd4', zeroline: false, ticks: 'outside', tickcolor: '#c5cbd4', automargin: true };

function plotConfig(p) {
  return {
    responsive: true, displaylogo: false,
    modeBarButtonsToRemove: ['lasso2d', 'select2d', 'autoScale2d'],
    toImageButtonOptions: { filename: safeName(p.title), scale: 2 },
  };
}

/* ================= renderers ================= */
function renderTS(p) {
  const traces = [], exp = [];
  const xm = p.xmode;
  const wells = p.wells.filter(w => S.data[w]);
  const vars = p.vars.filter(v => S.allVars.includes(v.name));
  const multiW = wells.length > 1, multiV = vars.length > 1;
  const stacked = p.axisMode === 'stacked';
  let npts = 0;
  vars.forEach((v, vi) => {
    wells.forEach(w => {
      const s = getSeries(w, v.name);
      if (!s) return;
      let x, y = s.y;
      if (xm === 'aligned') {
        const st = S.data[w].start, xs = [], ys = [];
        for (let i = 0; i < s.t.length; i++) { const dd = Math.round((s.t[i] - st) / DAY); if (dd >= 0) { xs.push(dd); ys.push(s.y[i]); } }
        x = xs; y = ys;
      } else if (xm === 'event') {
        const ev = S.data[w].event;
        if (ev === null) return; // well without a redrill
        x = s.t.map(t => Math.round((t - ev) / DAY));
      } else x = s.t.map(isoDate);
      npts += x.length;
      const fullName = multiW && multiV ? `${w} · ${v.name}` : multiW ? w : v.name;
      const name = stacked && multiW ? w : fullName;
      const color = multiW ? wellColor(w) : PALETTE[vi % PALETTE.length];
      const dash = multiW && multiV && !stacked ? DASHES[vi % DASHES.length] : multiW ? wellDash(w) : 'solid';
      traces.push({
        type: 'scatter', mode: p.style, x, y: y.map(v => Number.isFinite(v) ? v : null), name,
        yaxis: vi === 0 ? 'y' : 'y' + (vi + 1), connectgaps: false,
        legendgroup: stacked && multiW ? w : fullName, showlegend: !(stacked && multiW && vi > 0),
        line: { color, width: 1.6, dash }, marker: { color, size: 5 },
        hovertemplate: `%{y:,.4~g}<extra>${escapeHtml(fullName)}</extra>`,
      });
      exp.push({ name: fullName, x, y });
    });
  });
  if (npts > 40000) traces.forEach(t => { t.type = 'scattergl'; });
  if (traces.length > 10) traces.forEach(t => { t.hovertemplate = `${escapeHtml(t.name)}: %{y:,.4~g}<extra></extra>`; });

  const L = baseLayout(p);
  const xTitle = { aligned: 'Days since first oil', event: 'Days since redrill' }[xm];
  L.xaxis = { ...AXIS, type: xTitle ? 'linear' : 'date', title: xTitle ? { text: xTitle } : undefined };
  layoutYAxes(p, vars, L, multiW, traces);
  L.hovermode = traces.length <= 10 ? 'x unified' : 'closest';
  L.uirevision = `${p.id}|${xm}|${p.axisMode}|` + JSON.stringify(vars.map(v => [v.name, v.side, v.log, v.min, v.max]));
  const shade = flagShapes(p, xm);
  L.shapes = shade.shapes;
  if (xm === 'event') L.shapes = L.shapes.concat([{ type: 'line', xref: 'x', yref: 'paper', x0: 0, x1: 0, y0: 0, y1: 1,
    line: { color: '#1b1f24', width: 1.5, dash: 'dash' } }]);
  if (shade.label) {
    const right = p.axisMode === 'stacked' && vars[0] && vars[0].side === 'R' ? false : p.axisMode === 'stacked';
    L.annotations = (L.annotations || []).concat([{ xref: 'paper', yref: 'paper', x: right ? 1 : 0, xanchor: right ? 'right' : 'left',
      y: 1, yanchor: 'bottom', showarrow: false, text: shade.label, font: { size: 11, color: '#5d6672' } }]);
  }
  L.margin.t = p.axisMode === 'stacked' ? 26 : shade.label ? 30 : 18;
  if (S.focus && xm === 'date') L.shapes = L.shapes.concat([{ type: 'line', xref: 'x', yref: 'paper', x0: S.focus.date, x1: S.focus.date,
    y0: 0, y1: 1, line: { color: '#1b1f24', width: 1.5, dash: 'dash' } }]);
  const synced = (S.settings.sync && S.xrange && S.xrange[xm]) || (p.focus && S.focus && xm === 'date' && S.focus.range);
  if (synced) { L.xaxis.range = synced.slice(); L.xaxis.autorange = false; }

  // CSV: wide table on union of x
  const keys = new Map();
  exp.forEach((tr, ti) => tr.x.forEach((xv, i) => {
    if (!keys.has(xv)) keys.set(xv, new Array(exp.length).fill(''));
    const v = tr.y[i]; keys.get(xv)[ti] = Number.isFinite(v) ? v : '';
  }));
  const sorted = [...keys.keys()].sort((a, b) => (a < b ? -1 : a > b ? 1 : 0));
  S.exports[p.id] = { header: [xm === 'date' ? 'Date' : 'Day', ...exp.map(e => e.name)], rows: sorted.map(k => [k, ...keys.get(k)]) };

  const empty = traces.length ? '' : xm === 'event' && wells.length ? 'None of the selected wells has a redrill date.'
    : 'Select at least one well and one variable.';
  return { traces, layout: L, empty };
}

axisRange.n = 0;
/** Manual min/max → explicit full range. The missing bound comes from the plotted data (ys), because
    Plotly's partial autorange ('min'/'max') is unreliable together with automargin redraws. */
function axisRange(v, ys) {
  const num = x => (x === '' || x === null || x === undefined || !Number.isFinite(+x)) ? null : +x;
  const tr = x => x === null ? null : v.log ? (x > 0 ? Math.log10(x) : null) : x;
  let lo = tr(num(v.min)), hi = tr(num(v.max));
  if (lo === null && hi === null) return {};
  if (lo === null || hi === null) {
    let dmin = Infinity, dmax = -Infinity;
    for (const y of ys) {
      const t = tr(y);
      if (t === null || !Number.isFinite(t)) continue;
      if (t < dmin) dmin = t;
      if (t > dmax) dmax = t;
    }
    if (dmin > dmax) return {};
    if (lo === null) lo = Math.min(dmin, hi) - 0.05 * Math.abs(hi - Math.min(dmin, hi) || 1);
    if (hi === null) hi = Math.max(dmax, lo) + 0.05 * Math.abs(Math.max(dmax, lo) - lo || 1);
  }
  // fresh per-axis uirevision so a stored GUI state never overrides the manual range
  return { range: [lo, hi], autorange: false, uirevision: `fixed-${++axisRange.n}` };
}

/** One y-axis per variable. Overlay: axes side by side; stacked: one track per variable sharing x. */
function layoutYAxes(p, vars, L, multiW, traces) {
  const ysOf = vi => {
    const id = vi === 0 ? 'y' : 'y' + (vi + 1), out = [];
    for (const t of traces) if (t.yaxis === id) for (const y of t.y) if (y !== null) out.push(y);
    return out;
  };
  const key = i => i === 0 ? 'yaxis' : 'yaxis' + (i + 1);
  const n = vars.length;
  const tint = vi => multiW ? '#5d6672' : PALETTE[vi % PALETTE.length];
  const base = (v, vi) => ({
    ...AXIS, type: v.log ? 'log' : 'linear', side: v.side === 'R' ? 'right' : 'left',
    title: { text: axisTitle([v.name]), font: { size: 11, color: tint(vi) }, standoff: 6 },
    tickfont: { size: 10, color: tint(vi) }, showline: true, linecolor: multiW ? '#c5cbd4' : tint(vi),
    ...axisRange(v, (v.min !== '' && v.max !== '') ? [] : ysOf(vi)),
  });
  if (!n) { L.yaxis = { ...AXIS }; return; }
  if (p.axisMode === 'stacked') {
    const H = Math.max(HEIGHTS[p.height] || HEIGHTS.M, n * 150 + 80);
    L.height = H;
    const gap = Math.min(0.08, 26 / H), h = (1 - gap * (n - 1)) / n;
    L.annotations = L.annotations || [];
    vars.forEach((v, vi) => {
      const top = 1 - vi * (h + gap);
      // track label sits horizontally above the track; a vertical title would spill into neighbours
      L[key(vi)] = { ...base(v, vi), domain: [Math.max(0, top - h), top], anchor: 'x', title: { text: '' } };
      L.annotations.push({ xref: 'paper', yref: 'paper', x: v.side === 'R' ? 1 : 0, xanchor: v.side === 'R' ? 'right' : 'left',
        y: top, yanchor: 'bottom', showarrow: false, text: `<b>${escapeHtml(axisTitle([v.name]))}</b>`,
        font: { size: 11, color: tint(vi) } });
    });
    L.xaxis.anchor = n === 1 ? 'y' : 'y' + n;
    const anyR = vars.some(v => v.side === 'R'), anyL = vars.some(v => v.side !== 'R');
    L.margin.l = anyL ? 70 : 24; L.margin.r = anyR ? 70 : 24;
    return;
  }
  // overlay: innermost axis of each side anchors to the plot edge; the rest float outward
  const lefts = [], rights = [];
  vars.forEach((v, vi) => (v.side === 'R' ? rights : lefts).push(vi));
  const card = cardEl(p.id), el = card && card.querySelector('.chart');
  const width = Math.max(320, (el && el.clientWidth) || 900);
  const f = Math.min(0.14, 64 / (width - 120));
  const d0 = Math.max(0, lefts.length - 1) * f, d1 = 1 - Math.max(0, rights.length - 1) * f;
  L.xaxis.domain = [Math.min(d0, 0.6), Math.max(d1, 0.4)];
  vars.forEach((v, vi) => {
    const isR = v.side === 'R', k = (isR ? rights : lefts).indexOf(vi);
    const ax = { ...base(v, vi), showgrid: vi === 0 };
    if (vi > 0) ax.overlaying = 'y';
    if (k === 0) ax.anchor = 'x';
    else { ax.anchor = 'free'; ax.position = isR ? Math.min(1, L.xaxis.domain[1] + k * f) : Math.max(0, L.xaxis.domain[0] - k * f); }
    L[key(vi)] = ax;
  });
  L.margin.l = lefts.length ? 64 : 24; L.margin.r = rights.length ? 64 : 24;
}

function flagShapes(p, xm) {
  const f = p.flag || {};
  if (!f.var) return { shapes: [] };
  const w = f.well && S.data[f.well] ? f.well : p.wells.find(x => S.data[x]);
  const W = w && S.data[w];
  const col = W && W.cols[f.var];
  if (!col) return { shapes: [] };
  const val = +f.value;
  if (xm === 'event' && W.event === null) return { shapes: [] };
  const tx = t => xm === 'aligned' ? (t - W.start) / DAY : xm === 'event' ? (t - W.event) / DAY : isoDate(t);
  const shapes = [];
  let start = null;
  for (let i = 0; i <= W.n; i++) {
    const on = i < W.n && col[i] === val;
    if (on && start === null) start = W.t[i];
    if (!on && start !== null) {
      shapes.push({ type: 'rect', xref: 'x', yref: 'paper', x0: tx(start), x1: tx(W.t[i - 1] + DAY), y0: 0, y1: 1,
        fillcolor: 'rgba(110,118,130,0.16)', line: { width: 0 }, layer: 'below' });
      start = null;
    }
  }
  return { shapes, label: `Shaded: ${f.var} = ${val}  (${w})` };
}

function inRange(t, p) {
  const from = p.from ? Date.parse(p.from) : -Infinity;
  const to = p.to ? Date.parse(p.to) + DAY - 1 : Infinity;
  return t >= from && t <= to;
}

function renderXP(p) {
  const traces = [], rows = [];
  const wells = p.wells.filter(w => S.data[w]);
  const byTime = p.colorBy === 'time';
  let ymin = Infinity, ymax = -Infinity;
  wells.forEach(w => {
    const sx = getSeries(w, p.x), sy = getSeries(w, p.y);
    if (!sx || !sy) return;
    const X = [], Y = [], D = [], C = [];
    for (let i = 0; i < sx.t.length; i++) {
      const xv = sx.y[i], yv = sy.y[i], t = sx.t[i];
      if (!Number.isFinite(xv) || !Number.isFinite(yv) || !inRange(t, p)) continue;
      if ((p.logX && xv <= 0) || (p.logY && yv <= 0)) continue;
      const dt = new Date(t);
      const yr = dt.getUTCFullYear() + (t - Date.UTC(dt.getUTCFullYear(), 0, 1)) / (365.25 * DAY);
      X.push(xv); Y.push(yv); D.push(isoDate(t)); C.push(yr);
      ymin = Math.min(ymin, yr); ymax = Math.max(ymax, yr);
      rows.push([w, isoDate(t), xv, yv]);
    }
    const sym = byTime ? SYMBOLS[wellIndex(w) % SYMBOLS.length] : 'circle';
    traces.push({
      type: X.length > 8000 ? 'scattergl' : 'scatter', mode: 'markers', name: w, x: X, y: Y, customdata: D,
      marker: byTime ? { color: C, coloraxis: 'coloraxis', size: 6, symbol: sym, opacity: 0.75 }
                     : { color: wellColor(w), size: 6, opacity: 0.6, line: { width: 0 } },
      hovertemplate: `<b>${escapeHtml(w)}</b> %{customdata}<br>${escapeHtml(p.x)}: %{x:,.4~g}<br>${escapeHtml(p.y)}: %{y:,.4~g}<extra></extra>`,
    });
  });
  const L = baseLayout(p);
  L.xaxis = { ...AXIS, title: { text: `${p.x}${varUnit(p.x) ? ' [' + varUnit(p.x) + ']' : ''}` }, type: p.logX ? 'log' : 'linear' };
  L.yaxis = { ...AXIS, title: { text: `${p.y}${varUnit(p.y) ? ' [' + varUnit(p.y) + ']' : ''}` }, type: p.logY ? 'log' : 'linear' };
  L.hovermode = 'closest';
  L.uirevision = `${p.id}|${p.x}|${p.y}|${p.logX}|${p.logY}`;
  if (byTime && Number.isFinite(ymin)) {
    L.coloraxis = { colorscale: [[0, '#c9ddf5'], [0.5, '#2a78d6'], [1, '#0b2f5e']], cmin: ymin, cmax: ymax,
      colorbar: { title: { text: 'Year' }, thickness: 12, tickformat: 'd' } };
  }
  S.exports[p.id] = { header: ['Well', 'Date', p.x, p.y], rows };
  const has = traces.some(t => t.x.length);
  return { traces, layout: L, empty: has ? '' : 'No overlapping data for the selected wells / variables / period.' };
}

const STAT_FN = {
  mean: a => a.reduce((s, v) => s + v, 0) / a.length,
  median: a => { const b = [...a].sort((x, y) => x - y), m = b.length >> 1; return b.length % 2 ? b[m] : (b[m - 1] + b[m]) / 2; },
  sum: a => a.reduce((s, v) => s + v, 0),
  max: a => a.reduce((m, v) => v > m ? v : m, -Infinity),
  min: a => a.reduce((m, v) => v < m ? v : m, Infinity),
  last: a => a[a.length - 1],
  p90: a => { const b = [...a].sort((x, y) => x - y); return b[Math.min(b.length - 1, Math.floor(0.9 * (b.length - 1)))]; },
};
const STAT_LABEL = { mean: 'Mean', median: 'Median', sum: 'Sum (daily values)', max: 'Max', min: 'Min', last: 'Last value', p90: 'P90' };

/** Day windows around a well's redrill date: {normal?, before, after} → [first day, last day] in ms, clipped to the data. */
function eventWindows(w, days) {
  const W = S.data[w], ev = W.event;
  if (ev === null) return null;
  const last = W.t[W.n - 1], span = (days - 1) * DAY;
  const out = { before: [ev - days * DAY, ev - DAY], after: [ev, Math.min(last, ev + span)] };
  const hrs = W.cols['Production_Hours'];
  if (hrs) {
    // last day before the redrill whose trailing 90-day mean uptime was normal
    let sum = 0, n = 0, end = null;
    for (let i = 0; i < W.n && W.t[i] < ev; i++) {
      if (Number.isFinite(hrs[i])) { sum += hrs[i]; n++; }
      if (i >= 90 && Number.isFinite(hrs[i - 90])) { sum -= hrs[i - 90]; n--; }
      if (i >= 59 && n && sum / n >= NORMAL_UPTIME_H) end = W.t[i];
    }
    if (end !== null) out.normal = [end - span, end];
  }
  return out;
}

/** One value per window: mean of daily values, or for a ratio (SOR, water cut) the ratio of the summed components. */
function windowValue(w, name, stat, [a, b]) {
  const d = S.derived[name], W = S.data[w];
  const pick = nm => {
    const s = getSeries(w, nm, { res: 'D', noMA: true });
    const out = [];
    if (s) for (let i = 0; i < s.t.length; i++) if (s.t[i] >= a && s.t[i] <= b && Number.isFinite(s.y[i])) out.push(s.y[i]);
    return out;
  };
  if (d && d.kind === 'ratio') {
    const tot = names => names.reduce((acc, nm) => acc + pick(nm).reduce((x, y) => x + y, 0), 0);
    const num = tot(d.num), den = tot(d.den);
    const days = Math.round((Math.min(b, W.t[W.n - 1]) - Math.max(a, W.t[0])) / DAY) + 1;
    return { v: den > 0 ? num / den * d.scale : NaN, n: Math.max(0, days) };
  }
  const vals = pick(name);
  return { v: vals.length ? STAT_FN[stat](vals) : NaN, n: vals.length };
}

function renderEventST(p) {
  const wells = p.wells.filter(w => S.data[w] && S.data[w].event !== null);
  const keys = ['normal', 'before', 'after'].filter(k => k !== 'normal' || p.normal);
  const rows = [], traces = [];
  const isRatio = S.derived[p.var] && S.derived[p.var].kind === 'ratio';
  keys.forEach(k => {
    const X = [], Y = [], H = [];
    wells.forEach(w => {
      const win = eventWindows(w, p.window)[k];
      if (!win) return;
      const r = windowValue(w, p.var, p.stat, win);
      X.push(w); Y.push(Number.isFinite(r.v) ? r.v : null);
      H.push(`${isoDate(win[0])} → ${isoDate(win[1])} · ${r.n} d`);
      rows.push([w, EVENT_WINDOWS[k].label, isoDate(win[0]), isoDate(win[1]), r.n, Number.isFinite(r.v) ? r.v : '']);
    });
    traces.push({ type: 'bar', name: EVENT_WINDOWS[k].label, x: X, y: Y, customdata: H,
      marker: { color: EVENT_WINDOWS[k].color, line: { width: 0 } }, texttemplate: '%{y:,.3~s}', textposition: 'outside', cliponaxis: false,
      hovertemplate: `<b>%{x}</b> · ${EVENT_WINDOWS[k].label}<br>%{y:,.4~g}<br>%{customdata}<extra></extra>` });
  });
  const L = baseLayout(p);
  const unit = varUnit(p.var);
  const what = isRatio ? `${p.var} (window total)` : `${STAT_LABEL[p.stat]} of ${p.var}`;
  L.yaxis = { ...AXIS, title: { text: `${what}${unit ? ' [' + unit + ']' : ''}` }, type: p.logY ? 'log' : 'linear' };
  L.xaxis = { ...AXIS, type: 'category' };
  L.barmode = 'group'; L.bargap = 0.25; L.hovermode = 'closest';
  L.uirevision = `${p.id}|event|${p.var}|${p.stat}|${p.window}`;
  S.exports[p.id] = { header: ['Well', 'Window', 'From', 'To', 'Days', what], rows };
  const has = traces.some(t => t.y.length);
  return { traces, layout: L, empty: has ? '' : 'None of the selected wells has a redrill date.' };
}

function renderST(p) {
  if (p.period === 'event' && S.eventCol) return renderEventST(p);
  const wells = p.wells.filter(w => S.data[w]);
  const traces = [];
  const L = baseLayout(p);
  L.hovermode = 'closest';
  L.uirevision = `${p.id}|${p.chart}|${p.var}|${p.stat}`;
  const vals = {};
  wells.forEach(w => {
    const s = getSeries(w, p.var, { res: 'D', noMA: true });
    if (!s) return;
    const a = [], d = [];
    for (let i = 0; i < s.t.length; i++) if (Number.isFinite(s.y[i]) && inRange(s.t[i], p)) { a.push(s.y[i]); d.push(isoDate(s.t[i])); }
    vals[w] = { a, d };
  });
  const unit = varUnit(p.var);
  if (p.chart === 'box') {
    const rows = [];
    wells.forEach(w => {
      if (!vals[w] || !vals[w].a.length) return;
      traces.push({ type: 'box', name: w, y: vals[w].a, boxpoints: 'outliers', boxmean: true,
        marker: { color: wellColor(w), size: 3, opacity: 0.6 }, line: { width: 1.5 }, showlegend: false });
      vals[w].a.forEach((v, i) => rows.push([w, vals[w].d[i], v]));
    });
    L.yaxis = { ...AXIS, title: { text: `${p.var}${unit ? ' [' + unit + ']' : ''}` }, type: p.logY ? 'log' : 'linear' };
    L.xaxis = { ...AXIS, type: 'category' };
    S.exports[p.id] = { header: ['Well', 'Date', p.var], rows };
  } else {
    const X = [], Y = [], C = [], N = [];
    wells.forEach(w => {
      const a = vals[w] && vals[w].a;
      if (!a || !a.length) return;
      X.push(w); Y.push(STAT_FN[p.stat](a)); C.push(wellColor(w)); N.push(a.length);
    });
    traces.push({ type: 'bar', x: X, y: Y, customdata: N, marker: { color: C, line: { width: 0 } },
      texttemplate: '%{y:,.3~s}', textposition: 'outside', cliponaxis: false,
      hovertemplate: `<b>%{x}</b><br>${STAT_LABEL[p.stat]}: %{y:,.4~g}<br>n = %{customdata} days<extra></extra>` });
    L.yaxis = { ...AXIS, title: { text: `${STAT_LABEL[p.stat]} of ${p.var}${unit ? ' [' + unit + ']' : ''}` }, type: p.logY ? 'log' : 'linear' };
    L.xaxis = { ...AXIS, type: 'category' };
    L.bargap = 0.35;
    S.exports[p.id] = { header: ['Well', `${STAT_LABEL[p.stat]} of ${p.var}`, 'Days'], rows: X.map((w, i) => [w, Y[i], N[i]]) };
  }
  const has = traces.some(t => (t.y || []).length);
  return { traces, layout: L, empty: has ? '' : 'No data for the selected wells / period.' };
}

function renderPlot(p) {
  const card = cardEl(p.id);
  if (!card || !S.wells.length) return;
  const el = card.querySelector('.chart');
  const msg = card.querySelector('.chart-msg');
  const r = p.type === 'ts' ? renderTS(p) : p.type === 'xp' ? renderXP(p) : renderST(p);
  msg.textContent = r.empty || '';
  msg.classList.toggle('on', !!r.empty);
  const H = r.layout.height || HEIGHTS[p.height] || HEIGHTS.M;
  el.style.height = H + 'px';
  card.querySelector('.cfg').style.maxHeight = H + 8 + 'px';
  Plotly.react(el, r.traces, r.layout, plotConfig(p));
  if (!el._bgBound) {
    el._bgBound = true;
    el.on('plotly_relayout', ev => onRelayout(p.id, ev));
  }
}

function onRelayout(id, ev) {
  if (S.syncing || !S.settings.sync) return;
  const p = plotById(id);
  if (!p || p.type !== 'ts') return;
  let range = null, auto = false;
  if (ev['xaxis.autorange']) auto = true;
  else if (ev['xaxis.range[0]'] !== undefined) range = [ev['xaxis.range[0]'], ev['xaxis.range[1]']];
  else if (Array.isArray(ev['xaxis.range'])) range = ev['xaxis.range'];
  else return;
  S.xrange = S.xrange || {};
  S.xrange[p.xmode] = auto ? null : range;
  S.syncing = true;
  const jobs = [];
  for (const q of S.plots) {
    if (q.id === p.id || q.type !== 'ts' || q.xmode !== p.xmode) continue;
    const el = cardEl(q.id) && cardEl(q.id).querySelector('.chart');
    if (el && el.data) jobs.push(Plotly.relayout(el, auto ? { 'xaxis.autorange': true } : { 'xaxis.range': range.slice() }));
  }
  Promise.all(jobs).finally(() => { S.syncing = false; });
}

/* ================= plot config model ================= */
function hasVar(n) { return S.allVars.includes(n); }
function mkVar(name, side = 'L') { return { name, side, log: false, min: '', max: '' }; }
function firstVar(...cands) { return cands.find(hasVar) || S.allVars[0]; }

function newPlot(type, extra = {}) {
  const id = S.nextId++;
  const base = { id, type, title: `${TYPE_LABEL[type]} ${id}`, height: 'M', collapsed: false };
  if (type === 'ts') Object.assign(base, {
    wells: S.wells.slice(0, 1), vars: [mkVar(firstVar('Prod_Oil_rate_bbld'))],
    xmode: 'date', axisMode: 'overlay', style: 'lines', flag: { var: '', well: '', value: 0 },
  });
  if (type === 'xp') Object.assign(base, {
    wells: S.wells.slice(), x: firstVar('Inj_Steam_rate_bbld'), y: firstVar('Prod_Oil_rate_bbld'),
    colorBy: 'well', logX: false, logY: false, from: '', to: '',
  });
  if (type === 'st') Object.assign(base, {
    wells: S.wells.slice(), var: firstVar('Prod_Oil_rate_bbld'), stat: 'mean', chart: 'bar', logY: false, from: '', to: '',
    period: 'range', window: 180, normal: true,
  });
  return Object.assign(base, extra);
}

function sanitize(p) {
  p.wells = (p.wells || []).filter(w => S.data[w]);
  if (!p.wells.length && S.wells.length) p.wells = p.type === 'ts' ? [S.wells[0]] : S.wells.slice();
  if (p.type === 'ts') {
    p.axisMode = p.axisMode || 'overlay';
    // migrate v1 layouts ({axis:'L'|'R'} + plot-level logL/logR) to per-variable axes
    p.vars = (p.vars || []).filter(v => hasVar(v.name)).map(v => ({
      name: v.name, side: v.side || v.axis || 'L',
      log: v.log !== undefined ? !!v.log : !!(v.axis === 'R' ? p.logR : p.logL),
      min: v.min ?? '', max: v.max ?? '',
    }));
    delete p.logL; delete p.logR;
    if (!p.vars.length) p.vars = [mkVar(firstVar('Prod_Oil_rate_bbld'))];
    p.flag = p.flag || { var: '', well: '', value: 0 };
    if (p.flag.var && !hasVar(p.flag.var)) p.flag.var = '';
    if (p.flag.well && !S.data[p.flag.well]) p.flag.well = '';
  }
  if (p.type === 'xp') { if (!hasVar(p.x)) p.x = firstVar('Inj_Steam_rate_bbld'); if (!hasVar(p.y)) p.y = firstVar('Prod_Oil_rate_bbld'); }
  if (p.type === 'st') {
    if (!hasVar(p.var)) p.var = firstVar('Prod_Oil_rate_bbld');
    p.period = p.period === 'event' ? 'event' : 'range';
    p.window = [90, 180, 365].includes(+p.window) ? +p.window : 180;
    p.normal = p.normal !== false;
  }
  if (p.type === 'ts' && p.xmode === 'event' && !S.eventCol) p.xmode = 'date';
}

function defaultPlots() {
  S.plots = [];
  const w0 = S.wells[0];
  const v = [];
  if (hasVar('Prod_Oil_rate_bbld')) v.push(mkVar('Prod_Oil_rate_bbld', 'L'));
  if (hasVar('Inj_Steam_rate_bbld')) v.push(mkVar('Inj_Steam_rate_bbld', 'R'));
  if (hasVar('SOR')) v.push(mkVar('SOR', 'R'));
  S.plots.push(newPlot('ts', { title: `${w0} – Production & SOR`, vars: v.length ? v : undefined }));
  S.plots[0].vars = S.plots[0].vars || [mkVar(S.allVars[0])];
  S.plots.push(newPlot('ts', { title: 'Oil rate – all wells', wells: S.wells.slice(), vars: [mkVar(firstVar('Prod_Oil_rate_bbld'))] }));
  S.settings.res = S.settings.res || 'D';
  rebuildAll();
  persist();
}

/** Show the well and date the Profile Viewer was on: one reusable card at the top,
    zoomed to ±6 months around the date, with a marker line on every calendar-date plot. */
function applyFocus(f) {
  const ms = f && Date.parse(f.date);
  if (!f || !S.data[f.well] || !Number.isFinite(ms)) return;
  S.focus = { well: f.well, date: isoDate(ms), range: [isoDate(ms - FOCUS_HALF_SPAN), isoDate(ms + FOCUS_HALF_SPAN)] };
  let p = S.plots.find(q => q.focus && q.type === 'ts');
  if (!p) {
    const vars = ['Prod_Oil_rate_bbld', 'Prod_Water_rate_bbld', 'Inj_Steam_rate_bbld', 'Subcool_Min'].filter(hasVar).map(n => mkVar(n));
    p = newPlot('ts', { focus: true, axisMode: 'stacked', ...(vars.length ? { vars } : {}) });
    S.plots.unshift(p);
  }
  p.wells = [f.well];
  p.xmode = 'date';
  p.title = `${f.well} – from Profile Viewer`;
  if (S.settings.sync) S.xrange = { ...S.xrange, date: S.focus.range.slice() };
  rebuildAll();
  persist();
  // keep the card's header clear of the sticky toolbar
  const top = cardEl(p.id).getBoundingClientRect().top + window.scrollY - $('.sticky').offsetHeight - 8;
  window.scrollTo({ top: Math.max(0, top) });
}

/** Well of the top plot and the middle of its visible date range, for the Profile Viewer.
    A plot without a calendar-date axis falls back to the synced date zoom, else sends the well only. */
function topFocus() {
  const p = S.plots[0];
  const well = p && (p.wells || []).find(w => S.data[w]);
  if (!well) return null;
  let range = null;
  if (p.type === 'ts' && p.xmode === 'date') {
    const el = cardEl(p.id) && cardEl(p.id).querySelector('.chart');
    range = el && el._fullLayout && el._fullLayout.xaxis && el._fullLayout.xaxis.range;
  }
  if (!range && S.xrange && S.xrange.date) range = S.xrange.date;
  const a = range && Date.parse(range[0]), b = range && Date.parse(range[1]);
  return { well, date: Number.isFinite(a) && Number.isFinite(b) ? isoDate(Math.round((a + b) / 2 / DAY) * DAY) : null };
}

function plotById(id) { return S.plots.find(p => p.id === id); }
function cardEl(id) { return document.querySelector(`.card[data-id="${id}"]`); }

/* ================= layout save / load ================= */
function layoutSnapshot() {
  return {
    app: 'SAGD_Prod_Viewer', version: 1, savedAt: new Date().toISOString(), sourceFile: S.fileName,
    settings: { ...S.settings },
    plots: S.plots.map(p => JSON.parse(JSON.stringify(p))),
  };
}
function persist() { try { localStorage.setItem(STORE_KEY, JSON.stringify(layoutSnapshot())); } catch (e) { /* storage unavailable */ } }

function applyLayout(cfg, quiet) {
  if (!cfg || !Array.isArray(cfg.plots)) { toast('Not a valid layout file.', true); return; }
  Object.assign(S.settings, cfg.settings || {});
  S.plots = cfg.plots.filter(p => TYPE_LABEL[p.type]).map(p => ({ ...p }));
  S.nextId = Math.max(0, ...S.plots.map(p => +p.id || 0)) + 1;
  S.plots.forEach((p, i) => { if (!p.id) p.id = S.nextId + i; });
  S.nextId = Math.max(...S.plots.map(p => p.id), 0) + 1;
  S.plots.forEach(sanitize);
  S.xrange = {};
  syncToolbar();
  rebuildAll();
  persist();
  if (!quiet) toast(`Layout applied: ${S.plots.length} plots`);
}

function saveLayoutFile() {
  const blob = new Blob([JSON.stringify(layoutSnapshot(), null, 2)], { type: 'application/json' });
  download(`sagd_viewer_layout_${new Date().toISOString().slice(0, 10)}.json`, blob);
}

async function loadLayoutFile(file) {
  try {
    const cfg = JSON.parse(await file.text());
    if (!S.wells.length) { S.pendingLayout = cfg; toast('Layout stored — it will be applied when an Excel file is opened.'); return; }
    applyLayout(cfg);
  } catch (e) { toast('Could not read layout: ' + e.message, true); }
}

/* ================= DOM building ================= */
function $(sel, root = document) { return root.querySelector(sel); }
function h(tag, attrs = {}, ...kids) {
  const el = document.createElement(tag);
  for (const [k, v] of Object.entries(attrs || {})) {
    if (v === undefined || v === null || v === false) continue;
    if (k === 'class') el.className = v;
    else if (k.startsWith('on')) el.addEventListener(k.slice(2), v);
    else if (k === 'value') el.value = v;
    else if (k === 'checked') el.checked = !!v;
    else el.setAttribute(k, v === true ? '' : v);
  }
  for (const c of kids.flat()) if (c !== null && c !== undefined && c !== false) el.append(c instanceof Node ? c : document.createTextNode(c));
  return el;
}
function escapeHtml(s) { return String(s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c])); }
function safeName(s) { return String(s || 'plot').replace(/[\\/:*?"<>|]+/g, '_').trim() || 'plot'; }

function varGroup(n) {
  if (S.derived[n]) return S.derived[n].group || 'Derived';
  if (S.flagCols.has(n)) return 'Operation flags';
  if (/^Prod_|Production_Hours/i.test(n)) return 'Production';
  if (/Inj|Blanket/i.test(n)) return 'Injection';
  if (/Pressure|Choke/i.test(n)) return 'Pressure / Choke';
  if (/Frequency|Current|Torque|Emulsion/i.test(n)) return 'ESP / Pump';
  if (/Subcool/i.test(n)) return 'Subcool';
  if (/Temp|WHT/i.test(n)) return 'Temperature';
  return 'Other';
}
const GROUP_ORDER = ['Production', 'Derived', 'Injection', 'Pressure / Choke', 'ESP / Pump', 'Temperature', 'Subcool', 'Operation flags', 'Other'];

function varSelect(value, onChange, { placeholder, only } = {}) {
  const sel = h('select', { onchange: () => onChange(sel.value) });
  if (placeholder) sel.append(h('option', { value: '' }, placeholder));
  const groups = {};
  for (const v of (only || S.allVars)) (groups[varGroup(v)] = groups[varGroup(v)] || []).push(v);
  for (const g of GROUP_ORDER) {
    if (!groups[g]) continue;
    const og = h('optgroup', { label: g });
    groups[g].forEach(v => og.append(h('option', { value: v }, v)));
    sel.append(og);
  }
  sel.value = value || '';
  return sel;
}

function seg(options, value, onChange, small) {
  const el = h('div', { class: 'seg' + (small ? ' small' : '') });
  for (const [v, label] of options) {
    el.append(h('button', { class: String(v) === String(value) ? 'on' : '', type: 'button', onclick: () => onChange(v) }, label));
  }
  return el;
}

function section(title, links, ...body) {
  const head = h('h4', {}, title);
  if (links) head.append(h('span', { class: 'links' }, ...links));
  return h('div', {}, head, ...body);
}

function wellPicker(p, update) {
  const chips = h('div', { class: 'chips' });
  for (const w of S.wells) {
    const on = p.wells.includes(w);
    const chip = h('span', { class: 'chip' + (on ? ' on' : ''), title: w, onclick: () => {
      const set = new Set(p.wells);
      set.has(w) ? set.delete(w) : set.add(w);
      p.wells = S.wells.filter(x => set.has(x));
      update(true);
    } }, h('span', { class: 'dot', style: `background:${wellColor(w)}` }), w);
    chips.append(chip);
  }
  const prefixes = [...new Set(S.wells.map(w => (w.match(/^[A-Za-z]+/) || [''])[0]).filter(Boolean))];
  const links = [h('a', { onclick: () => { p.wells = S.wells.slice(); update(true); } }, 'All'),
    h('a', { onclick: () => { p.wells = []; update(true); } }, 'None')];
  if (prefixes.length > 1 && prefixes.length <= 6) {
    prefixes.forEach(pre => links.push(h('a', { title: `Only ${pre}-wells`, onclick: () => {
      p.wells = S.wells.filter(w => w.startsWith(pre)); update(true);
    } }, `${pre}*`)));
  }
  return section(`Wells (${p.wells.length}/${S.wells.length})`, links, chips);
}

function heightRow(p, update) {
  return h('div', { class: 'row' }, h('span', { class: 'tlabel' }, 'Height'),
    seg([['S', 'S'], ['M', 'M'], ['L', 'L'], ['XL', 'XL']], p.height, v => { p.height = v; update(true); }, true));
}

function dateRangeRows(p, update) {
  const from = h('input', { type: 'date', value: p.from || '', onchange: () => { p.from = from.value; update(false); } });
  const to = h('input', { type: 'date', value: p.to || '', onchange: () => { p.to = to.value; update(false); } });
  return section('Period', [h('a', { onclick: () => { p.from = ''; p.to = ''; update(true); } }, 'All')],
    h('div', { class: 'row' }, h('span', { class: 'tlabel' }, 'From'), from),
    h('div', { class: 'row' }, h('span', { class: 'tlabel' }, 'To    '), to));
}

function buildCfg(p) {
  const update = (rebuild) => { persist(); if (rebuild) fillCfg(p); renderPlot(p); };
  const parts = [];
  if (p.type === 'ts') {
    parts.push(wellPicker(p, update));
    const single = p.wells.length <= 1;
    const num = (val, ph, set) => h('input', { type: 'number', step: 'any', placeholder: ph, value: val,
      onchange: e => { set(e.target.value); update(false); } });
    const rows = p.vars.map((v, i) => h('div', { class: 'varblock' },
      h('div', { class: 'varrow' },
        h('span', { class: 'swatch', style: single ? `background:${PALETTE[i % PALETTE.length]}` : '' }),
        varSelect(v.name, val => { v.name = val; update(true); }),
        h('button', { class: 'xbtn', title: 'Remove', onclick: () => { p.vars.splice(i, 1); update(true); } }, '×')),
      h('div', { class: 'varopt' },
        seg([['L', 'Left'], ['R', 'Right']], v.side, a => { v.side = a; update(true); }, true),
        h('label', { class: 'chk' }, h('input', { type: 'checkbox', checked: v.log, onchange: e => { v.log = e.target.checked; update(false); } }), 'Log'),
        num(v.min, 'min', x => { v.min = x; }), num(v.max, 'max', x => { v.max = x; }))));
    rows.push(h('div', { class: 'varrow' }, varSelect('', val => {
      if (val) { p.vars.push(mkVar(val, p.vars.length % 2 ? 'R' : 'L')); update(true); }
    }, { placeholder: '+ Add variable (new axis)…' })));
    parts.push(section('Variables  (one y-axis each)', null, ...rows,
      h('div', { class: 'row' }, h('span', { class: 'tlabel' }, 'Y axes'),
        seg([['overlay', 'Overlay'], ['stacked', 'Stacked tracks']], p.axisMode, m => { p.axisMode = m; update(true); }, true)),
      h('div', { class: 'note' }, 'Leave min / max empty for auto. Overlay places axes side by side; Stacked gives each variable its own track.')));
    const xOpts = [['date', 'Calendar date'], ['aligned', 'Days since first oil']];
    if (S.eventCol) xOpts.push(['event', 'Days since redrill']);
    parts.push(section('X axis', null, seg(xOpts, p.xmode, v => { p.xmode = v; update(true); }, true),
      p.xmode === 'event' ? h('div', { class: 'note' }, `Day 0 = first day ${S.eventCol} = 1. Wells without a redrill are left out.`) : null));
    parts.push(section('Display', null,
      h('div', { class: 'row' }, seg([['lines', 'Lines'], ['markers', 'Markers'], ['lines+markers', 'Both']], p.style, v => { p.style = v; update(true); }, true)),
      heightRow(p, update)));
    const flags = S.rawCols.filter(c => S.flagCols.has(c));
    const f = p.flag;
    const wellSel = h('select', { onchange: e => { f.well = e.target.value; update(false); } },
      h('option', { value: '' }, `First selected well`), ...S.wells.map(w => h('option', { value: w }, w)));
    wellSel.value = f.well || '';
    parts.push(section('Background shading', null,
      h('div', { class: 'row' }, varSelect(f.var, v => { f.var = v; update(true); }, { placeholder: 'None', only: flags })),
      f.var ? h('div', { class: 'row' }, h('span', { class: 'tlabel' }, 'when ='), seg([[0, '0'], [1, '1']], f.value, v => { f.value = v; update(true); }, true), wellSel) : null,
      h('div', { class: 'note' }, 'Shades periods where the selected 0/1 flag equals the chosen value.')));
  }
  if (p.type === 'xp') {
    parts.push(wellPicker(p, update));
    parts.push(section('Axes', null,
      h('div', { class: 'row' }, h('span', { class: 'tlabel' }, 'X'), varSelect(p.x, v => { p.x = v; update(true); })),
      h('div', { class: 'row' }, h('span', { class: 'tlabel' }, 'Y'), varSelect(p.y, v => { p.y = v; update(true); })),
      h('div', { class: 'row' }, h('button', { class: 'btn', type: 'button', onclick: () => { [p.x, p.y] = [p.y, p.x]; update(true); } }, '⇄ Swap X/Y')),
      h('div', { class: 'row' },
        h('label', { class: 'chk' }, h('input', { type: 'checkbox', checked: p.logX, onchange: e => { p.logX = e.target.checked; update(false); } }), 'Log X'),
        h('label', { class: 'chk' }, h('input', { type: 'checkbox', checked: p.logY, onchange: e => { p.logY = e.target.checked; update(false); } }), 'Log Y'))));
    parts.push(section('Color by', null, seg([['well', 'Well'], ['time', 'Time (year)']], p.colorBy, v => { p.colorBy = v; update(true); }, true),
      h('div', { class: 'note' }, 'Uses the global resolution and moving average. "Time" colors points by date; wells are told apart by marker shape.')));
    parts.push(dateRangeRows(p, update));
    parts.push(section('Size', null, heightRow(p, update)));
  }
  if (p.type === 'st') {
    parts.push(wellPicker(p, update));
    const statSel = h('select', { onchange: e => { p.stat = e.target.value; update(false); } },
      ...Object.entries(STAT_LABEL).map(([k, l]) => h('option', { value: k }, l)));
    statSel.value = p.stat;
    parts.push(section('Variable', null,
      h('div', { class: 'row' }, varSelect(p.var, v => { p.var = v; update(true); })),
      p.period === 'event' ? null : h('div', { class: 'row' }, seg([['bar', 'Bar (statistic)'], ['box', 'Box (distribution)']], p.chart, v => { p.chart = v; update(true); }, true)),
      p.chart === 'bar' ? h('div', { class: 'row' }, h('span', { class: 'tlabel' }, 'Statistic'), statSel) : null,
      h('div', { class: 'row' }, h('label', { class: 'chk' }, h('input', { type: 'checkbox', checked: p.logY, onchange: e => { p.logY = e.target.checked; update(false); } }), 'Log Y')),
      h('div', { class: 'note' }, 'Computed from daily values within the period. Shut-in / zero filters apply; moving average and resolution do not.')));
    if (S.eventCol) {
      const ev = p.period === 'event';
      parts.push(section('Compare', null,
        h('div', { class: 'row' }, seg([['range', 'Date range'], ['event', 'Before / after redrill']], p.period,
          v => { p.period = v; if (v === 'event') p.chart = 'bar'; update(true); }, true)),
        ev ? h('div', { class: 'row' }, h('span', { class: 'tlabel' }, 'Window'),
          seg([[90, '90 d'], [180, '180 d'], [365, '365 d']], p.window, v => { p.window = v; update(true); }, true)) : null,
        ev ? h('div', { class: 'row' }, h('label', { class: 'chk' }, h('input', { type: 'checkbox', checked: p.normal,
          onchange: e => { p.normal = e.target.checked; update(false); } }), 'Last normal period')) : null,
        ev ? h('div', { class: 'note' }, `Just before = the window ending the day before the redrill; after = from the redrill day. ` +
          `Last normal = the window ending on the last day before the redrill whose 90-day mean Production_Hours was ≥ ${NORMAL_UPTIME_H} h. ` +
          'Ratios (SOR, water cut) are window totals.') : null));
    }
    if (p.period !== 'event' || !S.eventCol) parts.push(dateRangeRows(p, update));
    parts.push(section('Size', null, heightRow(p, update)));
  }
  return parts;
}

function fillCfg(p) {
  const cfg = cardEl(p.id).querySelector('.cfg');
  const scroll = cfg.scrollTop;
  cfg.replaceChildren(...buildCfg(p));
  cfg.scrollTop = scroll;
}

function buildCard(p) {
  const title = h('input', { class: 'title', type: 'text', value: p.title, title: 'Click to rename',
    onchange: () => { p.title = title.value; persist(); } });
  const card = h('section', { class: 'card' + (p.collapsed ? ' collapsed' : ''), 'data-id': p.id },
    h('div', { class: 'card-head' },
      title,
      h('span', { class: 'badge' }, TYPE_LABEL[p.type]),
      h('div', { class: 'icons' },
        h('button', { class: 'ibtn' + (p.collapsed ? '' : ' on'), title: 'Show / hide settings', onclick: e => {
          p.collapsed = !p.collapsed; card.classList.toggle('collapsed', p.collapsed); e.currentTarget.classList.toggle('on', !p.collapsed);
          persist(); resizeChart(p.id);
        } }, '⚙ Settings'),
        h('button', { class: 'ibtn', title: 'Move up', onclick: () => movePlot(p.id, -1) }, '↑'),
        h('button', { class: 'ibtn', title: 'Move down', onclick: () => movePlot(p.id, 1) }, '↓'),
        h('button', { class: 'ibtn', title: 'Duplicate', onclick: () => duplicatePlot(p.id) }, '⧉ Copy'),
        h('button', { class: 'ibtn', title: 'Download PNG', onclick: () => exportPNG(p) }, 'PNG'),
        h('button', { class: 'ibtn', title: 'Download plotted data as CSV', onclick: () => exportCSV(p) }, 'CSV'),
        h('button', { class: 'ibtn danger', title: 'Delete plot', onclick: () => removePlot(p.id) }, '✕'))),
    h('div', { class: 'card-body' },
      h('aside', { class: 'cfg' }),
      h('div', { class: 'chart-wrap' }, h('div', { class: 'chart' }), h('div', { class: 'chart-msg' }))));
  return card;
}

function addCard(p, scroll) {
  const card = buildCard(p);
  $('#plots').append(card);
  fillCfg(p);
  renderPlot(p);
  if (scroll) card.scrollIntoView({ behavior: 'smooth', block: 'start' });
}

function rebuildAll() {
  // no Plotly.purge here: a pending automargin redraw would then hit a cleared _fullLayout and throw
  $('#plots').replaceChildren();
  $('#plots').className = 'plots cols-' + (S.settings.cols || 1);
  S.plots.forEach(p => addCard(p));
}

function renderAll() { S.plots.forEach(renderPlot); }
// re-render (not just resize): overlay axis offsets depend on the chart width
function resizeChart(id) { setTimeout(() => { const p = plotById(id); if (p) renderPlot(p); }, 30); }

function movePlot(id, dir) {
  const i = S.plots.findIndex(p => p.id === id), j = i + dir;
  if (j < 0 || j >= S.plots.length) return;
  [S.plots[i], S.plots[j]] = [S.plots[j], S.plots[i]];
  const a = cardEl(id), b = cardEl(S.plots[i].id);
  dir < 0 ? b.before(a) : b.after(a);
  persist();
  a.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
}
function duplicatePlot(id) {
  const src = plotById(id);
  const p = JSON.parse(JSON.stringify(src));
  p.id = S.nextId++;
  p.title = src.title + ' (copy)';
  S.plots.splice(S.plots.indexOf(src) + 1, 0, p);
  const card = buildCard(p);
  cardEl(id).after(card);
  fillCfg(p); renderPlot(p); persist();
  card.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
}
function removePlot(id) {
  const c = cardEl(id);
  if (c) c.remove();
  S.plots = S.plots.filter(p => p.id !== id);
  delete S.exports[id];
  persist();
}

/* ================= export ================= */
function download(name, blob) {
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = name;
  document.body.append(a);
  a.click();
  setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 1500);
}
function exportPNG(p) {
  const el = cardEl(p.id).querySelector('.chart');
  Plotly.downloadImage(el, { format: 'png', filename: safeName(p.title), width: el.clientWidth, height: el.clientHeight, scale: 2 });
}
function exportCSV(p) {
  const ex = S.exports[p.id];
  if (!ex || !ex.rows.length) { toast('Nothing to export.', true); return; }
  const esc = v => { const s = typeof v === 'number' ? String(v) : String(v ?? ''); return /[",\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s; };
  const meta = `# ${p.title} | file: ${S.fileName} | resolution: ${S.settings.res} | MA: ${S.settings.ma} | exclude shut-in: ${S.settings.excludeShutin} | zero as gap: ${S.settings.zeroGap}`;
  const lines = [meta, ex.header.map(esc).join(','), ...ex.rows.map(r => r.map(esc).join(','))];
  download(`${safeName(p.title)}.csv`, new Blob(['﻿' + lines.join('\r\n')], { type: 'text/csv;charset=utf-8' }));
}

/* ================= toolbar & global events ================= */
function syncToolbar() {
  const st = S.settings;
  document.querySelectorAll('#resSeg button').forEach(b => b.classList.toggle('on', b.dataset.v === st.res));
  document.querySelectorAll('#colSeg button').forEach(b => b.classList.toggle('on', b.dataset.v === String(st.cols)));
  $('#maSel').value = String(st.ma);
  $('#shutinChk').checked = !!st.excludeShutin;
  $('#zeroChk').checked = !!st.zeroGap;
  $('#syncChk').checked = st.sync !== false;
}
function setSetting(k, v) {
  S.settings[k] = v;
  syncToolbar();
  persist();
  if (k === 'cols') { $('#plots').className = 'plots cols-' + v; S.plots.forEach(p => resizeChart(p.id)); return; }
  if (k === 'sync') { S.xrange = {}; return; }
  renderAll();
}

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

function handleFiles(files) {
  for (const f of files) {
    if (/\.json$/i.test(f.name)) loadLayoutFile(f);
    else if (/\.(xlsx|xlsm|xls|xlsb|csv)$/i.test(f.name)) { loadExcel(f); }
    else toast(`Unsupported file: ${f.name}`, true);
  }
}

function init() {
  syncToolbar();
  $('#fileInput').addEventListener('change', e => { handleFiles(e.target.files); e.target.value = ''; });
  $('#fileInput2').addEventListener('change', e => { handleFiles(e.target.files); e.target.value = ''; });
  $('#cfgInput').addEventListener('change', e => { handleFiles(e.target.files); e.target.value = ''; });
  $('#saveCfg').addEventListener('click', saveLayoutFile);
  linkHandoff($('#profileLink'), () => S.wells.length ? { fileName: S.fileName, sheets: S.sheets, focus: topFocus() } : null);
  document.querySelectorAll('#resSeg button').forEach(b => b.addEventListener('click', () => setSetting('res', b.dataset.v)));
  document.querySelectorAll('#colSeg button').forEach(b => b.addEventListener('click', () => setSetting('cols', +b.dataset.v)));
  $('#maSel').addEventListener('change', e => setSetting('ma', +e.target.value));
  $('#shutinChk').addEventListener('change', e => setSetting('excludeShutin', e.target.checked));
  $('#zeroChk').addEventListener('change', e => setSetting('zeroGap', e.target.checked));
  $('#syncChk').addEventListener('change', e => setSetting('sync', e.target.checked));
  document.querySelectorAll('[data-add]').forEach(b => b.addEventListener('click', () => {
    const p = newPlot(b.dataset.add);
    S.plots.push(p);
    addCard(p, true);
    persist();
  }));

  let rzTimer;
  window.addEventListener('resize', () => {
    clearTimeout(rzTimer);
    rzTimer = setTimeout(() => S.plots.filter(p => p.type === 'ts' && p.axisMode !== 'stacked' && p.vars.length > 2).forEach(renderPlot), 200);
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
loadFromOpener();
