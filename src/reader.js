'use strict';
/* Shared workbook reader for the SAGD viewers. Each sheet = one well; first "Date" column = timestamp. */

/* ================= fast .xlsx reader (zip + regex over sheet XML) =================
   SheetJS takes ~10 s on large numeric workbooks; this path reads only what we need.
   Any unexpected structure throws, and loadExcel falls back to SheetJS. */
async function inflateRaw(bytes) {
  const stream = new Blob([bytes]).stream().pipeThrough(new DecompressionStream('deflate-raw'));
  return new Uint8Array(await new Response(stream).arrayBuffer());
}

function readZipEntries(buf) {
  const dv = new DataView(buf), u8 = new Uint8Array(buf);
  let eocd = -1;
  for (let i = buf.byteLength - 22; i >= Math.max(0, buf.byteLength - 65557); i--) {
    if (dv.getUint32(i, true) === 0x06054b50) { eocd = i; break; }
  }
  if (eocd < 0) throw new Error('zip: end of central directory not found');
  const count = dv.getUint16(eocd + 10, true);
  let p = dv.getUint32(eocd + 16, true);
  if (p === 0xffffffff) throw new Error('zip64 not supported');
  const dec = new TextDecoder();
  const entries = {};
  for (let k = 0; k < count; k++) {
    if (dv.getUint32(p, true) !== 0x02014b50) throw new Error('zip: bad central header');
    const method = dv.getUint16(p + 10, true);
    const csize = dv.getUint32(p + 20, true);
    const nlen = dv.getUint16(p + 28, true), xlen = dv.getUint16(p + 30, true), clen = dv.getUint16(p + 32, true);
    const loc = dv.getUint32(p + 42, true);
    const name = dec.decode(u8.subarray(p + 46, p + 46 + nlen));
    const lnl = dv.getUint16(loc + 26, true), lxl = dv.getUint16(loc + 28, true);
    const start = loc + 30 + lnl + lxl;
    entries[name] = { method, data: u8.subarray(start, start + csize) };
    p += 46 + nlen + xlen + clen;
  }
  return entries;
}

function xmlUnescape(s) {
  return s.replace(/&(lt|gt|amp|quot|apos|#\d+|#x[0-9a-f]+);/gi, (m, e) => {
    const map = { lt: '<', gt: '>', amp: '&', quot: '"', apos: "'" };
    if (map[e.toLowerCase()]) return map[e.toLowerCase()];
    return String.fromCodePoint(e[1] === 'x' || e[1] === 'X' ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10));
  });
}

function colIndex(letters) {
  let n = 0;
  for (let i = 0; i < letters.length; i++) n = n * 26 + (letters.charCodeAt(i) - 64);
  return n - 1;
}

function parseSheetXml(xml, shared) {
  const rows = [];
  const re = /<c\b([^>]*?)(\/>|>([\s\S]*?)<\/c>)/g;
  let m;
  while ((m = re.exec(xml))) {
    const attrs = m[1];
    const ref = /\br="([A-Z]+)(\d+)"/.exec(attrs);
    if (!ref) throw new Error('cell without reference');
    const body = m[3];
    if (!body) continue;
    const r = +ref[2] - 1, c = colIndex(ref[1]);
    const tm = /\bt="(\w+)"/.exec(attrs);
    const t = tm ? tm[1] : 'n';
    let val = null;
    if (t === 'inlineStr') {
      const parts = body.match(/<t\b[^>]*>([\s\S]*?)<\/t>/g) || [];
      val = xmlUnescape(parts.map(x => x.replace(/<[^>]+>/g, '')).join(''));
    } else {
      const vm = /<v>([\s\S]*?)<\/v>/.exec(body);
      if (!vm) continue;
      const v = vm[1];
      if (t === 's') val = shared[+v];
      else if (t === 'str') val = xmlUnescape(v);
      else if (t === 'b') val = v === '1';
      else if (t === 'e') val = null;
      else val = +v;
    }
    (rows[r] = rows[r] || [])[c] = val;
  }
  const out = [];
  for (let i = 0; i < rows.length; i++) if (rows[i]) out.push(Array.from(rows[i], v => v === undefined ? null : v));
  return out;
}

async function fastReadXlsx(buf, onProgress) {
  const entries = readZipEntries(buf);
  const text = async (name) => {
    const e = entries[name];
    if (!e) return null;
    const bytes = e.method === 0 ? e.data : e.method === 8 ? await inflateRaw(e.data) : null;
    if (!bytes) throw new Error('zip: unsupported compression');
    return new TextDecoder().decode(bytes);
  };
  const wbXml = await text('xl/workbook.xml');
  const relXml = await text('xl/_rels/workbook.xml.rels');
  if (!wbXml || !relXml) throw new Error('not an xlsx workbook');
  const rels = {};
  for (const m of relXml.matchAll(/<Relationship\b([^>]*)\/?>/g)) {
    const id = /\bId="([^"]+)"/.exec(m[1]), tg = /\bTarget="([^"]+)"/.exec(m[1]);
    if (id && tg) rels[id[1]] = tg[1];
  }
  const shared = [];
  const ssXml = await text('xl/sharedStrings.xml');
  if (ssXml) for (const si of ssXml.matchAll(/<si>([\s\S]*?)<\/si>/g)) {
    const ts = si[1].replace(/<rPh\b[\s\S]*?<\/rPh>/g, '').match(/<t\b[^>]*>([\s\S]*?)<\/t>/g) || [];
    shared.push(xmlUnescape(ts.map(x => x.replace(/<[^>]+>/g, '')).join('')));
  }
  const sheets = [...wbXml.matchAll(/<sheet\b([^>]*)\/?>/g)].map(m => ({
    name: xmlUnescape((/\bname="([^"]*)"/.exec(m[1]) || [])[1] || ''),
    rid: (/\br:id="([^"]+)"/.exec(m[1]) || /\bid="([^"]+)"/.exec(m[1]) || [])[1],
  }));
  const out = [];
  for (let i = 0; i < sheets.length; i++) {
    const sh = sheets[i];
    let target = rels[sh.rid];
    if (!target) continue;
    target = target.startsWith('/') ? target.slice(1) : 'xl/' + target.replace(/^\.\//, '');
    if (onProgress) onProgress(i + 1, sheets.length, sh.name);
    const xml = await text(target);
    if (!xml) continue;
    const well = sheetRowsToWell(sh.name, parseSheetXml(xml, shared));
    if (well) out.push(well);
    await new Promise(r => setTimeout(r, 0)); // let the spinner breathe
  }
  return out;
}

/* ================= workbook parsing via SheetJS (fallback; runs in a Worker when possible) ================= */
function parseWorkbookBuffer(XLSX, buf) {
  var wb = XLSX.read(new Uint8Array(buf), { type: 'array', dense: true, cellDates: false });
  var out = [];
  for (var si = 0; si < wb.SheetNames.length; si++) {
    var name = wb.SheetNames[si];
    var rows = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, raw: true, defval: null, blankrows: false });
    var well = sheetRowsToWell(name, rows);
    if (well) out.push(well);
  }
  return out;
}

/** rows: array of arrays (first row = header). Returns {name, header, t, cols} or null. */
function sheetRowsToWell(name, rows) {
  function toMs(v) {
    if (v === null || v === undefined || v === '') return NaN;
    if (typeof v === 'number') return Math.round((v - 25569) * 86400000 / 60000) * 60000;
    if (v instanceof Date) return Date.UTC(v.getFullYear(), v.getMonth(), v.getDate(), v.getHours(), v.getMinutes());
    var s = String(v).trim();
    var m = s.match(/^(\d{4})[-\/.](\d{1,2})[-\/.](\d{1,2})/);
    if (m) return Date.UTC(+m[1], +m[2] - 1, +m[3]);
    var p = Date.parse(s);
    return isNaN(p) ? NaN : p;
  }
  function toNum(v) {
    if (typeof v === 'number') return v;
    if (typeof v === 'boolean') return v ? 1 : 0;
    if (v === null || v === undefined || v === '') return NaN;
    var n = parseFloat(String(v).replace(/,/g, ''));
    return isNaN(n) ? NaN : n;
  }
  {
    if (!rows.length || !rows[0]) return null;
    var header = rows[0].map(function (h, i) { return h === null || h === undefined ? 'Col_' + (i + 1) : String(h).trim(); });
    var di = header.findIndex(function (h) { return /date|time|일자|날짜/i.test(h); });
    if (di < 0) di = 0;
    var recs = [];
    for (var r = 1; r < rows.length; r++) {
      var t = toMs(rows[r][di]);
      if (!isNaN(t)) recs.push([t, rows[r]]);
    }
    if (recs.length < 2) return null;
    recs.sort(function (a, b) { return a[0] - b[0]; });
    var n = recs.length, tArr = new Float64Array(n), cols = {};
    for (var c = 0; c < header.length; c++) if (c !== di) cols[header[c]] = new Float64Array(n);
    for (var k = 0; k < n; k++) {
      tArr[k] = recs[k][0];
      var row = recs[k][1];
      for (var c2 = 0; c2 < header.length; c2++) if (c2 !== di) cols[header[c2]][k] = toNum(row[c2]);
    }
    // drop columns that are entirely non-numeric
    var keep = [];
    for (var h in cols) {
      var a = cols[h], any = false;
      for (var q = 0; q < n; q++) if (!isNaN(a[q])) { any = true; break; }
      if (any) keep.push(h); else delete cols[h];
    }
    return { name: name, header: keep, t: tArr, cols: cols };
  }
}

function parseInWorker(buf) {
  return new Promise((resolve, reject) => {
    let worker;
    try {
      const lib = document.getElementById('xlsx-lib').textContent;
      const code = lib + '\n' + parseWorkbookBuffer.toString() +
        '\nself.onmessage=function(e){try{postMessage({ok:true,r:parseWorkbookBuffer(XLSX,e.data)});}catch(err){postMessage({ok:false,err:String(err&&err.message||err)});}};';
      const url = URL.createObjectURL(new Blob([code], { type: 'text/javascript' }));
      worker = new Worker(url);
      URL.revokeObjectURL(url);
    } catch (e) { reject(e); return; }
    worker.onmessage = (e) => { worker.terminate(); e.data.ok ? resolve(e.data.r) : reject(new Error(e.data.err)); };
    worker.onerror = (e) => { worker.terminate(); reject(e); };
    worker.postMessage(buf, [buf]);
  });
}

/** Reads an Excel/CSV File into [{name, header, t, cols}]: fast zip reader first, SheetJS as fallback. */
async function readWorkbook(file, onStatus) {
  const status = onStatus || (() => {});
  const buf = await file.arrayBuffer();
  let sheets = null;
  if (/\.xls[xm]$/i.test(file.name) && typeof DecompressionStream === 'function') {
    try {
      sheets = await fastReadXlsx(buf, (i, n, nm) => { status(`Parsing sheet ${i}/${n}: ${nm}`); });
      if (!sheets.length) sheets = null;
    } catch (e) {
      console.warn('Fast reader failed, falling back to SheetJS:', e);
      sheets = null;
    }
  }
  if (!sheets) {
    status(`Reading ${file.name} with SheetJS (may take a while)…`);
    try {
      sheets = await parseInWorker(buf.slice(0));
    } catch (e) {
      console.warn('Worker parse failed, falling back to main thread:', e);
      await new Promise(r => setTimeout(r, 30));
      sheets = parseWorkbookBuffer(XLSX, buf);
    }
  }
  return sheets;
}

/* ================= hand-off between viewers =================
   A link opens the other viewer in a new tab; that tab asks its opener for the workbook
   already loaded here, so it shows the data without reading the Excel file again.
   Works on file:// pages too: postMessage needs no shared origin or storage. */
const HANDOFF = 'sagd-viewer-handoff';

/** Make `link` open its page in a new tab that receives getPayload() → {fileName, sheets, focus?}, when not null. */
function linkHandoff(link, getPayload) {
  const children = new Set();
  window.addEventListener('message', e => {
    const m = e.data;
    if (!m || m.type !== HANDOFF || !m.ask || !children.has(e.source)) return;
    const payload = getPayload();
    e.source.postMessage(payload ? { type: HANDOFF, fileName: payload.fileName, sheets: payload.sheets, focus: payload.focus || null }
      : { type: HANDOFF, none: true }, '*');
  });
  link.addEventListener('click', e => {
    if (!getPayload()) return; // nothing loaded: plain navigation
    const win = window.open(link.href, '_blank');
    if (!win) return; // popup blocked: fall back to the plain link
    e.preventDefault();
    children.add(win);
  });
}

/** If this tab was opened by the other viewer, ask it for its data. Resolves {fileName, sheets, focus} or null. */
function receiveHandoff(timeoutMs = 3000) {
  const parent = window.opener;
  if (!parent || parent === window) return Promise.resolve(null);
  return new Promise(resolve => {
    const done = v => { clearTimeout(timer); window.removeEventListener('message', onMsg); resolve(v); };
    const onMsg = e => {
      const m = e.data;
      if (e.source !== parent || !m || m.type !== HANDOFF || m.ask) return;
      done(m.none ? null : { fileName: m.fileName, sheets: m.sheets, focus: m.focus || null });
    };
    const timer = setTimeout(() => done(null), timeoutMs);
    window.addEventListener('message', onMsg);
    try { parent.postMessage({ type: HANDOFF, ask: true }, '*'); } catch (err) { done(null); }
  });
}
