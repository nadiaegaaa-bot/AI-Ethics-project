/* Data upload: lets the user upload an Excel (.xlsx) or CSV report on the Reports page to update the dashboard.
   - Runs entirely in the browser (no server). Uses DecompressionStream to unzip .xlsx and DOMParser for the XML.
   - Each accepted upload is stored in localStorage and re-applied on top of js/data.js every time the app loads,
     so uploads can be undone one by one or reset without touching the original data file.
   Must load after js/data.js and before js/app.js. */
(() => {
  const KEY = 'bulk2.uploads';
  const BASE = window.BULK_DATA;
  const clone = o => JSON.parse(JSON.stringify(o));
  const THRESH = BASE.reviewThreshold || 0.12;

  /* ---------------- persistence & applying ---------------- */
  const loadUploads = () => { try { return JSON.parse(localStorage.getItem(KEY)) || []; } catch { return []; } };
  const saveUploads = list => localStorage.setItem(KEY, JSON.stringify(list));

  function applyUpload(data, up) {
    const p = up.data;
    if (p.monthly && p.monthly.length) {
      if (up.mode === 'replace') data.monthly = clone(p.monthly);
      else {
        const idx = new Map(data.monthly.map((r, i) => [r.month + '|' + r.site, i]));
        p.monthly.forEach(r => {
          const k = r.month + '|' + r.site;
          if (idx.has(k)) data.monthly[idx.get(k)] = clone(r);
          else { idx.set(k, data.monthly.length); data.monthly.push(clone(r)); }
        });
      }
      data.monthly.sort((a, b) => a.month.localeCompare(b.month) || a.site.localeCompare(b.site));
    }
    ['benchmarks', 'peerMonthly', 'facts', 'sources', 'dictionary', 'guide'].forEach(k => { if (p[k] && p[k].length) data[k] = clone(p[k]); });
    if (p.dashboardSummary2025 && Object.keys(p.dashboardSummary2025).length) data.dashboardSummary2025 = clone(p.dashboardSummary2025);
    return data;
  }
  function build(uploads) {
    let data = clone(BASE);
    uploads.forEach(u => { data = applyUpload(data, u); });
    data.uploads = uploads.map(u => ({ id: u.id, name: u.name, date: u.date, mode: u.mode, summary: u.summary }));
    return data;
  }

  // Apply stored uploads before the dashboard starts.
  try { window.BULK_DATA = build(loadUploads()); } catch (e) { console.warn('Could not apply uploaded data; using the original data.', e); window.BULK_DATA = clone(BASE); window.BULK_DATA.uploads = []; }

  /* ---------------- .xlsx reading ---------------- */
  async function inflateRaw(bytes) {
    if (typeof DecompressionStream === 'undefined') throw new Error('This browser cannot read .xlsx files. Please update it or upload a CSV instead.');
    const ds = new DecompressionStream('deflate-raw');
    const out = new Response(new Blob([bytes]).stream().pipeThrough(ds));
    return new Uint8Array(await out.arrayBuffer());
  }
  async function unzip(buf) {
    const v = new DataView(buf), u8 = new Uint8Array(buf);
    let eocd = -1;
    for (let i = buf.byteLength - 22; i >= Math.max(0, buf.byteLength - 65557); i--) if (v.getUint32(i, true) === 0x06054b50) { eocd = i; break; }
    if (eocd < 0) throw new Error('This file is not a valid .xlsx workbook.');
    const count = v.getUint16(eocd + 10, true); let p = v.getUint32(eocd + 16, true);
    const files = {}, dec = new TextDecoder();
    for (let n = 0; n < count; n++) {
      if (v.getUint32(p, true) !== 0x02014b50) break;
      const method = v.getUint16(p + 10, true), csize = v.getUint32(p + 20, true);
      const nlen = v.getUint16(p + 28, true), elen = v.getUint16(p + 30, true), clen = v.getUint16(p + 32, true);
      const lho = v.getUint32(p + 42, true);
      const name = dec.decode(u8.subarray(p + 46, p + 46 + nlen));
      p += 46 + nlen + elen + clen;
      const start = lho + 30 + v.getUint16(lho + 26, true) + v.getUint16(lho + 28, true);
      const raw = u8.subarray(start, start + csize);
      files[name] = { method, raw };
    }
    return async name => {
      const f = files[name]; if (!f) return null;
      const bytes = f.method === 0 ? f.raw : f.method === 8 ? await inflateRaw(f.raw) : null;
      if (!bytes) throw new Error('Unsupported compression in workbook.');
      return dec.decode(bytes);
    };
  }
  const xml = s => new DOMParser().parseFromString(s, 'application/xml');
  const colIdx = ref => { let n = 0; for (const ch of ref.replace(/\d+/g, '')) n = n * 26 + ch.charCodeAt(0) - 64; return n - 1; };
  // Namespace-agnostic: workbooks may write <c> or <x:c> depending on the program that saved them.
  const byTag = (el, tag) => [...el.getElementsByTagNameNS('*', tag)];
  async function readXlsx(buf) {
    const get = await unzip(buf);
    const ssXml = await get('xl/sharedStrings.xml');
    const shared = ssXml ? byTag(xml(ssXml), 'si').map(si => byTag(si, 't').map(t => t.textContent).join('')) : [];
    const wb = xml(await get('xl/workbook.xml'));
    const rels = {}; byTag(xml(await get('xl/_rels/workbook.xml.rels')), 'Relationship').forEach(r => { rels[r.getAttribute('Id')] = r.getAttribute('Target'); });
    const sheets = {};
    for (const s of byTag(wb, 'sheet')) {
      const rid = s.getAttribute('r:id') || [...s.attributes].find(a => a.name.endsWith(':id'))?.value;
      let target = (rels[rid] || '').replace(/^\//, ''); if (!target.startsWith('xl/')) target = 'xl/' + target;
      const sx = await get(target); if (!sx) continue;
      const grid = [];
      byTag(xml(sx), 'row').forEach(row => {
        const r = +row.getAttribute('r') - 1, cells = [];
        byTag(row, 'c').forEach(c => {
          const t = c.getAttribute('t'), vEl = byTag(c, 'v')[0], isEl = byTag(c, 'is')[0];
          let val = null;
          if (t === 's' && vEl) val = shared[+vEl.textContent];
          else if (isEl) val = byTag(isEl, 't').map(x => x.textContent).join('');
          else if (vEl) { val = vEl.textContent; if (t !== 'str' && t !== 'b' && val !== '' && !isNaN(+val)) val = +val; }
          if (val !== null && val !== '') cells[colIdx(c.getAttribute('r'))] = val;
        });
        grid[r] = cells;
      });
      sheets[s.getAttribute('name')] = grid;
    }
    return sheets;
  }

  /* ---------------- CSV reading ---------------- */
  function readCsv(text) {
    text = text.replace(/^\ufeff/, '');
    const first = text.split(/\r?\n/)[0] || '';
    const delim = (first.match(/;/g) || []).length > (first.match(/,/g) || []).length ? ';' : (first.includes('\t') ? '\t' : ',');
    const rows = []; let row = [], cell = '', q = false;
    for (let i = 0; i < text.length; i++) {
      const ch = text[i];
      if (q) { if (ch === '"' && text[i + 1] === '"') { cell += '"'; i++; } else if (ch === '"') q = false; else cell += ch; }
      else if (ch === '"') q = true;
      else if (ch === delim) { row.push(cell); cell = ''; }
      else if (ch === '\n' || ch === '\r') { if (ch === '\r' && text[i + 1] === '\n') i++; row.push(cell); rows.push(row); row = []; cell = ''; }
      else cell += ch;
    }
    if (cell !== '' || row.length) { row.push(cell); rows.push(row); }
    return rows.filter(r => r.some(c => String(c).trim() !== ''));
  }

  /* ---------------- value helpers ---------------- */
  const MON = { jan: 1, feb: 2, mar: 3, apr: 4, may: 5, jun: 6, jul: 7, aug: 8, sep: 9, oct: 10, nov: 11, dec: 12, mai: 5, okt: 10, des: 12 };
  function toMonth(v) {
    if (v == null || v === '') return null;
    if (typeof v === 'number' || /^\d{4,5}(\.\d+)?$/.test(String(v).trim())) {
      const n = +v; if (n > 20000 && n < 80000) { const d = new Date(Date.UTC(1899, 11, 30) + Math.floor(n) * 864e5); return `${d.getUTCFullYear()}-${String(d.getUTCMonth() + 1).padStart(2, '0')}`; }
      return null;
    }
    const s = String(v).trim();
    let m = s.match(/^(\d{4})[-/.](\d{1,2})/); if (m) return `${m[1]}-${m[2].padStart(2, '0')}`;
    m = s.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})/); if (m) return `${m[3]}-${m[2].padStart(2, '0')}`; // dd.mm.yyyy
    m = s.match(/^([A-Za-zæøå]{3})[a-zæøå]*\.?\s+'?(\d{2,4})$/i);
    if (m && MON[m[1].toLowerCase()]) { const y = m[2].length === 2 ? '20' + m[2] : m[2]; return `${y}-${String(MON[m[1].toLowerCase()]).padStart(2, '0')}`; }
    return null;
  }
  function num(v) {
    if (v == null || v === '') return null;
    if (typeof v === 'number') return isFinite(v) ? v : null;
    let s = String(v).trim().replace(/\s|\u00a0/g, '').replace(/%$/, '');
    if (/^-?\d{1,3}(\.\d{3})+(,\d+)?$/.test(s)) s = s.replace(/\./g, '').replace(',', '.'); // 1.234,5
    else if (/^-?\d+,\d+$/.test(s)) s = s.replace(',', '.'); // 1234,5
    else s = s.replace(/,/g, '');
    const n = +s; return isFinite(n) ? n : null;
  }
  const norm = s => String(s ?? '').toLowerCase().replace(/[^a-z0-9%°]+/g, ' ').trim();

  /* ---------------- monthly rows ---------------- */
  const MONTHLY_COLS = [
    ['month', [/^month$/, /^date$/, /^period$/]],
    ['site', [/^site$/]],
    ['zone', [/zone/, /^country/]],
    ['days', [/^days$/]],
    ['hours', [/^hours$/]],
    ['itLoadMW', [/it avg load/, /it load/]],
    ['itMWh', [/^it energy/, /^it mwh/]],
    ['coolingMWh', [/^cooling mwh$/, /^cooling$/, /^cooling energy/]],
    ['otherMWh', [/^other/]],
    ['totalMWh', [/^total/]],
    ['pue', [/^pue$/]],
    ['tempC', [/outdoor/, /temp/]],
    ['priceNOKMWh', [/price/]],
    ['costNOK', [/cost/]],
    ['expectedCoolingMWh', [/^expected/]],
    ['residualMWh', [/residual mwh/, /^cooling residual/]],
    ['residualPct', [/residual %/, /residual pct/, /^residual$/]],
    ['alert', [/^alert/, /^flag/]],
    ['provenance', [/provenance/, /^source$/]],
  ];
  function findHeader(grid, test) {
    for (let r = 0; r < Math.min(grid.length, 30); r++) { const row = grid[r]; if (row && test((row || []).map(norm))) return r; }
    return -1;
  }
  function mapColumns(headerRow) {
    const map = {}, used = new Set();
    MONTHLY_COLS.forEach(([key, pats]) => {
      for (let i = 0; i < headerRow.length; i++) {
        const h = norm(headerRow[i]); if (!h || used.has(i)) continue;
        if (pats.some(p => p.test(h))) { map[key] = i; used.add(i); break; }
      }
    });
    return map;
  }
  const daysIn = m => new Date(+m.slice(0, 4), +m.slice(5), 0).getDate();
  function parseMonthly(grid, sheetName, current) {
    const warnings = [], out = [];
    const h = findHeader(grid, row => row.includes('month') && row.includes('site'));
    if (h < 0) return null;
    const map = mapColumns(grid[h]);
    const need = ['month', 'site', 'itMWh', 'coolingMWh'];
    const missing = need.filter(k => map[k] == null);
    if (missing.length) { warnings.push(`${sheetName}: missing required column(s): ${missing.join(', ')}.`); return { rows: [], warnings }; }
    const zoneOf = s => (current.monthly.find(r => r.site === s) || {}).zone || '';
    let estimated = 0;
    for (let r = h + 1; r < grid.length; r++) {
      const row = grid[r]; if (!row || !row.some(c => c !== undefined && String(c).trim() !== '')) continue;
      const g = k => (map[k] == null ? null : row[map[k]]);
      const month = toMonth(g('month')), site = String(g('site') ?? '').trim();
      const line = `${sheetName} row ${r + 1}`;
      if (!month || !site) { warnings.push(`${line}: skipped (no valid month or site).`); continue; }
      const it = num(g('itMWh')), cool = num(g('coolingMWh'));
      if (it == null || cool == null || it <= 0 || cool < 0) { warnings.push(`${line}: skipped (IT and cooling energy must be numbers, IT > 0).`); continue; }
      const other = num(g('otherMWh')) ?? 0;
      let total = num(g('totalMWh'));
      if (total == null) total = it + cool + other;
      else if (Math.abs(total - (it + cool + other)) > Math.max(1, total * 0.005)) warnings.push(`${line}: total (${total.toFixed(0)}) ≠ IT + cooling + other (${(it + cool + other).toFixed(0)}). The given total is kept.`);
      const days = num(g('days')) ?? daysIn(month), hours = num(g('hours')) ?? days * 24;
      const price = num(g('priceNOKMWh'));
      const temp = num(g('tempC'));
      let expected = num(g('expectedCoolingMWh'));
      if (expected == null) {
        // Workbook rule (Sources & Methods M05) without the unknown site adjustment.
        expected = temp != null ? it * (0.115 + 0.0055 * Math.max(temp - 8, 0)) : cool;
        estimated++;
      }
      const residualMWh = num(g('residualMWh')) ?? cool - expected;
      let residualPct = num(g('residualPct'));
      if (residualPct == null) residualPct = expected ? (cool - expected) / expected : 0;
      else if (Math.abs(residualPct) > 1.5) residualPct = residualPct / 100; // given as e.g. 15.6 (%)
      const alertIn = String(g('alert') ?? '').trim().toUpperCase();
      const pue = num(g('pue')) ?? total / it;
      out.push({
        month, site, zone: String(g('zone') ?? '').trim() || zoneOf(site), days, hours,
        itLoadMW: num(g('itLoadMW')) ?? +(it / hours).toFixed(4),
        itMWh: it, coolingMWh: cool, otherMWh: other, totalMWh: total, pue,
        tempC: temp ?? 0, priceNOKMWh: price ?? 0,
        costNOK: num(g('costNOK')) ?? (price != null ? total * price : 0),
        expectedCoolingMWh: expected, residualMWh, residualPct,
        alert: alertIn ? (alertIn.startsWith('REV') ? 'REVIEW' : 'Normal') : (residualPct > THRESH ? 'REVIEW' : 'Normal'),
        provenance: String(g('provenance') ?? '').trim() || 'UPLOADED',
        year: +month.slice(0, 4),
      });
      if (price == null) warnings.push(`${line}: no market price, so cost is set to 0.`);
    }
    if (estimated) warnings.push(`${estimated} row(s) had no "Expected cooling" column; it was estimated with the workbook formula (without site adjustment).`);
    return { rows: out, warnings };
  }

  /* ---------------- other workbook sheets (positional, same layout as the original workbook) ---------------- */
  function table(grid, firstHeader, keys) {
    const h = findHeader(grid, row => row[0] === norm(firstHeader));
    if (h < 0) return null;
    const out = [];
    for (let r = h + 1; r < grid.length; r++) {
      const row = grid[r]; if (!row || !row.some(c => c != null && c !== '')) continue;
      const o = {}; keys.forEach((k, i) => { if (k) o[k] = row[i] ?? null; }); out.push(o);
    }
    return out;
  }
  function parseOtherSheets(sheets, result) {
    const find = re => Object.keys(sheets).find(n => re.test(n));
    let n;
    if ((n = find(/competitor benchmark/i))) {
      const t = table(sheets[n], 'Operator', ['operator', 'facility', 'year', 'pue', 'pueType', 'energyMWh', 'energyScope', 'status', 'relevance', 'source', 'cautions']);
      if (t) { t.forEach(b => { if (typeof b.year === 'number') b.year = Math.round(b.year); if (b.pue != null && typeof b.pue !== 'number') b.pue = num(b.pue); }); result.benchmarks = t; }
    }
    if ((n = find(/peer.*scenario/i))) {
      const t = table(sheets[n], 'Month', ['month', 'peer', 'refYear', 'refPUE', 'pue', 'totalMWh', 'itMWh', 'provenance']);
      if (t) result.peerMonthly = t.map(r => ({ ...r, month: toMonth(r.month), refYear: Math.round(num(r.refYear) || 0), pue: num(r.pue), totalMWh: num(r.totalMWh), itMWh: num(r.itMWh) })).filter(r => r.month && r.peer);
    }
    if ((n = find(/company facts/i))) { const t = table(sheets[n], 'Entity', ['entity', 'period', 'metric', 'value', 'unit', 'source', 'classification', 'caveat']); if (t) result.facts = t.map(f => ({ ...f, period: f.period != null ? String(f.period) : f.period })); }
    if ((n = find(/sources/i))) { const t = table(sheets[n], 'Source ID', ['id', 'organization', 'document', 'period', 'type', 'use']); if (t) result.sources = t; }
    if ((n = find(/dictionary/i))) { const t = table(sheets[n], 'Sheet', ['sheet', 'fields', 'unit', 'method', 'notes']); if (t) result.dictionary = t; }
    if ((n = find(/start here/i))) { const g = sheets[n]; result.guide = g.slice(4).filter(r => r && r[1]).map(r => ({ step: r[0], text: r[1] })); }
    if ((n = find(/^dashboard$/i))) {
      const g = sheets[n], h = findHeader(g, row => row[0] === 'metric');
      if (h >= 0) {
        const sites = [1, 2, 3].map(i => g[h][i]).filter(Boolean), sum = {};
        for (let r = h + 1; r < h + 12 && g[r] && g[r][0]; r++) { sum[g[r][0]] = { scope: g[r][4] }; sites.forEach((s, i) => { sum[g[r][0]][s] = num(g[r][i + 1]); }); }
        if (Object.keys(sum).length) result.dashboardSummary2025 = sum;
      }
    }
  }

  /* ---------------- public API ---------------- */
  async function parseFile(file) {
    if (file.size > 15 * 1024 * 1024) throw new Error('The file is larger than 15 MB.');
    const name = file.name.toLowerCase(), current = window.BULK_DATA;
    const result = {}, warnings = [];
    if (name.endsWith('.xlsx')) {
      const sheets = await readXlsx(await file.arrayBuffer());
      const monthlySheet = Object.keys(sheets).find(n => /monthly/i.test(n) && !/peer/i.test(n)) ||
        Object.keys(sheets).find(n => findHeader(sheets[n], row => row.includes('month') && row.includes('site')) >= 0 && !/peer/i.test(n));
      if (monthlySheet) { const m = parseMonthly(sheets[monthlySheet], `"${monthlySheet}"`, current); if (m) { result.monthly = m.rows; warnings.push(...m.warnings); } }
      parseOtherSheets(sheets, result);
    } else if (name.endsWith('.csv') || name.endsWith('.txt')) {
      const m = parseMonthly(readCsv(await file.text()), 'CSV', current);
      if (!m) throw new Error('No header row with "Month" and "Site" columns was found.');
      result.monthly = m.rows; warnings.push(...m.warnings);
    } else if (name.endsWith('.xls')) {
      throw new Error('Old .xls files are not supported. In Excel, use File → Save As → Excel Workbook (.xlsx).');
    } else if (name.endsWith('.pdf')) {
      throw new Error('PDF reports cannot be read automatically. Upload the numbers as an Excel (.xlsx) or CSV file.');
    } else throw new Error('Please upload an Excel (.xlsx) or CSV file.');

    const parts = Object.keys(result).filter(k => Array.isArray(result[k]) ? result[k].length : Object.keys(result[k]).length);
    if (!parts.length) throw new Error('No dashboard data was found in this file. It needs a sheet with "Month" and "Site" columns, or the sheets of the Bulk data workbook.');

    // Compare with current data for the preview.
    const keyset = new Set(current.monthly.map(r => r.month + '|' + r.site));
    const m = result.monthly || [];
    const months = [...new Set(m.map(r => r.month))].sort();
    const summary = {
      monthlyRows: m.length,
      added: m.filter(r => !keyset.has(r.month + '|' + r.site)).length,
      updated: m.filter(r => keyset.has(r.month + '|' + r.site)).length,
      sites: [...new Set(m.map(r => r.site))],
      newSites: [...new Set(m.map(r => r.site))].filter(s => !current.monthly.some(r => r.site === s)),
      range: months.length ? [months[0], months.at(-1)] : null,
      flags: m.filter(r => r.alert === 'REVIEW').length,
      tables: parts.filter(k => k !== 'monthly'),
    };
    return { data: result, warnings, summary };
  }

  function apply(fileName, parsed, mode) {
    const list = loadUploads();
    list.push({ id: Date.now().toString(36), name: fileName, date: new Date().toISOString(), mode, data: parsed.data, summary: parsed.summary });
    try { saveUploads(list); } catch { throw new Error('The browser storage is full. Remove an earlier upload and try again.'); }
  }
  function remove(id) { saveUploads(loadUploads().filter(u => u.id !== id)); }
  function resetAll() { localStorage.removeItem(KEY); }

  window.BulkUpload = { parseFile, apply, remove, resetAll, list: () => loadUploads().map(u => ({ id: u.id, name: u.name, date: u.date, mode: u.mode, summary: u.summary })), base: BASE };
})();
